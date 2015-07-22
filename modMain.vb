Option Strict On

' This program inspects a .NET DLL or .Exe to determine the version.  
' This allows a 32-bit .NET application to call this program via the 
' command prompt to determine the version of a 64-bit DLL or Exe.
'
' -------------------------------------------------------------------------------
' Written by Matthew Monroe for the Department of Energy (PNNL, Richland, WA)
' Program started in April 2011
'
' E-mail: matthew.monroe@pnnl.gov or matt@alchemistmatt.com
' Website: http://omics.pnl.gov/ or http://www.sysbio.org/resources/staff/ or http://panomics.pnnl.gov/
' -------------------------------------------------------------------------------
' 

Module modMain
    Public Const PROGRAM_DATE As String = "July 22, 2015"

	Private mInputFilePath As String
	Private mVersionInfoFilePath As String = String.Empty
	Private mGenericDLL As Boolean = False

	Private mOutputFolderNameOrPath As String
	Private mParameterFilePath As String			' Not used by this program

	Private mOutputFolderAlternatePath As String				' Not used by this program
	Private mRecreateFolderHierarchyInAlternatePath As Boolean	' Not used by this program

	Private mRecurseFolders As Boolean
	Private mRecurseFoldersMaxLevels As Integer
	Private mLogMessagesToFile As Boolean
	Private mQuietMode As Boolean

	Private mShowResultsAtConsole As Boolean = False

	Private WithEvents mDLLVersionInspector As clsDLLVersionInspector
	Private mLastProgressReportTime As System.DateTime
	Private mLastProgressReportValue As Integer

	Public Function Main() As Integer
		' Returns 0 if no error, error code if an error

		Dim intReturnCode As Integer
		Dim objParseCommandLine As New clsParseCommandLine
		Dim blnProceed As Boolean


		intReturnCode = 0

		mOutputFolderNameOrPath = String.Empty
		mParameterFilePath = String.Empty

		mOutputFolderAlternatePath = String.Empty
		mRecreateFolderHierarchyInAlternatePath = False

		mRecurseFolders = False
		mRecurseFoldersMaxLevels = 0

		mLogMessagesToFile = False
		mQuietMode = False

		mShowResultsAtConsole = False


		Try
			blnProceed = False

			If objParseCommandLine.ParseCommandLine Then
				If SetOptionsUsingCommandLineParameters(objParseCommandLine) Then blnProceed = True
			End If

			If Not blnProceed OrElse _
			   objParseCommandLine.NeedToShowHelp OrElse _
			   objParseCommandLine.ParameterCount + objParseCommandLine.NonSwitchParameterCount = 0 OrElse _
			   String.IsNullOrWhiteSpace(mInputFilePath) Then
				ShowProgramHelp()
				intReturnCode = -1
			Else

				mDLLVersionInspector = New clsDLLVersionInspector

				Dim strVersionInfoFileName As String = String.Empty
				If Not String.IsNullOrWhiteSpace(mVersionInfoFilePath) Then
					Dim fiVersionInfoFile As System.IO.FileInfo
					fiVersionInfoFile = New System.IO.FileInfo(mVersionInfoFilePath)

					strVersionInfoFileName = fiVersionInfoFile.Name
					mOutputFolderNameOrPath = fiVersionInfoFile.Directory.FullName

					If mRecurseFolders AndAlso fiVersionInfoFile.Exists Then
						' Delete the existing version info file because we will be appending to it
						fiVersionInfoFile.Delete()
					End If				
				End If

				' Note: the following settings will be overridden if mParameterFilePath points to a valid parameter file that has these settings defined
				With mDLLVersionInspector
					.ShowMessages = True
					.LogMessagesToFile = False

					If mRecurseFolders AndAlso Not String.IsNullOrWhiteSpace(mOutputFolderNameOrPath) Then
						.AppendToVersionInfoFile = True
					Else
						.AppendToVersionInfoFile = False
					End If

					.GenericDLL = mGenericDLL

					.ShowResultsAtConsole = mShowResultsAtConsole

					.VersionInfoFileName = strVersionInfoFileName

					.IgnoreErrorsWhenUsingWildcardMatching = True
				End With

				If mRecurseFolders Then

					If mDLLVersionInspector.ProcessFilesAndRecurseFolders(mInputFilePath, mOutputFolderNameOrPath, mOutputFolderAlternatePath, mRecreateFolderHierarchyInAlternatePath, mParameterFilePath, mRecurseFoldersMaxLevels) Then
						intReturnCode = 0
					Else
						intReturnCode = mDLLVersionInspector.ErrorCode
					End If
				Else
					If mDLLVersionInspector.ProcessFilesWildcard(mInputFilePath, mOutputFolderNameOrPath, mParameterFilePath) Then
						intReturnCode = 0
					Else
						intReturnCode = mDLLVersionInspector.ErrorCode
						If intReturnCode <> 0 AndAlso Not mQuietMode Then
							Console.WriteLine("Error while processing: " & mDLLVersionInspector.GetErrorMessage())
						End If
					End If
				End If

				DisplayProgressPercent(mLastProgressReportValue, True)

			End If

		Catch ex As Exception
			ShowErrorMessage("Error occurred in modMain->Main: " & Environment.NewLine & ex.Message)
			intReturnCode = -1
		End Try

		Return intReturnCode

	End Function

	Private Sub DisplayProgressPercent(ByVal intPercentComplete As Integer, ByVal blnAddCarriageReturn As Boolean)
		If blnAddCarriageReturn Then
			Console.WriteLine()
        End If

        ' Note: Display of % complete is disabled for this program because it is extremely simple
        Return

        'If intPercentComplete > 100 Then intPercentComplete = 100
        'Console.Write("Processing: " & intPercentComplete.ToString & "% ")
        'If blnAddCarriageReturn Then
        '	Console.WriteLine()
        'End If

	End Sub

	Private Function GetAppVersion() As String
		Return System.Reflection.Assembly.GetExecutingAssembly.GetName.Version.ToString & " (" & PROGRAM_DATE & ")"
	End Function

	Private Function GetAppPath() As String
		Return System.Reflection.Assembly.GetExecutingAssembly().Location
	End Function

	Private Function SetOptionsUsingCommandLineParameters(ByVal objParseCommandLine As clsParseCommandLine) As Boolean
		' Returns True if no problems; otherwise, returns false

		Dim strValue As String = String.Empty
		Dim strValidParameters() As String = New String() {"I", "O", "G", "S", "C"}

		Try
			' Make sure no invalid parameters are present
			If objParseCommandLine.InvalidParametersPresent(strValidParameters) Then
				Return False
			Else
				With objParseCommandLine
					' Query objParseCommandLine to see if various parameters are present

					If .NonSwitchParameterCount > 0 Then
						mInputFilePath = .RetrieveNonSwitchParameter(0)
					End If

					If .RetrieveValueForParameter("I", strValue) Then
						mInputFilePath = String.Copy(strValue)
					End If

					If .RetrieveValueForParameter("O", strValue) Then
						mVersionInfoFilePath = String.Copy(strValue)
					End If

					If .RetrieveValueForParameter("G", strValue) Then
						mGenericDLL = True
					End If

					If .RetrieveValueForParameter("S", strValue) Then
						mRecurseFolders = True
						If Not Integer.TryParse(strValue, mRecurseFoldersMaxLevels) Then
							mRecurseFoldersMaxLevels = 0
						End If
					End If

					If .RetrieveValueForParameter("C", strValue) Then
						mShowResultsAtConsole = True
					End If

				End With

				Return True
			End If

		Catch ex As Exception
			ShowErrorMessage("Error parsing the command line parameters: " & Environment.NewLine & ex.Message)
		End Try

		Return False

	End Function

	Private Sub ShowErrorMessage(ByVal strMessage As String)
		Dim strSeparator As String = "------------------------------------------------------------------------------"

		Console.WriteLine()
		Console.WriteLine(strSeparator)
		Console.WriteLine(strMessage)
		Console.WriteLine(strSeparator)
		Console.WriteLine()

	End Sub

	Private Sub ShowProgramHelp()

		Try

			Console.WriteLine("This program inspects a .NET DLL or .Exe to determine the version.  This allows a 32-bit .NET application to call this program via the command prompt to determine the version of a 64-bit DLL or Exe.")
			Console.WriteLine("The program will create a VersionInfo file with the details about the DLL.")
			Console.WriteLine("Alternatively, you can search for all occurrences of a given DLL in a folder and its subdirectories (use switch /S).  In this mode, the DLL version will be displayed at the console.")
			Console.WriteLine()
			Console.WriteLine("Program syntax:" & Environment.NewLine & System.IO.Path.GetFileName(GetAppPath()))
			Console.WriteLine(" FilePath [/O:VersionInfoFilePath] [/G] [/C] [/S]")
			Console.WriteLine()
			Console.WriteLine("FilePath is the path to the .NET DLL or .NET Exe to inspect")
			Console.WriteLine("Use /O:VersionInfoFilePath to specify the path to the file to which this program should write the version info. If using /S, then use /O to specify the filename that will be created in the folder for which each DLL is found")
			Console.WriteLine()
			Console.WriteLine("Use /G to indicate that the DLL is a generic Windows DLL, and not a .NET DLL")
			Console.WriteLine("Use /C to display the version info in the console output instead of creating a VersionInfo file")
			Console.WriteLine("Use /S to search for all instances of the DLL in a folder and its subfolders (wildcards are allowed)")
			Console.WriteLine()

			Console.WriteLine("Program written by Matthew Monroe for the Department of Energy (PNNL, Richland, WA) in 2011")
			Console.WriteLine("Version: " & GetAppVersion())
			Console.WriteLine()

			Console.WriteLine("E-mail: matthew.monroe@pnnl.gov or matt@alchemistmatt.com")
			Console.WriteLine("Website: http://omics.pnl.gov/ or http://panomics.pnnl.gov/")
			Console.WriteLine()

			' Delay for 750 msec in case the user double clicked this file from within Windows Explorer (or started the program via a shortcut)
			System.Threading.Thread.Sleep(750)

		Catch ex As Exception
			ShowErrorMessage("Error displaying the program syntax: " & ex.Message)
		End Try

	End Sub

	Private Sub mDLLVersionInspector_ProgressChanged(ByVal taskDescription As String, ByVal percentComplete As Single) Handles mDLLVersionInspector.ProgressChanged
		Const PERCENT_REPORT_INTERVAL As Integer = 25
		Const PROGRESS_DOT_INTERVAL_MSEC As Integer = 250

		If percentComplete >= mLastProgressReportValue Then
			If mLastProgressReportValue > 0 Then
				Console.WriteLine()
			End If
			DisplayProgressPercent(mLastProgressReportValue, False)
			mLastProgressReportValue += PERCENT_REPORT_INTERVAL
			mLastProgressReportTime = DateTime.UtcNow
		Else
			If DateTime.UtcNow.Subtract(mLastProgressReportTime).TotalMilliseconds > PROGRESS_DOT_INTERVAL_MSEC Then
				mLastProgressReportTime = DateTime.UtcNow
				Console.Write(".")
			End If
		End If
	End Sub

	Private Sub mDLLVersionInspector_ProgressReset() Handles mDLLVersionInspector.ProgressReset
		mLastProgressReportTime = DateTime.UtcNow
		mLastProgressReportValue = 0
	End Sub
End Module
