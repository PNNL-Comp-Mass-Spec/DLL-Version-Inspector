Option Strict On

' This program inspects a .NET DLL or .Exe to determine the version.  
' This allows a 32-bit .NET application to call this program via the 
' command prompt to determine the version of a 64-bit DLL or Exe.
'
' -------------------------------------------------------------------------------
' Written by Matthew Monroe for the Department of Energy (PNNL, Richland, WA)
' Program started in April 2011
'
' E-mail: matthew.monroe@pnl.gov or matt@alchemistmatt.com
' Website: http://ncrr.pnl.gov/ or http://omics.pnl.gov or http://www.sysbio.org/resources/staff/
' -------------------------------------------------------------------------------
' 

Module modMain
	Public Const PROGRAM_DATE As String = "November 8, 2011"

	Private mFilePath As String = String.Empty
	Private mVersionInfoFilePath As String = String.Empty

	Public Function Main() As Integer
		' Returns 0 if no error, error code if an error

		Dim intReturnCode As Integer
		Dim objParseCommandLine As New clsParseCommandLine
		Dim blnProceed As Boolean

		intReturnCode = 0

		Try
			blnProceed = False

			If objParseCommandLine.ParseCommandLine Then
				If SetOptionsUsingCommandLineParameters(objParseCommandLine) Then blnProceed = True
			End If

			If Not blnProceed OrElse _
			   objParseCommandLine.NeedToShowHelp OrElse _
			   objParseCommandLine.ParameterCount + objParseCommandLine.NonSwitchParameterCount = 0 Then
				ShowProgramHelp()
				intReturnCode = -1
			Else

				If Not String.IsNullOrWhiteSpace(mFilePath) Then
					Dim blnSuccess As Boolean
					blnSuccess = DetermineVersion(mFilePath, mVersionInfoFilePath)

					If Not blnSuccess Then
						intReturnCode = 99
					End If

				Else
					ShowProgramHelp()
					intReturnCode = -1
				End If

			End If

		Catch ex As Exception
			ShowErrorMessage("Error occurred in modMain->Main: " & Environment.NewLine & ex.Message)
			intReturnCode = -1
		End Try

		Return intReturnCode

	End Function

	Private Function DetermineVersion(ByVal strFilePath As String, ByVal strVersionInfoFilePath As String) As Boolean

		Dim ioFileInfo As System.IO.FileInfo
		Dim strToolVersionInfo As String = String.Empty
		Dim blnSuccess As Boolean = False

		Try
			ioFileInfo = New System.IO.FileInfo(strFilePath)

			If Not ioFileInfo.Exists Then
				Dim strErrorMessage As String = "Error: File not found: " & strFilePath
				ShowErrorMessage(strErrorMessage)
				SaveVersionInfo(strFilePath, strVersionInfoFilePath, strToolVersionInfo, strErrorMessage)
				blnSuccess = False
			Else

				Dim oAssemblyName As System.Reflection.AssemblyName
				oAssemblyName = System.Reflection.Assembly.LoadFrom(ioFileInfo.FullName).GetName

				strToolVersionInfo = oAssemblyName.Name & ", Version=" & oAssemblyName.Version.ToString()

				blnSuccess = SaveVersionInfo(strFilePath, strVersionInfoFilePath, strToolVersionInfo)
			End If

		Catch ex As Exception
			' If you get an exception regarding .NET 4.0 not being able to read a .NET 1.0 runtime, then add these lines to the end of file AnalysisManagerProg.exe.config
			'  <startup useLegacyV2RuntimeActivationPolicy="true">
			'    <supportedRuntime version="v4.0" />
			'  </startup>
			Dim strErrorMessage As String = "Exception determining Assembly info for " & System.IO.Path.GetFileName(strFilePath) & ": " & ex.Message
			ShowErrorMessage(strErrorMessage)
			SaveVersionInfo(strFilePath, strVersionInfoFilePath, strToolVersionInfo, strErrorMessage)
			blnSuccess = False
		End Try

		Return blnSuccess

	End Function

	Private Function GetAppVersion() As String
		Return System.Reflection.Assembly.GetExecutingAssembly.GetName.Version.ToString & " (" & PROGRAM_DATE & ")"
	End Function

	Private Function GetAppPath() As String
		Return System.Reflection.Assembly.GetExecutingAssembly().Location
	End Function

	Private Function SaveVersionInfo(ByVal strFilePath As String, ByVal strVersionInfoFilePath As String, ByVal strToolVersionInfo As String) As Boolean
		Return SaveVersionInfo(strFilePath, strVersionInfoFilePath, strToolVersionInfo, String.Empty)
	End Function

	Private Function SaveVersionInfo(ByVal strFilePath As String, ByVal strVersionInfoFilePath As String, ByVal strToolVersionInfo As String, ByVal strErrorMessage As String) As Boolean

		Dim ioFileInfo As System.IO.FileInfo
		Dim srOutFile As System.IO.StreamWriter
		Dim blnSuccess As Boolean = False

		Try
			If String.IsNullOrWhiteSpace(strVersionInfoFilePath) Then
				' Auto-define the output file path
				Dim ioAppFileInfo As System.IO.FileInfo = New System.IO.FileInfo(GetAppPath())

				strVersionInfoFilePath = System.IO.Path.Combine(ioAppFileInfo.DirectoryName, System.IO.Path.GetFileNameWithoutExtension(strFilePath) & "_VersionInfo.txt")
			End If

			srOutFile = New System.IO.StreamWriter(New System.IO.FileStream(strVersionInfoFilePath, IO.FileMode.Create, IO.FileAccess.Write, IO.FileShare.Read))

			ioFileInfo = New System.IO.FileInfo(strFilePath)

			If ioFileInfo.Exists Then
				srOutFile.WriteLine("FileName=" & ioFileInfo.Name)
				srOutFile.WriteLine("Path=" & ioFileInfo.FullName)
			Else
				srOutFile.WriteLine("FileName=" & System.IO.Path.GetFileName(strFilePath))
				srOutFile.WriteLine("Path=" & strFilePath)
			End If

			srOutFile.WriteLine("Version=" & strToolVersionInfo)

			If Not String.IsNullOrWhiteSpace(strErrorMessage) Then
				srOutFile.WriteLine("Error=" & strErrorMessage)
			End If

			srOutFile.Close()

			blnSuccess = True
		Catch ex As Exception
			ShowErrorMessage("Exception writing the version info to the output file at " & strVersionInfoFilePath & ": " & ex.Message)
			blnSuccess = False
		End Try

		Return blnSuccess

	End Function

	Private Function SetOptionsUsingCommandLineParameters(ByVal objParseCommandLine As clsParseCommandLine) As Boolean
		' Returns True if no problems; otherwise, returns false

		Dim strValue As String = String.Empty
		Dim strValidParameters() As String = New String() {"I", "O"}

		Try
			' Make sure no invalid parameters are present
			If objParseCommandLine.InvalidParametersPresent(strValidParameters) Then
				Return False
			Else
				With objParseCommandLine
					' Query objParseCommandLine to see if various parameters are present

					If .NonSwitchParameterCount > 0 Then
						mFilePath = .RetrieveNonSwitchParameter(0)
					End If

					If .RetrieveValueForParameter("I", strValue) Then
						mFilePath = String.Copy(strValue)
					End If

					If .RetrieveValueForParameter("O", strValue) Then
						mVersionInfoFilePath = String.Copy(strValue)
					End If

				End With

				Return True
			End If

		Catch ex As Exception
			ShowErrorMessage("Error parsing the command line parameters: " & Environment.NewLine & ex.Message)
		End Try

		Return False

	End Function

	Private Sub ShowErrorMessage(strErrorMessage As String)
		Console.WriteLine()
		Console.WriteLine("=====================================================================")
		Console.WriteLine(strErrorMessage)
		Console.WriteLine("=====================================================================")
	End Sub

    Private Sub ShowProgramHelp()

        Try

			Console.WriteLine("This program inspects a .NET DLL or .Exe to determine the version.  This allows a 32-bit .NET application to call this program via the command prompt to determine the version of a 64-bit DLL or Exe.")
            Console.WriteLine()
            Console.WriteLine("Program syntax:" & Environment.NewLine & System.IO.Path.GetFileName(System.Reflection.Assembly.GetExecutingAssembly().Location))
			Console.WriteLine(" FilePath [/O:VersionInfoFilePath]")
            Console.WriteLine()
			Console.WriteLine("FilePath is the path to the .NET DLL or .NET Exe to inspect")
			Console.WriteLine("Use /O:VersionInfoFilePath to specify the path to the file to which this program should write the version info")
			Console.WriteLine()

            Console.WriteLine("Program written by Matthew Monroe for the Department of Energy (PNNL, Richland, WA) in 2011")
            Console.WriteLine("Version: " & GetAppVersion())
            Console.WriteLine()

            Console.WriteLine("E-mail: matthew.monroe@pnl.gov or matt@alchemistmatt.com")
            Console.WriteLine("Website: http://omics.pnl.gov/ or http://www.sysbio.org/resources/staff/")
            Console.WriteLine()

            ' Delay for 750 msec in case the user double clicked this file from within Windows Explorer (or started the program via a shortcut)
            System.Threading.Thread.Sleep(750)

        Catch ex As Exception
			ShowErrorMessage("Error displaying the program syntax: " & ex.Message)
        End Try

    End Sub

 
End Module
