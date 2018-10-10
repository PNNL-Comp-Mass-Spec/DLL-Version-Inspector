Option Strict On

Imports System.IO
Imports System.Threading
Imports PRISM

' This program inspects a .NET DLL or .Exe to determine the version.
' This allows a 32-bit .NET application to call this program via the
' command prompt to determine the version of a 64-bit DLL or Exe.
'
' -------------------------------------------------------------------------------
' Written by Matthew Monroe for the Department of Energy (PNNL, Richland, WA)
' Program started in April 2011
'
' E-mail: matthew.monroe@pnnl.gov or proteomics@pnnl.gov
' Website: https://omics.pnl.gov/ or https://panomics.pnnl.gov/
' -------------------------------------------------------------------------------
'

Module modMain
    Public Const PROGRAM_DATE As String = "October 9, 2018"

    Private mInputFilePath As String
    Private mVersionInfoFilePath As String = String.Empty
    Private mGenericDLL As Boolean = False

    Private mOutputFolderNameOrPath As String
    Private mParameterFilePath As String            ' Not used by this program

    Private mOutputFolderAlternatePath As String                ' Not used by this program
    Private mRecreateFolderHierarchyInAlternatePath As Boolean  ' Not used by this program

    Private mRecurseFolders As Boolean
    Private mRecurseFoldersMaxLevels As Integer

    Private mShowResultsAtConsole As Boolean = False

    Private mLastProgressReportTime As DateTime
    Private mLastProgressReportValue As Integer

    ''' <summary>
    ''' Main method
    ''' </summary>
    ''' <returns>0 if no error, error code if an error</returns>
    Public Function Main() As Integer

        Dim returnCode As Integer
        Dim commandLineParser As New clsParseCommandLine
        Dim proceed As Boolean

        mOutputFolderNameOrPath = String.Empty
        mParameterFilePath = String.Empty

        mOutputFolderAlternatePath = String.Empty
        mRecreateFolderHierarchyInAlternatePath = False

        mRecurseFolders = False
        mRecurseFoldersMaxLevels = 0

        mShowResultsAtConsole = False

        Try
            proceed = False

            If commandLineParser.ParseCommandLine Then
                If SetOptionsUsingCommandLineParameters(commandLineParser) Then proceed = True
            End If

            If Not proceed OrElse
               commandLineParser.NeedToShowHelp OrElse
               commandLineParser.ParameterCount + commandLineParser.NonSwitchParameterCount = 0 OrElse
               String.IsNullOrWhiteSpace(mInputFilePath) Then
                ShowProgramHelp()
                returnCode = -1
            Else

                Dim versionInfoFileName As String = String.Empty
                If Not String.IsNullOrWhiteSpace(mVersionInfoFilePath) Then
                    Dim versionInfoFile = New FileInfo(mVersionInfoFilePath)

                    versionInfoFileName = versionInfoFile.Name
                    mOutputFolderNameOrPath = versionInfoFile.Directory.FullName

                    If mRecurseFolders AndAlso versionInfoFile.Exists Then
                        ' Delete the existing version info file because we will be appending to it
                        versionInfoFile.Delete()
                    End If
                End If

                ' Note: the following settings will be overridden if mParameterFilePath points to a valid parameter file that has these settings defined
                Dim dllVersionInspector = New clsDLLVersionInspector() With {
                        .LogMessagesToFile = False,
                        .AppendToVersionInfoFile = False,
                        .GenericDLL = mGenericDLL,
                        .ShowResultsAtConsole = mShowResultsAtConsole,
                        .VersionInfoFileName = versionInfoFileName,
                        .IgnoreErrorsWhenUsingWildcardMatching = True
                        }

                AddHandler dllVersionInspector.ProgressUpdate, AddressOf DLLVersionInspector_ProgressUpdate
                AddHandler dllVersionInspector.ProgressReset, AddressOf DLLVersionInspector_ProgressReset

                If mRecurseFolders AndAlso Not String.IsNullOrWhiteSpace(mOutputFolderNameOrPath) Then
                    dllVersionInspector.AppendToVersionInfoFile = True
                End If

                If mRecurseFolders Then

                    If dllVersionInspector.ProcessFilesAndRecurseFolders(mInputFilePath, mOutputFolderNameOrPath, mOutputFolderAlternatePath, mRecreateFolderHierarchyInAlternatePath, mParameterFilePath, mRecurseFoldersMaxLevels) Then
                        returnCode = 0
                    Else
                        returnCode = dllVersionInspector.ErrorCode
                    End If
                Else
                    If dllVersionInspector.ProcessFilesWildcard(mInputFilePath, mOutputFolderNameOrPath, mParameterFilePath) Then
                        returnCode = 0
                    Else
                        returnCode = dllVersionInspector.ErrorCode
                        If returnCode <> 0 Then
                            Console.WriteLine("Error while processing: " & dllVersionInspector.GetErrorMessage())
                        End If
                    End If
                End If

                DisplayProgressPercent(mLastProgressReportValue, True)

            End If

        Catch ex As Exception
            ShowErrorMessage("Error occurred in modMain->Main", ex)
            returnCode = -1
        End Try

        Return returnCode

    End Function

    Private Sub DisplayProgressPercent(percentComplete As Integer, addCarriageReturn As Boolean)
        If addCarriageReturn Then
            Console.WriteLine()
        End If

        ' Note: Display of % complete is disabled for this program because it is extremely simple
        Return

        'If percentComplete > 100 Then percentComplete = 100
        'Console.Write("Processing: " & percentComplete.ToString & "% ")
        'If addCarriageReturn Then
        '    Console.WriteLine()
        'End If

    End Sub

    Private Function GetAppVersion() As String
        Return PRISM.FileProcessor.ProcessFoldersBase.GetAppVersion(PROGRAM_DATE)
    End Function

    Private Function GetAppPath() As String
        Return PRISM.FileProcessor.ProcessFoldersBase.GetAppPath()
    End Function

    Private Function SetOptionsUsingCommandLineParameters(commandLineParser As clsParseCommandLine) As Boolean
        ' Returns True if no problems; otherwise, returns false

        Dim value As String = String.Empty
        Dim validParameters = New List(Of String) From {"I", "O", "G", "S", "C"}

        Try
            ' Make sure no invalid parameters are present
            If commandLineParser.InvalidParametersPresent(validParameters) Then
                Return False
            End If

            ' Query commandLineParser to see if various parameters are present

            If commandLineParser.NonSwitchParameterCount > 0 Then
                mInputFilePath = commandLineParser.RetrieveNonSwitchParameter(0)
            End If

            If commandLineParser.RetrieveValueForParameter("I", value) Then
                mInputFilePath = String.Copy(value)
            End If

            If commandLineParser.RetrieveValueForParameter("O", value) Then
                mVersionInfoFilePath = String.Copy(value)
            End If

            mGenericDLL = commandLineParser.IsParameterPresent("G")
            mShowResultsAtConsole = commandLineParser.IsParameterPresent("C")

            If commandLineParser.RetrieveValueForParameter("S", value) Then
                mRecurseFolders = True
                If Not Integer.TryParse(value, mRecurseFoldersMaxLevels) Then
                    mRecurseFoldersMaxLevels = 0
                End If
            End If

            Return True

        Catch ex As Exception
            ShowErrorMessage("Error parsing the command line parameters", ex)
            Return False
        End Try

    End Function

    Private Sub ShowErrorMessage(message As String, ex As Exception)
        ConsoleMsgUtils.ShowError(message, ex)
    End Sub

    Private Sub ShowProgramHelp()

        Try

            Console.WriteLine(ConsoleMsgUtils.WrapParagraph(
                "This program inspects a .NET DLL or .Exe to determine the version. " &
                "This allows a 32-bit .NET application to call this program via the command prompt " &
                "to determine the version of a 64-bit DLL or Exe. " &
                "The program will create a VersionInfo file with the details about the DLL." &
                "Alternatively, you can search for all occurrences of a given DLL in a folder and its subdirectories (use switch /S). " &
                "In this mode, the DLL version will be displayed at the console."))
            Console.WriteLine()
            Console.WriteLine("Program syntax: " & Path.GetFileName(GetAppPath()))
            Console.WriteLine("  FilePath [/O:VersionInfoFilePath] [/G] [/C] [/S]")
            Console.WriteLine()
            Console.WriteLine("FilePath is the path to the .NET DLL or .NET Exe to inspect")
            Console.WriteLine()
            Console.WriteLine(ConsoleMsgUtils.WrapParagraph(
                "Use /O:VersionInfoFilePath to specify the path to the file to which this program should write the version info. " &
                "If using /S, then use /O to specify the filename that will be created in the folder for which each DLL is found"))
            Console.WriteLine()
            Console.WriteLine(ConsoleMsgUtils.WrapParagraph("Use /G to indicate that the DLL is a generic Windows DLL, and not a .NET DLL"))
            Console.WriteLine(ConsoleMsgUtils.WrapParagraph("Use /C to display the version info in the console output instead of creating a VersionInfo file"))
            Console.WriteLine(ConsoleMsgUtils.WrapParagraph("Use /S to search for all instances of the DLL in a folder and its subfolders (wildcards are allowed)"))
            Console.WriteLine()

            Console.WriteLine("Program written by Matthew Monroe for the Department of Energy (PNNL, Richland, WA) in 2011")
            Console.WriteLine("Version: " & GetAppVersion())
            Console.WriteLine()

            Console.WriteLine("E-mail: matthew.monroe@pnnl.gov or proteomics@pnnl.gov")
            Console.WriteLine("Website: https://omics.pnl.gov/ or https://panomics.pnnl.gov/")
            Console.WriteLine()

            ' Delay for 750 msec in case the user double clicked this file from within Windows Explorer (or started the program via a shortcut)
            Thread.Sleep(750)

        Catch ex As Exception
            ShowErrorMessage("Error displaying the program syntax", ex)
        End Try

    End Sub

    Private Sub DLLVersionInspector_ProgressUpdate(taskDescription As String, percentComplete As Single)
        Const PERCENT_REPORT_INTERVAL = 25
        Const PROGRESS_DOT_INTERVAL_MSEC = 250

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

    Private Sub DLLVersionInspector_ProgressReset()
        mLastProgressReportTime = DateTime.UtcNow
        mLastProgressReportValue = 0
    End Sub
End Module
