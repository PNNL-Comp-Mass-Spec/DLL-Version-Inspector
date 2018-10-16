Option Strict On

Imports System.IO
Imports System.Reflection
Imports PRISM
Imports PRISM.FileProcessor

''' <summary>
'''  This class will determine the version of a .NET DLL or a generic Windows DLL
'''  Written by Matthew Monroe for the Department of Energy (PNNL, Richland, WA)
'''  Started June 14, 2013
''' </summary>
Public Class clsDLLVersionInspector
    Inherits ProcessFilesBase

    Public Sub New()
        MyBase.mFileDate = "October 16, 2018"
        InitializeLocalVariables()
    End Sub

#Region "Constants and Enums"

    Public Const XML_SECTION_OPTIONS As String = "DLLVersionInspectorOptions"

    ' Error codes specialized for this class
    Public Enum eDLLVersionInspectorErrorCodes
        NoError = 0
        ErrorReadingInputFile = 1
        UnspecifiedError = -1
    End Enum

#End Region

#Region "Structures"

#End Region

#Region "Classwide Variables"
    Protected mAppendToVersionInfoFile As Boolean
    Protected mGenericDLL As Boolean
    Protected mShowResultsAtConsole As Boolean = False
    Protected mVersionInfoFileName As String
    Protected mLocalErrorCode As eDLLVersionInspectorErrorCodes
#End Region

#Region "Processing Options Interface Functions"

    Public Property AppendToVersionInfoFile As Boolean
        Get
            Return mAppendToVersionInfoFile
        End Get
        Set
            mAppendToVersionInfoFile = Value
        End Set
    End Property

    Public Property GenericDLL As Boolean
        Get
            Return mGenericDLL
        End Get
        Set
            mGenericDLL = Value
        End Set
    End Property

    Public Property ShowResultsAtConsole As Boolean
        Get
            Return mShowResultsAtConsole
        End Get
        Set
            mShowResultsAtConsole = Value
        End Set
    End Property

    Public Property VersionInfoFileName As String
        Get
            Return mVersionInfoFileName
        End Get
        Set
            mVersionInfoFileName = Value
        End Set
    End Property
#End Region

    ''' <summary>
    ''' Determines the version of a .NET DLL
    ''' </summary>
    ''' <param name="dllFilePath"></param>
    ''' <param name="outputDirectoryPath"></param>
    ''' <returns>True if success, false if an error</returns>
    ''' <remarks></remarks>
    Protected Function DetermineVersionDotNETDll(dllFilePath As String, outputDirectoryPath As String) As Boolean

        Dim toolVersionInfo As String = String.Empty

        Try
            Dim dllInfo = New FileInfo(dllFilePath)

            If Not dllInfo.Exists Then
                Dim errorMessage As String = "Error: File not found: " & dllFilePath
                ShowErrorMessage(errorMessage)
                SaveVersionInfo(dllFilePath, outputDirectoryPath, toolVersionInfo, errorMessage)
                Return False
            Else

                Dim oAssemblyName As AssemblyName
                oAssemblyName = Assembly.LoadFrom(dllInfo.FullName).GetName

                toolVersionInfo = oAssemblyName.Name & ", Version=" & oAssemblyName.Version.ToString()

                Dim success = SaveVersionInfo(dllFilePath, outputDirectoryPath, toolVersionInfo)
                Return success
            End If

        Catch ex As Exception
            ' If you get an exception regarding .NET 4.0 not being able to read a .NET 1.0 runtime, add these lines to the end of file AnalysisManagerProg.exe.config
            '  <startup useLegacyV2RuntimeActivationPolicy="true">
            '    <supportedRuntime version="v4.0" />
            '  </startup>
            Dim errorMessage As String = "Exception determining Assembly info for " & Path.GetFileName(dllFilePath) & ": " & ex.Message
            ShowErrorMessage(errorMessage)
            SaveVersionInfo(dllFilePath, outputDirectoryPath, toolVersionInfo, errorMessage)
            Return False
        End Try

    End Function

    ''' <summary>
    ''' Determines the version of a generic Windows DLL
    ''' </summary>
    ''' <param name="dllFilePath"></param>
    ''' <param name="outputDirectoryPath"></param>
    ''' <returns>True if success, false if an error</returns>
    ''' <remarks></remarks>
    Protected Function DetermineVersionGenericDLL(dllFilePath As String, outputDirectoryPath As String) As Boolean

        Dim toolVersionInfo As String = String.Empty

        Try
            Dim dllInfo = New FileInfo(dllFilePath)

            If Not dllInfo.Exists Then
                Dim errorMessage As String = "Error: File not found: " & dllFilePath
                ShowErrorMessage(errorMessage)
                SaveVersionInfo(dllFilePath, outputDirectoryPath, toolVersionInfo, errorMessage)
                Return False
            Else

                Dim oFileVersionInfo = FileVersionInfo.GetVersionInfo(dllFilePath)

                Dim fileName As String
                Dim dllVersion As String

                fileName = oFileVersionInfo.FileDescription
                If String.IsNullOrEmpty(fileName) Then
                    fileName = oFileVersionInfo.InternalName
                End If

                If String.IsNullOrEmpty(fileName) Then
                    fileName = oFileVersionInfo.FileName
                End If

                If String.IsNullOrEmpty(fileName) Then
                    fileName = dllInfo.Name
                End If

                dllVersion = oFileVersionInfo.FileVersion
                If String.IsNullOrEmpty(dllVersion) Then
                    dllVersion = oFileVersionInfo.ProductVersion
                End If

                If String.IsNullOrEmpty(dllVersion) Then
                    dllVersion = "??"
                End If

                toolVersionInfo = fileName & ", Version=" & dllVersion

                Dim success = SaveVersionInfo(dllFilePath, outputDirectoryPath, toolVersionInfo)
                Return success
            End If

        Catch ex As Exception
            Dim errorMessage As String = "Exception determining Version info for " & Path.GetFileName(dllFilePath) & ": " & ex.Message
            ShowErrorMessage(errorMessage)
            SaveVersionInfo(dllFilePath, outputDirectoryPath, toolVersionInfo, errorMessage)
            Return False
        End Try

    End Function

    Public Overrides Function GetDefaultExtensionsToParse() As IList(Of String)
        Dim extensionsToParse = New List(Of String) From {
                ".dll",
                ".exe"
                }

        Return extensionsToParse

    End Function

    Public Shared Function GetDefaultVersionInfoFileName(dllFileNameOrPath As String) As String
        Return Path.GetFileNameWithoutExtension(dllFileNameOrPath) & "_VersionInfo.txt"
    End Function

    Public Overrides Function GetErrorMessage() As String
        ' Returns "" if no error

        Dim errorMessage As String

        If ErrorCode = ProcessFilesErrorCodes.LocalizedError Or
           ErrorCode = ProcessFilesErrorCodes.NoError Then
            Select Case mLocalErrorCode
                Case eDLLVersionInspectorErrorCodes.NoError
                    errorMessage = ""

                Case eDLLVersionInspectorErrorCodes.ErrorReadingInputFile
                    errorMessage = "Error reading input file"

                Case eDLLVersionInspectorErrorCodes.UnspecifiedError
                    errorMessage = "Unspecified localized error"
                Case Else
                    ' This shouldn't happen
                    errorMessage = "Unknown error state"
            End Select
        Else
            errorMessage = GetBaseClassErrorMessage()
        End If

        Return errorMessage

    End Function

    Private Sub InitializeLocalVariables()
        mGenericDLL = False
        mShowResultsAtConsole = False
        mVersionInfoFileName = String.Empty
        mLocalErrorCode = eDLLVersionInspectorErrorCodes.NoError
    End Sub

    Public Function LoadParameterFileSettings(parameterFilePath As String) As Boolean

        Dim settingsFile As New XmlSettingsFileAccessor

        Try

            If parameterFilePath Is Nothing OrElse parameterFilePath.Length = 0 Then
                ' No parameter file specified; nothing to load
                Return True
            End If

            If Not File.Exists(parameterFilePath) Then
                ' See if parameterFilePath points to a file in the same directory as the application
                parameterFilePath = Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), Path.GetFileName(parameterFilePath))
                If Not File.Exists(parameterFilePath) Then
                    MyBase.SetBaseClassErrorCode(ProcessFilesErrorCodes.ParameterFileNotFound)
                    Return False
                End If
            End If

            If settingsFile.LoadSettings(parameterFilePath) Then
                If Not settingsFile.SectionPresent(XML_SECTION_OPTIONS) Then
                    ShowErrorMessage("The node '<section name=""" & XML_SECTION_OPTIONS & """> was not found in the parameter file: " & parameterFilePath)
                    MyBase.SetBaseClassErrorCode(ProcessFilesErrorCodes.InvalidParameterFile)
                    Return False
                Else
                    Me.GenericDLL = settingsFile.GetParam(XML_SECTION_OPTIONS, "GenericDLL", Me.GenericDLL)
                End If
            End If

        Catch ex As Exception
            HandleException("Error in LoadParameterFileSettings", ex)
            Return False
        End Try

        Return True

    End Function

    ''' <summary>
    ''' Main processing function -- Calls DetermineVersionGenericDLL or DetermineVersionDotNETDll
    ''' </summary>
    ''' <param name="inputFilePath"></param>
    ''' <param name="outputDirectoryPath"></param>
    ''' <param name="parameterFilePath"></param>
    ''' <param name="resetErrorCode"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Overloads Overrides Function ProcessFile(inputFilePath As String, outputDirectoryPath As String, parameterFilePath As String, resetErrorCode As Boolean) As Boolean
        ' Returns True if success, False if failure

        Dim inputFile As FileInfo
        Dim inputFilePathFull As String

        Dim success As Boolean

        If resetErrorCode Then
            SetLocalErrorCode(eDLLVersionInspectorErrorCodes.NoError)
        End If

        If Not LoadParameterFileSettings(parameterFilePath) Then
            ShowErrorMessage("Parameter file load error: " & parameterFilePath)

            If MyBase.ErrorCode = ProcessFilesErrorCodes.NoError Then
                MyBase.SetBaseClassErrorCode(ProcessFilesErrorCodes.InvalidParameterFile)
            End If
            Return False
        End If

        Try
            If inputFilePath Is Nothing OrElse inputFilePath.Length = 0 Then
                ShowMessage("Input file name is empty")
                MyBase.SetBaseClassErrorCode(ProcessFilesErrorCodes.InvalidInputFilePath)
                Return False
            End If

            Console.WriteLine()
            Console.WriteLine("Parsing " & Path.GetFileName(inputFilePath))

            ' Note that CleanupFilePaths() will update mOutputDirectoryPath, which is used by LogMessage()
            If Not CleanupFilePaths(inputFilePath, outputDirectoryPath) Then
                MyBase.SetBaseClassErrorCode(ProcessFilesErrorCodes.FilePathError)
                Return False
            End If

            MyBase.ResetProgress()

            Try
                ' Obtain the full path to the input file
                inputFile = New FileInfo(inputFilePath)
                inputFilePathFull = inputFile.FullName

                If mGenericDLL Then
                    success = DetermineVersionGenericDLL(inputFilePathFull, outputDirectoryPath)
                Else
                    success = DetermineVersionDotNETDll(inputFilePathFull, outputDirectoryPath)
                End If


                If success Then
                    ShowMessage(String.Empty, False)
                Else
                    SetLocalErrorCode(eDLLVersionInspectorErrorCodes.UnspecifiedError)
                    ShowErrorMessage("Error")
                End If

            Catch ex As Exception
                HandleException("Error calling DetermineVersionGenericDLL or DetermineVersionDotNETDll", ex)
            End Try


            Return success

        Catch ex As Exception
            HandleException("Error in ProcessFile", ex)
            Return False
        End Try


    End Function

    Private Function SaveVersionInfo(dllFilePath As String, outputDirectoryPath As String, toolVersionInfo As String) As Boolean
        Return SaveVersionInfo(dllFilePath, outputDirectoryPath, toolVersionInfo, String.Empty)
    End Function

    Private Function SaveVersionInfo(dllFilePath As String, outputDirectoryPath As String, toolVersionInfo As String, errorMessage As String) As Boolean

        ' This list tracks the DLL name, path, and version
        Dim versionInfo = New List(Of String)

        Try

            Dim dllFile = New FileInfo(dllFilePath)

            If dllFile.Exists Then
                versionInfo.Add("FileName=" & dllFile.Name)
                versionInfo.Add("Path=" & dllFile.FullName)
            Else
                versionInfo.Add("FileName=" & Path.GetFileName(dllFilePath))
                versionInfo.Add("Path=" & dllFilePath)
            End If

            versionInfo.Add("Version=" & toolVersionInfo)

            If Not String.IsNullOrWhiteSpace(errorMessage) Then
                versionInfo.Add("Error=" & errorMessage)
            End If

            If mShowResultsAtConsole Then
                For Each item In versionInfo
                    Console.WriteLine(item)
                Next
            Else

                If String.IsNullOrWhiteSpace(mVersionInfoFileName) Then
                    ' Auto-define the output file name
                    mVersionInfoFileName = GetDefaultVersionInfoFileName(dllFilePath)
                End If

                ' Auto-define the output file name
                If String.IsNullOrWhiteSpace(outputDirectoryPath) Then
                    Dim appFileInfo = New FileInfo(GetAppPath())
                    outputDirectoryPath = appFileInfo.DirectoryName
                End If

                Dim versionInfoFilePath = Path.Combine(outputDirectoryPath, mVersionInfoFileName)

                Try
                    Dim eFileMode As FileMode
                    If mAppendToVersionInfoFile Then
                        eFileMode = FileMode.Append
                    Else
                        eFileMode = FileMode.Create
                    End If

                    Using writer = New StreamWriter(New FileStream(versionInfoFilePath, eFileMode, FileAccess.Write, FileShare.Read))
                        For Each item In versionInfo
                            writer.WriteLine(item)
                        Next

                        If mAppendToVersionInfoFile Then
                            writer.WriteLine()
                        End If

                    End Using

                Catch ex As Exception
                    ShowErrorMessage("Exception writing the version info to the output file at " & versionInfoFilePath & ": " & ex.Message)
                    Return False
                End Try

            End If

        Catch ex As Exception
            ShowErrorMessage("Exception in SaveVersionInfo: " & ex.Message)
            Return False
        End Try

        Return True

    End Function

    Private Sub SetLocalErrorCode(eNewErrorCode As eDLLVersionInspectorErrorCodes)
        SetLocalErrorCode(eNewErrorCode, False)
    End Sub

    Private Sub SetLocalErrorCode(eNewErrorCode As eDLLVersionInspectorErrorCodes, leaveExistingErrorCodeUnchanged As Boolean)

        If leaveExistingErrorCodeUnchanged AndAlso mLocalErrorCode <> eDLLVersionInspectorErrorCodes.NoError Then
            ' An error code is already defined; do not change it
        Else
            mLocalErrorCode = eNewErrorCode

            If eNewErrorCode = eDLLVersionInspectorErrorCodes.NoError Then
                If MyBase.ErrorCode = ProcessFilesErrorCodes.LocalizedError Then
                    MyBase.SetBaseClassErrorCode(ProcessFilesErrorCodes.NoError)
                End If
            Else
                MyBase.SetBaseClassErrorCode(ProcessFilesErrorCodes.LocalizedError)
            End If
        End If

    End Sub

End Class
