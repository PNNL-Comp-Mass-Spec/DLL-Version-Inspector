Option Strict On

' This class will determine the version of a .NET DLL or a generic Windows DLL
'
' Written by Matthew Monroe for the Department of Energy (PNNL, Richland, WA)
'
' Started June 14, 2013

Public Class clsDLLVersionInspector
    Inherits clsProcessFilesBaseClass

    Public Sub New()
        MyBase.mFileDate = "June 14, 2013"
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
        Set(value As Boolean)
            mAppendToVersionInfoFile = value
        End Set
    End Property

    Public Property GenericDLL() As Boolean
        Get
            Return mGenericDLL
        End Get
        Set(value As Boolean)
            mGenericDLL = value
        End Set
    End Property

    Public ReadOnly Property LocalErrorCode() As eDLLVersionInspectorErrorCodes
        Get
            Return mLocalErrorCode
        End Get
    End Property

    Public Property ShowResultsAtConsole As Boolean
        Get
            Return mShowResultsAtConsole
        End Get
        Set(value As Boolean)
            mShowResultsAtConsole = value
        End Set
    End Property

    Public Property VersionInfoFileName As String
        Get
            Return mVersionInfoFileName
        End Get
        Set(value As String)
            mVersionInfoFileName = value
        End Set
    End Property
#End Region

    ''' <summary>
    ''' Determines the version of a .NET DLL
    ''' </summary>
    ''' <param name="strDLLFilePath"></param>
    ''' <param name="strOutputFolderPath"></param>
    ''' <param name="strVersionInfoFileName"></param>
    ''' <returns>True if success, false if an error</returns>
    ''' <remarks></remarks>
    Protected Function DetermineVersionDotNETDll(ByVal strDLLFilePath As String, ByVal strOutputFolderPath As String, ByVal strVersionInfoFileName As String) As Boolean

        Dim ioFileInfo As System.IO.FileInfo
        Dim strToolVersionInfo As String = String.Empty
        Dim blnSuccess As Boolean = False

        Try
            ioFileInfo = New System.IO.FileInfo(strDLLFilePath)

            If Not ioFileInfo.Exists Then
                Dim strErrorMessage As String = "Error: File not found: " & strDLLFilePath
                ShowErrorMessage(strErrorMessage)
                SaveVersionInfo(strDLLFilePath, strOutputFolderPath, strVersionInfoFileName, strToolVersionInfo, strErrorMessage)
                blnSuccess = False
            Else

                Dim oAssemblyName As System.Reflection.AssemblyName
                oAssemblyName = System.Reflection.Assembly.LoadFrom(ioFileInfo.FullName).GetName

                strToolVersionInfo = oAssemblyName.Name & ", Version=" & oAssemblyName.Version.ToString()

                blnSuccess = SaveVersionInfo(strDLLFilePath, strOutputFolderPath, strVersionInfoFileName, strToolVersionInfo)
            End If

        Catch ex As Exception
            ' If you get an exception regarding .NET 4.0 not being able to read a .NET 1.0 runtime, then add these lines to the end of file AnalysisManagerProg.exe.config
            '  <startup useLegacyV2RuntimeActivationPolicy="true">
            '    <supportedRuntime version="v4.0" />
            '  </startup>
            Dim strErrorMessage As String = "Exception determining Assembly info for " & System.IO.Path.GetFileName(strDLLFilePath) & ": " & ex.Message
            ShowErrorMessage(strErrorMessage)
            SaveVersionInfo(strDLLFilePath, strOutputFolderPath, strVersionInfoFileName, strToolVersionInfo, strErrorMessage)
            blnSuccess = False
        End Try

        Return blnSuccess

    End Function

    ''' <summary>
    ''' Determines the version of a generic Windows DLL
    ''' </summary>
    ''' <param name="strDLLFilePath"></param>
    ''' <param name="strOutputFolderPath"></param>
    ''' <param name="strVersionInfoFileName"></param>
    ''' <returns>True if success, false if an error</returns>
    ''' <remarks></remarks>
    Protected Function DetermineVersionGenericDLL(ByVal strDLLFilePath As String, ByVal strOutputFolderPath As String, ByVal strVersionInfoFileName As String) As Boolean

        Dim ioFileInfo As System.IO.FileInfo
        Dim strToolVersionInfo As String = String.Empty
        Dim blnSuccess As Boolean = False

        Try
            ioFileInfo = New System.IO.FileInfo(strDLLFilePath)

            If Not ioFileInfo.Exists Then
                Dim strErrorMessage As String = "Error: File not found: " & strDLLFilePath
                ShowErrorMessage(strErrorMessage)
                SaveVersionInfo(strDLLFilePath, strOutputFolderPath, strVersionInfoFileName, strToolVersionInfo, strErrorMessage)
                blnSuccess = False
            Else

                Dim oFileVersionInfo As System.Diagnostics.FileVersionInfo
                oFileVersionInfo = System.Diagnostics.FileVersionInfo.GetVersionInfo(strDLLFilePath)

                Dim strName As String
                Dim strVersion As String

                strName = oFileVersionInfo.FileDescription
                If String.IsNullOrEmpty(strName) Then
                    strName = oFileVersionInfo.InternalName
                End If

                If String.IsNullOrEmpty(strName) Then
                    strName = oFileVersionInfo.FileName
                End If

                If String.IsNullOrEmpty(strName) Then
                    strName = ioFileInfo.Name
                End If

                strVersion = oFileVersionInfo.FileVersion
                If String.IsNullOrEmpty(strVersion) Then
                    strVersion = oFileVersionInfo.ProductVersion
                End If

                If String.IsNullOrEmpty(strVersion) Then
                    strVersion = "??"
                End If

                strToolVersionInfo = strName & ", Version=" & strVersion

                blnSuccess = SaveVersionInfo(strDLLFilePath, strOutputFolderPath, strVersionInfoFileName, strToolVersionInfo)
            End If

        Catch ex As Exception
            Dim strErrorMessage As String = "Exception determining Version info for " & System.IO.Path.GetFileName(strDLLFilePath) & ": " & ex.Message
            ShowErrorMessage(strErrorMessage)
            SaveVersionInfo(strDLLFilePath, strOutputFolderPath, strVersionInfoFileName, strToolVersionInfo, strErrorMessage)
            blnSuccess = False
        End Try

        Return blnSuccess

    End Function

    Public Overrides Function GetDefaultExtensionsToParse() As String()
        Dim strExtensionsToParse(1) As String

        strExtensionsToParse(0) = ".dll"
        strExtensionsToParse(1) = ".exe"

        Return strExtensionsToParse

    End Function

    Public Shared Function GetDefaultVersionInfoFileName(ByVal strDllFileNameOrPath As String) As String
        Return System.IO.Path.GetFileNameWithoutExtension(strDllFileNameOrPath) & "_VersionInfo.txt"
    End Function

    Public Overrides Function GetErrorMessage() As String
        ' Returns "" if no error

        Dim strErrorMessage As String

        If MyBase.ErrorCode = clsProcessFilesBaseClass.eProcessFilesErrorCodes.LocalizedError Or
           MyBase.ErrorCode = clsProcessFilesBaseClass.eProcessFilesErrorCodes.NoError Then
            Select Case mLocalErrorCode
                Case eDLLVersionInspectorErrorCodes.NoError
                    strErrorMessage = ""

                Case eDLLVersionInspectorErrorCodes.ErrorReadingInputFile
                    strErrorMessage = "Error reading input file"

                Case eDLLVersionInspectorErrorCodes.UnspecifiedError
                    strErrorMessage = "Unspecified localized error"
                Case Else
                    ' This shouldn't happen
                    strErrorMessage = "Unknown error state"
            End Select
        Else
            strErrorMessage = MyBase.GetBaseClassErrorMessage()
        End If

        Return strErrorMessage

    End Function

    Private Sub InitializeLocalVariables()
        mGenericDLL = False
        mShowResultsAtConsole = False
        mVersionInfoFileName = String.Empty
        mLocalErrorCode = eDLLVersionInspectorErrorCodes.NoError
    End Sub

    Public Function LoadParameterFileSettings(ByVal strParameterFilePath As String) As Boolean

        Dim objSettingsFile As New XmlSettingsFileAccessor

        Try

            If strParameterFilePath Is Nothing OrElse strParameterFilePath.Length = 0 Then
                ' No parameter file specified; nothing to load
                Return True
            End If

            If Not System.IO.File.Exists(strParameterFilePath) Then
                ' See if strParameterFilePath points to a file in the same directory as the application
                strParameterFilePath = System.IO.Path.Combine(System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location), System.IO.Path.GetFileName(strParameterFilePath))
                If Not System.IO.File.Exists(strParameterFilePath) Then
                    MyBase.SetBaseClassErrorCode(clsProcessFilesBaseClass.eProcessFilesErrorCodes.ParameterFileNotFound)
                    Return False
                End If
            End If

            If objSettingsFile.LoadSettings(strParameterFilePath) Then
                If Not objSettingsFile.SectionPresent(XML_SECTION_OPTIONS) Then
                    ShowErrorMessage("The node '<section name=""" & XML_SECTION_OPTIONS & """> was not found in the parameter file: " & strParameterFilePath)
                    MyBase.SetBaseClassErrorCode(clsProcessFilesBaseClass.eProcessFilesErrorCodes.InvalidParameterFile)
                    Return False
                Else
                    Me.GenericDLL = objSettingsFile.GetParam(XML_SECTION_OPTIONS, "GenericDLL", Me.GenericDLL)
                End If
            End If

        Catch ex As Exception
            HandleException("Error in LoadParameterFileSettings", ex)
            Return False
        End Try

        Return True

    End Function

    ''' <summary>
    ''' Main processing function -- Calls InspectDLL
    ''' </summary>
    ''' <param name="strInputFilePath"></param>
    ''' <param name="strOutputFolderPath"></param>
    ''' <param name="strParameterFilePath"></param>
    ''' <param name="blnResetErrorCode"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Overloads Overrides Function ProcessFile(ByVal strInputFilePath As String, ByVal strOutputFolderPath As String, ByVal strParameterFilePath As String, ByVal blnResetErrorCode As Boolean) As Boolean
        ' Returns True if success, False if failure

        Dim ioFile As System.IO.FileInfo
        Dim strInputFilePathFull As String

        Dim blnSuccess As Boolean

        If blnResetErrorCode Then
            SetLocalErrorCode(eDLLVersionInspectorErrorCodes.NoError)
        End If

        If Not LoadParameterFileSettings(strParameterFilePath) Then
            ShowErrorMessage("Parameter file load error: " & strParameterFilePath)

            If MyBase.ErrorCode = clsProcessFilesBaseClass.eProcessFilesErrorCodes.NoError Then
                MyBase.SetBaseClassErrorCode(clsProcessFilesBaseClass.eProcessFilesErrorCodes.InvalidParameterFile)
            End If
            Return False
        End If

        Try
            If strInputFilePath Is Nothing OrElse strInputFilePath.Length = 0 Then
                ShowMessage("Input file name is empty")
                MyBase.SetBaseClassErrorCode(clsProcessFilesBaseClass.eProcessFilesErrorCodes.InvalidInputFilePath)
            Else

                Console.WriteLine()
                Console.WriteLine("Parsing " & System.IO.Path.GetFileName(strInputFilePath))

                ' Note that CleanupFilePaths() will update mOutputFolderPath, which is used by LogMessage()
                If Not CleanupFilePaths(strInputFilePath, strOutputFolderPath) Then
                    MyBase.SetBaseClassErrorCode(clsProcessFilesBaseClass.eProcessFilesErrorCodes.FilePathError)
                Else

                    MyBase.ResetProgress()

                    Try
                        ' Obtain the full path to the input file
                        ioFile = New System.IO.FileInfo(strInputFilePath)
                        strInputFilePathFull = ioFile.FullName

                        If mGenericDLL Then
                            blnSuccess = DetermineVersionGenericDLL(strInputFilePathFull, strOutputFolderPath, mVersionInfoFileName)
                        Else
                            blnSuccess = DetermineVersionDotNETDll(strInputFilePathFull, strOutputFolderPath, mVersionInfoFileName)
                        End If


                        If blnSuccess Then
                            ShowMessage(String.Empty, False)
                        Else
                            SetLocalErrorCode(eDLLVersionInspectorErrorCodes.UnspecifiedError)
                            ShowErrorMessage("Error")
                        End If

                    Catch ex As Exception
                        HandleException("Error calling InspectDLL", ex)
                    End Try
                End If
            End If
        Catch ex As Exception
            HandleException("Error in ProcessFile", ex)
        End Try

        Return blnSuccess

    End Function

    Private Function SaveVersionInfo(ByVal strDLLFilePath As String, ByVal strOutputFolderPath As String, ByVal strVersionInfoFileName As String, ByVal strToolVersionInfo As String) As Boolean
        Return SaveVersionInfo(strDLLFilePath, strOutputFolderPath, strVersionInfoFileName, strToolVersionInfo, String.Empty)
    End Function

    Private Function SaveVersionInfo(ByVal strDLLFilePath As String, ByVal strOutputFolderPath As String, ByVal strVersionInfoFileName As String, ByVal strToolVersionInfo As String, ByVal strErrorMessage As String) As Boolean

        Dim ioFileInfo As System.IO.FileInfo
        Dim lstInfo As Generic.List(Of String) = New Generic.List(Of String)

        Try

            ioFileInfo = New System.IO.FileInfo(strDLLFilePath)

            If ioFileInfo.Exists Then
                lstInfo.Add("FileName=" & ioFileInfo.Name)
                lstInfo.Add("Path=" & ioFileInfo.FullName)
            Else
                lstInfo.Add("FileName=" & System.IO.Path.GetFileName(strDLLFilePath))
                lstInfo.Add("Path=" & strDLLFilePath)
            End If

            lstInfo.Add("Version=" & strToolVersionInfo)

            If Not String.IsNullOrWhiteSpace(strErrorMessage) Then
                lstInfo.Add("Error=" & strErrorMessage)
            End If

            If mShowResultsAtConsole Then
                For Each item In lstInfo
                    Console.WriteLine(item)
                Next
            Else

                If String.IsNullOrWhiteSpace(strVersionInfoFileName) Then
                    ' Auto-define the output file name
                    strVersionInfoFileName = GetDefaultVersionInfoFileName(strDLLFilePath)
                End If

                ' Auto-define the output file name
                If String.IsNullOrWhiteSpace(strOutputFolderPath) Then
                    Dim ioAppFileInfo As System.IO.FileInfo = New System.IO.FileInfo(GetAppPath())
                    strOutputFolderPath = ioAppFileInfo.DirectoryName
                End If

                Dim strVersionInfoFilePath As String
                strVersionInfoFilePath = System.IO.Path.Combine(strOutputFolderPath, strVersionInfoFileName)

                Try
                    Dim eFileMode As IO.FileMode
                    If mAppendToVersionInfoFile Then
                        eFileMode = IO.FileMode.Append
                    Else
                        eFileMode = IO.FileMode.Create
                    End If

                    Using srOutFile As System.IO.StreamWriter = New System.IO.StreamWriter(New System.IO.FileStream(strVersionInfoFilePath, eFileMode, IO.FileAccess.Write, IO.FileShare.Read))
                        For Each item In lstInfo
                            srOutFile.WriteLine(item)
                        Next

                        If mAppendToVersionInfoFile Then
                            srOutFile.WriteLine()
                        End If

                    End Using

                Catch ex As Exception
                    ShowErrorMessage("Exception writing the version info to the output file at " & strVersionInfoFilePath & ": " & ex.Message)
                    Return False
                End Try

            End If


        Catch ex As Exception
            ShowErrorMessage("Exception in SaveVersionInfo: " & ex.Message)
            Return False
        End Try

        Return True

    End Function

    Private Sub SetLocalErrorCode(ByVal eNewErrorCode As eDLLVersionInspectorErrorCodes)
        SetLocalErrorCode(eNewErrorCode, False)
    End Sub

    Private Sub SetLocalErrorCode(ByVal eNewErrorCode As eDLLVersionInspectorErrorCodes, ByVal blnLeaveExistingErrorCodeUnchanged As Boolean)

        If blnLeaveExistingErrorCodeUnchanged AndAlso mLocalErrorCode <> eDLLVersionInspectorErrorCodes.NoError Then
            ' An error code is already defined; do not change it
        Else
            mLocalErrorCode = eNewErrorCode

            If eNewErrorCode = eDLLVersionInspectorErrorCodes.NoError Then
                If MyBase.ErrorCode = clsProcessFilesBaseClass.eProcessFilesErrorCodes.LocalizedError Then
                    MyBase.SetBaseClassErrorCode(clsProcessFilesBaseClass.eProcessFilesErrorCodes.NoError)
                End If
            Else
                MyBase.SetBaseClassErrorCode(clsProcessFilesBaseClass.eProcessFilesErrorCodes.LocalizedError)
            End If
        End If

    End Sub

End Class
