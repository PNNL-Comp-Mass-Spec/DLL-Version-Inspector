using System;
using System.Collections.Generic;
using System.Diagnostics;

using System.IO;
using System.Reflection;
using PRISM;
using PRISM.FileProcessor;

namespace DLLVersionInspector
{
    /// <summary>
    /// This class will determine the version of a .NET DLL or a generic Windows DLL
    /// Written by Matthew Monroe for the Department of Energy (PNNL, Richland, WA)
    /// Started June 14, 2013
    /// </summary>
    public class DLLVersionInspector : ProcessFilesBase
    {
        public DLLVersionInspector()
        {
            mFileDate = "June 20, 2024";
            InitializeLocalVariables();
        }

        public const string XML_SECTION_OPTIONS = "DLLVersionInspectorOptions";

        // Error codes specialized for this class
        public enum eDLLVersionInspectorErrorCodes
        {
            NoError = 0,
            ErrorReadingInputFile = 1,
            UnspecifiedError = -1
        }

        protected eDLLVersionInspectorErrorCodes mLocalErrorCode;

        public bool AppendToVersionInfoFile { get; set; }

        public bool GenericDLL { get; set; }

        public bool ShowResultsAtConsole { get; set; } = false;

        public string VersionInfoFileName { get; set; }

        /// <summary>
        /// Determines the version of a .NET DLL
        /// </summary>
        /// <param name="dllFilePath"></param>
        /// <param name="outputDirectoryPath"></param>
        /// <returns>True if success, false if an error</returns>
        /// <remarks></remarks>
        protected bool DetermineVersionDotNETDll(string dllFilePath, string outputDirectoryPath)
        {

            string toolVersionInfo = string.Empty;

            try
            {
                var dllInfo = new FileInfo(dllFilePath);

                if (!dllInfo.Exists)
                {
                    string errorMessage = "Error: File not found: " + dllFilePath;
                    ShowErrorMessage(errorMessage);
                    SaveVersionInfo(dllFilePath, outputDirectoryPath, toolVersionInfo, errorMessage);
                    return false;
                }
                else
                {

                    AssemblyName oAssemblyName;
                    oAssemblyName = Assembly.LoadFrom(dllInfo.FullName).GetName();

                    toolVersionInfo = oAssemblyName.Name + ", Version=" + oAssemblyName.Version.ToString();

                    bool success = SaveVersionInfo(dllFilePath, outputDirectoryPath, toolVersionInfo);
                    return success;
                }
            }

            catch (Exception ex)
            {
                // If you get an exception regarding .NET 4.0 not being able to read a .NET 1.0 runtime, add these lines to the end of file AnalysisManagerProg.exe.config
                // <startup useLegacyV2RuntimeActivationPolicy="true">
                // <supportedRuntime version="v4.0" />
                // </startup>
                string errorMessage = "Exception determining Assembly info for " + Path.GetFileName(dllFilePath) + ": " + ex.Message;
                ShowErrorMessage(errorMessage);
                SaveVersionInfo(dllFilePath, outputDirectoryPath, toolVersionInfo, errorMessage);
                return false;
            }
        }

        /// <summary>
        /// Determines the version of a generic Windows DLL
        /// </summary>
        /// <param name="dllFilePath"></param>
        /// <param name="outputDirectoryPath"></param>
        /// <returns>True if success, false if an error</returns>
        /// <remarks></remarks>
        protected bool DetermineVersionGenericDLL(string dllFilePath, string outputDirectoryPath)
        {
            string toolVersionInfo = string.Empty;

            try
            {
                var dllInfo = new FileInfo(dllFilePath);

                if (!dllInfo.Exists)
                {
                    string errorMessage = "Error: File not found: " + dllFilePath;
                    ShowErrorMessage(errorMessage);
                    SaveVersionInfo(dllFilePath, outputDirectoryPath, toolVersionInfo, errorMessage);
                    return false;
                }
                else
                {
                    var oFileVersionInfo = FileVersionInfo.GetVersionInfo(dllFilePath);

                    string fileName;
                    string dllVersion;

                    fileName = oFileVersionInfo.FileDescription;
                    if (string.IsNullOrEmpty(fileName))
                    {
                        fileName = oFileVersionInfo.InternalName;
                    }

                    if (string.IsNullOrEmpty(fileName))
                    {
                        fileName = oFileVersionInfo.FileName;
                    }

                    if (string.IsNullOrEmpty(fileName))
                    {
                        fileName = dllInfo.Name;
                    }

                    dllVersion = oFileVersionInfo.FileVersion;
                    if (string.IsNullOrEmpty(dllVersion))
                    {
                        dllVersion = oFileVersionInfo.ProductVersion;
                    }

                    if (string.IsNullOrEmpty(dllVersion))
                    {
                        dllVersion = "??";
                    }

                    toolVersionInfo = fileName + ", Version=" + dllVersion;

                    bool success = SaveVersionInfo(dllFilePath, outputDirectoryPath, toolVersionInfo);
                    return success;
                }
            }
            catch (Exception ex)
            {
                string errorMessage = "Exception determining Version info for " + Path.GetFileName(dllFilePath) + ": " + ex.Message;
                ShowErrorMessage(errorMessage);
                SaveVersionInfo(dllFilePath, outputDirectoryPath, toolVersionInfo, errorMessage);
                return false;
            }
        }

        public override IList<string> GetDefaultExtensionsToParse()
        {
            var extensionsToParse = new List<string>() { ".dll", ".exe" };

            return extensionsToParse;
        }

        public static string GetDefaultVersionInfoFileName(string dllFileNameOrPath)
        {
            return Path.GetFileNameWithoutExtension(dllFileNameOrPath) + "_VersionInfo.txt";
        }

        public override string GetErrorMessage()
        {
            // Returns "" if no error

            string errorMessage;

            if (ErrorCode == ProcessFilesErrorCodes.LocalizedError | ErrorCode == ProcessFilesErrorCodes.NoError)
            {
                switch (mLocalErrorCode)
                {
                    case eDLLVersionInspectorErrorCodes.NoError:
                        errorMessage = "";
                        break;

                    case eDLLVersionInspectorErrorCodes.ErrorReadingInputFile:
                        errorMessage = "Error reading input file";
                        break;

                    case eDLLVersionInspectorErrorCodes.UnspecifiedError:
                        errorMessage = "Unspecified localized error";
                        break;

                    default:
                        // This shouldn't happen
                        errorMessage = "Unknown error state";
                        break;
                }
            }
            else
            {
                errorMessage = GetBaseClassErrorMessage();
            }

            return errorMessage;
        }

        private void InitializeLocalVariables()
        {
            GenericDLL = false;
            ShowResultsAtConsole = false;
            VersionInfoFileName = string.Empty;
            mLocalErrorCode = eDLLVersionInspectorErrorCodes.NoError;
        }

        public bool LoadParameterFileSettings(string parameterFilePath)
        {
            var settingsFile = new XmlSettingsFileAccessor();

            try
            {
                if (parameterFilePath is null || parameterFilePath.Length == 0)
                {
                    // No parameter file specified; nothing to load
                    return true;
                }

                if (!File.Exists(parameterFilePath))
                {
                    // See if parameterFilePath points to a file in the same directory as the application
                    parameterFilePath = Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), Path.GetFileName(parameterFilePath));
                    if (!File.Exists(parameterFilePath))
                    {
                        SetBaseClassErrorCode(ProcessFilesErrorCodes.ParameterFileNotFound);
                        return false;
                    }
                }

                if (settingsFile.LoadSettings(parameterFilePath))
                {
                    if (!settingsFile.SectionPresent(XML_SECTION_OPTIONS))
                    {
                        ShowErrorMessage("The node '<section name=\"" + XML_SECTION_OPTIONS + "\"> was not found in the parameter file: " + parameterFilePath);
                        SetBaseClassErrorCode(ProcessFilesErrorCodes.InvalidParameterFile);
                        return false;
                    }
                    else
                    {
                        GenericDLL = settingsFile.GetParam(XML_SECTION_OPTIONS, "GenericDLL", GenericDLL);
                    }
                }
            }

            catch (Exception ex)
            {
                HandleException("Error in LoadParameterFileSettings", ex);
                return false;
            }

            return true;
        }

        /// <summary>
        /// Main processing function -- Calls DetermineVersionGenericDLL or DetermineVersionDotNETDll
        /// </summary>
        /// <param name="inputFilePath"></param>
        /// <param name="outputDirectoryPath"></param>
        /// <param name="parameterFilePath"></param>
        /// <param name="resetErrorCode"></param>
        /// <returns></returns>
        /// <remarks></remarks>
        public override bool ProcessFile(string inputFilePath, string outputDirectoryPath, string parameterFilePath, bool resetErrorCode)
        {
            // Returns True if success, False if failure

            FileInfo inputFile;
            string inputFilePathFull;

            var success = default(bool);

            if (resetErrorCode)
            {
                SetLocalErrorCode(eDLLVersionInspectorErrorCodes.NoError);
            }

            if (!LoadParameterFileSettings(parameterFilePath))
            {
                ShowErrorMessage("Parameter file load error: " + parameterFilePath);

                if (ErrorCode == ProcessFilesErrorCodes.NoError)
                {
                    SetBaseClassErrorCode(ProcessFilesErrorCodes.InvalidParameterFile);
                }
                return false;
            }

            try
            {
                if (inputFilePath is null || inputFilePath.Length == 0)
                {
                    ShowMessage("Input file name is empty");
                    SetBaseClassErrorCode(ProcessFilesErrorCodes.InvalidInputFilePath);
                    return false;
                }

                Console.WriteLine();
                Console.WriteLine("Parsing " + Path.GetFileName(inputFilePath));

                // Note that CleanupFilePaths() will update mOutputDirectoryPath, which is used by LogMessage()
                if (!CleanupFilePaths(ref inputFilePath, ref outputDirectoryPath))
                {
                    SetBaseClassErrorCode(ProcessFilesErrorCodes.FilePathError);
                    return false;
                }

                ResetProgress();

                try
                {
                    // Obtain the full path to the input file
                    inputFile = new FileInfo(inputFilePath);
                    inputFilePathFull = inputFile.FullName;

                    if (GenericDLL)
                    {
                        success = DetermineVersionGenericDLL(inputFilePathFull, outputDirectoryPath);
                    }
                    else
                    {
                        success = DetermineVersionDotNETDll(inputFilePathFull, outputDirectoryPath);
                    }


                    if (success)
                    {
                        ShowMessage(string.Empty, false);
                    }
                    else
                    {
                        SetLocalErrorCode(eDLLVersionInspectorErrorCodes.UnspecifiedError);
                        ShowErrorMessage("Error");
                    }
                }
                catch (Exception ex)
                {
                    HandleException("Error calling DetermineVersionGenericDLL or DetermineVersionDotNETDll", ex);
                }

                return success;
            }
            catch (Exception ex)
            {
                HandleException("Error in ProcessFile", ex);
                return false;
            }
        }

        private bool SaveVersionInfo(string dllFilePath, string outputDirectoryPath, string toolVersionInfo)
        {
            return SaveVersionInfo(dllFilePath, outputDirectoryPath, toolVersionInfo, string.Empty);
        }

        private bool SaveVersionInfo(string dllFilePath, string outputDirectoryPath, string toolVersionInfo, string errorMessage)
        {
            // This list tracks the DLL name, path, and version
            var versionInfo = new List<string>();

            try
            {
                var dllFile = new FileInfo(dllFilePath);

                if (dllFile.Exists)
                {
                    versionInfo.Add("FileName=" + dllFile.Name);
                    versionInfo.Add("Path=" + dllFile.FullName);
                }
                else
                {
                    versionInfo.Add("FileName=" + Path.GetFileName(dllFilePath));
                    versionInfo.Add("Path=" + dllFilePath);
                }

                versionInfo.Add("Version=" + toolVersionInfo);

                if (!string.IsNullOrWhiteSpace(errorMessage))
                {
                    versionInfo.Add("Error=" + errorMessage);
                }

                if (ShowResultsAtConsole)
                {
                    foreach (var item in versionInfo)
                        Console.WriteLine(item);
                }
                else
                {
                    if (string.IsNullOrWhiteSpace(VersionInfoFileName))
                    {
                        // Auto-define the output file name
                        VersionInfoFileName = GetDefaultVersionInfoFileName(dllFilePath);
                    }

                    // Auto-define the output file name
                    if (string.IsNullOrWhiteSpace(outputDirectoryPath))
                    {
                        var appFileInfo = new FileInfo(AppUtils.GetAppPath());
                        outputDirectoryPath = appFileInfo.DirectoryName;
                    }

                    string versionInfoFilePath = Path.Combine(outputDirectoryPath, VersionInfoFileName);

                    try
                    {
                        FileMode eFileMode;
                        if (AppendToVersionInfoFile)
                        {
                            eFileMode = FileMode.Append;
                        }
                        else
                        {
                            eFileMode = FileMode.Create;
                        }

                        using (var writer = new StreamWriter(new FileStream(versionInfoFilePath, eFileMode, FileAccess.Write, FileShare.Read)))
                        {
                            foreach (var item in versionInfo)
                                writer.WriteLine(item);

                            if (AppendToVersionInfoFile)
                            {
                                writer.WriteLine();
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        ShowErrorMessage("Exception writing the version info to the output file at " + versionInfoFilePath + ": " + ex.Message);
                        return false;
                    }
                }
            }
            catch (Exception ex)
            {
                ShowErrorMessage("Exception in SaveVersionInfo: " + ex.Message);
                return false;
            }

            return true;
        }

        private void SetLocalErrorCode(eDLLVersionInspectorErrorCodes eNewErrorCode)
        {
            SetLocalErrorCode(eNewErrorCode, false);
        }

        private void SetLocalErrorCode(eDLLVersionInspectorErrorCodes eNewErrorCode, bool leaveExistingErrorCodeUnchanged)
        {
            if (leaveExistingErrorCodeUnchanged && mLocalErrorCode != eDLLVersionInspectorErrorCodes.NoError)
            {
                // An error code is already defined; do not change it
            }
            else
            {
                mLocalErrorCode = eNewErrorCode;

                if (eNewErrorCode == eDLLVersionInspectorErrorCodes.NoError)
                {
                    if (ErrorCode == ProcessFilesErrorCodes.LocalizedError)
                    {
                        SetBaseClassErrorCode(ProcessFilesErrorCodes.NoError);
                    }
                }
                else
                {
                    SetBaseClassErrorCode(ProcessFilesErrorCodes.LocalizedError);
                }
            }
        }
    }
}