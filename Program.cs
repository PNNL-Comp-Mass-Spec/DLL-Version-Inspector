using System;
using System.Collections.Generic;
using System.IO;
using System.Threading;
using PRISM;

namespace DLLVersionInspector
{
    // This program inspects a .NET DLL or .Exe to determine the version.
    // This allows a 32-bit .NET application to call this program via the
    // command prompt to determine the version of a 64-bit DLL or Exe.
    //
    // -------------------------------------------------------------------------------
    // Written by Matthew Monroe for the Department of Energy (PNNL, Richland, WA)
    // Program started in April 2011
    //
    // E-mail: matthew.monroe@pnnl.gov or proteomics@pnnl.gov
    // Website: https://github.com/PNNL-Comp-Mass-Spec/ or https://panomics.pnnl.gov/ or https://www.pnnl.gov/integrative-omics
    // -------------------------------------------------------------------------------
    //

    static class Program
    {
        public const string PROGRAM_DATE = "June 20, 2024";

        private static string mInputFilePath;
        private static string mVersionInfoFilePath = string.Empty;
        private static bool mGenericDLL = false;

        private static string mOutputDirectoryPath;
        private static string mParameterFilePath;            // Not used by this program

        private static string mOutputDirectoryAlternatePath;              // Not used by this program
        private static bool mmRecreateDirectoryHierarchyInAlternatePath;  // Not used by this program

        private static bool mRecurseDirectories;
        private static int mMaxLevelsToRecurse;

        private static bool mShowResultsAtConsole = false;

        private static DateTime mLastProgressReportTime;
        private static int mLastProgressReportValue;

        /// <summary>
        /// Main method
        /// </summary>
        /// <returns>0 if no error, error code if an error</returns>
        public static int Main()
        {
            int returnCode;
            var commandLineParser = new clsParseCommandLine();
            bool proceed;

            mOutputDirectoryPath = string.Empty;
            mParameterFilePath = string.Empty;

            mOutputDirectoryAlternatePath = string.Empty;
            mmRecreateDirectoryHierarchyInAlternatePath = false;

            mRecurseDirectories = false;
            mMaxLevelsToRecurse = 0;

            mShowResultsAtConsole = false;

            try
            {
                proceed = false;

                if (commandLineParser.ParseCommandLine())
                {
                    if (SetOptionsUsingCommandLineParameters(commandLineParser))
                        proceed = true;
                }

                if (!proceed ||
                    commandLineParser.NeedToShowHelp ||
                    commandLineParser.ParameterCount + commandLineParser.NonSwitchParameterCount == 0 ||
                    string.IsNullOrWhiteSpace(mInputFilePath))
                {
                    ShowProgramHelp();
                    returnCode = -1;
                }
                else
                {
                    var versionInfoFileName = string.Empty;
                    if (!string.IsNullOrWhiteSpace(mVersionInfoFilePath))
                    {
                        var versionInfoFile = new FileInfo(mVersionInfoFilePath);

                        versionInfoFileName = versionInfoFile.Name;
                        mOutputDirectoryPath = versionInfoFile.Directory.FullName;

                        if (mRecurseDirectories && versionInfoFile.Exists)
                        {
                            // Delete the existing version info file because we will be appending to it
                            versionInfoFile.Delete();
                        }
                    }

                    // Note: the following settings will be overridden if mParameterFilePath points to a valid parameter file that has these settings defined
                    var dllVersionInspector = new DLLVersionInspector()
                    {
                        LogMessagesToFile = false,
                        AppendToVersionInfoFile = false,
                        GenericDLL = mGenericDLL,
                        ShowResultsAtConsole = mShowResultsAtConsole,
                        VersionInfoFileName = versionInfoFileName,
                        IgnoreErrorsWhenUsingWildcardMatching = true
                    };

                    dllVersionInspector.ProgressUpdate += DLLVersionInspector_ProgressUpdate;
                    dllVersionInspector.ProgressReset += DLLVersionInspector_ProgressReset;

                    if (mRecurseDirectories && !string.IsNullOrWhiteSpace(mOutputDirectoryPath))
                    {
                        dllVersionInspector.AppendToVersionInfoFile = true;
                    }

                    if (mRecurseDirectories)
                    {
                        dllVersionInspector.SkipConsoleWriteIfNoDebugListener = true;

                        if (dllVersionInspector.ProcessFilesAndRecurseDirectories(mInputFilePath, mOutputDirectoryPath, mOutputDirectoryAlternatePath,
                                                                                  mmRecreateDirectoryHierarchyInAlternatePath, mParameterFilePath,
                                                                                  mMaxLevelsToRecurse))
                        {
                            returnCode = 0;
                        }
                        else
                        {
                            returnCode = (int)dllVersionInspector.ErrorCode;
                        }
                    }
                    else if (dllVersionInspector.ProcessFilesWildcard(mInputFilePath, mOutputDirectoryPath, mParameterFilePath))
                    {
                        returnCode = 0;
                    }
                    else
                    {
                        returnCode = (int)dllVersionInspector.ErrorCode;
                        if (returnCode != 0)
                        {
                            Console.WriteLine("Error while processing: " + dllVersionInspector.GetErrorMessage());
                        }
                    }

                    DisplayProgressPercent(mLastProgressReportValue, true);
                }
            }
            catch (Exception ex)
            {
                ShowErrorMessage("Error occurred in modMain->Main", ex);
                returnCode = -1;
            }

            return returnCode;
        }

        private static void DisplayProgressPercent(int percentComplete, bool addCarriageReturn)
        {
            if (addCarriageReturn)
            {
                Console.WriteLine();
            }

            // Note: Display of % complete is disabled for this program because it is extremely simple
            return;

            // if (percentComplete > 100)
            // {
            //     percentComplete = 100;
            //     Console.Write("Processing: " + percentComplete.ToString() + "% ");
            //     if (addCarriageReturn)
            //     {
            //         Console.WriteLine();
            //     }
            // }
        }

        private static string GetAppVersion()
        {
            return AppUtils.GetAppVersion(PROGRAM_DATE);
        }

        private static string GetAppPath()
        {
            return AppUtils.GetAppPath();
        }

        private static bool SetOptionsUsingCommandLineParameters(clsParseCommandLine commandLineParser)
        {
            // Returns True if no problems; otherwise, returns false

            var value = string.Empty;
            var validParameters = new List<string>() { "I", "O", "G", "S", "C" };

            try
            {
                // Make sure no invalid parameters are present
                if (commandLineParser.InvalidParametersPresent(validParameters))
                {
                    return false;
                }

                // Query commandLineParser to see if various parameters are present

                if (commandLineParser.NonSwitchParameterCount > 0)
                {
                    mInputFilePath = commandLineParser.RetrieveNonSwitchParameter(0);
                }

                if (commandLineParser.RetrieveValueForParameter("I", out value))
                {
                    mInputFilePath = string.Copy(value);
                }

                if (commandLineParser.RetrieveValueForParameter("O", out value))
                {
                    mVersionInfoFilePath = string.Copy(value);
                }

                mGenericDLL = commandLineParser.IsParameterPresent("G");
                mShowResultsAtConsole = commandLineParser.IsParameterPresent("C");

                if (commandLineParser.RetrieveValueForParameter("S", out value))
                {
                    mRecurseDirectories = true;
                    if (!int.TryParse(value, out mMaxLevelsToRecurse))
                    {
                        mMaxLevelsToRecurse = 0;
                    }
                }

                return true;
            }
            catch (Exception ex)
            {
                ShowErrorMessage("Error parsing the command line parameters", ex);
                return false;
            }
        }

        private static void ShowErrorMessage(string message, Exception ex)
        {
            ConsoleMsgUtils.ShowError(message, ex);
        }

        private static void ShowProgramHelp()
        {
            try
            {
                Console.WriteLine(ConsoleMsgUtils.WrapParagraph(
                    "This program inspects a .NET DLL or .Exe to determine the version. " +
                    "This allows a 32-bit .NET application to call this program via the command prompt " +
                    "to determine the version of a 64-bit DLL or Exe. " +
                    "The program will create a VersionInfo file with the details about the DLL."));
                Console.WriteLine();
                Console.WriteLine(ConsoleMsgUtils.WrapParagraph(
                    "By default, uses the .NET framework to inspect the DLL or .Exe; " +
                    "use the /G switch if the DLL is a generic Windows DLL."));
                Console.WriteLine();
                Console.WriteLine(ConsoleMsgUtils.WrapParagraph(
                    "Alternatively, you can search for all occurrences of a given DLL in a directory and its subdirectories (use switch /S). " +
                    "In this mode, the DLL version will be displayed at the console."));
                Console.WriteLine();
                Console.WriteLine("Program syntax: " + Path.GetFileName(GetAppPath()));
                Console.WriteLine("  FilePath [/O:VersionInfoFilePath] [/G] [/C] [/S]");
                Console.WriteLine();
                Console.WriteLine("FilePath is the path to the .NET DLL or .NET Exe to inspect");
                Console.WriteLine();
                Console.WriteLine(ConsoleMsgUtils.WrapParagraph(
                    "Use /O:VersionInfoFilePath to specify the path to the file to which this program should write the version info. " +
                    "If using /S, use /O to specify the filename that will be created in the directory for which each DLL is found"));
                Console.WriteLine();
                Console.WriteLine(ConsoleMsgUtils.WrapParagraph("Use /G to indicate that the DLL is a generic Windows DLL, and not a .NET DLL"));
                Console.WriteLine();
                Console.WriteLine(ConsoleMsgUtils.WrapParagraph("Use /C to display the version info in the console output instead of creating a VersionInfo file"));
                Console.WriteLine();
                Console.WriteLine(ConsoleMsgUtils.WrapParagraph("Use /S to search for all instances of the DLL in a directory and its subdirectories (wildcards are allowed)"));
                Console.WriteLine();

                Console.WriteLine("Program written by Matthew Monroe for the Department of Energy (PNNL, Richland, WA) in 2011");
                Console.WriteLine("Version: " + GetAppVersion());
                Console.WriteLine();

                Console.WriteLine("E-mail: matthew.monroe@pnnl.gov or proteomics@pnnl.gov");
                Console.WriteLine("Website: https://github.com/PNNL-Comp-Mass-Spec/ or https://panomics.pnnl.gov/ or https://www.pnnl.gov/integrative-omics");
                Console.WriteLine();

                // Delay for 750 msec in case the user double clicked this file from within Windows Explorer (or started the program via a shortcut)
                Thread.Sleep(750);
            }
            catch (Exception ex)
            {
                ShowErrorMessage("Error displaying the program syntax", ex);
            }
        }

        private static void DLLVersionInspector_ProgressUpdate(string taskDescription, float percentComplete)
        {
            const int PERCENT_REPORT_INTERVAL = 25;
            const int PROGRESS_DOT_INTERVAL_MSEC = 250;

            if (percentComplete >= mLastProgressReportValue)
            {
                if (mLastProgressReportValue > 0)
                {
                    Console.WriteLine();
                }
                DisplayProgressPercent(mLastProgressReportValue, false);
                mLastProgressReportValue += PERCENT_REPORT_INTERVAL;
                mLastProgressReportTime = DateTime.UtcNow;
            }
            else if (DateTime.UtcNow.Subtract(mLastProgressReportTime).TotalMilliseconds > PROGRESS_DOT_INTERVAL_MSEC)
            {
                mLastProgressReportTime = DateTime.UtcNow;
                Console.Write(".");
            }
        }

        private static void DLLVersionInspector_ProgressReset()
        {
            mLastProgressReportTime = DateTime.UtcNow;
            mLastProgressReportValue = 0;
        }
    }
}