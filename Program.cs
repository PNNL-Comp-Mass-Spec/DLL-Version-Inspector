using System;
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

        private static DateTime mLastProgressReportTime;
        private static int mLastProgressReportValue;

        /// <summary>
        /// Main method
        /// </summary>
        /// <returns>0 if no error, error code if an error</returns>
        public static int Main(string[] args)
        {
            int returnCode;

            try
            {
                var parser = new CommandLineParser<CommandLineOptions>
                {
                    ProgramInfo = "This program inspects a .NET DLL or .Exe to determine the version. " +
                                  "This allows a 32-bit .NET application to call this program via the command prompt " +
                                  "to determine the version of a 64-bit DLL or Exe. " +
                                  "The program will create a VersionInfo file with the details about the DLL." +
                                  Environment.NewLine + Environment.NewLine +
                                  "By default, uses the .NET framework to inspect the DLL or .Exe; " +
                                  "use the /G switch if the DLL is a generic Windows DLL." +
                                  Environment.NewLine + Environment.NewLine +
                                  "Alternatively, you can search for all occurrences of a given DLL in a directory and its subdirectories (use switch /S). " +
                                  "In this mode, the DLL version will be displayed at the console." +
                                  Environment.NewLine + Environment.NewLine +
                                  "Version: " + AppUtils.GetAppVersion(PROGRAM_DATE),

                    ContactInfo = "Program written by Matthew Monroe for the Department of Energy (PNNL, Richland, WA) in 2011" +
                                  Environment.NewLine + Environment.NewLine +
                                  "E-mail: matthew.monroe@pnnl.gov or proteomics@pnnl.gov" + Environment.NewLine +
                                  "Website: https://github.com/PNNL-Comp-Mass-Spec/ or https://www.pnnl.gov/integrative-omics",

                    UsageExamples = { Path.GetFileName(AppUtils.GetAppPath()) + " FilePath [/O:VersionInfoFilePath] [/G] [/C] [/S]" },
                    DisableParameterFileSupport = true
                };

                var results = parser.ParseArgs(args, false);

                if (!results.Success)
                {
                    parser.PrintHelp(10, 64);
                    Console.WriteLine();

                    // Delay for 1500 msec in case the user double-clicked this file from within Windows Explorer (or started the program via a shortcut)
                    Thread.Sleep(1500);
                    return -1;
                }

                var options = results.ParsedResults;

                var versionInfoFileName = string.Empty;
                var outputDirectoryPath = string.Empty;
                if (!string.IsNullOrWhiteSpace(options.VersionInfoFilePath))
                {
                    var versionInfoFile = new FileInfo(options.VersionInfoFilePath);

                    versionInfoFileName = versionInfoFile.Name;
                    outputDirectoryPath = versionInfoFile.Directory.FullName;

                    if (options.RecurseDirectories && versionInfoFile.Exists)
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
                    GenericDLL = options.GenericDll,
                    ShowResultsAtConsole = options.ShowResultsAtConsole,
                    VersionInfoFileName = versionInfoFileName,
                    IgnoreErrorsWhenUsingWildcardMatching = true
                };

                dllVersionInspector.ProgressUpdate += DLLVersionInspector_ProgressUpdate;
                dllVersionInspector.ProgressReset += DLLVersionInspector_ProgressReset;

                if (options.RecurseDirectories && !string.IsNullOrWhiteSpace(outputDirectoryPath))
                {
                    dllVersionInspector.AppendToVersionInfoFile = true;
                }

                if (options.RecurseDirectories)
                {
                    dllVersionInspector.SkipConsoleWriteIfNoDebugListener = true;

                    if (dllVersionInspector.ProcessFilesAndRecurseDirectories(options.InputFilePath, outputDirectoryPath, string.Empty,
                            false, string.Empty,
                            options.MaxLevelsToRecurse))
                    {
                        returnCode = 0;
                    }
                    else
                    {
                        returnCode = (int)dllVersionInspector.ErrorCode;
                    }
                }
                else if (dllVersionInspector.ProcessFilesWildcard(options.InputFilePath, outputDirectoryPath, string.Empty))
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
            catch (Exception ex)
            {
                ShowErrorMessage("Error occurred in Program->Main", ex);
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

        private static void ShowErrorMessage(string message, Exception ex)
        {
            ConsoleMsgUtils.ShowError(message, ex);
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