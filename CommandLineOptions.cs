using PRISM;

namespace DLLVersionInspector
{
    internal class CommandLineOptions
    {
        [Option("I", ArgPosition = 1, Required = true, HelpText = "Path to the .NET DLL or .NET Exe to inspect", HelpShowsDefault = false)]
        public string InputFilePath { get; set; }

        [Option("O", HelpText = "(optional) Path to the file where version info should be written", HelpShowsDefault = false)]
        public string VersionInfoFilePath { get; set; }

        [Option("G", HelpText = "Specify to use generic Windows DLL checks instead of .NET Framework DLL checks")]
        public bool GenericDll { get; set; }

        [Option("C", HelpText = "Display the version info in the console output instead of creating a VersionInfo file")]
        public bool ShowResultsAtConsole { get; set; }

        [Option("S", ArgExistsProperty = nameof(RecurseDirectories), HelpText = "Search for all instances of the DLL in a directory and its subdirectories (wildcards are allowed). " +
            "Use /O with this to specify the filename that will be created in the directory for which each DLL is found", Min = 0)]
        public int MaxLevelsToRecurse { get; set; }
        public bool RecurseDirectories { get; set; }

        public CommandLineOptions()
        {
            InputFilePath = string.Empty;
            VersionInfoFilePath = string.Empty;
            GenericDll = false;
            ShowResultsAtConsole = false;
            MaxLevelsToRecurse = 0;
            RecurseDirectories = false;
        }
    }
}
