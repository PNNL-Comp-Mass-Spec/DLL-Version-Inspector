# DLL Version Inspector

This program inspects a .NET DLL or .Exe to determine the version.
This allows a 32-bit .NET application to call this program via the 
command prompt to determine the version of a 64-bit DLL or Exe. 
The program will create a VersionInfo file with the details about the DLL.

By default, uses the .NET framework to inspect the DLL or .Exe; 
use the /G switch if the DLL is a generic Windows DLL.

Alternatively, you can search for all occurrences of a given DLL 
in a directory and its subdirectories (use switch /S). In this mode, 
the DLL version will be displayed at the console.

## Console Switches

The DLL Version Inspector is a command line application.  Syntax:

```
DLLVersionInspector.exe
  FilePath [/O:VersionInfoFilePath] [/G] [/C] [/S]
```

FilePath is the path to the DLL or Exe to inspect

Use /O:VersionInfoFilePath to specify the path to the file to which this program
should write the version info. If using /S, then use /O to specify the filename
that will be created in the directory for which each DLL is found

Use /G to indicate that the DLL is a generic Windows DLL, and not a .NET DLL

Use /C to display the version info in the console output instead of creating a
VersionInfo file

Use /S to search for all instances of the DLL in a directory and its
subdirectories (wildcards are allowed)

## Contacts

Written by Matthew Monroe for the Department of Energy (PNNL, Richland, WA) \
E-mail: matthew.monroe@pnnl.gov or proteomics@pnnl.gov\
Website: https://omics.pnl.gov/ or https://panomics.pnnl.gov/ or https://github.com/PNNL-Comp-Mass-Spec/

## License

The DLL Version Inspector is licensed under the 2-Clause BSD License; 
you may not use this file except in compliance with the License.
You may obtain a copy of the License at https://opensource.org/licenses/BSD-2-Clause

Copyright 2018 Battelle Memorial Institute
