@echo off
Debug\net48\DLLVersionInspector.exe ..\LCMSFeatureFinder\LCMSFeatureFinder.exe /o:.\LCMSFeatureFinder_VersionInfo.txt
Debug\net48\DLLVersionInspector.exe ..\LCMSFeatureFinder\FeatureFinder.dll /o:.\FeatureFinderDLL_VersionInfo.txt
pause
