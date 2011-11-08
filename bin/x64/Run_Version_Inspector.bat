@echo off
DLLVersionInspector.exe ..\LCMSFeatureFinder\LCMSFeatureFinder.exe /o:.\LCMSFeatureFinder_VersionInfo.txt
DLLVersionInspector.exe ..\LCMSFeatureFinder\FeatureFinder.dll /o:.\FeatureFinderDLL_VersionInfo.txt
pause
