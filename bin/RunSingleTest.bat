@echo off
rem Mode (x86 or x64): %1
rem DLL folder: %2
rem DLL name: %3

if exist %2\%3_VersionInfo.txt  del %2\%3_VersionInfo.txt

@echo on
%1\DLLVersionInspector_%1.exe %2\%3.dll
@echo off

echo -------------------------------------------
echo Results for %1 examining %2
echo ----
if exist %2\%3_VersionInfo.txt (type %2\%3_VersionInfo.txt) else (echo No results)
echo -------------------------------------------
