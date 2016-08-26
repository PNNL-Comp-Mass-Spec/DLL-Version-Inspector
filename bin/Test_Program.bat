@echo off

cls

call RunSingleTest.bat x86 32bit_Dll_Examples UIMFLibrary
call RunSingleTest.bat x86 64bit_Dll_Examples UIMFLibrary
call RunSingleTest.bat x86 AnyCPU_DLL_Examples UIMFLibrary
pause
cls

call RunSingleTest.bat x64 32bit_Dll_Examples UIMFLibrary
call RunSingleTest.bat x64 64bit_Dll_Examples UIMFLibrary
call RunSingleTest.bat x64 AnyCPU_DLL_Examples UIMFLibrary

call RunSingleTest.bat x86 AnyCPU_DLL_Examples ThermoRawFileReader
