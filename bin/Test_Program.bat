@echo off

cls

rem This should report version 3.0.5654.32750
call RunSingleTest.bat x86 32bit_Dll_Examples UIMFLibrary "               3.0.5654.32750"

rem This should lead to error "Exception determining Assembly info for UIMFLibrary.dll ... An attempt was made to load a program with an incorrect format."
call RunSingleTest.bat x86 64bit_Dll_Examples UIMFLibrary "Exception determining Assembly info"

rem This should report version 3.0.5654.24060
call RunSingleTest.bat x86 AnyCPU_DLL_Examples UIMFLibrary "               3.0.5654.24060"

pause
cls

rem This should lead to error "Exception determining Assembly info for UIMFLibrary.dll ... An attempt was made to load a program with an incorrect format."
call RunSingleTest.bat x64 32bit_Dll_Examples UIMFLibrary "Exception determining Assembly info"

rem This should report version 3.0.5654.32746
call RunSingleTest.bat x64 64bit_Dll_Examples UIMFLibrary "               3.0.5654.32746"

rem This should report version 3.0.5654.24060
call RunSingleTest.bat x64 AnyCPU_DLL_Examples UIMFLibrary "               3.0.5654.24060"

rem This should report version 2.0.6068.26406
call RunSingleTest.bat x86 AnyCPU_DLL_Examples ThermoRawFileReader "                       2.0.6068.26406"
