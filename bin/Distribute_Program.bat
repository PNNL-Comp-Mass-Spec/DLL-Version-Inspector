@echo off
rem copy x86\Debug\net48\DLLVersionInspector.exe         x86\DLLVersionInspector_x86.exe /Y
rem copy x86\Debug\net48\DLLVersionInspector.exe.config  x86\DLLVersionInspector_x86.exe.config /Y

@echo on
if not exist C:\DMS_Programs\DLLVersionInspector mkdir   C:\DMS_Programs\DLLVersionInspector

xcopy x86\DLLVersionInspector_x86.exe        C:\DMS_Programs\DLLVersionInspector\ /Y /D
xcopy x86\DLLVersionInspector_x86.exe.config C:\DMS_Programs\DLLVersionInspector\ /Y /D
xcopy x86\PRISM.dll                          C:\DMS_Programs\DLLVersionInspector\ /Y /D

xcopy x86\DLLVersionInspector_x86.exe        \\PNL\projects\OmicsSW\DMS_Programs\AnalysisToolManagerDistribution\DLLVersionInspector\ /Y /D
xcopy x86\DLLVersionInspector_x86.exe.config \\PNL\projects\OmicsSW\DMS_Programs\AnalysisToolManagerDistribution\DLLVersionInspector\ /Y /D
xcopy x86\PRISM.dll                          \\PNL\projects\OmicsSW\DMS_Programs\AnalysisToolManagerDistribution\DLLVersionInspector\ /Y /D

@echo off
rem xcopy x86\DLLVersionInspector_x86.exe        F:\Documents\Projects\DataMining\DMS_Managers\Analysis_Manager\AM_Common\ /Y /D
rem xcopy x86\DLLVersionInspector_x86.exe        F:\Documents\Projects\DataMining\DMS_Managers\Analysis_Manager\AM_Program\bin\ /Y /D
rem xcopy x86\DLLVersionInspector_x86.exe.config F:\Documents\Projects\DataMining\DMS_Managers\Analysis_Manager\AM_Program\bin\ /Y /D
rem xcopy x86\DLLVersionInspector_x86.exe        F:\Documents\Projects\DataMining\DMS_Managers\Analysis_Manager\AM_Program\bin\Release\ /Y /D
rem xcopy x86\DLLVersionInspector_x86.exe        F:\Documents\Projects\DataMining\DMS_Managers\Capture_Task_Manager\CaptureTaskManager\bin\Debug\net8.0-windows\ /Y /D
rem xcopy x86\DLLVersionInspector_x86.exe.config F:\Documents\Projects\DataMining\DMS_Managers\Capture_Task_Manager\CaptureTaskManager\bin\Debug\net8.0-windows\ /Y /D
rem xcopy x86\DLLVersionInspector_x86.exe        F:\Documents\Projects\DataMining\DMS_Managers\Capture_Task_Manager\CaptureTaskManager\bin\Debug\net48\ /Y /D
rem xcopy x86\DLLVersionInspector_x86.exe.config F:\Documents\Projects\DataMining\DMS_Managers\Capture_Task_Manager\CaptureTaskManager\bin\Debug\net48\ /Y /D
rem xcopy x86\DLLVersionInspector_x86.exe        F:\Documents\Projects\DataMining\DMS_Managers\Capture_Task_Manager\RefLib\ /Y /D
rem xcopy x86\DLLVersionInspector_x86.exe        F:\Documents\Projects\DataMining\DMS_Managers\DMS_Update_Manager\DMSUpdateManagerConsole\bin\AnalysisToolManagerDistribution\ /Y /D
rem xcopy x86\DLLVersionInspector_x86.exe.config F:\Documents\Projects\DataMining\DMS_Managers\DMS_Update_Manager\DMSUpdateManagerConsole\bin\AnalysisToolManagerDistribution\ /Y /D
rem xcopy x86\DLLVersionInspector_x86.exe        F:\Documents\Projects\DataMining\DMS_Managers\DMS_Update_Manager\DMSUpdateManagerConsole\bin\DMS_Programs\ /Y /D
rem xcopy x86\DLLVersionInspector_x86.exe.config F:\Documents\Projects\DataMining\DMS_Managers\DMS_Update_Manager\DMSUpdateManagerConsole\bin\DMS_Programs\ /Y /D
rem xcopy x86\DLLVersionInspector_x86.exe        F:\Documents\Projects\DataMining\DMS_Managers\DMS_Update_Manager\DMSUpdateManagerConsole\bin\DMS_Programs\AnalysisToolManager1\ /Y /D
rem xcopy x86\DLLVersionInspector_x86.exe.config F:\Documents\Projects\DataMining\DMS_Managers\DMS_Update_Manager\DMSUpdateManagerConsole\bin\DMS_Programs\AnalysisToolManager1\ /Y /D

@echo off
rem copy x64\Debug\net48\DLLVersionInspector.exe         x64\DLLVersionInspector_x64.exe /Y
rem copy x64\Debug\net48\DLLVersionInspector.exe.config  x64\DLLVersionInspector_x64.exe.config /Y

@echo on
xcopy x64\DLLVersionInspector_x64.exe        C:\DMS_Programs\DLLVersionInspector\ /Y /D
xcopy x64\DLLVersionInspector_x64.exe.config C:\DMS_Programs\DLLVersionInspector\ /Y /D
xcopy x64\DLLVersionInspector_x64.exe        \\PNL\projects\OmicsSW\DMS_Programs\AnalysisToolManagerDistribution\DLLVersionInspector\ /Y /D
xcopy x64\DLLVersionInspector_x64.exe.config \\PNL\projects\OmicsSW\DMS_Programs\AnalysisToolManagerDistribution\DLLVersionInspector\ /Y /D

@echo off
rem xcopy x64\DLLVersionInspector_x64.exe        F:\Documents\Projects\DataMining\DMS_Managers\Analysis_Manager\AM_Common\ /Y /D
rem xcopy x64\DLLVersionInspector_x64.exe        F:\Documents\Projects\DataMining\DMS_Managers\Analysis_Manager\AM_Program\bin\ /Y /D
rem xcopy x64\DLLVersionInspector_x64.exe.config F:\Documents\Projects\DataMining\DMS_Managers\Analysis_Manager\AM_Program\bin\ /Y /D
rem xcopy x64\DLLVersionInspector_x64.exe        F:\Documents\Projects\DataMining\DMS_Managers\Analysis_Manager\AM_Program\bin\Release\ /Y /D
rem xcopy x64\DLLVersionInspector_x64.exe        F:\Documents\Projects\DataMining\DMS_Managers\Capture_Task_Manager\CaptureTaskManager\bin\Debug\net8.0-windows\ /Y /D
rem xcopy x64\DLLVersionInspector_x64.exe.config F:\Documents\Projects\DataMining\DMS_Managers\Capture_Task_Manager\CaptureTaskManager\bin\Debug\net8.0-windows\ /Y /D
rem xcopy x64\DLLVersionInspector_x64.exe        F:\Documents\Projects\DataMining\DMS_Managers\Capture_Task_Manager\CaptureTaskManager\bin\Debug\net48\ /Y /D
rem xcopy x64\DLLVersionInspector_x64.exe.config F:\Documents\Projects\DataMining\DMS_Managers\Capture_Task_Manager\CaptureTaskManager\bin\Debug\net48\ /Y /D
rem xcopy x64\DLLVersionInspector_x64.exe        F:\Documents\Projects\DataMining\DMS_Managers\Capture_Task_Manager\RefLib\ /Y /D
rem xcopy x64\DLLVersionInspector_x64.exe        F:\Documents\Projects\DataMining\DMS_Managers\DMS_Update_Manager\DMSUpdateManagerConsole\bin\AnalysisToolManagerDistribution\ /Y /D
rem xcopy x64\DLLVersionInspector_x64.exe.config F:\Documents\Projects\DataMining\DMS_Managers\DMS_Update_Manager\DMSUpdateManagerConsole\bin\AnalysisToolManagerDistribution\ /Y /D
rem xcopy x64\DLLVersionInspector_x64.exe        F:\Documents\Projects\DataMining\DMS_Managers\DMS_Update_Manager\DMSUpdateManagerConsole\bin\DMS_Programs\ /Y /D
rem xcopy x64\DLLVersionInspector_x64.exe.config F:\Documents\Projects\DataMining\DMS_Managers\DMS_Update_Manager\DMSUpdateManagerConsole\bin\DMS_Programs\ /Y /D
rem xcopy x64\DLLVersionInspector_x64.exe        F:\Documents\Projects\DataMining\DMS_Managers\DMS_Update_Manager\DMSUpdateManagerConsole\bin\DMS_Programs\AnalysisToolManager1\ /Y /D
rem xcopy x64\DLLVersionInspector_x64.exe.config F:\Documents\Projects\DataMining\DMS_Managers\DMS_Update_Manager\DMSUpdateManagerConsole\bin\DMS_Programs\AnalysisToolManager1\ /Y /D
pause
