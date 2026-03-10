@echo off
SET MT5WorkflowFolder=E:\Trading\MT5_Workflow_Manager
SET MT5Folder=C:\Program Files\Pepperstone_MT5_01
SET MT5MQL5Folder=C:\Users\msand\AppData\Roaming\MetaQuotes\Terminal\98196E2B1CEDEE516442D255B458C6C2\MQL5
SET QDMFolder=C:\QuantDataManager125
SET ExportDataFrom=2026.01.30
SET ExportDataTo=2026.01.30

Rem ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Rem Step 1: Update data in QuantDataManager
Rem ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------
python "%MT5WorkflowFolder%\Step1_Refresh_QDM_Data.py" "%QDMFolder%"\qdmcli.exe

echo QuantDataManager Data Update Complete

Rem ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Rem Step 2: Export tick data from QuantDataManager
Rem ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------
python "%MT5WorkflowFolder%\Step2_Export_Data_From_QDM.py" --date-from "%ExportDataFrom%" --date-to "%ExportDataTo%" --qdm-path "%QDMFolder%" --export-path "%MT5MQL5Folder%"\Files

echo Data Export from QuantDataManager Complete

Rem ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Rem Step 3: Launch MetaTrader 5, so the Service can import the tick data
Rem ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------
python "%MT5WorkflowFolder%\Step3_Start_MT5_Import.py" --mt5-path "%MT5Folder%"\terminal64.exe --wait --close-mt5

echo MetaTrader 5 Data Import Complete

pause 
