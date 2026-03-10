@echo off
SET MT5WorkflowFolder=E:\Trading\MT5_Workflow_Manager
SET MT5Folder=C:\Program Files\Pepperstone_MT5_01
SET MT5MQL5Folder=C:\Users\msand\AppData\Roaming\MetaQuotes\Terminal\98196E2B1CEDEE516442D255B458C6C2\MQL5
SET BacktestOutputFolder=E:\Trading\Analysis_Ouput
SET MT5BackTestFrom=2010.01.01
SET MT5BackTestTo=2025.12.31

cls

Rem ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Rem Step 4: Compile all the *.mt5 files, and move them to the PineappleStrats EA folder in MetaTrader 5
Rem ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------
python "%MT5WorkflowFolder%\Step4_Compile_MT5_EAs.py" -s "%BacktestOutputFolder%" -m "%MT5Folder%" -i "%MT5MQL5Folder%" -e "%MT5MQL5Folder%\Experts\Advisors\PineappleStrats"

echo Export Advisors Compilation Complete

Rem ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Rem Step 5: Back test all EAs in MetaTrader 5, and save reports
Rem ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------
python "%MT5WorkflowFolder%\Step5_MT5_Backtest.py" --mt5-terminal-path "%MT5Folder%\terminal64.exe" --report-dest-folder "%BacktestOutputFolder%" --model 4 --from-date "%MT5BackTestFrom%" --to-date "%MT5BackTestTo%"

echo Export Advisors Back Testing Complete

pause 
