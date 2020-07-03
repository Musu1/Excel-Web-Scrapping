'Create an instance of Excel
 Set ExcelApp = CreateObject("Excel.Application")

'Execute Macro Code
ExcelApp.Application.Run "'C:\Users\Muskan Agarwal\Downloads\try9 (1).xlsm'!Module2.Main"


'Prevent any App Launch Alerts (ie Update External Links)
  ExcelApp.DisplayAlerts = False


'End instance of Excel
  ExcelApp.Application.Quit

set ExcelApp = Nothing