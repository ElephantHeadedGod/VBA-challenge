Attribute VB_Name = "Module2"
Sub LoopWorksheet():
 Worksheets("2018").Activate
 Call StockLoop
 Worksheets("2019").Activate
 Call StockLoop
 Worksheets("2020").Activate
 Call StockLoop
End Sub
