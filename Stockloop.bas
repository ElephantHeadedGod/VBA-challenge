Attribute VB_Name = "Module1"
Sub StockLoop():
    Dim rng As Range 'set the range of cells to be formatted
    Dim i As Long 'i is the variable in the for loop that represents row number
    Dim rowA As String 'rowA is the value of the cells in column 1
    Dim opening As Double 'opening is the opening price of the stock
    Dim closing As Double 'closing is closing price of the stock
    Dim j As Long   'j is a variable to represent the new row number
    Dim k As Long ' k represents the new row number that increases every time the stock opens
    Dim StockVolume As LongLong 'StockVolume tells the volume of stock
    Dim m As Integer 'm represents a flag to indicate cumulative total of stock volume
    Dim x As Range ' x is the range of values to find the max of
    Dim y As Double 'y is the max
    Dim FndRng As Range 'range where you have to find the row number
    Dim maxRow As Long ' find the row number
    Dim FndRng2 As Range 'range where you have to find the row number
    Dim maxRow2 As Long ' find the row number
    Dim FndRng3 As Range 'range where you have to find the row number
    Dim maxRow3 As Long ' find the row number
    Dim z As Double 'z is the range of values to find the min of
    Dim a As Range 'a is the range of values to find the next max of
    Dim b As Double ' b is the value of the min for z
    Dim LRcolumn1 As Long
   ' LR = Last Row
    k = 1 'set k=1 to initialize it
    StockVolume = 0 'set StockVolume = 0 to initialize it
    
    m = 0 'set flag to 0 to initialize it

    LRcolumn1 = Cells(Rows.Count, 1).End(xlUp).Row
    
     For i = 2 To LRcolumn1  'create a for loop from 2 to the last row
        rowA = Cells(i, 1).Value 'set rowA variable equal to the value in cell with row i and column 1
       If m = 0 Then 'create if statement where flag = 0
         StockVolume = StockVolume + Cells(i, 7).Value 'set the volume of stocks equal to itself plus the value in cell with row i and column 7
       ElseIf m = 1 Then 'create else if for instance when flag is 1
         Cells(k, 12).Value = StockVolume ' put the stockvolume in cell with row k and column 12
         m = 0 'set flag back to 0
         StockVolume = 0 'set StockVolume back to 0
        End If
        If rowA <> Cells(i - 1, 1).Value Then 'create an if statement saying if rowA variable doesn't equal the value in the cell one row before it
            opening = Cells(i, 3).Value 'variable called opening to cell value in row i and column 3, it is the opening price of the stocks
            j = i 'set j equal to i
            k = k + 1 'set k equal to itself plus one
            Range("I" & k).Value = rowA 'put the value of rowA into cell column I and row k
        ElseIf IsEmpty(Cells(i + 1, 1).Value) Then
            Cells(k, 12).Value = StockVolume
        ElseIf rowA <> Cells(i + 1, 1).Value Then 'create an elseif statement that says if rowA variable does not equal cell in the next row and column one
            closing = Cells(i, 6).Value 'the closing value is the value taken from row i and column 6 cell
            m = 1 'set flag to 1
        End If
        If closing - opening = 0 Then
            Cells(k, 10).Value = 0
            Cells(k, 11).Value = 0
        ElseIf closing - opening <> 0 Then 'if closing - opening doesn't equal 0
            Cells(k, 10).Value = closing - opening ' then put the value of closing - opening in yearly change column
            Cells(k, 11).Value = ((closing - opening) / opening) * 100 'calculate and put in percent change
        End If
    Next i
    'set the range of cells for max/min to be calculated
    'find the max and put it in y,find the min put it in z
    'put the value of y/z/b in the new cell
    Dim LRcolumnK As Long
    LRcolumnK = Cells(Rows.Count, 11).End(xlUp).Row
    Dim percentChangeRange As String
    percentChangeRange = "K2:K" & LRcolumnK
    Set x = Range(percentChangeRange)
    'Set x = Range(Cells(11, 2).Address(), Cells(11, LRcolumnK).Address())
    'Set x = Range("K2:K3001")
    y = Application.WorksheetFunction.Max(x)
    
    Cells(2, 16).Value = y
    Set FndRng = x.Find(what:=y)
    maxRow = FndRng.Row
    Cells(2, 15).Value = Cells(maxRow, "I")
    z = Application.WorksheetFunction.Min(x)
    Cells(3, 16).Value = z
    Set FndRng2 = x.Find(what:=z)
    maxRow2 = FndRng2.Row
    Cells(3, 15).Value = Cells(maxRow2, "I")
      
    Dim LRcolumnL As Long
    LRcolumnL = Cells(Rows.Count, 12).End(xlUp).Row
    Dim percentChangeRange2 As String
    percentChangeRange2 = "L2:L" & LRcolumnL
    
    Set a = Range(percentChangeRange2)
    b = Application.WorksheetFunction.Max(a)
    Cells(4, 16).Value = b
    Set FndRng3 = a.Find(what:=b)
    maxRow3 = FndRng3.Row
    Cells(4, 15).Value = Cells(maxRow3, "I")
    Dim condition1 As FormatCondition, condition2 As FormatCondition ' set conditions as formatting conditions
    
    Dim LRcolumnJ As Long
    LRcolumnJ = Cells(Rows.Count, 10).End(xlUp).Row
    Dim percentChangeRange4 As String
    percentChangeRange4 = "J2:J" & LRcolumnJ
    Set rng = Range(percentChangeRange4)

  rng.FormatConditions.Delete 'delete the format already in excel
  'put the range of values to colr red and green
  Set condition1 = rng.FormatConditions.Add(xlCellValue, xlGreater, "=0")
  Set condition2 = rng.FormatConditions.Add(xlCellValue, xlLess, "=0")
  With condition1
    .Interior.Color = rgbGreen
   End With

   With condition2
    .Interior.Color = rgbRed
   End With
End Sub

