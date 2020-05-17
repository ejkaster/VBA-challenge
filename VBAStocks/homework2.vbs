Sub StockSummary()

' ----------------------------
' Loop through all worksheets
' ----------------------------
         
     Dim ws As Worksheet
     For Each ws In Worksheets


' Add summary table column headers

    Range("I1").Value = "Ticker"
    Range("J1").Value = "Yearly Change"
    Range("K1").Value = "Percent Change"
    Range("L1").Value = "Total Stock Volume"
 
' ----------------------------
' Add unique ticker row and total stock volume
' ----------------------------

  ' Set an initial variable for holding the ticker symbol name
  Dim Ticker_Name As String

  ' Set an initial variable for holding the total volume per stock
  Dim Stock_Total As Double
  Stock_Total = 0

  ' Set variables for pricing delta and % delta
        Dim Open_Price As Double
        Open_Price = 0
        Dim Close_Price As Double
        Close_Price = 0
        Dim Delta_Price As Double
        Delta_Price = 0
        Dim Delta_Percent As Double
        Delta_Percent = 0

  ' Keep track of the location for each ticker in the summary table
  Dim Summary_Table_Row As Integer
  Summary_Table_Row = 2

  ' Loop through rows in the ticker column
    Dim LastRow As Long
    LastRow = Cells(Rows.Count, 1).End(xlUp).Row
    Open_Price = Cells(2, 3).Value
  
    For i = 2 To LastRow

    ' Check if we are still within the same ticker name, if it is not...
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then

      ' Set the Ticker name
        Ticker_Name = Cells(i, 1).Value

      ' Calculate price and % delta
        Close_Price = Cells(i, 6).Value
        Delta_Price = Close_Price - Open_Price
        If Open_Price <> 0 Then
            Delta_Percent = Delta_Price / Open_Price
        Else
          ' Unlikely, but it needs to be checked to avoid program crushing
            MsgBox ("For " & Ticker_Name & ", Row " & Str(i) & ": Open Price =" & Open_Price & ". Fix <open> field manually and save the spreadsheet.")
        End If



      ' Add to the Stock Total
        Stock_Total = Stock_Total + Cells(i, 7).Value

      ' Print the Ticker Name in the Summary Table
        Range("I" & Summary_Table_Row).Value = Ticker_Name

      'Print the price delta in the summary table
        Range("J" & Summary_Table_Row).Value = Format(Delta_Price, "Standard")

      'Format price delta with conditional colors
      If (Delta_Price > 0) Then
        Range("J" & Summary_Table_Row).Interior.ColorIndex = 4
      ElseIf (Delta_Price <= 0) Then
        Range("J" & Summary_Table_Row).Interior.ColorIndex = 3
      End If

      'Print the % delta in the summary table
        Range("K" & Summary_Table_Row).Value = Format(Delta_Percent, "Percent")

      ' Print the Stock Total to the Summary Table
        Range("L" & Summary_Table_Row).Value = Stock_Total
 
      ' Add one to the summary table row
        Summary_Table_Row = Summary_Table_Row + 1
        Delta_Price = 0
        Close_Price = 0
        Open_Price = Cells(i + 1, 3).Value
        Stock_Total = 0

    ' If the cell immediately following a row is the same ticker...
    Else

      ' Add to the Brand Total
      Stock_Total = Stock_Total + Cells(i, 7).Value
      
    End If
    
    Next i

   Next ws

End Sub



