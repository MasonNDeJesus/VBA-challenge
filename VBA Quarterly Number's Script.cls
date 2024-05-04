VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub stocks():
   
    Dim i As Long
    Dim cell_vol As LongPtr
    Dim vol_total As Double
    Dim ticker As String
   
    Dim K As Long
   
    Dim ticker_close As Double
    Dim ticker_open As Double
    Dim price_change As Double
    Dim percent_change As Double
    Dim lastRow As Long
   
    Dim ws As Worksheet
   
    For Each ws In ThisWorkbook.Worksheets
    
        On Error Resume Next
   
        lastRow = ws.Cells(Rows.Count, "A").End(xlUp).Row
   
        vol_total = 0
        K = 2
       
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Quarterly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Volume"
       
        ticker_open = ws.Cells(2, 3).Value
       
        For i = 2 To lastRow
            cell_vol = ws.Cells(i, 7).Value
            ticker = ws.Cells(i, 1).Value
           
            If (ws.Cells(i + 1, 1).Value <> ticker) Then
                 
                 vol_total = vol_total + cell_vol
                 
                 ticker_close = ws.Cells(i, 6).Value
                 price_change = ticker_close - ticker_open
                 
                 If (ticker_open > 0) Then
                    percent_change = price_change / ticker_open
                Else
                    percent_change = 0
                End If
               
                ws.Cells(K, 9).Value = ticker
                ws.Cells(K, 10).Value = price_change
                ws.Cells(K, 11).Value = percent_change
                ws.Cells(K, 12).Value = vol_total
               
                If (price_change > 0) Then
                    ws.Cells(K, 10).Interior.ColorIndex = 4
                ElseIf (price_change < 0) Then
                    ws.Cells(K, 10).Interior.ColorIndex = 3
                Else
                    ws.Cells(K, 10).Interior.ColorIndex = 2
                End If
               
                vol_total = 0
                K = K + 1
                ticker_open = ws.Cells(i + 1, 3).Value
            Else
                vol_total = vol_total + cell_vol
            End If
        Next i
       
        ws.Columns("K:K").NumberFormat = "0.00%"
        ws.Columns("I:L").AutoFit
       
    Next ws
   
End Sub
