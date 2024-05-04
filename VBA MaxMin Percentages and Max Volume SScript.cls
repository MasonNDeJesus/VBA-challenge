VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub FindMaxMinValues()
    Dim dataPercentRange As Range
    Dim dataVolumeRange As Range
    Dim cell As Range
    Dim greatestPercentageIncrease As Double
    Dim greatestPercentageIncreaseCell As Range
    Dim greatestPercentageDecrease As Double
    Dim greatestPercentageDecreaseCell As Range
    Dim greatestTotalVolume As Double
    Dim greatestTotalVolumeCell As Range
    Dim ws As Worksheet

    For Each ws In ThisWorkbook.Worksheets
     
        ws.Range("Q1") = "Value"
        ws.Range("P1") = "Ticker"

        ws.Range("O2") = "Greatest % Increase"
        ws.Range("O3") = "Greatest % Decrease"
        ws.Range("O4") = "Greatest Total Volume"
    
        Set dataPercentRange = ws.Range("K2:K1501")
        Set dataVolumeRange = ws.Range("L2:L1501")
        
        greatestPercentageIncrease = 0
        greatestPercentageDecrease = 1
        
        For Each cell In dataPercentRange

            If IsNumeric(cell.Value) And cell.Value < greatestPercentageDecrease Then
                greatestPercentageDecrease = cell.Value
                Set greatestPercentageDecreaseCell = cell
            End If
        Next cell
        
        If Not greatestPercentageDecreaseCell Is Nothing Then
            ws.Range("Q3").Value = greatestPercentageDecrease

        Else

        End If
                 
        For Each cell In dataPercentRange
    
            If IsNumeric(cell.Value) And cell.Value > greatestPercentageIncrease Then
                greatestPercentageIncrease = cell.Value
                Set greatestPercentageIncreaseCell = cell
            End If
        Next cell
        
        If Not greatestPercentageIncreaseCell Is Nothing Then
            ws.Range("Q2").Value = greatestPercentageIncrease
        Else
        
        End If
        
        For Each cell In dataVolumeRange

        If IsNumeric(cell.Value) And cell.Value > greatestTotalVolume Then
            greatestTotalVolume = cell.Value
            Set greatestTotalVolumeCell = cell
        End If
        Next cell
        
        If Not greatestTotalVolumeCell Is Nothing Then
            ws.Range("Q4").Value = greatestTotalVolume
        Else
    
        End If
            
        ws.Cells(2, 17).NumberFormat = "0.00%"
        ws.Cells(3, 17).NumberFormat = "0.00%"
        ws.Columns("O:Q").AutoFit
    Next ws
    
End Sub


