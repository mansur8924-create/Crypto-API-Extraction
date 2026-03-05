Attribute VB_Name = "Module1"
' ***************************************************************
' Project: Crypto Market Intelligence Auditor
' Purpose: Automatically analyzes RSI and Trend data to flag
'          potential entry points.
' ***************************************************************

Sub AnalyzeMarketIntelligence()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim rsiValue As Double
    Dim assetName As String
    Dim alertCount As Integer
    
    ' Set the worksheet to the one containing your crypto data
    Set ws = ThisWorkbook.Sheets("Market Intelligence")
    
    ' Find the last row with data in the Asset column (Column E)
    lastRow = ws.Cells(ws.Rows.Count, "E").End(xlUp).Row
    alertCount = 0
    
    ' Loop through the data starting from the row after the header
    ' Based on your structure, data begins around row 6
    For i = 6 To lastRow
        assetName = ws.Cells(i, 5).Value ' Column E: Asset
        rsiValue = ws.Cells(i, 9).Value  ' Column I: RSI (14)
        
        ' Logic: Flag assets with RSI below 30 (Oversold condition)
        If rsiValue < 30 And rsiValue > 0 Then
            ' Highlight the row in a sophisticated light gold
            ws.Range(ws.Cells(i, 1), ws.Cells(i, 12)).Interior.Color = RGB(255, 240, 200)
            alertCount = alertCount + 1
        Else
            ' Clear any previous highlighting
            ws.Range(ws.Cells(i, 1), ws.Cells(i, 12)).Interior.ColorIndex = xlNone
        End If
    Next i
    
    ' Final Report to the user
    If alertCount > 0 Then
        MsgBox "Audit Complete. " & alertCount & " assets identified as 'Oversold' (RSI < 30).", vbInformation, "Market Intelligence"
    Else
        MsgBox "Audit Complete. No immediate RSI alerts found.", vbInformation, "Market Intelligence"
    End If
End Sub
