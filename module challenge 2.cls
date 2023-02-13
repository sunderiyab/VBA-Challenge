VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub AlphabeticalTesting()

Dim ws As Worksheet
For Each ws In ThisWorkbook.Worksheets
    Dim ticker As String
    Dim volume As Double
    Dim lastrow As Long
    Dim i As Long
    Dim rowCounter As Long
    
    rowCounter = 2
    volume = 0
    lastrow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
    ws.range("I1").value = "Ticker"
    ws.range("J1").value = "Open"
    ws.range("K1").value = "Close"
    ws.range("L1").value = "Yearly Change"
    ws.range("M1").value = "Percentage Change"
    ws.range("N1").value = "Total Stock Volume"
    
    For i = 2 To lastrow
        ticker = ws.Cells(i, 1).value
        
        If ws.Cells(i - 1, 1).value <> ticker Then
            ws.range("I" & rowCounter).value = ticker
            ws.range("J" & rowCounter).value = ws.Cells(i, 3).value
            volume = ws.Cells(i, 6).value
        ElseIf ws.Cells(i + 1, 1).value <> ticker Then
            ws.range("K" & rowCounter).value = ws.Cells(i, 6).value
            ws.range("L" & rowCounter).value = ws.range("K" & rowCounter).value - ws.range("J" & rowCounter).value
            ws.range("M" & rowCounter).value = ws.range("L" & rowCounter).value / ws.range("J" & rowCounter).value
            ws.range("M" & rowCounter).NumberFormat = "0.00%"
            ws.range("N" & rowCounter).value = volume
            rowCounter = rowCounter + 1
            volume = 0
        Else
            volume = volume + ws.Cells(i, 6).value
        End If
    Next i
Next ws


End Sub


Sub changecolors()

Dim ws As Worksheet
For Each ws In ThisWorkbook.Worksheets
Dim lastrow As Long
Dim i As Long

lastrow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
For i = 2 To lastrow
    
    'positive to green 4
    'negative to red 3
        If ws.Cells(i, 12) > 0 Then
    ws.Cells(i, 12).Interior.ColorIndex = 4
       
       Else
       ws.Cells(i, 12).Interior.ColorIndex = 3
       
    
End If

Next i

Next ws

End Sub


Sub functionality()

Dim ws As Worksheet
For Each ws In ThisWorkbook.Worksheets
Dim lastrow As Long
Dim i As Long
Dim highestValue As Double
Dim lowestValue As Double
Dim rng As range
Dim ticker As String
Dim volume As Double

    
ws.range("Q1").value = "Ticker"
ws.range("R1").value = "Value"
ws.range("P2").value = "greatest % increase"
ws.range("P3").value = " greatest % decrease"
ws.range("p4").value = "greatest volume"


lastrow = ws.Cells(ws.Rows.Count, "M").End(xlUp).Row
Set rng = ws.range("m2:M" & lastrow)

    lastrow = ws.Cells(ws.Rows.Count, "N").End(xlUp).Row
    highestValue = WorksheetFunction.Max(ws.range("M2:M" & lastrow))
    lowestValue = WorksheetFunction.Min(ws.range("M2:M" & lastrow))
    highestvolume = WorksheetFunction.Max(ws.range("n2:N" & lastrow))
    

For i = 2 To lastrow

    If ws.Cells(i, "m").value = highestValue Then
    ticker = ws.Cells(i, "I").value
    ws.range("Q2").value = ticker
    ws.range("R2").value = highestValue
   End If
   
    If ws.Cells(i, "M").value = lowestValue Then
    ticker = ws.Cells(i, "I").value
    ws.range("Q3").value = ticker
    ws.range("R3").value = lowestValue
    End If

    If ws.Cells(i, "N").value = highestvolume Then
                ticker = ws.Cells(i, "I").value
                ws.range("Q4").value = ticker
                volume = ws.Cells(i, "N").value
                ws.range("R4").value = volume
                ws.range("R2:R3").NumberFormat = "0.00%"
    
    End If
    Next i

    Next ws
    
    End Sub

    
  
