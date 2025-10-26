VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'Column based sheet creation
Sub Macro1()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim cell As Range
    
    ' Set the worksheet (adjust the sheet name if necessary)
    Set ws = ThisWorkbook.Sheets("Sheet1")
    Columns("G:G").Insert Shift:=xlToLeft
    ws.Cells(1, 7).Value = "Region"
    ws.Cells(1, 7).Font.Bold = True
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
    ' Loop through each row in the specified range
    For Each cell In ws.Range("F1:F" & lastRow)
        If cell.Value = "6520" Or cell.Value = "6521" Or cell.Value = "6522" Or cell.Value = "6523" Or cell.Value = "6525" Or _
           cell.Value = "652A" Or cell.Value = "652B" Or cell.Value = "652C" Or cell.Value = "652L" Or cell.Value = "652M" Or _
           cell.Value = "W237" Or cell.Value = "W270" Or cell.Value = "W274" Then
            cell.Offset(0, 1).Value = "BGSW"
            
        ElseIf cell.Value = "180B" Or cell.Value = "1810" Or cell.Value = "33D0" Or cell.Value = "5930" Or cell.Value = "65B0" Or _
           cell.Value = "9020" Or cell.Value = "IN10" Or cell.Value = "IN20" Then
            cell.Offset(0, 1).Value = "North"
            
        ElseIf cell.Value = "13DW" Or cell.Value = "15F0" Or cell.Value = "1720" Or cell.Value = "1760" Or cell.Value = "280W" Or _
           cell.Value = "54F9" Or cell.Value = "603V" Or cell.Value = "603W" Or cell.Value = "903C" Or cell.Value = "E13B" Or _
           cell.Value = "IN11" Or cell.Value = "IN12" Or cell.Value = "IN16" Or cell.Value = "IN17" Or cell.Value = "W002" Or _
           cell.Value = "W047" Or cell.Value = "W051" Or cell.Value = "W053" Or cell.Value = "W058" Or cell.Value = "W063" Or _
           cell.Value = "W069" Or cell.Value = "W070" Or cell.Value = "W279" Or cell.Value = "W285" Or cell.Value = "W063" Then
            cell.Offset(0, 1).Value = "South"
        
        End If
    Next cell
End Sub


