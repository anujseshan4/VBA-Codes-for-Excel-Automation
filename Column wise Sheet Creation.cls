VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'Sheet Creation and Column wise worksheet creation
Sub ExtractRows()
 Dim ws As Worksheet
 Dim wsSouth As Worksheet
 Dim wsNorth As Worksheet
 Dim wsBGSW As Worksheet
 Dim lastRow As Long
 Dim i As Long, j As Long
 
 ' Set the Main Worksheet
 Set ws = ThisWorkbook.Sheets("Sheet2")
 Columns("W:W").Insert Shift:=xlToLeft
 ws.Cells(1, 23).Value = "Region"
 ws.Cells(1, 23).Font.Bold = True
 lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
 
 ' Adding New Sheets
 Set wsSouth = Sheets.Add(After:=Sheets(Sheets.Count))
 wsSouth.name = "South"
 Set wsNorth = Sheets.Add(After:=Sheets(Sheets.Count))
 wsNorth.name = "North"
 Set wsBGSW = Sheets.Add(After:=Sheets(Sheets.Count))
 wsBGSW.name = "BGSW"
 
 ' Header Rows
 ws.Rows(1).Copy wsSouth.Rows(1)
 ws.Rows(1).Copy wsNorth.Rows(1)
 ws.Rows(1).Copy wsBGSW.Rows(1)
 
 
 lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
 
 'Region Assignment
 For Each cell In ws.Range("V1:V" & lastRow)
 If cell.Value Like "6520*" Or cell.Value Like "6521*" Or cell.Value Like "6522*" Or cell.Value Like "6523*" Or cell.Value Like "6525*" Or _
 cell.Value Like "652A*" Or cell.Value Like "652B*" Or cell.Value Like "652C*" Or cell.Value Like "652L*" Or cell.Value Like "652M*" Or _
 cell.Value Like "W237*" Or cell.Value Like "W270*" Or cell.Value Like "W274*" Or cell.Value Like "652T*" Then
 cell.Offset(0, 1).Value = "BGSW"
 
 ElseIf cell.Value Like "180B*" Or cell.Value Like "1810*" Or cell.Value Like "33D0*" Or cell.Value Like "5930*" Or cell.Value Like "65B0*" Or _
 cell.Value Like "9020*" Or cell.Value Like "IN10*" Or cell.Value Like "IN20*" Then
 cell.Offset(0, 1).Value = "North"
 
 ElseIf cell.Value Like "13DW*" Or cell.Value Like "15F0*" Or cell.Value Like "1720*" Or cell.Value Like "1760*" Or cell.Value Like "280W*" Or _
 cell.Value Like "54F9*" Or cell.Value Like "603V*" Or cell.Value Like "603W*" Or cell.Value Like "903C*" Or cell.Value Like "E13B*" Or _
 cell.Value Like "IN11*" Or cell.Value Like "IN12*" Or cell.Value Like "IN16*" Or cell.Value Like "IN17*" Or cell.Value Like "W002*" Or _
 cell.Value Like "W047*" Or cell.Value Like "W051*" Or cell.Value Like "W053*" Or cell.Value Like "W058*" Or cell.Value Like "W063*" Or _
 cell.Value Like "W069*" Or cell.Value Like "W070*" Or cell.Value Like "W074*" Or cell.Value Like "W279*" Or cell.Value Like "W285*" Or cell.Value Like "7460" Then
 cell.Offset(0, 1).Value = "South"
 
 End If
 Next cell
 
 ' Extractor
 For i = 2 To lastRow
 If ws.Cells(i, 23).Value = "South" Then
 j = wsSouth.Cells(wsSouth.Rows.Count, "A").End(xlUp).Row + 1
 ws.Rows(i).Copy wsSouth.Rows(j)
 End If
 
 If ws.Cells(i, 23).Value = "North" Then
 j = wsNorth.Cells(wsNorth.Rows.Count, "A").End(xlUp).Row + 1
 ws.Rows(i).Copy wsNorth.Rows(j)
 End If
 
 If ws.Cells(i, 23).Value = "BGSW" Then
 j = wsBGSW.Cells(wsBGSW.Rows.Count, "A").End(xlUp).Row + 1
 ws.Rows(i).Copy wsBGSW.Rows(j)
 End If
 Next i
End Sub


