VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
' Report Generation - New Workbooks

Sub Data_Splitter()
Application.ScreenUpdating = False

Dim ws As Worksheet
Set ws = ThisWorkbook.Sheets("Awarding Summary")

Dim helper As Worksheet
Set helper = ThisWorkbook.Sheets("Helper")

Dim nwb As Workbook
Dim nsh As Worksheet

' Get unique supervisors

ws.Range("A:A").Clear
ws.AutoFilterMode = False
ws.Range("AB:AB").Copy helper.Range("A1")
helper.Range("A:A").RemoveDuplicates 1, xlYes

Dim I As Long ' Using Long instead of Integer to avoid overflow
    For I = 2 To Application.CountA(helper.Range("A:A"))
        ws.UsedRange.AutoFilter 21, helper.Range("A" & I).Value

        Set nwb = Workbooks.Add
        Set nsh = nwb.Sheets(1)

        ws.UsedRange.SpecialCells(xlCellTypeVisible).Copy nsh.Range("A1")
        nsh.UsedRange.EntireColumn.ColumnWidth = 15
        nsh.Cells.EntireColumn.AutoFit

        ' Construct the file path for saving the workbook
        Dim folderPath As String
        folderPath = helper.Range("H6").Value & "\"
        Dim fileName As String
        fileName = helper.Range("A" & I).Value & ".xlsx"

        ' Save the workbook
        nwb.SaveAs folderPath & fileName
        nwb.Close False
        ws.AutoFilterMode = False
    Next I

    helper.Range("A:A").Clear
    MsgBox "Done"
    End Sub


