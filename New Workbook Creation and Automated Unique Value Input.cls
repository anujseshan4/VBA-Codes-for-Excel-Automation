VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'New Workbook Creation with Name Creation

Sub Mobility_and_Hospitality()

    Dim ws As Worksheet
    Dim ws2 As Worksheet

    ' Plant wise Split Dimensions
    Dim data_sh As Worksheet
    Dim folderPath As String
    Dim loc_sh As Worksheet
    Dim lastRow As Long
    Dim uniqueNames As Object
    Dim cell As Range

    ' Pivot Table Range and Workspace
    Set ws = ThisWorkbook.Worksheets("Sheet1")
    
    ' Location for Sheet Creation
    Set data_sh = ThisWorkbook.Sheets("Sheet1")
    Set loc_sh = ThisWorkbook.Sheets("Location")
    Set uniqueNames = CreateObject("Scripting.Dictionary")
    Application.ScreenUpdating = False

    ' Retrieve Unique names from Department Codes
    lastRow = data_sh.Cells(data_sh.Rows.Count, "K").End(xlUp).Row
    For Each cell In data_sh.Range("K2:K" & lastRow)
        If Not uniqueNames.exists(cell.Value) And cell.Value Like "TRA_*" Then
            uniqueNames.Add cell.Value, cell.Value
        End If
    Next cell

    ' Use a For Loop to copy the unique rows to a New File
    Dim name As Variant
    For Each name In uniqueNames.keys
        Dim newWb As Workbook
        Set newWb = Workbooks.Add
        Dim newSh As Worksheet
        Set newSh = newWb.Sheets(1)

        ' Copy header row from original data sheet
        data_sh.Rows(1).Copy newSh.Rows(1)

        ' Filter and copy data rows matching current 'name'
        Dim destRow As Long
        destRow = 2
        For Each cell In data_sh.Range("K2:K" & lastRow)
            If cell.Value = name Then
                data_sh.Rows(cell.Row).Copy newSh.Rows(destRow)
                destRow = destRow + 1
            End If
        Next cell

        ' Save the new workbook
        folderPath = loc_sh.Range("H6").Value & "\"
        Dim fileName As String
        fileName = name & ".xlsx"
        newWb.SaveAs folderPath & fileName
        newWb.Close SaveChanges:=True
    Next name
End Sub


