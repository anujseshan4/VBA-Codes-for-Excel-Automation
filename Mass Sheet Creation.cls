VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub AddingSheets()
    AddAndRenameSheet "BGSW"
    AddAndRenameSheet "South"
    AddAndRenameSheet "North"
    AddAndRenameSheet "Robert Bosch GmbH"
End Sub

Sub AddAndRenameSheet(sheetName As String)
    Sheets.Add(After:=ActiveSheet).name = sheetName
End Sub
