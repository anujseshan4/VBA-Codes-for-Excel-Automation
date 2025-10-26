VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
' Multiple Sheet Creation and sending reports over email
Sub Split_Data_in_workbooks()
    Application.ScreenUpdating = False

    Dim data_sh As Worksheet
    Set data_sh = ThisWorkbook.Sheets("Data")

    Dim setting_Sh As Worksheet
    Set setting_Sh = ThisWorkbook.Sheets("Settings")

    Dim nwb As Workbook
    Dim nsh As Worksheet

    ' Get unique supervisors
    setting_Sh.Range("A:A").Clear
    data_sh.AutoFilterMode = False
    data_sh.Range("U:U").Copy setting_Sh.Range("A1")
    setting_Sh.Range("A:A").RemoveDuplicates 1, xlYes

    Dim I As Long ' Using Long instead of Integer to avoid overflow
    For I = 2 To Application.CountA(setting_Sh.Range("A:A"))
        data_sh.UsedRange.AutoFilter 21, setting_Sh.Range("A" & I).Value

        Set nwb = Workbooks.Add
        Set nsh = nwb.Sheets(1)

        data_sh.UsedRange.SpecialCells(xlCellTypeVisible).Copy nsh.Range("A1")
        nsh.UsedRange.EntireColumn.ColumnWidth = 15
        nsh.Cells.EntireColumn.AutoFit

        ' Construct the file path for saving the workbook
        Dim folderPath As String
        folderPath = setting_Sh.Range("H6").Value & "\"
        Dim fileName As String
        fileName = setting_Sh.Range("A" & I).Value & ".xlsx"

        ' Save the workbook
        nwb.SaveAs folderPath & fileName
        nwb.Close False
        data_sh.AutoFilterMode = False
    Next I

    setting_Sh.Range("A:A").Clear
    MsgBox "Done"

    Dim OutApp As Object
    Dim msg As Object
    Dim sh As Worksheet

    ' Declare the loop variable
    Set sh = ThisWorkbook.Sheets("Mail")
    Set OutApp = CreateObject("Outlook.application")

    For I = 5 To sh.Range("A" & Rows.Count).End(xlUp).Row
        If sh.Range("C" & I).Value <> "" And sh.Range("B" & I).Value <> "" Then
            Set msg = OutApp.CreateItem(0)
            'msg.To =
            msg.CC = sh.Range("F" & I).Value
            msg.BCC = sh.Range("A" & I).Value
            msg.Subject = sh.Range("G" & I).Value
            msg.Body = sh.Range("H" & I).Value

            ' Construct the file path for the attachment
            Dim attachmentPath As String
            attachmentPath = "C:\Users\ECN3KOR\Desktop\ECN3KOR\Deeraj's Folder\Rate Cards\" & sh.Range("B" & I).Value & ".xlsx"

            ' Check if the attachment file exists before adding it
            If Dir(attachmentPath) <> "" Then
            msg.Attachments.Add attachmentPath
            Else
                MsgBox "File not found: " & attachmentPath, vbExclamation
            End If
            
        End If

            msg.Display
        
    Next I
    MsgBox "Done"
End Sub

Sub Send_mails()
Dim sh As Worksheet
Set sh = ThisWorkbook.Sheets("Mail")
Dim I As Integer

Dim D As Object
Dim msg As Object

Set D = CreateObject("Outlook.application")

Dim L As Integer
L = sh.Range("C" & Application.Rows.Count).End(xlUp).Row

For I = 5 To L
   
    If UCase(sh.Range("B" & I).Value) <> "" Then
        
        Set msg = D.CreateItem(0)
        
        If sh.Range("C" & I).Value <> "" Then msg.SentonBehalfofName = sh.Range("C" & I).Value
        
        msg.To = sh.Range("D" & I).Value
        msg.CC = sh.Range("E" & I).Value
        msg.Subject = sh.Range("F" & I).Value
        msg.Body = sh.Range("G" & I).Value
        
        If sh.Range("A" & I).Value <> "" Then
            'msg.Attachments.Add sh.Range("A" & I).Value
        End If
        
        If sh.Range("I1").Value = 1 Then
            msg.send
        Else
            'msg.display
        End If
        
    End If

Next I

MsgBox " Process completed !!", vbInformation

End Sub

Sub Get_File_Path()

Dim file_path As String
file_path = Application.GetOpenFilename(MultiSelect:=False)
If file_path <> "False" Then
    Selection.Value = file_path
End If

End Sub


