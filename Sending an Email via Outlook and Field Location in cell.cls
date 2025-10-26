VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sending an Attachment via Email

Option Explicit

Sub Emails()
    Dim OutlookApp As Object
    Dim OutlookMail As Object
    Dim rng As ListObject
    Dim Recipient As String
    Dim CCRecipient As String
    Dim AttachmentPath As String
    Dim AttachmentPath2 As String
    Dim Body As String
    Dim Subject As String
    Dim i As Long
    
    Set rng = ThisWorkbook.Sheets("Sheet1").ListObjects("Tabular1")

    ' Loop through each row in the table
    For i = 1 To rng.ListRows.Count
        Set OutlookApp = CreateObject("Outlook.Application")

        ' Read values from the table
        Recipient = rng.ListRows(i).Range.Cells(1, 5).Value ' To
        CCRecipient = rng.ListRows(i).Range.Cells(1, 8).Value ' CC
        Subject = "Appreciation, Heartiest Congratulations" ' Subject
        Body = "Hello Boschler, Evac Boschler"
        AttachmentPath = rng.ListRows(i).Range.Cells(1, 9).Value
        AttachmentPath2 = rng.ListRows(i).Range.Cells(1, 10).Value

        ' Create a new email
        Set OutlookMail = OutlookApp.CreateItem(0)

        ' Set email properties
        With OutlookMail
            .To = Recipient
            .cc = CCRecipient
            .Subject = Subject
            .Body = Body

            ' Check if the attachment paths are not empty
            ' Check if the attachment paths are not empty
        If AttachmentPath <> "" Then
            .Attachments.Add AttachmentPath
        End If

        If AttachmentPath2 <> "" Then
            .Attachments.Add AttachmentPath2
        End If


            ' Display the email (you can use .Send instead of .Display to send the email)
            .Display
        End With

        ' Clean up
        Set OutlookMail = Nothing
    Next i

    ' Release Outlook object
    Set OutlookApp = Nothing
End Sub



