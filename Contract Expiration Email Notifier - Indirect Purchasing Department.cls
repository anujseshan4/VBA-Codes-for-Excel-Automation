VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'Value based Emailer

Sub FilteredEmail()
    Dim OutlookApp As Object
    Dim OutlookMail As Object
    Dim setupsht As Worksheet
    Dim MailBody As String
    Dim Recipient As String
    Dim Subject As String
    Dim I As Long
    
    Set OutlookApp = GetObject(, "Outlook.Application")
        Set OutlookApp = CreateObject("Outlook.Application")

  
    Set setupsht = ThisWorkbook.Worksheets("MailAutomation")
    
    For I = 2 To setupsht.Range("A" & setupsht.Rows.Count).End(xlUp).Row
    
        MailBody = "Respected " & setupsht.Range("B" & I).Value & "," & vbLf & vbLf & "We regret to inform you that your contract has expired, please kindly renew the contract to avail our services" & vbLf & vbLf & "Yours Sincerely," & vbLf & vbLf & "XYZ" & vbLf & "GSC PUI-AP2"
        Recipient = setupsht.Range("F" & I).Value
        Subject = "Email Subject for Row"
        
        ' The Main Code Factor
        If setupsht.Range("G" & I).Value > 0 Then
            Set OutlookMail = OutlookApp.CreateItem(0)
            
            With OutlookMail
                .To = Recipient
                .Subject = Subject
                .Body = MailBody
                ' You can add more email properties or attachments as needed
                .Attachments.Add ("C:\Users\ECN3KOR\Downloads\contactdetailsconfirmation.msg")
        '.Send
                .Display
            End With
        Else
            Set OutlookMail = OutlookApp.CreateItem(0)

            With OutlookMail
                .To = Recipient
                .Subject = "Contractual Obligations"
                .Body = "Respected " & setupsht.Range("B" & I).Value & "," & vbLf & vbLf & "We are extremely happy to inform you that you have the adequate time of " & setupsht.Range("G" & I).Value & " days, please kindly renew the contract to avail our services and a chance of extra incentives" & vbLf & vbLf & "Yours Sincerely," & vbLf & "XYZ" & vbLf & "GSC PUI-AP2"
                ' .Attachments.Add "C:\Path\To\AlternativeAttachment.pdf"
                
                ' Send the email
                .Display
            End With
        End If
        
        Set OutlookMail = Nothing
    Next I
    
    Set OutlookApp = Nothing
End Sub


