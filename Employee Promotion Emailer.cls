VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
VBA Mailer

Sub Mails()

Dim OutApp As Object
Dim msg As Object
Dim I As Integer
Dim sh As Worksheet
Dim AttachmentPath As String

Set sh = ThisWorkbook.Sheets("Sheet1")
Set OutApp = CreateObject("Outlook.application")


For I = 2 To Range("A" & Rows.Count).End(xlUp).Row
Set msg = OutApp.createitem(0)
    msg.To = sh.Range("F" & I).Value
    msg.CC = sh.Range("G" & I).Value & ";" & sh.Range("H" & I).Value & ";" & sh.Range("I" & I).Value & ";" & sh.Range("J" & I).Value
    msg.Subject = "TP1 Process - 2024"
    msg.body = "Dear" & " " & sh.Range("D" & I).Value & "," & vbCrLf & vbCrLf & _
                 "Firstly, happy to inform you that you are nominated for TP1 process. As a next step, there will be CBD (competency based dialog) tentatively set to begin by end of June 2024 and continuing July." & vbCrLf & _
                 "As part of the preparation, we are enabling the following measures:" & vbCrLf & vbCrLf & _
                 "• Update one-pager document, to enable you to consolidate your profile using the attached template as reference for your conversation in the CBD process. Please have this document reviewed with your GRM/SL leader on or before 24 May 2024." & vbCrLf & _
                 "• Face-to-face discussion with your neutral manager will be arranged for you to get to know your panel for CBD discussion. (EY partner not included) in CW 23/24." & vbCrLf & _
                 "• Suggestion that you speak to your GRM/SL leader for inputs and guidance for CBD preparation." & vbCrLf & vbCrLf & _
                 "Wishing you all the best, and please feel free to reach out to me for any clarification." & vbCrLf & vbCrLf & _
                 "Mit freundlichen Grüßen / Best regards" & vbCrLf & _
                 "Pete Mitchell"
    AttachmentPath = sh.Range("L" & I).Value
    If AttachmentPath <> "" Then
            msg.Attachments.Add AttachmentPath
        End If
    'msg.body = "Hello " & sh.Range("D" & I).Value & "," & vbLf & vbLf & "Firstly, happy to inform you that you are nominated for TP1 process.  As a next step, there will be CBD (competency based dialog) tentatively set to begin by end of June 2024 and continuing July." & vbLf & "As part of the preparation, we are enabling the following measures:" & vbLf & vbLf & Chr(149) & " " & "Update one-pager document, to enable you to consolidate your profile using the attached template as reference for your conversation in the CBD process. Please have this document reviewed with your GRM/SL leader on or before 24 May 2024." & vbLf & Chr(149) & " " & "Face-to-face discussion with your neutral manager will be arranged for you to get to know your panel for CBD discussion. ( EY partner not included) in CW 23/24." & vbLf & Chr(149) & " " & "Suggestion that you speak to your GRM/SL leader for inputs and guidance for CBD preparation." & vbLf & vbLf & "Wishing you all the best and please feel free to reach out to"
    msg.display
    'msg.Send
    
    Next I
    End Sub


