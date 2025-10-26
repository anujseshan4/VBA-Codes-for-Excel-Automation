VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'Scheduling an Outlook Meeting
Sub MailAutomation()
    Dim OutApp As Outlook.Application, Outmeet As Outlook.AppointmentItem
    Dim I As Long, setupsht As Worksheet
    
    Set setupsht = Worksheets("Setup")
    
    For I = 2 To Range("A" & Rows.Count).End(xlUp).Row
        Set OutApp = Outlook.Application
        Set Outmeet = OutApp.CreateItem(olAppointmentItem)
        
        With Outmeet
            .Subject = "#BuildingGreatPlaceToWork Survey Participation"
            .RequiredAttendees = setupsht.Range("F" & I).Value
            .Start = setupsht.Range("D" & I).Value
            .Duration = setupsht.Range("E" & I).Value
            .Importance = olImportanceHigh
            .Body = "Hello " & setupsht.Range("B" & I).Value & "," & vbLf & vbLf & "Thank you for your participation in the #BuildingGreatPlaceToWork Survey and sharing your insights with us to make a difference." & vbLf & vbLf & "We are at 36% participation in COB location and 45% participation in Bangalore. Hence encouraging the one who is yet to take the survey to take out some time and share your inputs with us." & vbLf & vbLf & "Your feedback/insights will really help and – Together, We Build, #Great Place to Work" & vbLf & vbLf & "Warm Regards," & vbLf & "GS/HRL-IN"
            .Location = "Adugodi - 207 - 2 Floor - HR Room"
            .MeetingStatus = olMeeting
            .ReminderMinutesBeforeStart = 15
            .Display
            '.Send
        End With
        
    Next I
End Sub

