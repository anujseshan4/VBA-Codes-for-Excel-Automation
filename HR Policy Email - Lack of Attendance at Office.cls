VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
HTML Table Attachment

Sub HybridWorkPolicy()
    Dim OutApp As Object
    Dim msg As Object
    Dim I As Integer
    Dim ws As Worksheet
    Dim LastRow As Long
    Dim tableHTML As String
    
    Set ws = ThisWorkbook.Sheets("Sheet1")
    Set OutApp = CreateObject("Outlook.application")
    
    LastRow = ws.Range("A" & Rows.Count).End(xlUp).Row
    lastColumn = WorksheetFunction.Min(ws.Cells(1, Columns.Count).End(xlToLeft).Column, ws.Columns("Q").Column)
    
    For I = 2 To LastRow
        Set msg = OutApp.CreateItem(0)
        msg.To = ws.Range("X" & I).Value
        msg.CC = ws.Range("Y" & I).Value & "; " & ws.Range("Z" & I).Value & "; " & ws.Range("AA" & I).Value
        msg.Subject = "Alignment to Smart Hybrid Work Policy"
        
        ' Check if value in column B is part of value in column D
        If InStr(1, ws.Range("D" & I).Value, ws.Range("B" & I).Value, vbTextCompare) > 0 Then
            ' Create table HTML with the entire row data
            tableHTML = "<table border='3'><tr>"
            For j = 1 To lastColumn
                tableHTML = tableHTML & "<th>" & ws.Cells(1, j).Value & "</th>"
            Next j
            tableHTML = tableHTML & "</tr><tr>"
            For j = 1 To lastColumn
                tableHTML = tableHTML & "<td>" & ws.Cells(I, j).Value & "</td>"
            Next j
            tableHTML = tableHTML & "</tr></table>"
        Else
            tableHTML = "<p>No matching data found.</p>"
        End If
        
        msg.HTMLBody = "<p>Dear " & ws.Range("D" & I).Value & ",</p>" & _
            "<p>Our Smart Hybrid Work Program was launched last July. Since the launch of the program, Boschlers have enthusiastically returned to our dynamic campuses." & _
            " Our spaces are now vibrant hubs of community engagement and creativity. From thought-provoking technology events to immersive learning workshops, captivating art-in-motion sessions, bustling marketplaces, inspiring roadshows, and more - there's something for everyone at BGSW!</p>" & _
            "<p>However, we observe that you have not been able to align with our Smart Hybrid Work Program for the months of <b> Aug 2024 </b>. Your presence in office has been consistently short of 3 days/week.</p>" & _
            tableHTML & _
            "<p>You are hereby immediately requested to align with the Smart Hybrid Work Program. Please reach out to your respective manager or HRBP for further support and information." & _
            " Great things happen when we join forces. Join us as we embrace this innovative way of working and build a brighter tomorrow together at BGSW!</p>" & _
            "<p><b>Please note:</b></p>" & _
            "<ol>" & _
            "<li>Your leave as per 'Leave Tool' has been factored in for office presence. If you were on leave for some days in June but have not applied in 'Leave Tool' yet, then please do so before 18th of Jul. Such days would not be factored in for office presence, if not applied yet.</li>" & _
            "<li>Your business travel as per 'Swift Tool' has been factored in for your office presence.</li>" & _
            "<li>You can check your presence in office <a href=""https://banvm034.apac.bosch.com/SwipeInformation/Home/Index""> here</a><br></br></li>" & _
            "<li>If you are required to work from Client Location, attending an External Official Conference / Workshop / Training or BLC, please update 'Business Reason' as reason on 'Leave Tool' to ensure your presence for official work is updated at all times.</li>" & _
            "<li>You are required to apply for planned leave in advance in the Leave Tool.</li>" & _
            "<li>To view the Smart Hybrid Work Policy, please refer to this link <a href=""https://kor2vm0022.apac.bosch.com/stages/#/workspace/705/_vv/process/guidance/_u_FdwAQtPMy2VM9Utg9ZZA""> here</a></li>" & _
            "<li>To view the Smart Hybrid Work Policy FAQs, please refer to this link <a href=""https://kor2vm0022.apac.bosch.com/stages/#/workspace/705/_vv/process/guidance/_jhlrQDDAFbWM2Phx-rK14A""> here</a></li>" & _
            "</ol>"
        
        msg.Display
        'msg.Send
    Next I
End Sub



