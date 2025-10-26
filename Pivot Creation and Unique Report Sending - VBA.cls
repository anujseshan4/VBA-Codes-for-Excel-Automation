VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'Prerequisites: Location Sheet cell H6 (Has the File Directory) and Location Sheet Columns A and B (Department and Email ID)

Sub CreatePivotTable()
    Dim PivotTable As PivotTable
    Dim PivotCache As PivotCache
    Dim PivotRange As Range
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
    Sheets.Add(After:=Sheets("Sheet1")).name = "Pivot Table"
    Set ws2 = ThisWorkbook.Worksheets("Pivot Table")
    Set PivotRange = ws.Range("A1:W20000")
    
    ' Location for Sheet Creation
    Set data_sh = ThisWorkbook.Sheets("Sheet1")
    Set loc_sh = ThisWorkbook.Sheets("Location")
    Set uniqueNames = CreateObject("Scripting.Dictionary") ' Use Scripting Dictionary for unique names
    
    ' Remove Duplicates
    loc_sh.Range("A:A").Clear
    ws.AutoFilterMode = False
    ws.Range("V:V").Copy loc_sh.Range("A1")
    ws.Range("W:W").Copy loc_sh.Range("B1")
    loc_sh.Range("A1:B100000").RemoveDuplicates 1, xlYes
    
    ' Create Pivot Cache
    Set PivotCache = ThisWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=PivotRange)
    Set PivotTable = PivotCache.CreatePivotTable(TableDestination:=ws2.Range("D4"), TableName:="Pivot Table")
    
    With PivotTable
        .PivotFields("Department").Orientation = xlRowField
        .PivotFields("PUI Resp.").Orientation = xlColumnField
        .PivotFields("PO Date").Orientation = xlPageField
        .PivotFields("Supplier").Orientation = xlRowField
        .AddDataField .PivotFields("Quot Item INR"), "Total Sales", xlSum
    End With
    MsgBox "Done"
    Application.ScreenUpdating = False
    
    ' Identifying Unique Name
    lastRow = data_sh.Cells(data_sh.Rows.Count, "V").End(xlUp).Row
    For Each cell In data_sh.Range("V2:V" & lastRow)
        If Not uniqueNames.Exists(cell.Value) Then
            uniqueNames.Add cell.Value, cell.Value
        End If
    Next cell

    ' Loop to copy unique rows to a New File
    Dim name As Variant
    For Each name In uniqueNames.Keys
        Dim newWb As Workbook
        Set newWb = Workbooks.Add
        Dim newSh As Worksheet
        Set newSh = newWb.Sheets(1)

        ' Copy header row from original data sheet
        data_sh.Rows(1).Copy newSh.Rows(1)

        ' Filter and copy data rows matching current 'name'
        Dim destRow As Long
        destRow = 2
        For Each cell In data_sh.Range("V2:V" & lastRow)
            If cell.Value = name Then
                data_sh.Rows(cell.Row).Copy newSh.Rows(destRow)
                destRow = destRow + 1
            End If
        Next cell

        ' Sort data in the new workbook based on Date
        With newSh.Sort
            .SortFields.Clear
            .SortFields.Add Key:=newSh.Range("P:P"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
            .SetRange newSh.UsedRange
            .Header = xlYes
            .Orientation = xlTopToBottom
            .SortMethod = xlPinYin
            .Apply
        End With

        ' Conditional Formatting
        Dim lastRowNewSh As Long
        lastRowNewSh = newSh.Cells(newSh.Rows.Count, "K").End(xlUp).Row
        If lastRowNewSh > 1 Then
            Dim rngDuplicate As Range
            Set rngDuplicate = newSh.Range("K2:K" & lastRowNewSh)
            Set rngDuplicate2 = newSh.Range("N2:N" & lastRowNewSh)
            Set rngDuplicate3 = newSh.Range("P2:P" & lastRowNewSh)
            Set rngDuplicate4 = newSh.Range("S2:S" & lastRowNewSh)

            ' Clear existing Conditional Formatting
            rngDuplicate.FormatConditions.Delete
            rngDuplicate2.FormatConditions.Delete
            rngDuplicate3.FormatConditions.Delete
            rngDuplicate4.FormatConditions.Delete

            ' Apply new conditional formatting to highlight duplicates
            With rngDuplicate.FormatConditions.AddUniqueValues
                .DupeUnique = xlDuplicate
                .Font.Color = -16383844
                .Interior.Color = 13551615
            End With

            With rngDuplicate2.FormatConditions.AddUniqueValues
                .DupeUnique = xlDuplicate
                .Font.Color = -16383844
                .Interior.Color = 13551615
            End With

            With rngDuplicate3.FormatConditions.AddUniqueValues
                .DupeUnique = xlDuplicate
                .Font.Color = -16383844
                .Interior.Color = 13551615
            End With

            With rngDuplicate4.FormatConditions.AddUniqueValues
                .DupeUnique = xlDuplicate
                .Font.Color = -16383844
                .Interior.Color = 13551615
            End With
        End If

        ' Save the new workbook
        folderPath = loc_sh.Range("H6").Value & "\"
        Dim fileName As String
        fileName = name & ".xlsx"
        newWb.SaveAs folderPath & fileName

        ' Create Pivot Table in the new workbook
        Dim ptSheet As Worksheet
        Set ptSheet = newWb.Sheets.Add(After:=newWb.Sheets(newWb.Sheets.Count))
        ptSheet.name = "PivotTable"

        Dim ptCache As PivotCache
        Dim pt As PivotTable
        Dim ptRange As Range
        Dim SlicerCache As SlicerCache
        Dim Slicer As Slicer
        Dim TimelineSlicer As Slicer

        ' Assuming your data starts from column A in the new sheet
        Set ptRange = newSh.Range("A1").CurrentRegion

        ' Create Pivot Cache and Table
        Set ptCache = newWb.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=ptRange)
        Set pt = ptCache.CreatePivotTable(TableDestination:=ptSheet.Range("A3"), TableName:="PivotTable")
        Set SlicerCache = newWb.SlicerCaches.Add(pt, "PUI Resp.")

        ' Add fields to Pivot Table
        With pt
        .PivotFields("Department").Orientation = xlRowField
        .PivotFields("PUI Resp.").Orientation = xlColumnField
        .PivotFields("PO Date").Orientation = xlPageField
        .PivotFields("Supplier").Orientation = xlRowField
        .AddDataField .PivotFields("Quot Item INR"), "Total Value", xlSum
    End With
        
        Set Slicer = SlicerCache.Slicers.Add(ptSheet, , "PUI Resp.", "PUI Resp. Slicer", 100, 100, 200, 200)
        newWb.SlicerCaches.Add2(ptSheet.PivotTables("PivotTable"), "PO Date", , xlTimeline).Slicers.Add ptSheet, , "PO Date", "PO Date", 221.25, 449.25, 262.5, 108
        Slicer.Shape.Width = 150
        Slicer.Shape.Height = 200
        
        On Error Resume Next
        ptSheet.Shapes.Range(Array("PO Date")).Select
        ptSheet.Shapes.AddChart2(201, xlColumnClustered).Select
        ActiveChart.SetSourceData (Range("ptRange"))
        ptSheet.Shapes("Chart 3").IncrementLeft 342.75
        ptSheet.Shapes("Chart 3").IncrementTop -120
        On Error GoTo 0
        
        newWb.Close SaveChanges:=True
    Next name
    
' Emailing
    
    Dim OutApp As Outlook.Application
    Dim msg As Outlook.MailItem
    Dim sharedAccount As Outlook.Account
    Dim attachmentPath As String
    Dim strbody As String

    strbody = "<BODY style='font-size:14pt; font-family:Calibri'>" & _
              "Dear Valued Partners,<p> Greetings from S2C-IN! <p> We’re thrilled to offer you a fantastic opportunity to optimize your purchasing by bundling orders currently processed via the Easy RFQ channel, designed for non-recurring orders. <p>" & _
              "<p> For cumulative orders exceeding INR 5.2 lakh over 1 to 3 years, moving to Request Sourcing or eCatalogue brings you: <p>" & _
              "<li> Better Pricing </li>" & _
              "<li> Stronger Supplier Relationships </li>" & _
              "<li> Faster PO Creation </li>" & _
              "<li> Enhanced Service Levels </li>" & _
              "<li> Full Compliance: This ensures that your recurring or high-volume orders are not only negotiated effectively but also remain fully compliant!</li>" & _
              "<p> Planned Next Steps: </p>" & _
              "<li> Review the attached report (Jan-Aug 2024) for potential split cases.</li>" & _
              "<li> Contact the Category Experts (details in the report) to plan future orders via the Right channels And leverage Price agreements. </li>" & _
              "<p> By Bundling, you'll prevent splits and turn these orders into bundled volumes for maximum benefit! </p>" & _
              "<p> Thanks for your support in optimizing processes! Let's achieve bigger and better Results together! </p>" & _
              "<b>Best Regards, <br>S2C-IN Team<p></BODY>"

    Set OutApp = New Outlook.Application
    For I = 2 To loc_sh.Range("A" & Rows.Count).End(xlUp).Row
        If loc_sh.Range("A" & I).Value <> "" And loc_sh.Range("B" & I).Value <> "-" Then
            Set msg = OutApp.CreateItem(olMailItem)
            
    For Each sharedAccount In OutApp.Session.Accounts
        If sharedAccount.DisplayName = "anujseshan4@gmail.com" Then
            msg.SendUsingAccount = sharedAccount
            Exit For
        End If
    Next sharedAccount
    
            msg.To = loc_sh.Range("B" & I).Value
            msg.Subject = "Easy RFQ - Bundling Potential Cases"
            attachmentPath = loc_sh.Range("H6").Value & "\" & loc_sh.Range("A" & I).Value & ".xlsx"
            Dim AttachmentPath2 As String
            AttachmentPath2 = "C:\Users\ECN3KOR\Desktop\File Explorer\Easy RFQ Split Cases\S2C CE List - IN.xlsx"
            msg.HTMLBody = strbody & "<img src='C:\Users\ECN3KOR\Desktop\File Explorer\Easy RFQ Split Cases\GS Image.png' height='10%'>"
            msg.Attachments.Add attachmentPath
            msg.Attachments.Add AttachmentPath2
            msg.Display
        End If
    Next I
    
    MsgBox "Done"
    Application.ScreenUpdating = True
End Sub


