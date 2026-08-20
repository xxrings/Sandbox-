Attribute VB_Name = "JobTrackingImport"
Option Compare Database
Option Explicit

' Importer for OpticalAdmin_FE.accdb.
' Requires linked tables: Consult, JobOrder, ShipmentTracking, ImportAudit.
' Workbook sheets required: CTM, InProcess, Tracking, Shipping.
' It imports the sheets locally, validates them, then updates canonical records.

Private Const STAGE_CTM As String = "stgDailyCTM"
Private Const STAGE_INPROCESS As String = "stgDailyInProcess"
Private Const STAGE_TRACKING As String = "stgDailyTracking"
Private Const STAGE_SHIPPING As String = "stgDailyShipping"

Public Sub InstallJobTrackingImporter()
    Dim db As DAO.Database: Set db = CurrentDb
    On Error Resume Next
    db.Execute "CREATE TABLE ImportRun (ImportRunID AUTOINCREMENT CONSTRAINT PK_ImportRun PRIMARY KEY, ImportStarted DATETIME, SourceFile TEXT(255), Status TEXT(30), CTMRows LONG, InProcessRows LONG, TrackingRows LONG, ShippingRows LONG, Notes LONGTEXT)"
    db.Execute "CREATE TABLE ImportIssue (ImportIssueID AUTOINCREMENT CONSTRAINT PK_ImportIssue PRIMARY KEY, ImportRunID LONG, Severity TEXT(20), SourceSheet TEXT(30), SourceKey TEXT(100), IssueText LONGTEXT)"
    ' Preserve the date that tracking was recorded; ClosedDate remains Date Shipped.
    db.Execute "ALTER TABLE ShipmentTracking ADD COLUMN TrackingDate DATETIME"
    On Error GoTo 0
    MsgBox "Installer complete. Run ImportDailyJobTrackingWorkbook to select and validate a workbook.", vbInformation
End Sub

Public Sub ImportDailyJobTrackingWorkbook()
    Dim sourceFile As String, runID As Long, db As DAO.Database
    sourceFile = PickExcelFile()
    If Len(sourceFile) = 0 Then Exit Sub
    Set db = CurrentDb
    ClearStageTables db
    ImportSheet sourceFile, "CTM", STAGE_CTM
    ImportSheet sourceFile, "InProcess", STAGE_INPROCESS
    ImportSheet sourceFile, "Tracking", STAGE_TRACKING
    ImportSheet sourceFile, "Shipping", STAGE_SHIPPING
    runID = CreateRun(db, sourceFile)
    ValidateImport db, runID
    If CountBlockingIssues(db, runID) = 0 Then
        UpdateRun db, runID, "Validated", "No blocking issues. Review ImportIssue, then run CommitDailyJobTrackingImport."
        MsgBox "Import is staged and validated. Review ImportIssue, then run CommitDailyJobTrackingImport.", vbInformation
    Else
        UpdateRun db, runID, "Needs review", "Blocking issues found. Correct source data and re-import; do not commit."
        MsgBox "Import staged, but blocking issues were found. Open ImportIssue and correct them before committing.", vbExclamation
    End If
End Sub

Public Sub CommitDailyJobTrackingImport()
    Dim db As DAO.Database, runID As Long
    Set db = CurrentDb
    runID = LatestRunID(db)
    If runID = 0 Then Err.Raise vbObjectError + 900, , "No staged import exists."
    If CountBlockingIssues(db, runID) > 0 Then Err.Raise vbObjectError + 901, , "This import has blocking issues. Review ImportIssue first."
    If MsgBox("This will update Consult, JobOrder, and ShipmentTracking. Confirm that you reviewed the staged data and have a backup.", vbYesNo + vbExclamation) <> vbYes Then Exit Sub

    UpsertConsults db
    UpsertJobOrders db
    ApplyInProcessUpdates db
    ApplyTrackingUpdates db
    ApplyShippingUpdates db, runID
    WriteAudit db, runID
    UpdateRun db, runID, "Committed", "Validated daily workbook committed to canonical tables."
    MsgBox "Daily job-tracking import committed. Review ImportAudit and ImportIssue.", vbInformation
End Sub

Private Sub ImportSheet(ByVal sourceFile As String, ByVal sheetName As String, ByVal localTable As String)
    On Error GoTo Failed
    DoCmd.TransferSpreadsheet acImport, acSpreadsheetTypeExcel12Xml, localTable, sourceFile, True, sheetName & "$"
    Exit Sub
Failed:
    Err.Raise vbObjectError + 910, , "Could not import sheet '" & sheetName & "'. It must exist and have headers in row 1. " & Err.Description
End Sub

Private Sub ClearStageTables(ByVal db As DAO.Database)
    Dim tableName As Variant
    For Each tableName In Array(STAGE_CTM, STAGE_INPROCESS, STAGE_TRACKING, STAGE_SHIPPING)
        On Error Resume Next
        db.TableDefs.Delete CStr(tableName)
        On Error GoTo 0
    Next tableName
End Sub

Private Function CreateRun(ByVal db As DAO.Database, ByVal sourceFile As String) As Long
    Dim q As DAO.Recordset
    db.Execute "INSERT INTO ImportRun (ImportStarted, SourceFile, Status, CTMRows, InProcessRows, TrackingRows, ShippingRows) VALUES (Now(), '" & SqlText(sourceFile) & "', 'Staged', " & RowCount(db, STAGE_CTM) & ", " & RowCount(db, STAGE_INPROCESS) & ", " & RowCount(db, STAGE_TRACKING) & ", " & RowCount(db, STAGE_SHIPPING) & ")", dbFailOnError
    Set q = db.OpenRecordset("SELECT @@IDENTITY AS NewID")
    CreateRun = CLng(q!NewID): q.Close
End Function

Private Sub ValidateImport(ByVal db As DAO.Database, ByVal runID As Long)
    Dim consultField As String, poField As String
    consultField = FieldRef(db, STAGE_CTM, "S", "Consult#", "Consult Number", "ConsultNumber")
    If consultField = "Null" Then AddIssue db, runID, "Block", "CTM", "", "Required field missing: Consult# (or Consult Number).": Exit Sub
    AddIssuesForSql db, runID, "Block", "CTM", "Missing consult number", "SELECT 'CTM', '', 'Missing Consult Number' FROM [" & STAGE_CTM & "] AS S WHERE " & consultField & " Is Null"
    AddIssuesForSql db, runID, "Block", "CTM", "Duplicate consult number", "SELECT 'CTM', CStr(" & consultField & "), 'Duplicate Consult Number in CTM sheet' FROM [" & STAGE_CTM & "] AS S GROUP BY " & consultField & " HAVING Count(*) > 1"

    If RowCount(db, STAGE_INPROCESS) > 0 Then
        consultField = FieldRef(db, STAGE_INPROCESS, "S", "Consult#", "Consult Number", "ConsultNumber")
        poField = FieldRef(db, STAGE_INPROCESS, "S", "PO", "PO Number", "PONumber", "PurchaseOrderNumber")
        If consultField = "Null" Or poField = "Null" Then AddIssue db, runID, "Block", "InProcess", "", "InProcess sheet needs both Consult Number and PO Number."
    End If

    If TableExists(db, STAGE_TRACKING) Then
        consultField = FieldRef(db, STAGE_TRACKING, "S", "Consult#", "Consult Number", "ConsultNumber")
        poField = FieldRef(db, STAGE_TRACKING, "S", "PO", "PO Number", "PONumber", "PurchaseOrderNumber")
        If consultField = "Null" And poField = "Null" Then AddIssue db, runID, "Block", "Tracking", "", "Tracking sheet needs Consult Number or PO Number."
    End If
    If TableExists(db, STAGE_SHIPPING) Then
        consultField = FieldRef(db, STAGE_SHIPPING, "S", "Consult#", "Consult Number", "ConsultNumber")
        poField = FieldRef(db, STAGE_SHIPPING, "S", "PO", "P#O#", "PO Number", "PurchaseOrderNumber")
        If consultField = "Null" And poField = "Null" Then AddIssue db, runID, "Block", "Shipping", "", "Shipping sheet needs Consult Number or PO Number."
    End If
End Sub

Private Sub UpsertConsults(ByVal db As DAO.Database)
    Dim cn As String, patient As String, last5 As String, created As String, service As String, status As String, county As String
    cn = FieldRef(db, STAGE_CTM, "S", "Consult#", "Consult Number", "ConsultNumber")
    patient = FieldRef(db, STAGE_CTM, "S", "Patient"): last5 = FieldRef(db, STAGE_CTM, "S", "Last 5", "Last5")
    created = FieldRef(db, STAGE_CTM, "S", "Created Date", "Order Date"): service = FieldRef(db, STAGE_CTM, "S", "Service")
    status = FieldRef(db, STAGE_CTM, "S", "Status"): county = FieldRef(db, STAGE_CTM, "S", "County")
    db.Execute "UPDATE Consult AS C INNER JOIN [" & STAGE_CTM & "] AS S ON C.ConsultNumber=" & cn & " SET C.PatientDisplay=" & patient & ", C.Last5=" & last5 & ", C.CreatedDate=" & created & ", C.ServiceName=" & service & ", C.ConsultStatus=" & status & ", C.County=" & county, dbFailOnError
    db.Execute "INSERT INTO Consult (ConsultNumber,PatientDisplay,Last5,CreatedDate,ServiceName,ConsultStatus,County) SELECT " & cn & "," & patient & "," & last5 & "," & created & "," & service & "," & status & "," & county & " FROM [" & STAGE_CTM & "] AS S LEFT JOIN Consult AS C ON C.ConsultNumber=" & cn & " WHERE C.ConsultNumber Is Null", dbFailOnError
End Sub

Private Sub UpsertJobOrders(ByVal db As DAO.Database)
    Dim cn As String, orderDate As String, locationCode As String, county As String
    cn = FieldRef(db, STAGE_CTM, "S", "Consult#", "Consult Number", "ConsultNumber")
    orderDate = FieldRef(db, STAGE_CTM, "S", "Created Date", "Order Date")
    locationCode = "Left(" & FieldRef(db, STAGE_CTM, "S", "From Service", "Location") & ",3)"
    county = FieldRef(db, STAGE_CTM, "S", "County")
    db.Execute "UPDATE JobOrder AS J INNER JOIN [" & STAGE_CTM & "] AS S ON J.ConsultNumber=" & cn & " SET J.OrderDate=Nz(" & orderDate & ",J.OrderDate), J.LocationCode=Nz(" & locationCode & ",J.LocationCode), J.County=Nz(" & county & ",J.County)", dbFailOnError
    db.Execute "INSERT INTO JobOrder (ConsultNumber,OrderDate,LocationCode,County,JobStatus) SELECT " & cn & "," & orderDate & "," & locationCode & "," & county & ",'Open' FROM [" & STAGE_CTM & "] AS S LEFT JOIN JobOrder AS J ON J.ConsultNumber=" & cn & " WHERE J.JobOrderID Is Null", dbFailOnError
End Sub

Private Sub ApplyInProcessUpdates(ByVal db As DAO.Database)
    Dim cn As String, po As String
    cn = FieldRef(db, STAGE_INPROCESS, "S", "Consult#", "Consult Number", "ConsultNumber")
    po = FieldRef(db, STAGE_INPROCESS, "S", "PO", "PO Number", "PONumber", "PurchaseOrderNumber")
    If cn = "Null" Or po = "Null" Then Exit Sub
    db.Execute "UPDATE JobOrder AS J INNER JOIN [" & STAGE_INPROCESS & "] AS S ON J.ConsultNumber=" & cn & " SET J.PONumber=" & po & " WHERE " & po & " Is Not Null", dbFailOnError
End Sub

Private Sub ApplyTrackingUpdates(ByVal db As DAO.Database)
    Dim cn As String, po As String, tracking As String, trackingDate As String
    cn = FieldRef(db, STAGE_TRACKING, "S", "Consult#", "Consult Number", "ConsultNumber")
    po = FieldRef(db, STAGE_TRACKING, "S", "PO", "PO Number", "PONumber", "PurchaseOrderNumber")
    tracking = FieldRef(db, STAGE_TRACKING, "S", "Tracking Number", "TrackingNumber")
    trackingDate = FieldRef(db, STAGE_TRACKING, "S", "Tracking Date", "TrackingDate", "Date Tracked")
    If tracking = "Null" Then Exit Sub
    db.Execute "INSERT INTO ShipmentTracking (JobOrderID,TrackingNumber,TrackingDate) SELECT J.JobOrderID," & tracking & "," & trackingDate & " FROM [" & STAGE_TRACKING & "] AS S INNER JOIN JobOrder AS J ON (" & cn & " Is Not Null AND J.ConsultNumber=" & cn & ") OR (" & cn & " Is Null AND " & po & " Is Not Null AND J.PONumber=" & po & ") WHERE NOT EXISTS (SELECT * FROM ShipmentTracking AS T WHERE T.JobOrderID=J.JobOrderID AND Nz(T.TrackingNumber,'')=Nz(" & tracking & ",''))", dbFailOnError
End Sub

Private Sub ApplyShippingUpdates(ByVal db As DAO.Database, ByVal runID As Long)
    Dim cn As String, po As String, shipped As String, tracking As String, invoice As String, frame As String, lens As String
    cn = FieldRef(db, STAGE_SHIPPING, "S", "Consult#", "Consult Number", "ConsultNumber")
    po = FieldRef(db, STAGE_SHIPPING, "S", "PO", "P#O#", "PO Number", "PurchaseOrderNumber")
    shipped = FieldRef(db, STAGE_SHIPPING, "S", "Date Shipped", "Shipped Date")
    tracking = FieldRef(db, STAGE_SHIPPING, "S", "Tracking Number", "TrackingNumber")
    invoice = FieldRef(db, STAGE_SHIPPING, "S", "Invoice Number"): frame = FieldRef(db, STAGE_SHIPPING, "S", "Frame Name")
    lens = FieldRef(db, STAGE_SHIPPING, "S", "Lens Description")
    If shipped = "Null" Then AddIssue db, runID, "Block", "Shipping", "", "Required field missing: Date Shipped.": Exit Sub
    db.Execute "INSERT INTO ShipmentTracking (JobOrderID,ShippedDate,TrackingNumber,InvoiceNumber,FrameName,LensDescription) SELECT J.JobOrderID," & shipped & "," & tracking & "," & invoice & "," & frame & "," & lens & " FROM [" & STAGE_SHIPPING & "] AS S INNER JOIN JobOrder AS J ON (" & cn & " Is Not Null AND J.ConsultNumber=" & cn & ") OR (" & cn & " Is Null AND " & po & " Is Not Null AND J.PONumber=" & po & ") WHERE NOT EXISTS (SELECT * FROM ShipmentTracking AS T WHERE T.JobOrderID=J.JobOrderID AND T.ShippedDate=" & shipped & ")", dbFailOnError
    db.Execute "UPDATE JobOrder AS J INNER JOIN [" & STAGE_SHIPPING & "] AS S ON (" & cn & " Is Not Null AND J.ConsultNumber=" & cn & ") OR (" & cn & " Is Null AND " & po & " Is Not Null AND J.PONumber=" & po & ") SET J.ClosedDate=" & shipped & ", J.JobStatus='Closed', J.FrameName=Nz(" & frame & ",J.FrameName), J.LensDescription=Nz(" & lens & ",J.LensDescription) WHERE " & shipped & " Is Not Null", dbFailOnError
End Sub

Private Function FieldRef(ByVal db As DAO.Database, ByVal tableName As String, ByVal aliasName As String, ParamArray choices()) As String
    Dim choice As Variant, f As DAO.Field, wanted As String
    For Each choice In choices
        wanted = CleanHeader(CStr(choice))
        For Each f In db.TableDefs(tableName).Fields
            If CleanHeader(f.Name) = wanted Then
                FieldRef = aliasName & ".[" & Replace(f.Name, "]", "]]" ) & "]"
                Exit Function
            End If
        Next f
    Next choice
    FieldRef = "Null"
End Function

' Access may preserve hidden spaces or normalize punctuation in an imported
' Excel header. Compare normalized names, but return the actual Access name.
Private Function CleanHeader(ByVal value As String) As String
    value = Replace(value, Chr$(160), " ")
    value = Replace(value, " ", "")
    value = Replace(value, "_", "")
    value = Replace(value, "-", "")
    value = Replace(value, "#", "")
    CleanHeader = UCase$(Trim$(value))
End Function

Private Function HasField(ByVal db As DAO.Database, ByVal tableName As String, ByVal fieldName As String) As Boolean
    On Error GoTo Missing
    Dim f As DAO.Field: Set f = db.TableDefs(tableName).Fields(fieldName): HasField = True
Missing:
End Function

Private Function TableExists(ByVal db As DAO.Database, ByVal tableName As String) As Boolean
    On Error GoTo Missing
    Dim t As DAO.TableDef: Set t = db.TableDefs(tableName): TableExists = True
Missing:
End Function

Private Function PickExcelFile() As String
    With Application.FileDialog(3)
        .Title = "Select daily JobTracking workbook": .Filters.Clear: .Filters.Add "Excel files", "*.xlsx;*.xlsm;*.xls"
        If .Show Then PickExcelFile = .SelectedItems(1)
    End With
End Function

Private Function RowCount(ByVal db As DAO.Database, ByVal tableName As String) As Long
    RowCount = CLng(db.OpenRecordset("SELECT Count(*) AS N FROM [" & tableName & "]")!N)
End Function

Private Function LatestRunID(ByVal db As DAO.Database) As Long
    LatestRunID = Nz(db.OpenRecordset("SELECT Max(ImportRunID) AS N FROM ImportRun")!N, 0)
End Function

Private Function CountBlockingIssues(ByVal db As DAO.Database, ByVal runID As Long) As Long
    CountBlockingIssues = CLng(db.OpenRecordset("SELECT Count(*) AS N FROM ImportIssue WHERE ImportRunID=" & runID & " AND Severity='Block'")!N)
End Function

Private Sub AddIssue(ByVal db As DAO.Database, ByVal runID As Long, ByVal severity As String, ByVal sheet As String, ByVal sourceKey As String, ByVal message As String)
    db.Execute "INSERT INTO ImportIssue (ImportRunID,Severity,SourceSheet,SourceKey,IssueText) VALUES (" & runID & ",'" & SqlText(severity) & "','" & SqlText(sheet) & "','" & SqlText(sourceKey) & "','" & SqlText(message) & "')"
End Sub

Private Sub AddIssuesForSql(ByVal db As DAO.Database, ByVal runID As Long, ByVal severity As String, ByVal sheet As String, ByVal message As String, ByVal selectSql As String)
    Dim rs As DAO.Recordset: Set rs = db.OpenRecordset(selectSql)
    Do While Not rs.EOF: AddIssue db, runID, severity, sheet, Nz(rs.Fields(1).Value, ""), Nz(rs.Fields(2).Value, message): rs.MoveNext: Loop
    rs.Close
End Sub

Private Sub UpdateRun(ByVal db As DAO.Database, ByVal runID As Long, ByVal status As String, ByVal notes As String)
    db.Execute "UPDATE ImportRun SET Status='" & SqlText(status) & "', Notes='" & SqlText(notes) & "' WHERE ImportRunID=" & runID
End Sub

Private Sub WriteAudit(ByVal db As DAO.Database, ByVal runID As Long)
    On Error Resume Next
    db.Execute "INSERT INTO ImportAudit (ImportedAt,ImportType,SourceFile,Status,Detail) SELECT Now(),'Daily Job Tracking',SourceFile,'Committed','ImportRunID=' & ImportRunID FROM ImportRun WHERE ImportRunID=" & runID
    On Error GoTo 0
End Sub

Private Function SqlText(ByVal value As String) As String: SqlText = Replace(value, "'", "''"): End Function
