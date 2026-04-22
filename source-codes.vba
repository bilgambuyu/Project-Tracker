Option Explicit

' =========================
' AUTO INIT CHECK
' =========================
Sub EnsureSetup()

    If Not SheetExists("Data") Or Not TableExists("ProjectTable") Then
        SetupTracker
    End If
    
    ' Ensure we have at least 100 projects
    On Error Resume Next
    If WorksheetFunction.CountA(GetSheet("Data").Range("A:A")) - 1 < 100 Then
        GenerateDummyProjects
    End If
    On Error GoTo 0

End Sub


' =========================
' MAIN SETUP (SAFE)
' =========================
Sub SetupTracker()

    On Error GoTo ErrHandler
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    RecreateSheet "Data"
    RecreateSheet "Dashboard"
    RecreateSheet "Lists"
    
    SetupDataSheet
    SetupLists
    ApplyValidation
    GenerateDummyProjects
    CreateExecutiveDashboard
    
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    
    MsgBox "Executive Dashboard ready with " & WorksheetFunction.CountA(GetSheet("Data").Range("A:A")) - 1 & " projects.", vbInformation
    Exit Sub

ErrHandler:
    MsgBox "Setup error: " & Err.Description, vbCritical
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True

End Sub


' =========================
' GENERATE 100+ DUMMY PROJECTS
' =========================
Sub GenerateDummyProjects()

    Dim ws As Worksheet
    Set ws = GetSheet("Data")
    
    If ws Is Nothing Then Exit Sub
    
    Dim projectNames As Variant
    Dim countries As Variant
    Dim regions As Variant
    Dim statuses As Variant
    Dim managers As Variant
    
    ' Arrays for realistic dummy data
    projectNames = Array("Digital Transformation", "Infrastructure Upgrade", "Cloud Migration", _
                        "AI Implementation", "Security Enhancement", "Mobile App Development", _
                        "Data Center Consolidation", "ERP Integration", "CRM Optimization", _
                        "Network Modernization", "Cybersecurity Initiative", "IoT Deployment", _
                        "Blockchain Solution", "Edge Computing", "5G Implementation", _
                        "Virtual Desktop Infrastructure", "Disaster Recovery", "Compliance Update", _
                        "Training Program", "Quality Assurance", "R&D Initiative", _
                        "Market Expansion", "Product Launch", "Customer Experience", _
                        "Supply Chain Optimization", "Sustainability Project", "Merger Integration", _
                        "Process Automation", "Talent Acquisition", "Brand Refresh")
    
    countries = Array("USA", "Canada", "UK", "Germany", "France", "Japan", "Australia", _
                     "Brazil", "India", "China", "South Africa", "Mexico", "Spain", _
                     "Italy", "Netherlands", "Singapore", "UAE", "Saudi Arabia", "Sweden", _
                     "Norway", "Denmark", "Poland", "Turkey", "Israel", "Chile", "Argentina")
    
    regions = Array("Americas", "Americas", "Europe", "Europe", "Europe", "Asia", "Asia Pacific", _
                   "Americas", "Asia", "Asia", "Africa", "Americas", "Europe", "Europe", _
                   "Europe", "Asia Pacific", "Middle East", "Middle East", "Europe", "Europe", _
                   "Europe", "Europe", "Middle East", "Middle East", "Americas", "Americas")
    
    statuses = Array("Planned", "Ongoing", "Completed", "Delayed")
    managers = Array("John Smith", "Sarah Johnson", "Michael Brown", "Emily Davis", "David Wilson", _
                    "Lisa Anderson", "Robert Taylor", "Maria Garcia", "James Martinez", "Patricia Lee")
    
    Dim i As Long
    Dim startDate As Date
    Dim endDate As Date
    Dim budget As Double
    Dim progress As Integer
    Dim statusIndex As Integer
    
    ' Clear existing data except headers
    On Error Resume Next
    If ws.ListObjects.Count > 0 Then
        ws.ListObjects(1).DataBodyRange.ClearContents
    End If
    On Error GoTo 0
    
    ' Generate 120 projects
    For i = 1 To 120
        ' Project ID
        ws.Cells(i + 1, 1).Value = "PRJ-" & Format(i, "000")
        
        ' Project Name (with variation)
        If i Mod 3 = 0 Then
            ws.Cells(i + 1, 2).Value = projectNames(Int((UBound(projectNames) + 1) * Rnd)) & " " & Year(Date) - Int(Rnd * 3)
        Else
            ws.Cells(i + 1, 2).Value = projectNames(Int((UBound(projectNames) + 1) * Rnd)) & " " & Chr(65 + Int(26 * Rnd))
        End If
        
        ' Country and Region
        Dim countryIndex As Integer
        countryIndex = Int((UBound(countries) + 1) * Rnd)
        ws.Cells(i + 1, 3).Value = countries(countryIndex)
        ws.Cells(i + 1, 4).Value = regions(countryIndex)
        
        ' Status and dates
        statusIndex = Int((UBound(statuses) + 1) * Rnd)
        ws.Cells(i + 1, 5).Value = statuses(statusIndex)
        
        startDate = DateSerial(2023 + Int(Rnd * 3), 1 + Int(Rnd * 12), 1 + Int(Rnd * 28))
        ws.Cells(i + 1, 6).Value = startDate
        
        Select Case statuses(statusIndex)
            Case "Completed"
                endDate = startDate + Int(Rnd * 365) + 30
                progress = 100
            Case "Ongoing"
                endDate = startDate + Int(Rnd * 365) + 180
                progress = 20 + Int(Rnd * 70)
            Case "Delayed"
                endDate = startDate + Int(Rnd * 180) + 90
                progress = 30 + Int(Rnd * 50)
            Case Else ' Planned
                endDate = startDate + Int(Rnd * 365) + 90
                progress = 0 + Int(Rnd * 20)
        End Select
        
        ws.Cells(i + 1, 7).Value = endDate
        
        ' Budget (between $100k and $10M)
        budget = 100000 + (Rnd * 9900000)
        ws.Cells(i + 1, 8).Value = Round(budget, 2)
        
        ' Manager
        ws.Cells(i + 1, 9).Value = managers(Int((UBound(managers) + 1) * Rnd))
        
        ' Progress %
        ws.Cells(i + 1, 10).Value = progress
        
    Next i
    
    ' Refresh table
    On Error Resume Next
    If ws.ListObjects.Count > 0 Then
        ws.ListObjects(1).Resize ws.Range("A1:J" & i + 1)
    End If
    On Error GoTo 0

End Sub


' =========================
' SHEET HELPERS
' =========================
Function SheetExists(name As String) As Boolean
    On Error Resume Next
    SheetExists = Not Worksheets(name) Is Nothing
    On Error GoTo 0
End Function

Function GetSheet(name As String) As Worksheet
    On Error Resume Next
    Set GetSheet = Worksheets(name)
    On Error GoTo 0
    
    ' If sheet doesn't exist, create it
    If GetSheet Is Nothing Then
        On Error Resume Next
        Set GetSheet = Worksheets.Add(After:=Worksheets(Worksheets.Count))
        GetSheet.Name = name
        On Error GoTo 0
    End If
End Function

Sub RecreateSheet(name As String)
    On Error Resume Next
    Worksheets(name).Delete
    On Error GoTo 0
    Worksheets.Add(After:=Worksheets(Worksheets.Count)).Name = name
End Sub


' =========================
' DATA SHEET
' =========================
Sub SetupDataSheet()

    Dim ws As Worksheet
    Set ws = GetSheet("Data")
    
    If ws Is Nothing Then Exit Sub
    
    ws.Cells.Clear
    
    ws.Range("A1:J1").Value = Array("Project ID", "Project Name", "Country", "Region", _
                                   "Status", "Start Date", "End Date", _
                                   "Budget_USD", "Manager", "Progress_Percent")
    ws.Rows(1).Font.Bold = True
    ws.Rows(1).Interior.Color = RGB(68, 114, 196)
    ws.Rows(1).Font.Color = RGB(255, 255, 255)
    
    ws.Columns("A:J").AutoFit
    
    Dim tbl As ListObject
    Set tbl = ws.ListObjects.Add(xlSrcRange, ws.Range("A1:J2"), , xlYes)
    tbl.Name = "ProjectTable"
    tbl.TableStyle = "TableStyleMedium9"

End Sub


' =========================
' LISTS
' =========================
Sub SetupLists()

    Dim ws As Worksheet
    Set ws = GetSheet("Lists")
    
    If ws Is Nothing Then Exit Sub
    
    ws.Cells.Clear
    
    ' Status list
    ws.Range("A1").Value = "Status"
    ws.Range("A1").Font.Bold = True
    ws.Range("A2:A5").Value = Application.Transpose(Array("Planned", "Ongoing", "Completed", "Delayed"))
    
    ' Regions
    ws.Range("B1").Value = "Regions"
    ws.Range("B1").Font.Bold = True
    ws.Range("B2:B7").Value = Application.Transpose(Array("Africa", "Asia", "Europe", "Americas", "Asia Pacific", "Middle East"))
    
    ' Managers list for validation
    ws.Range("C1").Value = "Managers"
    ws.Range("C1").Font.Bold = True
    ws.Range("C2:C11").Value = Application.Transpose(Array("John Smith", "Sarah Johnson", "Michael Brown", _
                                                          "Emily Davis", "David Wilson", "Lisa Anderson", _
                                                          "Robert Taylor", "Maria Garcia", "James Martinez", "Patricia Lee"))
    
    ws.Columns("A:C").AutoFit

End Sub


' =========================
' VALIDATION
' =========================
Sub ApplyValidation()

    Dim ws As Worksheet
    Set ws = GetSheet("Data")
    
    If ws Is Nothing Then Exit Sub
    
    ' Status validation
    On Error Resume Next
    ws.Range("E2:E500").Validation.Delete
    ws.Range("E2:E500").Validation.Add Type:=xlValidateList, Formula1:="=Lists!$A$2:$A$5"
    
    ' Region validation
    ws.Range("D2:D500").Validation.Delete
    ws.Range("D2:D500").Validation.Add Type:=xlValidateList, Formula1:="=Lists!$B$2:$B$7"
    
    ' Manager validation
    ws.Range("I2:I500").Validation.Delete
    ws.Range("I2:I500").Validation.Add Type:=xlValidateList, Formula1:="=Lists!$C$2:$C$11"
    
    ' Progress validation (0-100)
    ws.Range("J2:J500").Validation.Delete
    ws.Range("J2:J500").Validation.Add Type:=xlValidateWholeNumber, _
        AlertStyle:=xlValidAlertStop, _
        Formula1:=0, Formula2:=100
    On Error GoTo 0

End Sub


' =========================
' EXECUTIVE DASHBOARD WITH VISUALIZATIONS
' =========================
Sub CreateExecutiveDashboard()

    Dim ws As Worksheet
    Set ws = GetSheet("Dashboard")
    
    If ws Is Nothing Then Exit Sub
    
    ws.Cells.Clear
    ws.Activate
    
    ' Company Header
    With ws.Range("A1")
        .Value = "🏢 EXECUTIVE PROJECT DASHBOARD"
        .Font.Size = 24
        .Font.Bold = True
        .Font.Color = RGB(0, 51, 102)
    End With
    
    ws.Range("A2").Value = "Real-Time Project Portfolio Overview"
    ws.Range("A2").Font.Size = 12
    ws.Range("A2").Font.Color = RGB(100, 100, 100)
    
    ' Date and Time
    ws.Range("I1").Value = "As of:"
    ws.Range("I1").Font.Bold = True
    ws.Range("J1").Formula = "=NOW()"
    ws.Range("J1").NumberFormat = "mmmm dd, yyyy hh:mm AM/PM"
    ws.Range("J1").Font.Bold = True
    
    ' ===== SECTION 1: KEY METRICS CARDS =====
    ws.Range("A4").Value = "📊 KEY PERFORMANCE INDICATORS"
    ws.Range("A4").Font.Size = 14
    ws.Range("A4").Font.Bold = True
    ws.Range("A4").Interior.Color = RGB(68, 114, 196)
    ws.Range("A4").Font.Color = RGB(255, 255, 255)
    ws.Range("A4:I4").Merge
    
    ' Row 6: Main KPIs
    Dim kpiRow As Long
    kpiRow = 6
    
    ' KPI 1: Total Projects
    With ws.Range("A" & kpiRow & ":C" & kpiRow)
        .Merge
        .Value = "TOTAL PROJECTS"
        .Font.Bold = True
        .Font.Size = 10
        .Interior.Color = RGB(52, 73, 94)
        .Font.Color = RGB(255, 255, 255)
        .HorizontalAlignment = xlCenter
    End With
    
    ws.Range("A" & kpiRow + 1 & ":C" & kpiRow + 1).Merge
    ws.Range("A" & kpiRow + 1).Formula = "=COUNTA(Data!A:A)-1"
    ws.Range("A" & kpiRow + 1).Font.Size = 28
    ws.Range("A" & kpiRow + 1).Font.Bold = True
    ws.Range("A" & kpiRow + 1).HorizontalAlignment = xlCenter
    
    ' KPI 2: Total Budget
    With ws.Range("D" & kpiRow & ":F" & kpiRow)
        .Merge
        .Value = "TOTAL BUDGET"
        .Font.Bold = True
        .Font.Size = 10
        .Interior.Color = RGB(52, 73, 94)
        .Font.Color = RGB(255, 255, 255)
        .HorizontalAlignment = xlCenter
    End With
    
    ws.Range("D" & kpiRow + 1 & ":F" & kpiRow + 1).Merge
    ws.Range("D" & kpiRow + 1).Formula = "=SUM(Data!H:H)"
    ws.Range("D" & kpiRow + 1).NumberFormat = "$#,##0"
    ws.Range("D" & kpiRow + 1).Font.Size = 28
    ws.Range("D" & kpiRow + 1).Font.Bold = True
    ws.Range("D" & kpiRow + 1).HorizontalAlignment = xlCenter
    
    ' KPI 3: Avg Budget/Project
    With ws.Range("G" & kpiRow & ":I" & kpiRow)
        .Merge
        .Value = "AVG BUDGET/PROJECT"
        .Font.Bold = True
        .Font.Size = 10
        .Interior.Color = RGB(52, 73, 94)
        .Font.Color = RGB(255, 255, 255)
        .HorizontalAlignment = xlCenter
    End With
    
    ws.Range("G" & kpiRow + 1 & ":I" & kpiRow + 1).Merge
    ws.Range("G" & kpiRow + 1).Formula = "=IFERROR(AVERAGE(Data!H:H),0)"
    ws.Range("G" & kpiRow + 1).NumberFormat = "$#,##0"
    ws.Range("G" & kpiRow + 1).Font.Size = 28
    ws.Range("G" & kpiRow + 1).Font.Bold = True
    ws.Range("G" & kpiRow + 1).HorizontalAlignment = xlCenter
    
    ' Row 9: Status KPIs
    kpiRow = 9
    
    ' Completed
    With ws.Range("A" & kpiRow & ":C" & kpiRow)
        .Merge
        .Value = "COMPLETED"
        .Font.Bold = True
        .Interior.Color = RGB(46, 204, 113)
        .Font.Color = RGB(255, 255, 255)
        .HorizontalAlignment = xlCenter
    End With
    
    ws.Range("A" & kpiRow + 1 & ":C" & kpiRow + 1).Merge
    ws.Range("A" & kpiRow + 1).Formula = "=COUNTIF(Data!E:E,""Completed"")"
    ws.Range("A" & kpiRow + 1).Font.Size = 20
    ws.Range("A" & kpiRow + 1).Font.Bold = True
    ws.Range("A" & kpiRow + 1).HorizontalAlignment = xlCenter
    
    ' Ongoing
    With ws.Range("D" & kpiRow & ":F" & kpiRow)
        .Merge
        .Value = "ONGOING"
        .Font.Bold = True
        .Interior.Color = RGB(52, 152, 219)
        .Font.Color = RGB(255, 255, 255)
        .HorizontalAlignment = xlCenter
    End With
    
    ws.Range("D" & kpiRow + 1 & ":F" & kpiRow + 1).Merge
    ws.Range("D" & kpiRow + 1).Formula = "=COUNTIF(Data!E:E,""Ongoing"")"
    ws.Range("D" & kpiRow + 1).Font.Size = 20
    ws.Range("D" & kpiRow + 1).Font.Bold = True
    ws.Range("D" & kpiRow + 1).HorizontalAlignment = xlCenter
    
    ' Delayed
    With ws.Range("G" & kpiRow & ":I" & kpiRow)
        .Merge
        .Value = "DELAYED"
        .Font.Bold = True
        .Interior.Color = RGB(231, 76, 60)
        .Font.Color = RGB(255, 255, 255)
        .HorizontalAlignment = xlCenter
    End With
    
    ws.Range("G" & kpiRow + 1 & ":I" & kpiRow + 1).Merge
    ws.Range("G" & kpiRow + 1).Formula = "=COUNTIF(Data!E:E,""Delayed"")"
    ws.Range("G" & kpiRow + 1).Font.Size = 20
    ws.Range("G" & kpiRow + 1).Font.Bold = True
    ws.Range("G" & kpiRow + 1).HorizontalAlignment = xlCenter
    
    ' ===== SECTION 2: CHARTS AND VISUALIZATIONS =====
    Dim chartRow As Long
    chartRow = 13
    
    ws.Range("A" & chartRow).Value = "📈 PERFORMANCE VISUALIZATIONS"
    ws.Range("A" & chartRow).Font.Size = 14
    ws.Range("A" & chartRow).Font.Bold = True
    ws.Range("A" & chartRow).Interior.Color = RGB(68, 114, 196)
    ws.Range("A" & chartRow).Font.Color = RGB(255, 255, 255)
    ws.Range("A" & chartRow & ":I" & chartRow).Merge
    
    ' Create supporting data for charts
    CreateChartData
    
    ' Chart 1: Status Distribution (Donut Chart)
    CreateStatusDonutChart ws, chartRow + 2
    
    ' Chart 2: Regional Budget Allocation (Bar Chart)
    CreateRegionalBudgetChart ws, chartRow + 2
    
    ' Chart 3: Progress Tracker (Gauge-style)
    CreateProgressGauge ws, chartRow + 2
    
    ' Chart 4: Top Projects by Budget
    CreateTopProjectsChart ws, chartRow + 20
    
    ' ===== SECTION 3: REGIONAL SUMMARY TABLE =====
    Dim tableRow As Long
    tableRow = 42
    
    ws.Range("A" & tableRow).Value = "🌍 REGIONAL PERFORMANCE SUMMARY"
    ws.Range("A" & tableRow).Font.Size = 14
    ws.Range("A" & tableRow).Font.Bold = True
    ws.Range("A" & tableRow).Interior.Color = RGB(68, 114, 196)
    ws.Range("A" & tableRow).Font.Color = RGB(255, 255, 255)
    ws.Range("A" & tableRow & ":I" & tableRow).Merge
    
    ' Create summary table
    CreateRegionalSummaryTable ws, tableRow + 1
    
    ' ===== SECTION 4: MANAGER PERFORMANCE =====
    Dim managerRow As Long
    managerRow = 55
    
    ws.Range("A" & managerRow).Value = "👥 TOP PERFORMING MANAGERS"
    ws.Range("A" & managerRow).Font.Size = 14
    ws.Range("A" & managerRow).Font.Bold = True
    ws.Range("A" & managerRow).Interior.Color = RGB(68, 114, 196)
    ws.Range("A" & managerRow).Font.Color = RGB(255, 255, 255)
    ws.Range("A" & managerRow & ":I" & managerRow).Merge
    
    ' Create manager summary
    CreateManagerSummaryTable ws, managerRow + 1
    
    ' ===== SECTION 5: RECENT PROJECTS =====
    Dim recentRow As Long
    recentRow = 68
    
    ws.Range("A" & recentRow).Value = "🔄 RECENT PROJECT ACTIVITY"
    ws.Range("A" & recentRow).Font.Size = 14
    ws.Range("A" & recentRow).Font.Bold = True
    ws.Range("A" & recentRow).Interior.Color = RGB(68, 114, 196)
    ws.Range("A" & recentRow).Font.Color = RGB(255, 255, 255)
    ws.Range("A" & recentRow & ":I" & recentRow).Merge
    
    ' Create recent projects view
    CreateRecentProjectsView ws, recentRow + 1
    
    ' Format the dashboard
    FormatExecutiveDashboard ws
    
    ' Add refresh button
    AddRefreshButton ws
    
    ' Initial data refresh
    RefreshAllData ws

End Sub

Sub CreateChartData()
    
    Dim ws As Worksheet
    Set ws = GetSheet("Dashboard")
    
    If ws Is Nothing Then Exit Sub
    
    ' Create hidden sheet for chart data
    Dim chartDataWs As Worksheet
    On Error Resume Next
    Set chartDataWs = Worksheets("ChartData")
    If chartDataWs Is Nothing Then
        Set chartDataWs = Worksheets.Add
        chartDataWs.Name = "ChartData"
    End If
    On Error GoTo 0
    
    chartDataWs.Cells.Clear
    chartDataWs.Visible = xlSheetVeryHidden
    
    ' Status distribution data
    chartDataWs.Range("A1").Value = "Status"
    chartDataWs.Range("B1").Value = "Count"
    chartDataWs.Range("A2").Value = "Completed"
    chartDataWs.Range("B2").Formula = "=COUNTIF(Data!E:E,""Completed"")"
    chartDataWs.Range("A3").Value = "Ongoing"
    chartDataWs.Range("B3").Formula = "=COUNTIF(Data!E:E,""Ongoing"")"
    chartDataWs.Range("A4").Value = "Delayed"
    chartDataWs.Range("B4").Formula = "=COUNTIF(Data!E:E,""Delayed"")"
    chartDataWs.Range("A5").Value = "Planned"
    chartDataWs.Range("B5").Formula = "=COUNTIF(Data!E:E,""Planned"")"
    
    ' Regional budget data
    chartDataWs.Range("D1").Value = "Region"
    chartDataWs.Range("E1").Value = "Budget"
    chartDataWs.Range("D2").Value = "Americas"
    chartDataWs.Range("E2").Formula = "=SUMIF(Data!D:D,""Americas"",Data!H:H)"
    chartDataWs.Range("D3").Value = "Europe"
    chartDataWs.Range("E3").Formula = "=SUMIF(Data!D:D,""Europe"",Data!H:H)"
    chartDataWs.Range("D4").Value = "Asia"
    chartDataWs.Range("E4").Formula = "=SUMIF(Data!D:D,""Asia"",Data!H:H)"
    chartDataWs.Range("D5").Value = "Asia Pacific"
    chartDataWs.Range("E5").Formula = "=SUMIF(Data!D:D,""Asia Pacific"",Data!H:H)"
    chartDataWs.Range("D6").Value = "Middle East"
    chartDataWs.Range("E6").Formula = "=SUMIF(Data!D:D,""Middle East"",Data!H:H)"
    chartDataWs.Range("D7").Value = "Africa"
    chartDataWs.Range("E7").Formula = "=SUMIF(Data!D:D,""Africa"",Data!H:H)"
    
    ' Overall progress
    chartDataWs.Range("G1").Value = "Metric"
    chartDataWs.Range("H1").Value = "Value"
    chartDataWs.Range("G2").Value = "Overall Progress"
    chartDataWs.Range("H2").Formula = "=AVERAGE(Data!J:J)"
    chartDataWs.Range("G3").Value = "Completion Rate"
    chartDataWs.Range("H3").Formula = "=ROUND(COUNTIF(Data!E:E,""Completed"")/(COUNTA(Data!A:A)-1)*100,1)"

End Sub

Sub CreateStatusDonutChart(ws As Worksheet, topRow As Long)
    
    On Error Resume Next
    
    ' Delete existing chart
    Dim ch As ChartObject
    For Each ch In ws.ChartObjects
        If ch.Name = "StatusChart" Then ch.Delete
    Next ch
    
    ' Create new chart
    Dim chartObj As ChartObject
    Set chartObj = ws.ChartObjects.Add(Left:=50, Width:=350, Top:=topRow * ws.Rows(topRow).Height, Height:=250)
    chartObj.Name = "StatusChart"
    
    With chartObj.Chart
        .SetSourceData Source:=Worksheets("ChartData").Range("A1:B5")
        .ChartType = xlDoughnut
        .HasTitle = True
        .ChartTitle.Text = "Project Status Distribution"
        .ChartTitle.Font.Size = 12
        .ChartTitle.Font.Bold = True
        .ApplyDataLabels
        .SetElement (msoElementDataLabelOutSideEnd)
    End With
    
    On Error GoTo 0

End Sub

Sub CreateRegionalBudgetChart(ws As Worksheet, topRow As Long)
    
    On Error Resume Next
    
    ' Delete existing chart
    Dim ch As ChartObject
    For Each ch In ws.ChartObjects
        If ch.Name = "RegionalChart" Then ch.Delete
    Next ch
    
    ' Create new chart
    Dim chartObj As ChartObject
    Set chartObj = ws.ChartObjects.Add(Left:=450, Width:=400, Top:=topRow * ws.Rows(topRow).Height, Height:=250)
    chartObj.Name = "RegionalChart"
    
    With chartObj.Chart
        .SetSourceData Source:=Worksheets("ChartData").Range("D1:E7")
        .ChartType = xlColumnClustered
        .HasTitle = True
        .ChartTitle.Text = "Budget by Region (USD)"
        .ChartTitle.Font.Size = 12
        .ChartTitle.Font.Bold = True
        .Axes(xlValue).NumberFormat = "$#,##0"
    End With
    
    On Error GoTo 0

End Sub

Sub CreateProgressGauge(ws As Worksheet, topRow As Long)
    
    On Error Resume Next
    
    ' Delete existing chart
    Dim ch As ChartObject
    For Each ch In ws.ChartObjects
        If ch.Name = "ProgressChart" Then ch.Delete
    Next ch
    
    ' Create new chart
    Dim chartObj As ChartObject
    Set chartObj = ws.ChartObjects.Add(Left:=900, Width:=350, Top:=topRow * ws.Rows(topRow).Height, Height:=250)
    chartObj.Name = "ProgressChart"
    
    With chartObj.Chart
        .SetSourceData Source:=Worksheets("ChartData").Range("G1:H3")
        .ChartType = xlColumnClustered
        .HasTitle = True
        .ChartTitle.Text = "Overall Performance Metrics"
        .ChartTitle.Font.Size = 12
        .ChartTitle.Font.Bold = True
    End With
    
    On Error GoTo 0

End Sub

Sub CreateTopProjectsChart(ws As Worksheet, topRow As Long)
    
    On Error Resume Next
    
    ' Create temporary data for top 5 projects
    Dim tempWs As Worksheet
    Set tempWs = Worksheets("ChartData")
    
    tempWs.Range("J1").Value = "Project"
    tempWs.Range("K1").Value = "Budget"
    
    ' Get top 5 projects by budget
    tempWs.Range("J2").Formula = "=INDEX(Data!B:B,MATCH(LARGE(Data!H:H,1),Data!H:H,0))"
    tempWs.Range("K2").Formula = "=LARGE(Data!H:H,1)"
    tempWs.Range("J3").Formula = "=INDEX(Data!B:B,MATCH(LARGE(Data!H:H,2),Data!H:H,0))"
    tempWs.Range("K3").Formula = "=LARGE(Data!H:H,2)"
    tempWs.Range("J4").Formula = "=INDEX(Data!B:B,MATCH(LARGE(Data!H:H,3),Data!H:H,0))"
    tempWs.Range("K4").Formula = "=LARGE(Data!H:H,3)"
    tempWs.Range("J5").Formula = "=INDEX(Data!B:B,MATCH(LARGE(Data!H:H,4),Data!H:H,0))"
    tempWs.Range("K5").Formula = "=LARGE(Data!H:H,4)"
    tempWs.Range("J6").Formula = "=INDEX(Data!B:B,MATCH(LARGE(Data!H:H,5),Data!H:H,0))"
    tempWs.Range("K6").Formula = "=LARGE(Data!H:H,5)"
    
    ' Delete existing chart
    Dim ch As ChartObject
    For Each ch In ws.ChartObjects
        If ch.Name = "TopProjectsChart" Then ch.Delete
    Next ch
    
    ' Create new chart
    Dim chartObj As ChartObject
    Set chartObj = ws.ChartObjects.Add(Left:=50, Width:=1200, Top:=topRow * ws.Rows(topRow).Height, Height:=300)
    chartObj.Name = "TopProjectsChart"
    
    With chartObj.Chart
        .SetSourceData Source:=tempWs.Range("J1:K6")
        .ChartType = xlBarClustered
        .HasTitle = True
        .ChartTitle.Text = "Top 5 Projects by Budget"
        .ChartTitle.Font.Size = 12
        .ChartTitle.Font.Bold = True
        .Axes(xlValue).NumberFormat = "$#,##0"
    End With
    
    On Error GoTo 0

End Sub

Sub CreateRegionalSummaryTable(ws As Worksheet, startRow As Long)
    
    ' Headers
    ws.Range("A" & startRow).Value = "Region"
    ws.Range("B" & startRow).Value = "Total Projects"
    ws.Range("C" & startRow).Value = "Total Budget"
    ws.Range("D" & startRow).Value = "Avg Budget"
    ws.Range("E" & startRow).Value = "Completion Rate"
    ws.Range("F" & startRow).Value = "On-Time Rate"
    
    ws.Range("A" & startRow & ":F" & startRow).Font.Bold = True
    ws.Range("A" & startRow & ":F" & startRow).Interior.Color = RGB(68, 114, 196)
    ws.Range("A" & startRow & ":F" & startRow).Font.Color = RGB(255, 255, 255)
    
    Dim regions As Variant
    regions = Array("Americas", "Europe", "Asia", "Asia Pacific", "Middle East", "Africa")
    
    Dim i As Integer
    For i = 0 To UBound(regions)
        Dim rowNum As Long
        rowNum = startRow + i + 1
        
        ws.Range("A" & rowNum).Value = regions(i)
        ws.Range("B" & rowNum).Formula = "=COUNTIFS(Data!D:D,""" & regions(i) & """)"
        ws.Range("C" & rowNum).Formula = "=SUMIF(Data!D:D,""" & regions(i) & """,Data!H:H)"
        ws.Range("C" & rowNum).NumberFormat = "$#,##0"
        ws.Range("D" & rowNum).Formula = "=IFERROR(AVERAGEIF(Data!D:D,""" & regions(i) & """,Data!H:H),0)"
        ws.Range("D" & rowNum).NumberFormat = "$#,##0"
        ws.Range("E" & rowNum).Formula = "=IFERROR(ROUND(COUNTIFS(Data!D:D,""" & regions(i) & """,Data!E:E,""Completed"")/COUNTIF(Data!D:D,""" & regions(i) & """)*100,1),0) & ""%"""
        ws.Range("F" & rowNum).Formula = "=IFERROR(ROUND(COUNTIFS(Data!D:D,""" & regions(i) & """,Data!E:E,""Ongoing"")/COUNTIF(Data!D:D,""" & regions(i) & """)*100,1),0) & ""%"""
    Next i
    
    ws.Range("A" & startRow & ":F" & startRow + UBound(regions) + 1).Borders.LineStyle = xlContinuous

End Sub

Sub CreateManagerSummaryTable(ws As Worksheet, startRow As Long)
    
    ' Headers
    ws.Range("A" & startRow).Value = "Manager"
    ws.Range("B" & startRow).Value = "Projects"
    ws.Range("C" & startRow).Value = "Total Budget"
    ws.Range("D" & startRow).Value = "Avg Progress"
    ws.Range("E" & startRow).Value = "Completed"
    ws.Range("F" & startRow).Value = "Performance"
    
    ws.Range("A" & startRow & ":F" & startRow).Font.Bold = True
    ws.Range("A" & startRow & ":F" & startRow).Interior.Color = RGB(68, 114, 196)
    ws.Range("A" & startRow & ":F" & startRow).Font.Color = RGB(255, 255, 255)
    
    Dim managers As Variant
    managers = Array("John Smith", "Sarah Johnson", "Michael Brown", "Emily Davis", "David Wilson", _
                    "Lisa Anderson", "Robert Taylor", "Maria Garcia", "James Martinez", "Patricia Lee")
    
    Dim i As Integer
    For i = 0 To UBound(managers)
        Dim rowNum As Long
        rowNum = startRow + i + 1
        
        ws.Range("A" & rowNum).Value = managers(i)
        ws.Range("B" & rowNum).Formula = "=COUNTIF(Data!I:I,""" & managers(i) & """)"
        ws.Range("C" & rowNum).Formula = "=SUMIF(Data!I:I,""" & managers(i) & """,Data!H:H)"
        ws.Range("C" & rowNum).NumberFormat = "$#,##0"
        ws.Range("D" & rowNum).Formula = "=IFERROR(ROUND(AVERAGEIF(Data!I:I,""" & managers(i) & """,Data!J:J),1),0) & ""%"""
        ws.Range("E" & rowNum).Formula = "=COUNTIFS(Data!I:I,""" & managers(i) & """,Data!E:E,""Completed"")"
        ws.Range("F" & rowNum).Formula = "=IFERROR(ROUND(E" & rowNum & "/B" & rowNum & "*100,1),0) & ""%"""
        
        ' Color code performance
        Dim perfValue As Double
        perfValue = 0
        On Error Resume Next
        perfValue = ws.Range("F" & rowNum).Value
        On Error GoTo 0
        
        If perfValue >= 80 Then
            ws.Range("F" & rowNum).Interior.Color = RGB(46, 204, 113)
        ElseIf perfValue >= 60 Then
            ws.Range("F" & rowNum).Interior.Color = RGB(241, 196, 15)
        Else
            ws.Range("F" & rowNum).Interior.Color = RGB(231, 76, 60)
        End If
    Next i
    
    ws.Range("A" & startRow & ":F" & startRow + UBound(managers) + 1).Borders.LineStyle = xlContinuous

End Sub

Sub CreateRecentProjectsView(ws As Worksheet, startRow As Long)
    
    ' Headers
    ws.Range("A" & startRow).Value = "Project ID"
    ws.Range("B" & startRow).Value = "Project Name"
    ws.Range("C" & startRow).Value = "Status"
    ws.Range("D" & startRow).Value = "Progress"
    ws.Range("E" & startRow).Value = "Manager"
    ws.Range("F" & startRow).Value = "Budget"
    
    ws.Range("A" & startRow & ":F" & startRow).Font.Bold = True
    ws.Range("A" & startRow & ":F" & startRow).Interior.Color = RGB(68, 114, 196)
    ws.Range("A" & startRow & ":F" & startRow).Font.Color = RGB(255, 255, 255)
    
    ' Show last 10 projects
    ws.Range("A" & startRow + 1).Formula = "=IFERROR(INDEX(Data!A:A,COUNTA(Data!A:A)-9),"")"
    ws.Range("B" & startRow + 1).Formula = "=IFERROR(INDEX(Data!B:B,COUNTA(Data!A:A)-9),"")"
    ws.Range("C" & startRow + 1).Formula = "=IFERROR(INDEX(Data!E:E,COUNTA(Data!A:A)-9),"")"
    ws.Range("D" & startRow + 1).Formula = "=IFERROR(INDEX(Data!J:J,COUNTA(Data!A:A)-9),"") & ""%"""
    ws.Range("E" & startRow + 1).Formula = "=IFERROR(INDEX(Data!I:I,COUNTA(Data!A:A)-9),"")"
    ws.Range("F" & startRow + 1).Formula = "=IFERROR(INDEX(Data!H:H,COUNTA(Data!A:A)-9),0)"
    ws.Range("F" & startRow + 1).NumberFormat = "$#,##0"
    
    ' Copy down for 10 rows
    ws.Range("A" & startRow + 1 & ":F" & startRow + 1).AutoFill Destination:=ws.Range("A" & startRow + 1 & ":F" & startRow + 10), Type:=xlFillDefault
    
    ws.Range("A" & startRow & ":F" & startRow + 10).Borders.LineStyle = xlContinuous

End Sub

Sub FormatExecutiveDashboard(ws As Worksheet)
    
    ' Auto-fit columns
    ws.Columns("A:I").AutoFit
    
    ' Add conditional formatting for status
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, "C").End(xlUp).Row
    
    ' Format status colors
    ws.Range("C" & 70 & ":C" & lastRow).FormatConditions.Delete
    
    ' Completed - Green
    With ws.Range("C" & 70 & ":C" & lastRow).FormatConditions.Add(xlCellValue, xlEqual, "Completed")
        .Interior.Color = RGB(46, 204, 113)
        .Font.Color = RGB(255, 255, 255)
    End With
    
    ' Ongoing - Blue
    With ws.Range("C" & 70 & ":C" & lastRow).FormatConditions.Add(xlCellValue, xlEqual, "Ongoing")
        .Interior.Color = RGB(52, 152, 219)
        .Font.Color = RGB(255, 255, 255)
    End With
    
    ' Delayed - Red
    With ws.Range("C" & 70 & ":C" & lastRow).FormatConditions.Add(xlCellValue, xlEqual, "Delayed")
        .Interior.Color = RGB(231, 76, 60)
        .Font.Color = RGB(255, 255, 255)
    End With
    
    ' Planned - Orange
    With ws.Range("C" & 70 & ":C" & lastRow).FormatConditions.Add(xlCellValue, xlEqual, "Planned")
        .Interior.Color = RGB(241, 196, 15)
        .Font.Color = RGB(0, 0, 0)
    End With

End Sub

Sub AddRefreshButton(ws As Worksheet)
    
    ' Add a refresh button to the dashboard
    Dim btn As Button
    On Error Resume Next
    ws.Buttons("RefreshButton").Delete
    On Error GoTo 0
    
    Set btn = ws.Buttons.Add(Left:=ws.Range("I3").Left, Top:=ws.Range("I3").Top, Width:=100, Height:=30)
    btn.Name = "RefreshButton"
    btn.Caption = "🔄 Refresh Data"
    btn.OnAction = "RefreshDashboard"
    
    ' Add quick action buttons
    Dim addBtn As Button
    Set addBtn = ws.Buttons.Add(Left:=ws.Range("I4").Left, Top:=ws.Range("I4").Top, Width:=100, Height:=30)
    addBtn.Name = "AddProjectButton"
    addBtn.Caption = "➕ Add Project"
    addBtn.OnAction = "AddProjectForm"
    
    Dim statsBtn As Button
    Set statsBtn = ws.Buttons.Add(Left:=ws.Range("I5").Left, Top:=ws.Range("I5").Top, Width:=100, Height:=30)
    statsBtn.Name = "StatsButton"
    statsBtn.Caption = "📊 Show Stats"
    statsBtn.OnAction = "ShowStatistics"
    
    Dim exportBtn As Button
    Set exportBtn = ws.Buttons.Add(Left:=ws.Range("I6").Left, Top:=ws.Range("I6").Top, Width:=100, Height:=30)
    exportBtn.Name = "ExportButton"
    exportBtn.Caption = "💾 Export Data"
    exportBtn.OnAction = "ExportToCSV"

End Sub

Sub RefreshAllData(ws As Worksheet)
    
    ' Refresh all formulas and charts
    ws.Calculate
    
    ' Refresh chart data
    If SheetExists("ChartData") Then
        Worksheets("ChartData").Calculate
    End If
    
    ' Update timestamp
    ws.Range("J1").Value = Now
    
    ' Refresh all pivot-like tables by recalculating
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
    ' Force refresh of all formulas
    ws.Range("A1:J" & lastRow).Calculate

End Sub


' =========================
' REAL-TIME DATA UPDATE HANDLERS
' =========================

' Worksheet change event to trigger real-time updates
Private Sub Worksheet_Change(ByVal Target As Range)
    
    ' This should be placed in the Data worksheet module
    ' For now, we'll call it from the add/update functions
    
End Sub

Sub UpdateDashboardRealTime()
    
    ' Real-time dashboard update
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    On Error Resume Next
    
    Dim ws As Worksheet
    Set ws = GetSheet("Dashboard")
    
    If Not ws Is Nothing Then
        ' Refresh all formulas
        ws.Calculate
        
        ' Update timestamp
        ws.Range("J1").Value = Now
        
        ' Refresh chart data if exists
        If SheetExists("ChartData") Then
            Worksheets("ChartData").Calculate
        End If
        
        ' Refresh all charts by reactivating them
        Dim ch As ChartObject
        For Each ch In ws.ChartObjects
            ch.Chart.Refresh
        Next ch
    End If
    
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    
End Sub


' =========================
' ADD PROJECT WITH REAL-TIME UPDATE
' =========================
Sub AddProjectForm()

    ' Create a simple input form using InputBox
    Dim projectName As String
    Dim country As String
    Dim status As String
    Dim budget As Double
    Dim manager As String
    
    projectName = InputBox("Enter Project Name:", "Add New Project")
    If projectName = "" Then Exit Sub
    
    country = InputBox("Enter Country:", "Add New Project")
    If country = "" Then Exit Sub
    
    status = InputBox("Enter Status (Planned/Ongoing/Completed/Delayed):", "Add New Project")
    If status = "" Then Exit Sub
    
    budget = InputBox("Enter Budget (USD):", "Add New Project", 100000)
    If budget <= 0 Then Exit Sub
    
    manager = InputBox("Enter Manager Name:", "Add New Project")
    If manager = "" Then Exit Sub
    
    AddProjectRealTime projectName, country, status, budget, manager

End Sub

Sub AddProjectRealTime(pName As String, pCountry As String, pStatus As String, pBudget As Double, pManager As String)

    EnsureSetup
    
    Dim ws As Worksheet
    Set ws = GetSheet("Data")
    
    If ws Is Nothing Then
        MsgBox "Data sheet not found!", vbExclamation
        Exit Sub
    End If
    
    ' Turn off screen updating for performance
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
    
    ' Auto-determine region
    Dim region As String
    region = GetRegionForCountry(pCountry)
    
    ' Auto-generate dates
    Dim startDate As Date
    Dim endDate As Date
    Dim progress As Integer
    
    startDate = Date
    Select Case pStatus
        Case "Completed"
            endDate = startDate + Int(Rnd * 90) + 30
            progress = 100
        Case "Ongoing"
            endDate = startDate + Int(Rnd * 180) + 90
            progress = Int(Rnd * 80) + 10
        Case "Delayed"
            endDate = startDate + Int(Rnd * 90) + 30
            progress = Int(Rnd * 60) + 10
        Case Else ' Planned
            endDate = startDate + Int(Rnd * 180) + 90
            progress = 0
    End Select
    
    ' Add the project
    ws.Cells(lastRow, 1).Value = "PRJ-" & Format(lastRow - 1, "000")
    ws.Cells(lastRow, 2).Value = pName
    ws.Cells(lastRow, 3).Value = pCountry
    ws.Cells(lastRow, 4).Value = region
    ws.Cells(lastRow, 5).Value = pStatus
    ws.Cells(lastRow, 6).Value = startDate
    ws.Cells(lastRow, 7).Value = endDate
    ws.Cells(lastRow, 8).Value = pBudget
    ws.Cells(lastRow, 9).Value = pManager
    ws.Cells(lastRow, 10).Value = progress
    
    ' Refresh table
    On Error Resume Next
    If ws.ListObjects.Count > 0 Then
        ws.ListObjects(1).Resize ws.Range("A1:J" & lastRow)
    End If
    On Error GoTo 0
    
    ' Update dashboard in real-time
    UpdateDashboardRealTime
    
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    
    MsgBox "✅ Project '" & pName & "' added successfully! Dashboard updated.", vbInformation

End Sub


' =========================
' UPDATE PROJECT STATUS (FOR REAL-TIME UPDATES)
' =========================
Sub UpdateProjectStatus()
    
    Dim projectID As String
    Dim newStatus As String
    Dim newProgress As Integer
    
    projectID = InputBox("Enter Project ID to update:", "Update Project Status")
    If projectID = "" Then Exit Sub
    
    ' Find the project
    Dim ws As Worksheet
    Set ws = GetSheet("Data")
    
    If ws Is Nothing Then Exit Sub
    
    Dim foundRow As Long
    foundRow = 0
    
    Dim i As Long
    For i = 2 To ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        If ws.Cells(i, 1).Value = projectID Then
            foundRow = i
            Exit For
        End If
    Next i
    
    If foundRow = 0 Then
        MsgBox "Project ID not found!", vbExclamation
        Exit Sub
    End If
    
    ' Get new status
    newStatus = InputBox("Enter new status (Planned/Ongoing/Completed/Delayed):", "Update Status", ws.Cells(foundRow, 5).Value)
    If newStatus = "" Then Exit Sub
    
    ' Get new progress
    newProgress = InputBox("Enter new progress (0-100):", "Update Progress", ws.Cells(foundRow, 10).Value)
    If newProgress < 0 Or newProgress > 100 Then
        MsgBox "Progress must be between 0 and 100!", vbExclamation
        Exit Sub
    End If
    
    ' Update the values
    Application.ScreenUpdating = False
    
    ws.Cells(foundRow, 5).Value = newStatus
    ws.Cells(foundRow, 10).Value = newProgress
    
    ' If completed, set progress to 100
    If newStatus = "Completed" Then
        ws.Cells(foundRow, 10).Value = 100
        ws.Cells(foundRow, 7).Value = Date ' Set end date to today
    End If
    
    ' Update dashboard in real-time
    UpdateDashboardRealTime
    
    Application.ScreenUpdating = True
    
    MsgBox "✅ Project " & projectID & " updated successfully!", vbInformation

End Sub


' =========================
' REFRESH DASHBOARD (MANUAL TRIGGER)
' =========================
Sub RefreshDashboard()

    UpdateDashboardRealTime
    MsgBox "✅ Dashboard refreshed with latest data!", vbInformation

End Sub


' =========================
' HELPER FUNCTIONS
' =========================
Function GetRegionForCountry(country As String) As String
    
    Dim europeanCountries As Variant
    Dim asianCountries As Variant
    Dim americanCountries As Variant
    Dim africanCountries As Variant
    Dim middleEastCountries As Variant
    
    europeanCountries = Array("UK", "Germany", "France", "Spain", "Italy", "Netherlands", "Sweden", "Norway", "Denmark", "Poland", "Turkey")
    asianCountries = Array("Japan", "China", "India", "Singapore")
    americanCountries = Array("USA", "Canada", "Brazil", "Mexico", "Chile", "Argentina")
    africanCountries = Array("South Africa")
    middleEastCountries = Array("UAE", "Saudi Arabia", "Israel")
    
    Dim c As Variant
    
    For Each c In europeanCountries
        If country = c Then GetRegionForCountry = "Europe": Exit Function
    Next c
    
    For Each c In asianCountries
        If country = c Then GetRegionForCountry = "Asia": Exit Function
    Next c
    
    For Each c In americanCountries
        If country = c Then GetRegionForCountry = "Americas": Exit Function
    Next c
    
    For Each c In africanCountries
        If country = c Then GetRegionForCountry = "Africa": Exit Function
    Next c
    
    For Each c In middleEastCountries
        If country = c Then GetRegionForCountry = "Middle East": Exit Function
    Next c
    
    GetRegionForCountry = "Global"

End Function


' =========================
' EXPORT TO CSV
' =========================
Sub ExportToCSV()

    Dim ws As Worksheet
    Set ws = GetSheet("Data")
    
    If ws Is Nothing Then
        MsgBox "Data sheet not found!", vbExclamation
        Exit Sub
    End If
    
    If ws.ListObjects.Count = 0 Then
        MsgBox "No data to export!", vbExclamation
        Exit Sub
    End If
    
    Dim fileName As String
    fileName = ThisWorkbook.Path & "\ProjectExport_" & Format(Now, "yyyymmdd_hhmmss") & ".csv"
    
    ' Copy data to new workbook
    ws.ListObjects(1).Range.Copy
    Dim newWB As Workbook
    Set newWB = Workbooks.Add
    newWB.Worksheets(1).Range("A1").PasteSpecial xlPasteValues
    
    ' Save as CSV
    Application.DisplayAlerts = False
    newWB.SaveAs fileName, xlCSV
    newWB.Close
    Application.DisplayAlerts = True
    
    MsgBox "Data exported to: " & fileName, vbInformation

End Sub


' =========================
' SHOW STATISTICS
' =========================
Sub ShowStatistics()

    On Error GoTo ErrHandler
    
    ' Ensure setup is complete first
    EnsureSetup
    
    Dim wsData As Worksheet
    
    ' Safely get the Data worksheet
    Set wsData = GetSheet("Data")
    
    If wsData Is Nothing Then
        MsgBox "Data sheet not found. Please run SetupTracker first.", vbExclamation
        Exit Sub
    End If
    
    ' Check if there's data in column A
    Dim lastRow As Long
    lastRow = wsData.Cells(wsData.Rows.Count, 1).End(xlUp).Row
    
    If lastRow <= 1 Then
        MsgBox "No project data found. Please generate projects first.", vbInformation
        Exit Sub
    End If
    
    Dim totalProjects As Long
    Dim completed As Long
    Dim ongoing As Long
    Dim delayed As Long
    Dim planned As Long
    Dim totalBudget As Double
    Dim avgProgress As Double
    
    ' Calculate statistics with error handling for each
    On Error Resume Next
    
    totalProjects = WorksheetFunction.CountA(wsData.Range("A:A")) - 1
    
    ' Use CountIf for accurate counting
    completed = WorksheetFunction.CountIf(wsData.Range("E:E"), "Completed")
    ongoing = WorksheetFunction.CountIf(wsData.Range("E:E"), "Ongoing")
    delayed = WorksheetFunction.CountIf(wsData.Range("E:E"), "Delayed")
    planned = WorksheetFunction.CountIf(wsData.Range("E:E"), "Planned")
    
    totalBudget = WorksheetFunction.Sum(wsData.Range("H:H"))
    avgProgress = WorksheetFunction.Average(wsData.Range("J:J"))
    
    On Error GoTo ErrHandler
    
    ' Check if we have any projects
    If totalProjects = 0 Then
        MsgBox "No projects found in the database. Please add some projects first.", vbInformation
        Exit Sub
    End If
    
    ' Create formatted message
    Dim msg As String
    msg = "📊 PROJECT STATISTICS" & vbCrLf & vbCrLf
    msg = msg & "Total Projects: " & totalProjects & vbCrLf
    msg = msg & "├─ Completed: " & completed & " (" & FormatPercent(completed / totalProjects, 1) & ")" & vbCrLf
    msg = msg & "├─ Ongoing: " & ongoing & " (" & FormatPercent(ongoing / totalProjects, 1) & ")" & vbCrLf
    msg = msg & "├─ Delayed: " & delayed & " (" & FormatPercent(delayed / totalProjects, 1) & ")" & vbCrLf
    msg = msg & "└─ Planned: " & planned & " (" & FormatPercent(planned / totalProjects, 1) & ")" & vbCrLf
    msg = msg & vbCrLf
    msg = msg & "💰 Financial Summary:" & vbCrLf
    msg = msg & "├─ Total Budget: " & FormatCurrency(totalBudget, 0) & vbCrLf
    msg = msg & "└─ Average Budget/Project: " & FormatCurrency(totalBudget / totalProjects, 0) & vbCrLf
    msg = msg & vbCrLf
    msg = msg & "📈 Performance:" & vbCrLf
    msg = msg & "└─ Average Progress: " & Format(avgProgress, "0.0") & "%"
    
    MsgBox msg, vbInformation, "Project Statistics - " & Format(Date, "mmmm dd, yyyy")
    
    Exit Sub

ErrHandler:
    MsgBox "Error calculating statistics: " & Err.Description & vbCrLf & _
           "Please ensure the Data sheet has valid project information.", vbCritical
    Resume Next

End Sub

' Helper function for percentage formatting
Function FormatPercent(value As Double, decimals As Integer) As String
    If value = 0 Then
        FormatPercent = "0%"
    Else
        FormatPercent = Format(value * 100, "0." & String(decimals, "0")) & "%"
    End If
End Function

' Helper function for currency formatting
Function FormatCurrency(value As Double, decimals As Integer) As String
    If value = 0 Then
        FormatCurrency = "$0"
    Else
        FormatCurrency = "$" & Format(value, "#,##" & IIf(decimals > 0, "." & String(decimals, "0"), ""))
    End If
End Function


' =========================
' DELETE ALL PROJECTS
' =========================
Sub DeleteAllProjects()

    If MsgBox("Are you sure you want to delete ALL projects? This cannot be undone!", _
              vbYesNo + vbCritical, "Confirm Delete") = vbYes Then
        
        Dim ws As Worksheet
        Set ws = GetSheet("Data")
        
        If ws Is Nothing Then
            MsgBox "Data sheet not found!", vbExclamation
            Exit Sub
        End If
        
        On Error Resume Next
        If ws.ListObjects.Count > 0 Then
            ws.ListObjects(1).DataBodyRange.ClearContents
            ws.ListObjects(1).Resize ws.Range("A1:J2")
        End If
        On Error GoTo 0
        
        GenerateDummyProjects
        UpdateDashboardRealTime
        
        MsgBox "All projects have been regenerated and dashboard updated!", vbInformation
        
    End If

End Sub


' =========================
' TABLE CHECK
' =========================
Function TableExists(name As String) As Boolean

    Dim ws As Worksheet, t As ListObject
    
    For Each ws In Worksheets
        For Each t In ws.ListObjects
            If t.Name = name Then
                TableExists = True
                Exit Function
            End If
        Next t
    Next ws

End Function