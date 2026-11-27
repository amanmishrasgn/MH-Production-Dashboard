Option Explicit

' ==============================================================================
' POWERPOINT & EXCEL INTERACTIVE REMARKS MACRO
' Compatible with PowerPoint & Excel on Windows & macOS (Office 365 / 2021 / 2019)
'
' PURPOSE:
'   1. PowerPoint Presentation Mode (F5):
'      - Click any date button / hotspot -> Remarks for that chart's commodity pop up.
'      - Oil remarks show on Oil slide, Gas on Gas slide, WI on WI slide.
'      - Compact popup box fits inside top-right graph area without hiding chart.
'      - Bullet points on each remark with only complex name in bold.
'      - Click popup card -> Closes instantly.
'   2. Live Excel & Offline Standalone Fallback:
'      - Reads live from Remarks_Data sheet when Excel workbook is open.
'      - Automatically falls back to embedded shortfall data when Excel is closed.
'
' HOW TO USE IN POWERPOINT:
'   1. Open PowerPoint (.pptm or .pptx)
'   2. Press Alt + F11 (or Fn + Option + F11 on Mac) -> File -> Import File -> select this .bas file
'   3. In PowerPoint, press Alt + F8 -> Select "SetupAllSlides" (or "SetupInteractiveSlide") -> Click "Run"
'   4. Press F5 (Slide Show) and click any date button!
'
' HOW TO USE IN EXCEL:
'   1. Open 100_Days_Actual_vs_Projection.xlsx
'   2. Press Alt + F11 -> File -> Import File -> select this .bas file
'   3. Test with: TestRemarks, TestGasRemarks, TestWIRemarks
' ==============================================================================

' Constants defined explicitly to eliminate any missing reference or Option Explicit compile errors
Private Const msoShapeRectangle As Long = 1
Private Const msoShapeRoundedRectangle As Long = 5
Private Const msoTrue As Long = -1
Private Const msoFalse As Long = 0
Private Const msoBringToFront As Long = 0
Private Const msoAnchorNone As Long = 1
Private Const msoAnchorCenter As Long = 2
Private Const msoAnchorTop As Long = 1
Private Const msoAnchorMiddle As Long = 3
Private Const msoShadow21 As Long = 21
Private Const ppMouseClick As Long = 1
Private Const ppActionRunMacro As Long = 8
Private Const ppViewNormal As Long = 9
Private Const ppViewSlide As Long = 1
Private Const ppAlignLeft As Long = 1
Private Const ppAlignCenter As Long = 2

Private Type RemarkRecord
    DateStr As String
    Commodity As String
    Complex As String
    RemarkText As String
    LossVal As Double
    Unit As String
End Type

Private m_Remarks() As RemarkRecord
Private m_Dates As Variant
Private m_Count As Long
Private m_Initialized As Boolean

' ==============================================================================
' 1. SLIDE SHOW INTERACTION HANDLERS (PowerPoint)
' ==============================================================================

Public Sub ShowDateRemarks(ByVal oShp As Object)
    '
    ' Triggered when any date button or chart hotspot is clicked during Slide Show (F5)
    '
    Dim sld As Object
    Dim dtStr As String
    Dim comm As String
    Dim remarksText As String
    Dim popCard As Object
    Dim chartShp As Object
    Dim chartL As Single, chartT As Single, chartW As Single, chartH As Single
    Dim sldW As Single, sldH As Single
    Dim i As Long

    On Error GoTo ErrHandler

    Set sld = oShp.Parent
    dtStr = GetDateFromShape(oShp)
    If Len(dtStr) = 0 Then Exit Sub

    ' Determine commodity for this specific slide/button (Oil / Gas / WI)
    comm = GetCommodityFromShapeOrSlide(oShp)

    ' Build single-commodity remarks (bullet points only, no section banners)
    remarksText = BuildSingleCommodityRemarks(dtStr, comm)

    ' Get chart bounds to position the card inside top-right of the graph area
    sldW = ActivePresentation.PageSetup.SlideWidth
    sldH = ActivePresentation.PageSetup.SlideHeight

    Set chartShp = Nothing
    On Error Resume Next
    For i = 1 To sld.Shapes.Count
        If sld.Shapes(i).HasChart Then
            Set chartShp = sld.Shapes(i)
            Exit For
        End If
    Next i
    On Error GoTo ErrHandler

    If Not chartShp Is Nothing Then
        chartL = chartShp.Left
        chartT = chartShp.Top
        chartW = chartShp.Width
        chartH = chartShp.Height
    Else
        chartL = sldW * 0.08
        chartT = sldH * 0.18
        chartW = sldW * 0.84
        chartH = sldH * 0.62
    End If

    Set popCard = GetOrCreatePopupCard(sld, chartL, chartT, chartW, chartH)

    ' Put text and apply compact formatting
    With popCard.TextFrame
        .WordWrap = msoTrue
        .MarginLeft = 6: .MarginRight = 6: .MarginTop = 5: .MarginBottom = 5
        .VerticalAnchor = msoAnchorTop
        .TextRange.Text = remarksText
    End With

    ' Format text: Date header bold, complex name at start of each remark bold, rest normal
    Call FormatPopupText(popCard)

    popCard.Visible = msoTrue
    popCard.ZOrder msoBringToFront
    Exit Sub

ErrHandler:
    MsgBox "Remarks Notice: " & Err.Description, vbInformation, "Remarks"
End Sub

Public Sub HideRemarksPopup(Optional ByVal oShp As Object = Nothing)
    '
    ' Dismisses popup card when clicked
    '
    Dim sld As Object
    Dim popCard As Object
    Dim closeBtn As Object

    On Error Resume Next
    If Not oShp Is Nothing Then
        Set sld = oShp.Parent
    Else
        Set sld = GetCurrentSlide()
    End If

    If Not sld Is Nothing Then
        Set popCard = sld.Shapes("Remarks_Popup_Card")
        If Not popCard Is Nothing Then popCard.Visible = msoFalse
        Set closeBtn = sld.Shapes("Remarks_Popup_Close")
        If Not closeBtn Is Nothing Then closeBtn.Visible = msoFalse
    End If
End Sub

Public Sub CloseActivePopup()
    HideRemarksPopup
End Sub

' ==============================================================================
' 2. UNIVERSAL REMARKS FOR DATA POINT (Excel & PowerPoint Compatible)
' ==============================================================================

Public Sub ShowRemarksForDataPoint(Optional ByVal chartTitle As String = "Oil Production", Optional ByVal dateValue As String = "")
    Dim commodity As String
    Dim msg As String
    Dim normDate As String

    Call InitializeEmbeddedData

    If Len(dateValue) = 0 Then
        If IsArray(m_Dates) Then
            If UBound(m_Dates) >= LBound(m_Dates) Then dateValue = CStr(m_Dates(LBound(m_Dates))) Else dateValue = "01-Oct-2026"
        Else
            dateValue = "01-Oct-2026"
        End If
    End If

    ' Determine commodity from chart title
    If InStr(1, chartTitle, "Gas", vbTextCompare) > 0 Then
        commodity = "Gas"
    ElseIf InStr(1, chartTitle, "Water", vbTextCompare) > 0 Or InStr(1, chartTitle, "WI", vbTextCompare) > 0 Then
        commodity = "WI"
    Else
        commodity = "Oil"
    End If

    normDate = NormalizeDate(dateValue)

    ' 1. Live lookup from open Excel workbook
    msg = GetRemarksFromExcel(normDate, commodity)

    ' 2. Fallback to embedded data if Excel is not open
    If Len(msg) = 0 Then
        msg = GetRemarksFromEmbedded(normDate, commodity)
    End If

    If Len(msg) = 0 Then
        msg = "(No " & commodity & " shortfall remarks recorded for " & normDate & ")"
    End If

    MsgBox "Remarks for " & normDate & " (" & commodity & "):" & vbCrLf & vbCrLf & msg, _
           vbInformation, "Daily Remarks — " & normDate
End Sub

' ==============================================================================
' 3. ONE-CLICK TEST MACROS (Run from Alt + F8)
' ==============================================================================

Public Sub TestRemarks()
    TestRemarksOil
End Sub

Public Sub TestRemarksOil()
    Dim msg As String
    msg = BuildSingleCommodityRemarks("01-Oct-2026", "Oil")
    MsgBox msg, vbInformation, "Oil Remarks Preview — 01-Oct-2026"
End Sub

Public Sub TestGasRemarks()
    Dim msg As String
    msg = BuildSingleCommodityRemarks("01-Oct-2026", "Gas")
    MsgBox msg, vbInformation, "Gas Remarks Preview — 01-Oct-2026"
End Sub

Public Sub TestWIRemarks()
    Dim msg As String
    msg = BuildSingleCommodityRemarks("01-Oct-2026", "WI")
    MsgBox msg, vbInformation, "WI Remarks Preview — 01-Oct-2026"
End Sub

' ==============================================================================
' 4. ONE-CLICK SLIDE SETUP MACROS (PowerPoint)
' ==============================================================================

Public Sub SetupInteractiveSlide()
    '
    ' Auto-detects whether active slide is Oil, Gas, or WI and creates interactive buttons
    '
    Dim sld As Object
    Dim comm As String

    Set sld = GetCurrentSlide()
    If sld Is Nothing Then
        MsgBox "Could not identify active slide. Please click on the slide first.", vbExclamation, "No Active Slide"
        Exit Sub
    End If

    comm = DetectSlideCommodity(sld)
    Call SetupSlideForCommodity(sld, comm, True)
End Sub

Public Sub SetupOilSlide()
    Dim sld As Object
    Set sld = GetCurrentSlide()
    If Not sld Is Nothing Then Call SetupSlideForCommodity(sld, "Oil", True)
End Sub

Public Sub SetupGasSlide()
    Dim sld As Object
    Set sld = GetCurrentSlide()
    If Not sld Is Nothing Then Call SetupSlideForCommodity(sld, "Gas", True)
End Sub

Public Sub SetupWISlide()
    Dim sld As Object
    Set sld = GetCurrentSlide()
    If Not sld Is Nothing Then Call SetupSlideForCommodity(sld, "WI", True)
End Sub

Public Sub SetupAllSlides()
    '
    ' Sets up interactive buttons across ALL slides in the presentation in one click!
    '
    Dim pres As Object
    Dim sld As Object
    Dim comm As String
    Dim count As Long

    On Error GoTo ErrHandler
    Set pres = ActivePresentation
    If pres Is Nothing Then Exit Sub

    count = 0
    For Each sld In pres.Slides
        comm = DetectSlideCommodity(sld)
        Call SetupSlideForCommodity(sld, comm, False)
        count = count + 1
    Next sld

    MsgBox "Interactive buttons successfully configured across all " & count & " slides!" & vbCrLf & vbCrLf & _
           "1. Press F5 to start Slide Show mode." & vbCrLf & _
           "2. Slide 1 will show Oil remarks, Slide 2 will show Gas remarks, etc." & vbCrLf & _
           "3. Remarks box is compact inside the graph area with bold complex names.", _
           vbInformation, "Setup Complete — All Slides"
    Exit Sub

ErrHandler:
    MsgBox "Setup Error: " & Err.Description, vbCritical, "Setup Failed"
End Sub

Public Sub SetupSlideForCommodity(ByVal sld As Object, ByVal commodity As String, Optional ByVal showMsg As Boolean = True)
    Dim chartShp As Object
    Dim dates As Variant
    Dim numDates As Long
    Dim i As Long
    Dim chartL As Single, chartT As Single, chartW As Single, chartH As Single
    Dim sldW As Single, sldH As Single
    Dim btnW As Single, btnH As Single, btnGap As Single
    Dim maxCols As Long, colIdx As Long, rowIdx As Long
    Dim btnL As Single, btnT As Single
    Dim themeColor As Long, themeBorder As Long

    On Error GoTo ErrHandler

    ' Tag the slide with its commodity
    On Error Resume Next
    sld.Tags.Delete "Commodity"
    sld.Tags.Add "Commodity", commodity
    On Error GoTo ErrHandler

    sldW = ActivePresentation.PageSetup.SlideWidth
    sldH = ActivePresentation.PageSetup.SlideHeight

    ' Locate chart if present
    Set chartShp = Nothing
    On Error Resume Next
    If ActiveWindow.Selection.ShapeRange.Count > 0 Then
        Set chartShp = ActiveWindow.Selection.ShapeRange(1)
    End If
    If chartShp Is Nothing Then
        For i = 1 To sld.Shapes.Count
            If sld.Shapes(i).HasChart Then
                Set chartShp = sld.Shapes(i)
                Exit For
            End If
        Next i
    End If
    On Error GoTo ErrHandler

    If Not chartShp Is Nothing Then
        chartL = chartShp.Left
        chartT = chartShp.Top
        chartW = chartShp.Width
        chartH = chartShp.Height
    Else
        chartL = sldW * 0.08
        chartT = sldH * 0.18
        chartW = sldW * 0.84
        chartH = sldH * 0.62
    End If

    Call CleanUpOldElements(sld)
    Call InitializeEmbeddedData
    dates = GetAllDates()
    numDates = UBound(dates) - LBound(dates) + 1
    If numDates <= 0 Then Exit Sub

    maxCols = 10
    If numDates < maxCols Then maxCols = numDates
    btnGap = 4
    btnW = (chartW - (maxCols - 1) * btnGap) / maxCols
    If btnW < 45 Then btnW = 45
    If btnW > 70 Then btnW = 70
    btnH = 22

    btnT = chartT + chartH + 10
    If btnT + btnH * 2 > sldH - 10 Then
        btnT = sldH - (btnH * 2 + 15)
    End If

    ' Color theme by commodity
    Select Case UCase(commodity)
        Case "GAS":
            themeColor = RGB(197, 90, 17)     ' Warm Bronze/Orange for Gas
            themeBorder = RGB(140, 60, 10)
        Case "WI":
            themeColor = RGB(46, 117, 182)    ' Cyan/Blue for WI
            themeBorder = RGB(30, 80, 130)
        Case Else:
            themeColor = RGB(31, 78, 120)     ' Deep Blue for Oil
            themeBorder = RGB(20, 50, 85)
    End Select

    For i = LBound(dates) To UBound(dates)
        Dim fullDate As String
        Dim shortLabel As String
        Dim btn As Object

        fullDate = dates(i)
        shortLabel = GetShortDateLabel(fullDate)

        colIdx = (i - LBound(dates)) Mod maxCols
        rowIdx = Int((i - LBound(dates)) / maxCols)
        btnL = chartL + colIdx * (btnW + btnGap)

        Set btn = sld.Shapes.AddShape(msoShapeRoundedRectangle, _
                                     btnL, btnT + rowIdx * (btnH + btnGap), _
                                     btnW, btnH)
        With btn
            .Name = "DateBtn_" & Replace(Replace(fullDate, "-", "_"), " ", "_")
            .AlternativeText = fullDate & "|" & commodity
            .Adjustments.Item(1) = 0.2
            .Fill.Solid
            .Fill.ForeColor.RGB = themeColor
            .Line.ForeColor.RGB = themeBorder
            .Line.Weight = 1

            With .TextFrame
                .WordWrap = msoFalse
                .MarginLeft = 2: .MarginRight = 2: .MarginTop = 1: .MarginBottom = 1
                .VerticalAnchor = msoAnchorMiddle
                With .TextRange
                    .Text = shortLabel
                    .Font.Name = "Calibri"
                    .Font.Size = 9
                    .Font.Bold = msoTrue
                    .Font.Color.RGB = RGB(255, 255, 255)
                    .ParagraphFormat.Alignment = ppAlignCenter
                End With
            End With

            .ActionSettings(ppMouseClick).Action = ppActionRunMacro
            .ActionSettings(ppMouseClick).Run = "ShowDateRemarks"
        End With
    Next i

    Call CreateInitialPopupCard(sld, chartL, chartT, chartW, chartH)

    If showMsg Then
        MsgBox "Interactive Buttons created for " & numDates & " dates (" & commodity & " Chart)!" & vbCrLf & vbCrLf & _
               "1. Press F5 to start Slide Show mode." & vbCrLf & _
               "2. Click any date button to see " & commodity & " shortfall remarks inside the graph." & vbCrLf & _
               "3. Click popup card to close.", _
               vbInformation, "Setup Complete — " & commodity
    End If
    Exit Sub

ErrHandler:
    MsgBox "Setup Error: " & Err.Description, vbCritical, "Setup Failed"
End Sub

Public Sub SetupInteractiveHotspotsOnChart()
    '
    ' Creates transparent click hotspots directly overlaying each bar on the chart
    '
    Dim sld As Object
    Dim chartShp As Object
    Dim dates As Variant
    Dim numDates As Long
    Dim i As Long
    Dim chartL As Single, chartT As Single, chartW As Single, chartH As Single
    Dim colW As Single
    Dim comm As String

    On Error GoTo ErrHandler

    Set sld = GetCurrentSlide()
    If sld Is Nothing Then Exit Sub

    comm = DetectSlideCommodity(sld)

    If ActiveWindow.Selection.ShapeRange.Count > 0 Then
        Set chartShp = ActiveWindow.Selection.ShapeRange(1)
    Else
        For i = 1 To sld.Shapes.Count
            If sld.Shapes(i).HasChart Then
                Set chartShp = sld.Shapes(i)
                Exit For
            End If
        Next i
    End If

    If chartShp Is Nothing Then
        MsgBox "Please select or place a chart on this slide first!", vbExclamation, "No Chart Found"
        Exit Sub
    End If

    chartL = chartShp.Left
    chartT = chartShp.Top
    chartW = chartShp.Width
    chartH = chartShp.Height

    Call CleanUpOldElements(sld)
    Call InitializeEmbeddedData
    dates = GetAllDates()
    numDates = UBound(dates) - LBound(dates) + 1

    colW = chartW / numDates

    For i = LBound(dates) To UBound(dates)
        Dim fullDate As String
        Dim hotspot As Object
        fullDate = dates(i)

        Set hotspot = sld.Shapes.AddShape(msoShapeRectangle, _
                                         chartL + (i - LBound(dates)) * colW, _
                                         chartT + chartH * 0.15, _
                                         colW, chartH * 0.75)
        With hotspot
            .Name = "Hotspot_" & Replace(Replace(fullDate, "-", "_"), " ", "_")
            .AlternativeText = fullDate & "|" & comm
            .Fill.Solid
            .Fill.ForeColor.RGB = RGB(0, 112, 192)
            .Fill.Transparency = 1#
            .Line.Visible = msoFalse

            .ActionSettings(ppMouseClick).Action = ppActionRunMacro
            .ActionSettings(ppMouseClick).Run = "ShowDateRemarks"
        End With
    Next i

    Call CreateInitialPopupCard(sld, chartL, chartT, chartW, chartH)

    MsgBox "Transparent hotspots created over " & comm & " chart columns!" & vbCrLf & _
           "In Slide Show mode (F5), clicking anywhere on a chart bar will pop up " & comm & " remarks.", _
           vbInformation, "Hotspots Active — " & comm
    Exit Sub

ErrHandler:
    MsgBox "Hotspot Error: " & Err.Description, vbCritical, "Error"
End Sub

Public Sub RemoveInteractiveElements()
    Dim sld As Object
    On Error Resume Next
    Set sld = GetCurrentSlide()
    If Not sld Is Nothing Then
        Call CleanUpOldElements(sld)
        MsgBox "Interactive elements removed from current slide.", vbInformation, "Cleaned Up"
    End If
End Sub

' ==============================================================================
' 5. REMARKS COMPILATION & COMMODITY FILTERING LOGIC
' ==============================================================================

Public Function BuildSingleCommodityRemarks(ByVal targetDateStr As String, ByVal commodity As String) As String
    '
    ' Builds clean, bulleted remarks for ONE commodity only (Oil / Gas / WI)
    ' Reads exclusively from Excel Remarks_Data sheet - no embedded fallback
    '
    Dim normDate As String
    Dim remTxt As String
    Dim msg As String

    normDate = NormalizeDate(targetDateStr)

    ' Read remarks ONLY from Excel Remarks_Data sheet
    remTxt = GetRemarksFromExcel(normDate, commodity)

    ' Header line: Just the date (no section banners)
    msg = normDate & vbCrLf

    If Len(remTxt) > 0 Then
        msg = msg & remTxt
    Else
        msg = msg & ChrW(8226) & " No remarks recorded for this date" & vbCrLf
    End If

    BuildSingleCommodityRemarks = msg
End Function

Public Function BuildMultiCommodityRemarks(ByVal targetDateStr As String) As String
    Dim normDate As String
    Dim oilRem As String, gasRem As String, wiRem As String
    Dim msg As String

    normDate = NormalizeDate(targetDateStr)

    ' Read remarks ONLY from Excel Remarks_Data sheet - no embedded fallback
    oilRem = GetRemarksFromExcel(normDate, "Oil")
    gasRem = GetRemarksFromExcel(normDate, "Gas")
    wiRem = GetRemarksFromExcel(normDate, "WI")

    msg = normDate & vbCrLf & vbCrLf
    If Len(oilRem) > 0 Then msg = msg & oilRem Else msg = msg & ChrW(8226) & " No remarks recorded for Oil" & vbCrLf
    If Len(gasRem) > 0 Then msg = msg & gasRem Else msg = msg & ChrW(8226) & " No remarks recorded for Gas" & vbCrLf
    If Len(wiRem) > 0 Then msg = msg & wiRem Else msg = msg & ChrW(8226) & " No remarks recorded for WI" & vbCrLf

    BuildMultiCommodityRemarks = msg
End Function

Private Function GetRemarksFromEmbedded(ByVal targetDate As String, ByVal commodity As String) As String
    Dim i As Long
    Dim res As String
    Dim prefix As String
    Dim lossStr As String

    res = ""
    For i = 0 To m_Count - 1
        With m_Remarks(i)
            If UCase(.DateStr) = UCase(targetDate) And UCase(.Commodity) = UCase(commodity) Then
                If Len(.RemarkText) > 0 Then
                    If .LossVal <> 0 Then
                        If UCase(commodity) = "GAS" Then
                            lossStr = " (" & Format(.LossVal, "+0.000;-0.000;0") & " " & .Unit & ")"
                        Else
                            lossStr = " (" & Format(.LossVal, "+#,##0;-#,##0;0") & " " & .Unit & ")"
                        End If
                    Else
                        lossStr = ""
                    End If

                    If Len(.Complex) > 0 And UCase(.Complex) <> "NONE" And UCase(.Complex) <> "MH TOTAL" Then
                        prefix = .Complex & ": "
                    Else
                        prefix = ""
                    End If

                    res = res & ChrW(8226) & " " & prefix & .RemarkText & lossStr & vbCrLf
                End If
            End If
        End With
    Next i

    GetRemarksFromEmbedded = res
End Function

Private Function GetRemarksFromExcel(ByVal targetDate As String, ByVal commodity As String) As String
    Dim xlApp As Object
    Dim wb As Object
    Dim ws As Object
    Dim lastRow As Long
    Dim i As Long
    Dim res As String

    On Error GoTo ErrExcel

    Set xlApp = GetObject(, "Excel.Application")
    If xlApp Is Nothing Then Exit Function

    Set wb = Nothing
    On Error Resume Next
    Set wb = xlApp.Workbooks("100_Days_Actual_vs_Projection.xlsx")
    If wb Is Nothing Then Set wb = xlApp.Workbooks("100_Days_Actual_vs_Projection.xlsm")
    If wb Is Nothing Then Set wb = xlApp.ActiveWorkbook
    On Error GoTo ErrExcel

    If wb Is Nothing Then Exit Function

    Set ws = Nothing
    On Error Resume Next
    Set ws = wb.Worksheets("Remarks_Data")
    On Error GoTo ErrExcel
    If ws Is Nothing Then Exit Function

    lastRow = ws.Cells(ws.Rows.Count, 1).End(-4162).Row ' -4162 = xlUp
    res = ""

    For i = 3 To lastRow
        Dim cellDate As String
        Dim cellComm As String
        cellDate = Trim(CStr(ws.Cells(i, 1).Value))
        cellComm = Trim(CStr(ws.Cells(i, 2).Value))

        If UCase(cellDate) = UCase(targetDate) And UCase(cellComm) = UCase(commodity) Then
            Dim compName As String, remTxt As String, lossVal As Variant, unitTxt As String
            compName = Trim(CStr(ws.Cells(i, 3).Value))
            remTxt = Trim(CStr(ws.Cells(i, 4).Value))
            lossVal = ws.Cells(i, 5).Value
            unitTxt = Trim(CStr(ws.Cells(i, 6).Value))

            If Len(remTxt) > 0 Then
                Dim lStr As String
                If IsNumeric(lossVal) And lossVal <> 0 Then
                    If UCase(commodity) = "GAS" Then
                        lStr = " (" & Format(CDbl(lossVal), "+0.000;-0.000;0") & " " & unitTxt & ")"
                    Else
                        lStr = " (" & Format(CDbl(lossVal), "+#,##0;-#,##0;0") & " " & unitTxt & ")"
                    End If
                Else
                    lStr = ""
                End If

                Dim pfx As String
                If Len(compName) > 0 And UCase(compName) <> "NONE" And UCase(compName) <> "MH TOTAL" Then
                    pfx = compName & ": "
                Else
                    pfx = ""
                End If

                res = res & ChrW(8226) & " " & pfx & remTxt & lStr & vbCrLf
            End If
        End If
    Next i

    GetRemarksFromExcel = res
    Exit Function

ErrExcel:
    GetRemarksFromExcel = ""
End Function

Private Function GetCommodityFromShapeOrSlide(ByVal oShp As Object) As String
    Dim alt As String
    Dim parts As Variant
    Dim sld As Object

    On Error Resume Next
    alt = oShp.AlternativeText
    If InStr(alt, "|") > 0 Then
        parts = Split(alt, "|")
        If UBound(parts) >= 1 Then
            If Len(parts(1)) > 0 Then
                GetCommodityFromShapeOrSlide = parts(1)
                Exit Function
            End If
        End If
    End If

    Set sld = oShp.Parent
    GetCommodityFromShapeOrSlide = DetectSlideCommodity(sld)
End Function

Public Function DetectSlideCommodity(ByVal sld As Object) As String
    Dim shp As Object
    Dim txt As String
    Dim i As Long
    Dim sIdx As Long, sName As String

    On Error Resume Next

    ' 1. Check slide tag if previously assigned
    If Len(sld.Tags("Commodity")) > 0 Then
        DetectSlideCommodity = sld.Tags("Commodity")
        Exit Function
    End If

    ' 2. Inspect charts on slide for Title or Series Name
    For i = 1 To sld.Shapes.Count
        Set shp = sld.Shapes(i)
        If shp.HasChart Then
            If shp.Chart.HasTitle Then
                txt = UCase(shp.Chart.ChartTitle.Text)
                If InStr(txt, "GAS") > 0 Then
                    DetectSlideCommodity = "Gas": Exit Function
                ElseIf InStr(txt, "WATER") > 0 Or InStr(txt, "WI") > 0 Or InStr(txt, "INJECTION") > 0 Then
                    DetectSlideCommodity = "WI": Exit Function
                ElseIf InStr(txt, "OIL") > 0 Then
                    DetectSlideCommodity = "Oil": Exit Function
                End If
            End If

            For sIdx = 1 To shp.Chart.SeriesCollection.Count
                sName = UCase(shp.Chart.SeriesCollection(sIdx).Name)
                If InStr(sName, "GAS") > 0 Then
                    DetectSlideCommodity = "Gas": Exit Function
                ElseIf InStr(sName, "WATER") > 0 Or InStr(sName, "WI") > 0 Then
                    DetectSlideCommodity = "WI": Exit Function
                ElseIf InStr(sName, "OIL") > 0 Then
                    DetectSlideCommodity = "Oil": Exit Function
                End If
            Next sIdx
        End If
    Next i

    ' 3. Inspect slide text shapes
    For i = 1 To sld.Shapes.Count
        Set shp = sld.Shapes(i)
        If shp.HasTextFrame Then
            txt = UCase(shp.TextFrame.TextRange.Text)
            If InStr(txt, "GAS PRODUCTION") > 0 Or InStr(txt, "GAS ACTUAL") > 0 Then
                DetectSlideCommodity = "Gas": Exit Function
            ElseIf InStr(txt, "WATER INJECTION") > 0 Or InStr(txt, "WI PRODUCTION") > 0 Then
                DetectSlideCommodity = "WI": Exit Function
            ElseIf InStr(txt, "OIL PRODUCTION") > 0 Or InStr(txt, "OIL ACTUAL") > 0 Then
                DetectSlideCommodity = "Oil": Exit Function
            End If
        End If
    Next i

    ' Default to Oil
    DetectSlideCommodity = "Oil"
End Function

Private Function GetDateFromShape(ByVal shp As Object) As String
    Dim tag As String
    Dim shpName As String
    Dim parts As Variant

    On Error Resume Next
    tag = shp.AlternativeText
    If Len(tag) > 0 Then
        If InStr(tag, "|") > 0 Then
            parts = Split(tag, "|")
            GetDateFromShape = NormalizeDate(CStr(parts(0)))
            Exit Function
        Else
            GetDateFromShape = NormalizeDate(tag)
            Exit Function
        End If
    End If

    shpName = shp.Name
    If Left(shpName, 8) = "DateBtn_" Or Left(shpName, 8) = "Hotspot_" Then
        Dim dtPart As String
        dtPart = Mid(shpName, 9)
        dtPart = Replace(dtPart, "_", "-")
        GetDateFromShape = NormalizeDate(dtPart)
        Exit Function
    End If

    GetDateFromShape = ""
End Function

Private Function NormalizeDate(ByVal rawStr As String) As String
    Dim cleaned As String
    cleaned = Trim(rawStr)

    If InStr(cleaned, "-") > 0 Or InStr(cleaned, "/") > 0 Or InStr(cleaned, ".") > 0 Then
        If IsDate(cleaned) Then
            NormalizeDate = Format(CDate(cleaned), "dd-mmm-yyyy")
            Exit Function
        End If
    End If

    Call InitializeEmbeddedData
    If IsArray(m_Dates) Then
        Dim i As Long
        For i = LBound(m_Dates) To UBound(m_Dates)
            If InStr(1, m_Dates(i), cleaned, vbTextCompare) > 0 Or _
               InStr(1, cleaned, m_Dates(i), vbTextCompare) > 0 Then
                NormalizeDate = CStr(m_Dates(i))
                Exit Function
            End If
        Next i
    End If

    NormalizeDate = cleaned
End Function

Private Function GetShortDateLabel(ByVal fullDateStr As String) As String
    If IsDate(fullDateStr) Then
        GetShortDateLabel = Format(CDate(fullDateStr), "dd-mmm")
    Else
        Dim parts As Variant
        parts = Split(fullDateStr, "-")
        If UBound(parts) >= 1 Then
            GetShortDateLabel = parts(0) & "-" & parts(1)
        Else
            GetShortDateLabel = fullDateStr
        End If
    End If
End Function

' ==============================================================================
' 6. COMPACT POPUP CARD & SELECTIVE BOLD FORMATTING
' ==============================================================================

Private Function GetOrCreatePopupCard(ByVal sld As Object, ByVal chartL As Single, ByVal chartT As Single, ByVal chartW As Single, ByVal chartH As Single) As Object
    Dim pop As Object
    Dim popW As Single, popH As Single, popL As Single, popT As Single

    On Error Resume Next
    Set pop = sld.Shapes("Remarks_Popup_Card")
    On Error GoTo 0

    ' Compact size: width ~ 290 pt, height ~ 130 pt
    ' Positioned inside the chart area in the top-right corner so chart lines and bars remain visible
    popW = chartW * 0.38
    If popW > 320 Then popW = 320
    If popW < 260 Then popW = 260

    ' Increased height to accommodate more remarks with reduced line spacing
    popH = chartH * 0.42
    If popH > 180 Then popH = 180
    If popH < 110 Then popH = 110

    popL = chartL + chartW - popW - 12
    popT = chartT + 18

    If pop Is Nothing THEN
        Set pop = sld.Shapes.AddShape(msoShapeRoundedRectangle, popL, popT, popW, popH)
        With pop
            .Name = "Remarks_Popup_Card"
            .Adjustments.Item(1) = 0.04
            .Fill.Solid
            .Fill.ForeColor.RGB = RGB(255, 255, 255)
            .Line.ForeColor.RGB = RGB(31, 78, 120)
            .Line.Weight = 1.25

            .Shadow.Type = msoShadow21
            .Shadow.Visible = msoTrue
            .Shadow.ForeColor.RGB = RGB(160, 160, 160)

            With .TextFrame
                .WordWrap = msoTrue
                .MarginLeft = 6: .MarginRight = 6: .MarginTop = 5: .MarginBottom = 5
                .VerticalAnchor = msoAnchorTop
            End With

            .ActionSettings(ppMouseClick).Action = ppActionRunMacro
            .ActionSettings(ppMouseClick).Run = "HideRemarksPopup"
        End With
    Else
        pop.Left = popL
        pop.Top = popT
        pop.Width = popW
        pop.Height = popH
    End If

    Set GetOrCreatePopupCard = pop
End Function

Private Sub FormatPopupText(ByVal popCard As Object)
    '
    ' Applies styling:
    ' 1. Date header on line 1 is bold and deep blue.
    ' 2. Remark lines start with bullet points.
    ' 3. ONLY the complex prefix (e.g. "• MHN:") is bold, rest of the remark is normal font.
    '
    Dim p As Long
    Dim para As Object
    Dim txt As String
    Dim colPos As Long

    On Error Resume Next
    With popCard.TextFrame.TextRange
        For p = 1 To .Paragraphs.Count
            Set para = .Paragraphs(p)
            txt = para.Text

            If p = 1 Then
                ' Header line: Date, LEFT-ALIGNED, compact spacing
                para.Font.Name = "Calibri"
                para.Font.Size = 7
                para.Font.Bold = msoTrue
                para.Font.Color.RGB = RGB(0, 0, 0)
                para.ParagraphFormat.Alignment = ppAlignLeft
                para.ParagraphFormat.LineSpacingRule = 2  ' msoLineSpaceExact
                para.ParagraphFormat.LineSpacing = 8#
                para.ParagraphFormat.SpaceBefore = 0
                para.ParagraphFormat.SpaceAfter = 0
            Else
                ' Remark line: 6pt Calibri, LEFT-ALIGNED, tight exact spacing
                para.Font.Name = "Calibri"
                para.Font.Size = 6
                para.Font.Bold = msoFalse
                para.Font.Color.RGB = RGB(0, 0, 0)
                para.ParagraphFormat.Alignment = ppAlignLeft
                para.ParagraphFormat.LineSpacingRule = 2  ' msoLineSpaceExact
                para.ParagraphFormat.LineSpacing = 8#
                para.ParagraphFormat.SpaceBefore = 0
                para.ParagraphFormat.SpaceAfter = 0

                ' Bold ONLY the complex name in the starting (e.g. "• MHN:")
                colPos = InStr(txt, ":")
                If colPos > 0 And colPos <= 18 Then
                    para.Characters(1, colPos).Font.Bold = msoTrue
                    para.Characters(1, colPos).Font.Color.RGB = RGB(20, 50, 85)
                End If
            End If
        Next p
    End With
    On Error GoTo 0
End Sub

Private Sub CreateInitialPopupCard(ByVal sld As Object, ByVal chartL As Single, ByVal chartT As Single, ByVal chartW As Single, ByVal chartH As Single)
    Dim pop As Object
    Set pop = GetOrCreatePopupCard(sld, chartL, chartT, chartW, chartH)
    pop.Visible = msoFalse
End Sub

Private Sub CleanUpOldElements(ByVal sld As Object)
    Dim i As Long
    Dim shp As Object

    For i = sld.Shapes.Count To 1 Step -1
        Set shp = sld.Shapes(i)
        If Left(shp.Name, 8) = "DateBtn_" Or _
           Left(shp.Name, 8) = "Hotspot_" Or _
           shp.Name = "Remarks_Popup_Card" Or _
           shp.Name = "Remarks_Popup_Close" Then
            shp.Delete
        End If
    Next i
End Sub

Private Function GetCurrentSlide() As Object
    On Error Resume Next
    If SlideShowWindows.Count > 0 Then
        Set GetCurrentSlide = SlideShowWindows(1).View.Slide
        Exit Function
    End If
    If ActiveWindow.ViewType = ppViewSlide Or ActiveWindow.ViewType = ppViewNormal Then
        Set GetCurrentSlide = ActiveWindow.View.Slide
        Exit Function
    End If
    If ActivePresentation.Slides.Count > 0 Then
        Set GetCurrentSlide = ActivePresentation.Slides(1)
        Exit Function
    End If
    On Error GoTo 0
End Function

Private Function GetAllDates() As Variant
    Call InitializeEmbeddedData
    GetAllDates = m_Dates
End Function

' ==============================================================================
' 7. EMBEDDED DATA INITIALIZER
' ==============================================================================

Private Sub InitializeEmbeddedData()
    If m_Initialized Then Exit Sub

    m_Dates = Array("01-Oct-2026", "02-Oct-2026", "03-Oct-2026", "04-Oct-2026", "05-Oct-2026", "06-Oct-2026", "07-Oct-2026", "08-Oct-2026", "09-Oct-2026", "10-Oct-2026", "11-Oct-2026", "12-Oct-2026", "13-Oct-2026", "14-Oct-2026", "15-Oct-2026", "16-Oct-2026", "17-Oct-2026", "18-Oct-2026", "19-Oct-2026", "20-Oct-2026", "21-Oct-2026", "22-Oct-2026", "23-Oct-2026", "24-Oct-2026", "25-Oct-2026")
    m_Count = 0
    ReDim m_Remarks(0 To 392)

    AddRec "01-Oct-2026", "Oil", "MHN", "NA#4yH: SSSV control line leakage.Wireline cargo picked up at NA Plat...", -158.000#, "BOPD"
    AddRec "01-Oct-2026", "Oil", "NQO", "Separator V-1220 taken offline for replacement of ILT lower tapping...", -60.000#, "BOPD"
    AddRec "01-Oct-2026", "Oil", "NQO", "NX#6: clsd for EMG POOH by wireline team for 1 Hr", -10.000#, "BOPD"
    AddRec "01-Oct-2026", "Oil", "BHS", "SP: Plat SD occurred low pneumatics in SDP", -150.000#, "BOPD"
    AddRec "01-Oct-2026", "Oil", "BHS", "12"" SE-MSP WFL leakage rectification job by MSV SEAMEC AGASTYA: Vessel...", -450.000#, "BOPD"
    AddRec "01-Oct-2026", "Oil", "BHS", "IH: WF cargo flowing via 14"" IH-SH WF since . Less liquid receipt at...", -250.000#, "BOPD"
    AddRec "01-Oct-2026", "Oil", "WO-16", "WO-16A#3/#5/#9 , WO-5 #1 & #3 & B-121A#1 clsd for Solvent treatment job...", -1773.000#, "BOPD"
    AddRec "01-Oct-2026", "Oil", "ICP", "RS12#12 (O): flowing with less temperature, Observed annulus-tubing...", -104.000#, "BOPD"
    AddRec "01-Oct-2026", "Oil", "ICP", "ICD#9: 2.25"" held up @4920 FT (MMG-2@3763 ft and MMG-3@5768 FT) jarred...", -15.000#, "BOPD"
    AddRec "01-Oct-2026", "Gas", "MHN", "NA#4yH: SSSV control line leakage.Wireline cargo picked up at NA Plat...", -0.020#, "MMSCMD"
    AddRec "01-Oct-2026", "Gas", "BHS", "SP: Plat SD occurred low pneumatics in SDP", -0.010#, "MMSCMD"
    AddRec "01-Oct-2026", "Gas", "BHS", "12"" SE-MSP WFL leakage rectification job by MSV SEAMEC AGASTYA: Vessel...", -0.030#, "MMSCMD"
    AddRec "01-Oct-2026", "Gas", "WO-16", "WO-16A#3/#5/#9 , WO-5 #1 & #3 & B-121A#1 clsd for Solvent treatment job...", -0.201#, "MMSCMD"
    AddRec "02-Oct-2026", "Oil", "MHN", "NA#4yH S/choke installation: was clsd SSSV control line leakage", -70.000#, "BOPD"
    AddRec "02-Oct-2026", "Oil", "MHN", "N19#4H: Acid job carried out by S/Nidhi", -35.000#, "BOPD"
    AddRec "02-Oct-2026", "Oil", "NQO", "N14: 14"" N14-NQO W/F line sterilization job done by GS Roopa, #8 clsd...", -6.000#, "BOPD"
    AddRec "02-Oct-2026", "Oil", "NQO", "LD#8Y: found clsd from SSSV, opened", -15.000#, "BOPD"
    AddRec "02-Oct-2026", "Oil", "BHS", "IH: WF cargo flowing via 14"" IH-SH WF since . Less liquid receipt at...", -250.000#, "BOPD"
    AddRec "02-Oct-2026", "Oil", "WO-16", "WO-16A#3/#5/#9 , WO-5 #1 & #3 & B-121A#1 clsd for Solvent treatment job...", -40.000#, "BOPD"
    AddRec "02-Oct-2026", "Gas", "MHN", "NA#4yH S/choke installation: was clsd SSSV control line leakage", -0.020#, "MMSCMD"
    AddRec "03-Oct-2026", "Oil", "MHN", "PGC-B: Tripped high LPC DE X&Y axis radial vibration", -25.000#, "BOPD"
    AddRec "03-Oct-2026", "Oil", "NQO", "NHA#4: Carried out hole prove", -24.000#, "BOPD"
    AddRec "03-Oct-2026", "Oil", "NQO", "Backpressure prod high CIM pressure caused by MHN PGC hot wash,...", -100.000#, "BOPD"
    AddRec "03-Oct-2026", "Oil", "BHS", "12"" SE-MSP WFL leakage rectification job by MSV SEAMEC AGASTYA:...", -420.000#, "BOPD"
    AddRec "03-Oct-2026", "Oil", "BHS", "IH: WF cargo flowing via 14"" IH-SH WF since . Less liquid receipt at...", -250.000#, "BOPD"
    AddRec "03-Oct-2026", "Oil", "ICP", "ICD #11: FGS carried out; #2: Hole probed with 2.65"" GC up to 6023 ft,...", -21.000#, "BOPD"
    AddRec "03-Oct-2026", "Oil", "ICP", "RS12#8 (O): Replaced GLV at MMG-4 (ft) and MMG-1 (ft) after multiple...", -150.000#, "BOPD"
    AddRec "03-Oct-2026", "Gas", "MHN", "PGC-B: Tripped high LPC DE X&Y axis radial vibration", -0.015#, "MMSCMD"
    AddRec "03-Oct-2026", "Gas", "NQO", "Backpressure prod high CIM pressure caused by MHN PGC hot wash,...", -0.050#, "MMSCMD"
    AddRec "03-Oct-2026", "Gas", "NQO", "NR#6 (free gas well): Found pin hole leakage from PFA welded flange,...", -0.051#, "MMSCMD"
    AddRec "03-Oct-2026", "Gas", "BHS", "12"" SE-MSP WFL leakage rectification job by MSV SEAMEC AGASTYA:...", -0.019#, "MMSCMD"
    AddRec "04-Oct-2026", "Oil", "MHN", "NA# 3zH: 2.35""GC run up to storm choke depth (514 ft", -30.000#, "BOPD"
    AddRec "04-Oct-2026", "Oil", "NQO", "NHA#4H: P/O EMG & recorded FBHPT gradients from 1375m to surface", -7.000#, "BOPD"
    AddRec "04-Oct-2026", "Oil", "NQO", "NQG-PGC-A (6th PGC): Stopped for 6hrs for NQG-PGC-B discharge XSDV...", -150.000#, "BOPD"
    AddRec "04-Oct-2026", "Oil", "BHS", "12"" SE-MSP WFL leakage rectification job by MSV SEAMEC AGASTYA: Vessel...", -400.000#, "BOPD"
    AddRec "04-Oct-2026", "Oil", "BHS", "IH: WF cargo flowing via 14"" IH-SH WF since . Less liquid receipt at...", -250.000#, "BOPD"
    AddRec "04-Oct-2026", "Oil", "SHP", "S1-6#1Z(ABAN-8): Flowback started post DPRAS job by WSV GS Roopa", 330.000#, "BOPD"
    AddRec "04-Oct-2026", "Oil", "ICP", "ZB #7ZH: DPRAS job carried out by GS Roopa", -22.000#, "BOPD"
    AddRec "04-Oct-2026", "Gas", "NQO", "NR#6 (free gas well): Found pin hole leakage from PFA welded flange,...", -0.070#, "MMSCMD"
    AddRec "04-Oct-2026", "Gas", "NQO", "NQG-PGC-A (6th PGC): Stopped for 6hrs for NQG-PGC-B discharge XSDV...", -0.198#, "MMSCMD"
    AddRec "04-Oct-2026", "Gas", "BHS", "12"" SE-MSP WFL leakage rectification job by MSV SEAMEC AGASTYA: Vessel...", -0.020#, "MMSCMD"
    AddRec "04-Oct-2026", "Gas", "SHP", "S1-6#1Z(ABAN-8): Flowback started post DPRAS job by WSV GS Roopa", 0.060#, "MMSCMD"
    AddRec "05-Oct-2026", "Oil", "NQO", "N17#8H: DPRAS job carried out by GS Roopa", -20.000#, "BOPD"
    AddRec "05-Oct-2026", "Oil", "BHS", "12"" SE-MSP WFL leakage rectification job by MSV SEAMEC AGASTYA: Vessel...", -467.000#, "BOPD"
    AddRec "05-Oct-2026", "Oil", "BHS", "IH: WF cargo flowing via 14"" IH-SH WF since . Less liquid receipt at...", -250.000#, "BOPD"
    AddRec "05-Oct-2026", "Oil", "BHS", "IP#9: clsd for GC run & EMG hang(Wireline job", -25.000#, "BOPD"
    AddRec "05-Oct-2026", "Oil", "BHS", "IG#5: Observed GIdP & temperature in decreasing trend since 1900 hrs", -27.000#, "BOPD"
    AddRec "05-Oct-2026", "Oil", "SHP", "SHP PGC-C (7th PGC) given ventilated S/D leakage in 2nd stage cooler...", -60.000#, "BOPD"
    AddRec "05-Oct-2026", "Oil", "SHP", "SV#5: clsd high W/C as per base instruction", -70.000#, "BOPD"
    AddRec "05-Oct-2026", "Oil", "ICP", "RS12#8 (O): CTU Job by WSV Celestial cancelled leakage in bowen...", -150.000#, "BOPD"
    AddRec "05-Oct-2026", "Oil", "ICP", "ICD #4: EMG parked Ft", -15.000#, "BOPD"
    AddRec "05-Oct-2026", "Gas", "NQO", "NR#6 (free gas well): Found pin hole leakage from PFA welded flange,...", -0.070#, "MMSCMD"
    AddRec "05-Oct-2026", "Gas", "BHS", "12"" SE-MSP WFL leakage rectification job by MSV SEAMEC AGASTYA: Vessel...", -0.025#, "MMSCMD"
    AddRec "05-Oct-2026", "Gas", "SHP", "SHP PGC-C (7th PGC) given ventilated S/D leakage in 2nd stage cooler...", -0.040#, "MMSCMD"
    AddRec "05-Oct-2026", "Gas", "SHP", "SV#5: clsd high W/C as per base instruction", -0.005#, "MMSCMD"
    AddRec "06-Oct-2026", "Oil", "MHN", "NA#1xH: RIH 2"" GC cleared up to 62 ft", -15.000#, "BOPD"
    AddRec "06-Oct-2026", "Oil", "NQO", "N12: Suspected deoptimization in multiple wells as per SCADA", -100.000#, "BOPD"
    AddRec "06-Oct-2026", "Oil", "NQO", "NHA#5H: clsd for EMG park for FGS study", -6.000#, "BOPD"
    AddRec "06-Oct-2026", "Oil", "NQO", "LE#9: Gas injection clsd for 1.5 hrs for ACV replacement job", -12.000#, "BOPD"
    AddRec "06-Oct-2026", "Oil", "BHS", "12"" SE-MSP WFL leakage rectification job by MSV SEAMEC AGASTYA: Vessel...", -467.000#, "BOPD"
    AddRec "06-Oct-2026", "Oil", "BHS", "IH: WF cargo flowing via 14"" IH-SH WF since . Less liquid receipt at...", -250.000#, "BOPD"
    AddRec "06-Oct-2026", "Oil", "BHS", "RS-1: Plat SD as per SCADA observation", -185.000#, "BOPD"
    AddRec "06-Oct-2026", "Oil", "BHS", "NWIS-R TG tripped turbine enclosure combustible gas level hi-hi alarm...", -110.000#, "BOPD"
    AddRec "06-Oct-2026", "Oil", "ICP", "RS12 #8 CTU Job by WSV GS Roopa: RIH CT with technical water, observed...", -180.000#, "BOPD"
    AddRec "06-Oct-2026", "Gas", "NQO", "NR#6 (free gas well): Part spool replaced and well opened", -0.023#, "MMSCMD"
    AddRec "06-Oct-2026", "Gas", "BHS", "12"" SE-MSP WFL leakage rectification job by MSV SEAMEC AGASTYA: Vessel...", -0.025#, "MMSCMD"
    AddRec "06-Oct-2026", "Gas", "BHS", "RS-1: Plat SD as per SCADA observation", -0.030#, "MMSCMD"
    AddRec "06-Oct-2026", "Gas", "BHS", "NWIS-R TG tripped turbine enclosure combustible gas level hi-hi alarm...", -0.127#, "MMSCMD"
    AddRec "07-Oct-2026", "Oil", "NQO", "#All: High CIM pressure less gas received by MHN", -300.000#, "BOPD"
    AddRec "07-Oct-2026", "Oil", "NQO", "10"" LB- NQP WF line pigging: Used 9.75"" & 10"" Scrapper pigs", -45.000#, "BOPD"
    AddRec "07-Oct-2026", "Oil", "BHS", "12"" SE-MSP WFL leakage rectification job by MSV SEAMEC AGASTYA: Vessel...", -417.000#, "BOPD"
    AddRec "07-Oct-2026", "Oil", "BHS", "IH: WF cargo flowing via 14"" IH-SH WF since . Less prod receipt at SHP...", -250.000#, "BOPD"
    AddRec "07-Oct-2026", "Oil", "SHP", "S1-6#1Z(ABAN-8): After observation of reduced liquid rate, choke was...", 154.000#, "BOPD"
    AddRec "07-Oct-2026", "Oil", "SHP", "SV#7: Development Well completed in L-III reservoir by Rig Jindal Explorer", 93.000#, "BOPD"
    AddRec "07-Oct-2026", "Oil", "SHP", "SHP PGC-C (7th PGC) given ventilated S/D , leakage in 2nd stage gas...", -100.000#, "BOPD"
    AddRec "07-Oct-2026", "Oil", "SHP", "RS-15: SHMFS team working at RS-15 plat reported leakage at wellhead...", -80.000#, "BOPD"
    AddRec "07-Oct-2026", "Oil", "ICP", "RS12 #8: Acid spot job carried out by WSV GS Roopa, Well flowing with...", -150.000#, "BOPD"
    AddRec "07-Oct-2026", "Gas", "BHS", "12"" SE-MSP WFL leakage rectification job by MSV SEAMEC AGASTYA: Vessel...", -0.020#, "MMSCMD"
    AddRec "07-Oct-2026", "Gas", "BHS", "IH: WF cargo flowing via 14"" IH-SH WF since . Less prod receipt at SHP...", -0.020#, "MMSCMD"
    AddRec "07-Oct-2026", "Gas", "SHP", "S1-6#2(FG): clsd to avoid backpressure at SH 7th PGC unavailability", -0.060#, "MMSCMD"
    AddRec "07-Oct-2026", "Gas", "SHP", "SHP PGC-C (7th PGC) given ventilated S/D , leakage in 2nd stage gas...", -0.080#, "MMSCMD"
    AddRec "07-Oct-2026", "Gas", "SHP", "RS-15: SHMFS team working at RS-15 plat reported leakage at wellhead...", -0.050#, "MMSCMD"
    AddRec "08-Oct-2026", "Oil", "NQO", "N17: Found XSDV 80% clsd at LA end, instrumentation team mobilised...", -122.000#, "BOPD"
    AddRec "08-Oct-2026", "Oil", "NQO", "NQO#All: High CIM pressure less gas received by MHN, pig loaded at NQG,...", -100.000#, "BOPD"
    AddRec "08-Oct-2026", "Oil", "BHS", "IH: WF cargo flowing via 14"" IH-SH WF since . Less prod receipt at SHP...", -250.000#, "BOPD"
    AddRec "08-Oct-2026", "Oil", "BHS", "RS21 #6: clsd to utilize the available Gas Injection optimally to...", -75.000#, "BOPD"
    AddRec "08-Oct-2026", "Oil", "SHP", "RS-15: On arrival , found leakage from #3H PFA line; RS-15#1H/2/3H/8 clsd", -290.000#, "BOPD"
    AddRec "08-Oct-2026", "Oil", "SHP", "SV#7: Development Well completed in L-III reservoir by Rig Jindal Explorer", 100.000#, "BOPD"
    AddRec "08-Oct-2026", "Oil", "SHP", "S1-6#1Z(ST job by ABAN-VIII): Post Nitrified DPRAS job by GS Roopa ,...", 147.000#, "BOPD"
    AddRec "08-Oct-2026", "Gas", "NQO", "NQO#All: High CIM pressure less gas received by MHN, pig loaded at NQG,...", -0.039#, "MMSCMD"
    AddRec "08-Oct-2026", "Gas", "BHS", "IH: WF cargo flowing via 14"" IH-SH WF since . Less prod receipt at SHP...", -0.020#, "MMSCMD"
    AddRec "08-Oct-2026", "Gas", "BHS", "RS21 #6: clsd to utilize the available Gas Injection optimally to...", -0.005#, "MMSCMD"
    AddRec "08-Oct-2026", "Gas", "SHP", "RS-15: On arrival , found leakage from #3H PFA line; RS-15#1H/2/3H/8 clsd", -0.145#, "MMSCMD"
    AddRec "09-Oct-2026", "Oil", "MHN", "TG-1 & TG-2 tripped simultaneously engine flame-out , resulting in PSD", -7940.000#, "BOPD"
    AddRec "09-Oct-2026", "Oil", "MHN", "MHN PSD: All wells of N2-N2A/NV/NJ/ND and N14 sector de-optimised drop...", -1040.000#, "BOPD"
    AddRec "09-Oct-2026", "Oil", "BHS", "12"" SE-MSP WFL leakage rectification job by MSV SEAMEC AGASTYA: Vessel...", -130.000#, "BOPD"
    AddRec "09-Oct-2026", "Oil", "BHS", "IH: WF cargo flowing via 14"" IH-SH WF since . Less prod receipt at SHP...", -250.000#, "BOPD"
    AddRec "09-Oct-2026", "Oil", "BHS", "BHS RS-18 (All wells except #12(self)): MLGIP Down by 56 KSC PSD at MHN", -320.000#, "BOPD"
    AddRec "09-Oct-2026", "Oil", "SCA", "12'' IF-SCA WF sub sea line leakage arrest job by MSV Seamec Paladin:...", -256.000#, "BOPD"
    AddRec "09-Oct-2026", "Oil", "SHP", "RS-15: RS-15: Rectified #3H leaking spool at SHW and shifted to RS-15....", -90.000#, "BOPD"
    AddRec "09-Oct-2026", "Oil", "SHP", "IM#12H: clsd for EMG hang(Wireline job", -38.000#, "BOPD"
    AddRec "09-Oct-2026", "Oil", "ICP", "RS2#7Z (O): Observed leakage from PFA, well clsd", -90.000#, "BOPD"
    AddRec "09-Oct-2026", "Gas", "MHN", "TG-1 & TG-2 tripped simultaneously engine flame-out , resulting in PSD", -0.865#, "MMSCMD"
    AddRec "09-Oct-2026", "Gas", "MHN", "MHN PSD: All wells of N2-N2A/NV/NJ/ND and N14 sector de-optimised drop...", -0.150#, "MMSCMD"
    AddRec "09-Oct-2026", "Gas", "NQO", "Received excess gas from MHN after PSD revival unavailability of MHN...", -0.250#, "MMSCMD"
    AddRec "09-Oct-2026", "Gas", "BHS", "12"" SE-MSP WFL leakage rectification job by MSV SEAMEC AGASTYA: Vessel...", -0.020#, "MMSCMD"
    AddRec "09-Oct-2026", "Gas", "BHS", "IH: WF cargo flowing via 14"" IH-SH WF since . Less prod receipt at SHP...", -0.020#, "MMSCMD"
    AddRec "09-Oct-2026", "Gas", "BHS", "BHS RS-18 (All wells except #12(self)): MLGIP Down by 56 KSC PSD at MHN", -0.045#, "MMSCMD"
    AddRec "09-Oct-2026", "Gas", "SCA", "12'' IF-SCA WF sub sea line leakage arrest job by MSV Seamec Paladin:...", -0.027#, "MMSCMD"
    AddRec "09-Oct-2026", "Gas", "SHP", "RS-15: RS-15: Rectified #3H leaking spool at SHW and shifted to RS-15....", -0.025#, "MMSCMD"
    AddRec "10-Oct-2026", "Oil", "MHN", "PGC-D (IVth PGC) started and loaded", -90.000#, "BOPD"
    AddRec "10-Oct-2026", "Oil", "NQO", "N2-N2A/N14 sector: Observed MLGIP drop by 3 KSC MHN PGC shut down", -120.000#, "BOPD"
    AddRec "10-Oct-2026", "Oil", "BHS", "IH: WF cargo flowing via 14"" IH-SH WF since . Less liquid receipt at...", -250.000#, "BOPD"
    AddRec "10-Oct-2026", "Oil", "BHS", "RS-1#4: clsd for rocking", -17.000#, "BOPD"
    AddRec "10-Oct-2026", "Oil", "BHS", "RS-21#6: Observed well flowing with sub optimal temperature degree...", -55.000#, "BOPD"
    AddRec "10-Oct-2026", "Oil", "SCA", "12'' IF-SCA WF sub sea line leakage arrest job by MSV Seamec Paladin:...", -561.000#, "BOPD"
    AddRec "10-Oct-2026", "Oil", "WO-16", "WO-16A#5: clsd intermittently to maintain separator level constraint of...", -230.000#, "BOPD"
    AddRec "10-Oct-2026", "Oil", "SHP", "IM#12H: clsd for EMG run (Wireline job", -36.000#, "BOPD"
    AddRec "10-Oct-2026", "Gas", "MHN", "PGC-D (IVth PGC) started and loaded", -0.300#, "MMSCMD"
    AddRec "10-Oct-2026", "Gas", "NQO", "N2-N2A/N14 sector: Observed MLGIP drop by 3 KSC MHN PGC shut down", -0.040#, "MMSCMD"
    AddRec "10-Oct-2026", "Gas", "BHS", "IH: WF cargo flowing via 14"" IH-SH WF since . Less liquid receipt at...", -0.020#, "MMSCMD"
    AddRec "10-Oct-2026", "Gas", "SCA", "12'' IF-SCA WF sub sea line leakage arrest job by MSV Seamec Paladin:...", -0.037#, "MMSCMD"
    AddRec "10-Oct-2026", "WI", "MNW", "MNW: MNW TG-1 and MNP TG-4 running in island mode", -57000.000#, "BOPD"
    AddRec "10-Oct-2026", "WI", "MNW", "MNW: MNW TG-1 and MNP TG-4 running in island mode", -57000.000#, "BOPD"
    AddRec "11-Oct-2026", "Oil", "NQO", "12"" N24-NQD WF line pigging: Used 12"" scrapper pig without brush", -58.000#, "BOPD"
    AddRec "11-Oct-2026", "Oil", "BHS", "RS-21#6: Observed well flowing with sub optimal temperature C to 32C", -56.000#, "BOPD"
    AddRec "11-Oct-2026", "Oil", "BHS", "IH: WF cargo flowing via 14"" IH-SH WF since . Less liquid receipt at...", -250.000#, "BOPD"
    AddRec "11-Oct-2026", "Oil", "BHS", "RS-1: Landing temperature down by 3 deg C", -100.000#, "BOPD"
    AddRec "11-Oct-2026", "Oil", "SCA", "12'' IF-SCA WF sub sea line leakage arrest job by MSV Seamec Paladin:...", -475.000#, "BOPD"
    AddRec "11-Oct-2026", "Oil", "SHP", "18"" RS-15 WFL PIGGING: #12/2/6H/5H clsd for PIGGING operation under the...", -115.000#, "BOPD"
    AddRec "11-Oct-2026", "Oil", "SHP", "SV#7: clsd for NDPRAS+XLD Stimulation job by WSV GS ROOPA", -30.000#, "BOPD"
    AddRec "11-Oct-2026", "Oil", "ICP", "ICD #5: Acid spotting job carried out by Samudra Nidhi", -20.000#, "BOPD"
    AddRec "11-Oct-2026", "Oil", "CL-7", "Cluster 7/8: a) 101th offtake: Tanker Swarna Brahmaputra, dispatch...", -100.000#, "BOPD"
    AddRec "11-Oct-2026", "Gas", "BHS", "IH: WF cargo flowing via 14"" IH-SH WF since . Less liquid receipt at...", -0.020#, "MMSCMD"
    AddRec "11-Oct-2026", "Gas", "BHS", "RS-1: Landing temperature down by 3 deg C", -0.010#, "MMSCMD"
    AddRec "11-Oct-2026", "Gas", "SCA", "12'' IF-SCA WF sub sea line leakage arrest job by MSV Seamec Paladin:...", -0.029#, "MMSCMD"
    AddRec "11-Oct-2026", "Gas", "SHP", "18"" RS-15 WFL PIGGING: #12/2/6H/5H clsd for PIGGING operation under the...", -0.025#, "MMSCMD"
    AddRec "11-Oct-2026", "WI", "MNW", "MNW: PGC-D (4TH PGC) tripped EBV338 drive module fault", 0.000#, "BOPD"
    AddRec "11-Oct-2026", "WI", "MNW", "MNW: MIP tripped PSD", -57000.000#, "BOPD"
    AddRec "11-Oct-2026", "WI", "MNW", "MNW: MIP tripped PSD", -57000.000#, "BOPD"
    AddRec "12-Oct-2026", "Oil", "MHN", "Fire incidence in NQG PGC cooler area", -3750.000#, "BOPD"
    AddRec "12-Oct-2026", "Oil", "NQO", "FSD occurred rupture and fire in NQG-PGC-C 3rd stage cooler discharge...", -31390.000#, "BOPD"
    AddRec "12-Oct-2026", "Oil", "BHS", "IH: WF cargo flowing via 14"" IH-SH WF since . Less liquid receipt at...", -250.000#, "BOPD"
    AddRec "12-Oct-2026", "Gas", "MHN", "Fire incidence in NQG PGC cooler area", -0.610#, "MMSCMD"
    AddRec "12-Oct-2026", "Gas", "NQO", "FSD occurred rupture and fire in NQG-PGC-C 3rd stage cooler discharge...", -2.640#, "MMSCMD"
    AddRec "12-Oct-2026", "Gas", "BHS", "IH: WF cargo flowing via 14"" IH-SH WF since . Less liquid receipt at...", -0.020#, "MMSCMD"
    AddRec "12-Oct-2026", "WI", "MNW", "MNW: MIP tripped PSD", -57000.000#, "BOPD"
    AddRec "12-Oct-2026", "WI", "MNW", "MNW: MIP tripped PSD", -57000.000#, "BOPD"
    AddRec "13-Oct-2026", "Oil", "MHN", "NQO revival is in progress", -2500.000#, "BOPD"
    AddRec "13-Oct-2026", "Oil", "NQO", "NQO Fire Shut down: Revival is in progress. 145 (out of 246) wells opened", -12800.000#, "BOPD"
    AddRec "13-Oct-2026", "Oil", "BHS", "IH: WF cargo flowing via 14"" IH-SH WF since . Less liquid receipt at...", -250.000#, "BOPD"
    AddRec "13-Oct-2026", "Oil", "BHS", "RS-18#3, #7, 10: found flowing with suboptimal temperature, after...", -85.000#, "BOPD"
    AddRec "13-Oct-2026", "Oil", "WO-16", "WO-16A#5: clsd intermittently to maintain separator level constraint of...", -300.000#, "BOPD"
    AddRec "13-Oct-2026", "Oil", "ICP", "ST#9Z (O): Observed increase in W/C from 55% to 63%", -96.000#, "BOPD"
    AddRec "13-Oct-2026", "Gas", "MHN", "NQO revival is in progress", -0.200#, "MMSCMD"
    AddRec "13-Oct-2026", "Gas", "NQO", "NQO Fire Shut down: Revival is in progress. 145 (out of 246) wells opened", -1.447#, "MMSCMD"
    AddRec "13-Oct-2026", "Gas", "BHS", "IH: WF cargo flowing via 14"" IH-SH WF since . Less liquid receipt at...", -0.020#, "MMSCMD"
    AddRec "13-Oct-2026", "Gas", "BHS", "RS-18#3, #7, 10: found flowing with suboptimal temperature, after...", -0.020#, "MMSCMD"
    AddRec "13-Oct-2026", "Gas", "WO-16", "WO-16A#5: clsd intermittently to maintain separator level constraint of...", -0.015#, "MMSCMD"
    AddRec "14-Oct-2026", "Oil", "MHN", "N11, NS lift gas catered by NQO is not available post NQO FSD", -1300.000#, "BOPD"
    AddRec "14-Oct-2026", "Oil", "NQO", "NQO Fire Shut down: Revival is in progress. 161 (out of 244) wells...", -9100.000#, "BOPD"
    AddRec "14-Oct-2026", "Oil", "BHS", "IH: WF cargo flowing via 14"" IH-SH WF since . Less liquid receipt at...", -250.000#, "BOPD"
    AddRec "14-Oct-2026", "Oil", "WO-16", "WO-16A#5, WO-5A#3 Well clsd intermittently to maintain separator level...", -268.000#, "BOPD"
    AddRec "14-Oct-2026", "Oil", "ICP", "ICD #8 (O): Acid spot job carried out by WSV Celestial", -25.000#, "BOPD"
    AddRec "14-Oct-2026", "Gas", "MHN", "N11, NS lift gas catered by NQO is not available post NQO FSD", -0.130#, "MMSCMD"
    AddRec "14-Oct-2026", "Gas", "NQO", "NQO Fire Shut down: Revival is in progress. 161 (out of 244) wells...", -1.248#, "MMSCMD"
    AddRec "14-Oct-2026", "Gas", "BHS", "IH: WF cargo flowing via 14"" IH-SH WF since . Less liquid receipt at...", -0.020#, "MMSCMD"
    AddRec "14-Oct-2026", "Gas", "WO-16", "WO-16A#5, WO-5A#3 Well clsd intermittently to maintain separator level...", -0.011#, "MMSCMD"
    AddRec "15-Oct-2026", "Oil", "MHN", "N11#12 was opened after gas lift normalization from NQO; however, the...", -1076.000#, "BOPD"
    AddRec "15-Oct-2026", "Oil", "NQO", "NQO Fire Shut down: Revival is in progress. 161 (out of 244) wells...", -7700.000#, "BOPD"
    AddRec "15-Oct-2026", "Oil", "BHS", "IH: WF cargo flowing via 14"" IH-SH WF since . Less liquid receipt at...", -250.000#, "BOPD"
    AddRec "15-Oct-2026", "Oil", "WO-16", "WO-16A#5, WO-5A#3 Well clsd intermittently to maintain separator level...", -288.000#, "BOPD"
    AddRec "15-Oct-2026", "Gas", "MHN", "N11#12 was opened after gas lift normalization from NQO; however, the...", -0.090#, "MMSCMD"
    AddRec "15-Oct-2026", "Gas", "NQO", "NQO Fire Shut down: Revival is in progress. 161 (out of 244) wells...", -1.077#, "MMSCMD"
    AddRec "15-Oct-2026", "Gas", "BHS", "IH: WF cargo flowing via 14"" IH-SH WF since . Less liquid receipt at...", -0.020#, "MMSCMD"
    AddRec "15-Oct-2026", "Gas", "WO-16", "WO-16A#5, WO-5A#3 Well clsd intermittently to maintain separator level...", -0.012#, "MMSCMD"
    AddRec "15-Oct-2026", "WI", "MNW", "MNW: MIP tripped PSD", -57000.000#, "BOPD"
    AddRec "15-Oct-2026", "WI", "MNW", "MNW: MIP tripped PSD", -57000.000#, "BOPD"
    AddRec "16-Oct-2026", "Oil", "MHN", "N11#12 was opened yesterday after gas lift normalisation from NQO;...", -1005.000#, "BOPD"
    AddRec "16-Oct-2026", "Oil", "NQO", "NQO Fire Shut down: Revival is in progress. 165 (out of 244) wells opened", -7416.000#, "BOPD"
    AddRec "16-Oct-2026", "Oil", "BHS", "IH: WF cargo flowing via 14"" IH-SH WF since . Less liquid receipt at...", -250.000#, "BOPD"
    AddRec "16-Oct-2026", "Oil", "BHS", "SF#7SS: clsd for GC run", -10.000#, "BOPD"
    AddRec "16-Oct-2026", "Oil", "BHS", "BHS PGC-A & B: non-availability of both units, RS-16#1,9,11,SA#9SS clsd...", -325.000#, "BOPD"
    AddRec "16-Oct-2026", "Oil", "WO-16", "WO-16A#5, WO-5A#3 Well clsd intermittently to maintain separator level...", -336.000#, "BOPD"
    AddRec "16-Oct-2026", "Oil", "SHP", "Carried out well testing of SV#7 after stimulation job", 60.000#, "BOPD"
    AddRec "16-Oct-2026", "Oil", "ICP", "RS2 #7: PFA replaced", -25.000#, "BOPD"
    AddRec "16-Oct-2026", "Gas", "MHN", "N11#12 was opened yesterday after gas lift normalisation from NQO;...", -0.090#, "MMSCMD"
    AddRec "16-Oct-2026", "Gas", "NQO", "NQO Fire Shut down: Revival is in progress. 165 (out of 244) wells opened", -1.103#, "MMSCMD"
    AddRec "16-Oct-2026", "Gas", "BHS", "IH: WF cargo flowing via 14"" IH-SH WF since . Less liquid receipt at...", -0.020#, "MMSCMD"
    AddRec "16-Oct-2026", "Gas", "BHS", "BHS PGC-A & B: non-availability of both units, RS-16#1,9,11,SA#9SS clsd...", -1.124#, "MMSCMD"
    AddRec "16-Oct-2026", "Gas", "SHP", "Carried out well testing of SV#7 after stimulation job", 0.018#, "MMSCMD"
    AddRec "16-Oct-2026", "Gas", "WO-16", "WO-16A#5, WO-5A#3 Well clsd intermittently to maintain separator level...", -0.014#, "MMSCMD"
    AddRec "16-Oct-2026", "WI", "MNW", "MNW: MIP tripped PSD", -57000.000#, "BOPD"
    AddRec "16-Oct-2026", "WI", "ICW", "ICW MIP-1650: Pump tripped heavy seal leakage in DE side", -27021.000#, "BOPD"
    AddRec "16-Oct-2026", "WI", "MNW", "MNW: MIP tripped PSD", -57000.000#, "BOPD"
    AddRec "16-Oct-2026", "WI", "ICW", "ICW MIP-1650: Pump tripped heavy seal leakage in DE side", -27021.000#, "BOPD"
    AddRec "17-Oct-2026", "Oil", "MHN", "N11#12 was opened after gas lift normalization from NQO; however, the...", -1005.000#, "BOPD"
    AddRec "17-Oct-2026", "Oil", "NQO", "NQO Fire Shut down: Revival is in progress. 160 (out of 244) wells...", -7704.000#, "BOPD"
    AddRec "17-Oct-2026", "Oil", "BHS", "IH: WF cargo flowing via 14"" IH-SH WF since . Less liquid receipt at...", -250.000#, "BOPD"
    AddRec "17-Oct-2026", "Oil", "BHS", "SF#7SS: clsd for wireline job", -10.000#, "BOPD"
    AddRec "17-Oct-2026", "Oil", "BHS", "BHS PGC-A & B: non availability of both units, RS-16#1,9,11,SA#9SS clsd...", -205.000#, "BOPD"
    AddRec "17-Oct-2026", "Oil", "WO-16", "WO-16A#5: clsd intermittently to maintain separator level constraint of...", -386.000#, "BOPD"
    AddRec "17-Oct-2026", "Oil", "ICP", "ICD #9: Acid spot job carried out by WSV AM Passion", -36.000#, "BOPD"
    AddRec "17-Oct-2026", "Oil", "ICP", "ICD #4: Acid spot job carried out by WSV AM Passion", -23.000#, "BOPD"
    AddRec "17-Oct-2026", "Gas", "MHN", "N11#12 was opened after gas lift normalization from NQO; however, the...", -0.090#, "MMSCMD"
    AddRec "17-Oct-2026", "Gas", "NQO", "NQO Fire Shut down: Revival is in progress. 160 (out of 244) wells...", -1.183#, "MMSCMD"
    AddRec "17-Oct-2026", "Gas", "BHS", "IH: WF cargo flowing via 14"" IH-SH WF since . Less liquid receipt at...", -0.020#, "MMSCMD"
    AddRec "17-Oct-2026", "Gas", "BHS", "BHS PGC-A & B: non availability of both units, RS-16#1,9,11,SA#9SS clsd...", -0.638#, "MMSCMD"
    AddRec "17-Oct-2026", "Gas", "WO-16", "WO-16A#5: clsd intermittently to maintain separator level constraint of...", -0.014#, "MMSCMD"
    AddRec "17-Oct-2026", "Gas", "WO-16", "PGC-A tripped ; WO-16A#2/6/7 clsd", -0.075#, "MMSCMD"
    AddRec "17-Oct-2026", "WI", "MNW", "MNW: MIP tripped PSD", -57000.000#, "BOPD"
    AddRec "17-Oct-2026", "WI", "ICW", "ICW: MIP-1650: Pump tripped heavy seal leakage in DE side", -55853.000#, "BOPD"
    AddRec "17-Oct-2026", "WI", "MNW", "MNW: MIP tripped PSD", -57000.000#, "BOPD"
    AddRec "17-Oct-2026", "WI", "ICW", "ICW: MIP-1650: Pump tripped heavy seal leakage in DE side", -55853.000#, "BOPD"
    AddRec "18-Oct-2026", "Oil", "MHN", "N11#12 was opened after gas lift normalization from NQO; however, the...", -1005.000#, "BOPD"
    AddRec "18-Oct-2026", "Oil", "MHN", "NB#4xH: Storm choke servicing / replacement job in progress", -28.000#, "BOPD"
    AddRec "18-Oct-2026", "Oil", "NQO", "NQO Fire Shut down: Revival is in progress. 159 (out of 244 strings)...", -7343.000#, "BOPD"
    AddRec "18-Oct-2026", "Oil", "BHS", "IH: WF cargo flowing via 14"" IH-SH WF since . Less liquid receipt at...", -250.000#, "BOPD"
    AddRec "18-Oct-2026", "Oil", "BHS", "SF#9: clsd for wireline job", -32.000#, "BOPD"
    AddRec "18-Oct-2026", "Oil", "WO-16", "WO-16A#5: clsd intermittently to maintain separator level constraint of...", -445.000#, "BOPD"
    AddRec "18-Oct-2026", "Oil", "SHP", "SHG PGC-E(7TH PGC) tripped GG centre vibration high", -30.000#, "BOPD"
    AddRec "18-Oct-2026", "Gas", "MHN", "N11#12 was opened after gas lift normalization from NQO; however, the...", -0.090#, "MMSCMD"
    AddRec "18-Oct-2026", "Gas", "NQO", "NQO Fire Shut down: Revival is in progress. 159 (out of 244 strings)...", -1.273#, "MMSCMD"
    AddRec "18-Oct-2026", "Gas", "BHS", "IH: WF cargo flowing via 14"" IH-SH WF since . Less liquid receipt at...", -0.020#, "MMSCMD"
    AddRec "18-Oct-2026", "Gas", "SHP", "SHG PGC-E(7TH PGC) tripped GG centre vibration high", -0.015#, "MMSCMD"
    AddRec "18-Oct-2026", "Gas", "WO-16", "WO-16A#5: clsd intermittently to maintain separator level constraint of...", -0.012#, "MMSCMD"
    AddRec "18-Oct-2026", "WI", "MNW", "MNW: MIP tripped PSD", -57000.000#, "BOPD"
    AddRec "18-Oct-2026", "WI", "ICW", "ICW: MIP-1650: Pump tripped heavy seal leakage in DE side", -55853.000#, "BOPD"
    AddRec "18-Oct-2026", "WI", "MNW", "MNW: MIP tripped PSD", -57000.000#, "BOPD"
    AddRec "18-Oct-2026", "WI", "ICW", "ICW: MIP-1650: Pump tripped heavy seal leakage in DE side", -55853.000#, "BOPD"
    AddRec "19-Oct-2026", "Oil", "MHN", "N11#12 was opened after gas lift normalization from NQO; however, the...", -1005.000#, "BOPD"
    AddRec "19-Oct-2026", "Oil", "MHN", "NA & N9 deoptimised NQO PGC-C tripping, GI resumed from NQO side", -929.000#, "BOPD"
    AddRec "19-Oct-2026", "Oil", "MHN", "NB # 4xH: handover to FORE SIGHT DRILLER IX RIG for solvent Job and...", -140.000#, "BOPD"
    AddRec "19-Oct-2026", "Oil", "MHN", "NB#3XH SS: Diverted to MOL post WOR by FORE SIGHT DRILLER IX RIG", 99.000#, "BOPD"
    AddRec "19-Oct-2026", "Oil", "MHN", "N21#9H: clsd for PLO tool installation , unsuccessful", -21.000#, "BOPD"
    AddRec "19-Oct-2026", "Oil", "NQO", "NQO Fire Shut down: Revival is in progress. 159 (out of 244 strings) opened", -7343.000#, "BOPD"
    AddRec "19-Oct-2026", "Oil", "NQO", "NQP-PGC-C (3rd PGC) tripped GG lube oil supply pressure low S/D...", -2800.000#, "BOPD"
    AddRec "19-Oct-2026", "Oil", "BHS", "IH: WF cargo flowing via 14"" IH-SH WF since . Less liquid receipt at...", -250.000#, "BOPD"
    AddRec "19-Oct-2026", "Oil", "BHS", "SF#9: clsd for wireline job", -20.000#, "BOPD"
    AddRec "19-Oct-2026", "Oil", "BHS", "RS-1#4: found flowing with suboptimal temperature", -40.000#, "BOPD"
    AddRec "19-Oct-2026", "Oil", "WO-16", "B-119/B121 plat shut down since 0800 hrs SDP primary pump failure", -325.000#, "BOPD"
    AddRec "19-Oct-2026", "Oil", "WO-16", "WO-16A#5: clsd intermittently to maintain separator level constraint of...", -110.000#, "BOPD"
    AddRec "19-Oct-2026", "Gas", "MHN", "N11#12 was opened after gas lift normalization from NQO; however, the...", -0.090#, "MMSCMD"
    AddRec "19-Oct-2026", "Gas", "MHN", "NA & N9 deoptimised NQO PGC-C tripping, GI resumed from NQO side", -0.114#, "MMSCMD"
    AddRec "19-Oct-2026", "Gas", "NQO", "NQO Fire Shut down: Revival is in progress. 159 (out of 244 strings) opened", -1.273#, "MMSCMD"
    AddRec "19-Oct-2026", "Gas", "NQO", "NQP-PGC-C (3rd PGC) tripped GG lube oil supply pressure low S/D...", -0.100#, "MMSCMD"
    AddRec "19-Oct-2026", "Gas", "BHS", "IH: WF cargo flowing via 14"" IH-SH WF since . Less liquid receipt at...", -0.020#, "MMSCMD"
    AddRec "19-Oct-2026", "Gas", "WO-16", "B-119/B121 plat shut down since 0800 hrs SDP primary pump failure", -0.300#, "MMSCMD"
    AddRec "19-Oct-2026", "Gas", "WO-16", "WO-16A#5: clsd intermittently to maintain separator level constraint of...", -0.012#, "MMSCMD"
    AddRec "19-Oct-2026", "WI", "MNW", "MNW: MIP tripped PSD", -57000.000#, "BOPD"
    AddRec "19-Oct-2026", "WI", "ICW", "ICW: MIP-1650: Pump tripped heavy seal leakage in DE side", -31416.000#, "BOPD"
    AddRec "19-Oct-2026", "WI", "MNW", "MNW: MIP tripped PSD", -57000.000#, "BOPD"
    AddRec "19-Oct-2026", "WI", "ICW", "ICW: MIP-1650: Pump tripped heavy seal leakage in DE side", -31416.000#, "BOPD"
    AddRec "20-Oct-2026", "Oil", "MHN", "N11#12 was opened after gas lift normalization from NQO; however, the...", -1005.000#, "BOPD"
    AddRec "20-Oct-2026", "Oil", "MHN", "NB # 4xH / FD IX RIG: Storm choke installed", -198.000#, "BOPD"
    AddRec "20-Oct-2026", "Oil", "MHN", "N21#9H: Found well not flowing", -62.000#, "BOPD"
    AddRec "20-Oct-2026", "Oil", "NQO", "N10#: Wells found deoptimized", -45.000#, "BOPD"
    AddRec "20-Oct-2026", "Oil", "NQO", "LA#: Wells found deoptimized", -40.000#, "BOPD"
    AddRec "20-Oct-2026", "Oil", "NQO", "NQO Fire Shut down: Revival is in progress. 159 (out of 244 strings) opened", -7455.000#, "BOPD"
    AddRec "20-Oct-2026", "Oil", "BHS", "IH: WF cargo flowing via 14"" IH-SH WF since . Less liquid receipt at...", -250.000#, "BOPD"
    AddRec "20-Oct-2026", "Oil", "BHS", "SP#9: flowing with sub-optimal parameter", -20.000#, "BOPD"
    AddRec "20-Oct-2026", "Oil", "WO-16", "WO-16A#5: clsd intermittently to maintain separator level constraint of...", -424.000#, "BOPD"
    AddRec "20-Oct-2026", "Oil", "SHP", "SHP PGC-C(7th PGC) tripped GG lube oil level low", -20.000#, "BOPD"
    AddRec "20-Oct-2026", "Oil", "SHP", "II#7: clsd for storm choke servicing", -30.000#, "BOPD"
    AddRec "20-Oct-2026", "Gas", "MHN", "N11#12 was opened after gas lift normalization from NQO; however, the...", -0.090#, "MMSCMD"
    AddRec "20-Oct-2026", "Gas", "NQO", "NQO Fire Shut down: Revival is in progress. 159 (out of 244 strings) opened", -1.323#, "MMSCMD"
    AddRec "20-Oct-2026", "Gas", "BHS", "IH: WF cargo flowing via 14"" IH-SH WF since . Less liquid receipt at...", -0.020#, "MMSCMD"
    AddRec "20-Oct-2026", "Gas", "SHP", "Post pigging of 8"" B-157-B-127 WFL, observed gain in gas prod of 0.050...", 0.050#, "MMSCMD"
    AddRec "20-Oct-2026", "Gas", "SHP", "SHP PGC-C(7th PGC) tripped GG lube oil level low", -0.010#, "MMSCMD"
    AddRec "20-Oct-2026", "Gas", "WO-16", "WO-16A#5: clsd intermittently to maintain separator level constraint of...", -0.018#, "MMSCMD"
    AddRec "20-Oct-2026", "WI", "MNW", "MNW: MIP tripped PSD", -57000.000#, "BOPD"
    AddRec "20-Oct-2026", "WI", "MNW", "MNW: MIP tripped PSD", -57000.000#, "BOPD"
    AddRec "21-Oct-2026", "Oil", "MHN", "N11#12 was opened after gas lift normalization from NQO; however, the...", -1005.000#, "BOPD"
    AddRec "21-Oct-2026", "Oil", "MHN", "NB # 4xH / FD IX RIG: A section valve replaced by Rig FD-9 after...", -113.000#, "BOPD"
    AddRec "21-Oct-2026", "Oil", "MHN", "N21#9H: Retrieved 3 ""extended storm choke with 3 1/2"" slip lock ft", -34.000#, "BOPD"
    AddRec "21-Oct-2026", "Oil", "NQO", "NQO Fire Shut down: Revival is in progress. 156 (out of 244 strings)...", -7919.000#, "BOPD"
    AddRec "21-Oct-2026", "Oil", "BHS", "IH: WF cargo flowing via 14"" IH-SH WF since . Less liquid receipt at...", -250.000#, "BOPD"
    AddRec "21-Oct-2026", "Oil", "WO-16", "WO-16A#5: clsd intermittently to maintain separator level constraint of...", -440.000#, "BOPD"
    AddRec "21-Oct-2026", "Gas", "MHN", "N11#12 was opened after gas lift normalization from NQO; however, the...", -0.090#, "MMSCMD"
    AddRec "21-Oct-2026", "Gas", "NQO", "NQO Fire Shut down: Revival is in progress. 156 (out of 244 strings)...", -1.453#, "MMSCMD"
    AddRec "21-Oct-2026", "Gas", "BHS", "IH: WF cargo flowing via 14"" IH-SH WF since . Less liquid receipt at...", -0.020#, "MMSCMD"
    AddRec "21-Oct-2026", "Gas", "WO-16", "WO-16A#5: clsd intermittently to maintain separator level constraint of...", -0.019#, "MMSCMD"
    AddRec "21-Oct-2026", "WI", "MNW", "MNW: MIP tripped PSD", -57000.000#, "BOPD"
    AddRec "21-Oct-2026", "WI", "MNW", "MNW: MIP tripped PSD", -57000.000#, "BOPD"
    AddRec "22-Oct-2026", "Oil", "MHN", "N11#12 was opened after gas lift normalization from NQO; however, the...", -1005.000#, "BOPD"
    AddRec "22-Oct-2026", "Oil", "NQO", "NQO Fire Shut down: Revival is in progress. 155 (out of 244 strings)...", -7785.000#, "BOPD"
    AddRec "22-Oct-2026", "Oil", "BHS", "IH: WF cargo flowing via 14"" IH-SH WF since . Less liquid receipt at...", -250.000#, "BOPD"
    AddRec "22-Oct-2026", "Oil", "BHS", "SI#5,8,9: Found flowing with suboptimal parameters, optimized", -20.000#, "BOPD"
    AddRec "22-Oct-2026", "Oil", "BHS", "SG#9: Found flowing cold, optimized", -40.000#, "BOPD"
    AddRec "22-Oct-2026", "Oil", "WO-16", "WO-16A#5: clsd intermittently to maintain separator level constraint of...", -451.000#, "BOPD"
    AddRec "22-Oct-2026", "Oil", "SHP", "SS#6SS: Post pigging, leakage was observed from the test header clamp...", -100.000#, "BOPD"
    AddRec "22-Oct-2026", "Oil", "SHP", "S1-4#2Y: clsd for Well stimulation (Nitrified DPRAS with XLD) job", 100.000#, "BOPD"
    AddRec "22-Oct-2026", "Gas", "MHN", "N11#12 was opened after gas lift normalization from NQO; however, the...", -0.090#, "MMSCMD"
    AddRec "22-Oct-2026", "Gas", "NQO", "NQO Fire Shut down: Revival is in progress. 155 (out of 244 strings)...", -1.453#, "MMSCMD"
    AddRec "22-Oct-2026", "Gas", "BHS", "IH: WF cargo flowing via 14"" IH-SH WF since . Less liquid receipt at...", -0.020#, "MMSCMD"
    AddRec "22-Oct-2026", "Gas", "BHS", "SI#5,8,9: Found flowing with suboptimal parameters, optimized", -0.010#, "MMSCMD"
    AddRec "22-Oct-2026", "Gas", "BHS", "SG#9: Found flowing cold, optimized", -0.010#, "MMSCMD"
    AddRec "22-Oct-2026", "Gas", "SHP", "SS#6SS: Post pigging, leakage was observed from the test header clamp...", -0.014#, "MMSCMD"
    AddRec "22-Oct-2026", "Gas", "SHP", "S1-4#2Y: clsd for Well stimulation (Nitrified DPRAS with XLD) job", 0.080#, "MMSCMD"
    AddRec "22-Oct-2026", "Gas", "WO-16", "WO-16A#5: clsd intermittently to maintain separator level constraint of...", -0.019#, "MMSCMD"
    AddRec "22-Oct-2026", "WI", "MNW", "MNW: After MNW TG-1 & TG-2 synchronization; Water injection resumed", 24500.000#, "BOPD"
    AddRec "22-Oct-2026", "WI", "MNW", "MNW: After MNW TG-1 & TG-2 synchronization; Water injection resumed", 24500.000#, "BOPD"
    AddRec "23-Oct-2026", "Oil", "MHN", "N11#12 was opened after gas lift normalization from NQO; however, the...", -1005.000#, "BOPD"
    AddRec "23-Oct-2026", "Oil", "NQO", "N22#1: Nitrified DPRAS + XLD Job carried out by WSV Celestial, Well...", -73.000#, "BOPD"
    AddRec "23-Oct-2026", "Oil", "NQO", "NQO Fire Shut down: Revival is in progress. 155 (out of 244 strings)...", -7846.000#, "BOPD"
    AddRec "23-Oct-2026", "Oil", "NQO", "N22#5Z Well diverted to main line", 280.000#, "BOPD"
    AddRec "23-Oct-2026", "Oil", "BHS", "IH: WF cargo flowing via 14"" IH-SH WF since . Less liquid receipt at...", -250.000#, "BOPD"
    AddRec "23-Oct-2026", "Oil", "WO-16", "WO-16A#5: clsd intermittently to maintain separator level constraint of...", -451.000#, "BOPD"
    AddRec "23-Oct-2026", "Oil", "SHP", "S1-4#2Y(SF): Observed liquid gain post stimulation job (Nitrified DPRAS...", 249.000#, "BOPD"
    AddRec "23-Oct-2026", "Oil", "ICP", "ICD #9 (O): GLV deplugged from MMG-2", -18.000#, "BOPD"
    AddRec "23-Oct-2026", "Gas", "MHN", "N11#12 was opened after gas lift normalization from NQO; however, the...", -0.090#, "MMSCMD"
    AddRec "23-Oct-2026", "Gas", "NQO", "NQO Fire Shut down: Revival is in progress. 155 (out of 244 strings)...", -1.460#, "MMSCMD"
    AddRec "23-Oct-2026", "Gas", "BHS", "IH: WF cargo flowing via 14"" IH-SH WF since . Less liquid receipt at...", -0.020#, "MMSCMD"
    AddRec "23-Oct-2026", "Gas", "SHP", "S1-4#2Y(SF): Observed liquid gain post stimulation job (Nitrified DPRAS...", 0.107#, "MMSCMD"
    AddRec "23-Oct-2026", "Gas", "WO-16", "WO-16A#5: clsd intermittently to maintain separator level constraint of...", -0.019#, "MMSCMD"
    AddRec "24-Oct-2026", "Oil", "MHN", "N11#12 was opened after gas lift normalization from NQO; however, the...", -1005.000#, "BOPD"
    AddRec "24-Oct-2026", "Oil", "MHN", "N9 #1zH: found de-optimised; multiple revival attempts were carried out", -253.000#, "BOPD"
    AddRec "24-Oct-2026", "Oil", "NQO", "N22#1: Observed flash gain of 800 bbl after DPRAS job", 800.000#, "BOPD"
    AddRec "24-Oct-2026", "Oil", "NQO", "NQO Fire Shut down: Revival is in progress. 155 (out of 244 strings) opened", -7980.000#, "BOPD"
    AddRec "24-Oct-2026", "Oil", "NQO", "N22#5Z Observed flash gain of 800 Bbl after DPRAS job", 800.000#, "BOPD"
    AddRec "24-Oct-2026", "Oil", "BHS", "IH: WF cargo flowing via 14"" IH-SH WF since . Less liquid receipt at...", -250.000#, "BOPD"
    AddRec "24-Oct-2026", "Oil", "WO-16", "WO-16A#5: clsd intermittently to maintain separator level constraint of...", -451.000#, "BOPD"
    AddRec "24-Oct-2026", "Oil", "ICP", "ID#8H (O): Acid spot job carried out by WSV AM Passion", -10.000#, "BOPD"
    AddRec "24-Oct-2026", "Oil", "ICP", "ICD#9 (O): Tried multiple times with jarring to install GLV in MMG-2...", -50.000#, "BOPD"
    AddRec "24-Oct-2026", "Gas", "MHN", "N11#12 was opened after gas lift normalization from NQO; however, the...", -0.090#, "MMSCMD"
    AddRec "24-Oct-2026", "Gas", "MHN", "N9 #1zH: found de-optimised; multiple revival attempts were carried out", -0.020#, "MMSCMD"
    AddRec "24-Oct-2026", "Gas", "NQO", "NQO Fire Shut down: Revival is in progress. 155 (out of 244 strings) opened", -1.471#, "MMSCMD"
    AddRec "24-Oct-2026", "Gas", "BHS", "IH: WF cargo flowing via 14"" IH-SH WF since . Less liquid receipt at...", -0.020#, "MMSCMD"
    AddRec "24-Oct-2026", "Gas", "WO-16", "WO-16A#5: clsd intermittently to maintain separator level constraint of...", -0.019#, "MMSCMD"
    AddRec "25-Oct-2026", "Oil", "MHN", "N11#12 was opened after gas lift normalisation from NQO; however, the...", -1005.000#, "BOPD"
    AddRec "25-Oct-2026", "Oil", "MHN", "N9 #1zH: found de-optimised and unable to sustain pressure buildup", -253.000#, "BOPD"
    AddRec "25-Oct-2026", "Oil", "NQO", "N22#6H: clsd to reduce flaring", -57.000#, "BOPD"
    AddRec "25-Oct-2026", "Oil", "NQO", "NQO Fire Shut down: Revival is in progress. 155 (out of 244 strings) opened", -7980.000#, "BOPD"
    AddRec "25-Oct-2026", "Oil", "BHS", "IH: WF cargo flowing via 14"" IH-SH WF since . Less liquid receipt at...", -250.000#, "BOPD"
    AddRec "25-Oct-2026", "Oil", "WO-16", "WO-16A#5: clsd intermittently to maintain separator level constraint of...", -535.000#, "BOPD"
    AddRec "25-Oct-2026", "Oil", "SHP", "IN: Leakage reported by Discovery-1 on west side of IN jacket", 73.000#, "BOPD"
    AddRec "25-Oct-2026", "Oil", "ICP", "ICD#9: Tried multiple times with jarring to install GLV in MMG-2 (3773...", -55.000#, "BOPD"
    AddRec "25-Oct-2026", "Gas", "MHN", "N11#12 was opened after gas lift normalisation from NQO; however, the...", -0.090#, "MMSCMD"
    AddRec "25-Oct-2026", "Gas", "MHN", "N9 #1zH: found de-optimised and unable to sustain pressure buildup", -0.020#, "MMSCMD"
    AddRec "25-Oct-2026", "Gas", "NQO", "NQO Fire Shut down: Revival is in progress. 155 (out of 244 strings) opened", -1.471#, "MMSCMD"
    AddRec "25-Oct-2026", "Gas", "BHS", "IH: WF cargo flowing via 14"" IH-SH WF since . Less liquid receipt at...", -0.020#, "MMSCMD"
    AddRec "25-Oct-2026", "Gas", "WO-16", "WO-16A#5: clsd intermittently to maintain separator level constraint of...", -0.029#, "MMSCMD"

    If m_Count > 0 Then
        ReDim Preserve m_Remarks(0 To m_Count - 1)
    Else
        ReDim m_Remarks(0 To 0)
    End If

    m_Initialized = True
End Sub

Private Sub AddRec(ByVal d As String, ByVal c As String, ByVal comp As String, ByVal r As String, ByVal l As Double, ByVal u As String)
    m_Remarks(m_Count).DateStr = d
    m_Remarks(m_Count).Commodity = c
    m_Remarks(m_Count).Complex = comp
    m_Remarks(m_Count).RemarkText = r
    m_Remarks(m_Count).LossVal = l
    m_Remarks(m_Count).Unit = u
    m_Count = m_Count + 1
End Sub
