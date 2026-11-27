#!/usr/bin/env python3
"""
Generate VBA Macro with Embedded Remarks

This script reads remarks from Excel and generates a .bas file
that contains all remarks embedded directly in the VBA code.

Usage:
    python3 generate_remarks_macro.py

Output:
    Remarks_Embedded_Macro.bas - Ready to import into PowerPoint
"""

import openpyxl
from pathlib import Path


def load_remarks_from_excel(excel_path):
    """Load all remarks from Excel Remarks_Data sheet"""
    remarks = []

    try:
        wb = openpyxl.load_workbook(excel_path, data_only=True)

        if 'Remarks_Data' not in wb.sheetnames:
            print(f"❌ 'Remarks_Data' sheet not found!")
            return remarks

        ws = wb['Remarks_Data']

        # Skip header rows (1-2), start from row 3
        for row_idx in range(3, ws.max_row + 1):
            date_val = ws.cell(row_idx, 1).value
            comm_val = ws.cell(row_idx, 2).value
            complex_val = ws.cell(row_idx, 3).value
            remark_val = ws.cell(row_idx, 4).value
            loss_val = ws.cell(row_idx, 5).value
            unit_val = ws.cell(row_idx, 6).value

            if date_val and comm_val and remark_val:
                try:
                    loss_num = float(loss_val) if loss_val else 0
                except (ValueError, TypeError):
                    loss_num = 0

                remarks.append({
                    'date': str(date_val).strip(),
                    'commodity': str(comm_val).strip(),
                    'complex': str(complex_val).strip() if complex_val else '',
                    'remark': str(remark_val).strip(),
                    'loss': loss_num,
                    'unit': str(unit_val).strip() if unit_val else ''
                })

        return remarks

    except Exception as e:
        print(f"❌ Error loading remarks: {e}")
        return []


def escape_vba_string(text):
    """Escape text for VBA string literal - double all quotes"""
    return text.replace('"', '""')


def generate_bas_file(remarks, output_path):
    """Generate .bas file with embedded remarks in pipe-delimited format"""

    # Build the VBA code
    bas_code = '''Attribute VB_Name = "Remarks_Display"
Option Explicit

Private Const msoShapeRoundedRectangle As Long = 5
Private Const msoTrue As Long = -1
Private Const msoFalse As Long = 0
Private Const msoBringToFront As Long = 0
Private Const msoAnchorTop As Long = 1
Private Const ppMouseClick As Long = 1
Private Const ppActionRunMacro As Long = 8
Private Const ppAlignLeft As Long = 1

' PUBLIC CALLABLE MACROS (these show in Alt+F8)

Public Sub SetupAllSlides()
    MsgBox "Date buttons are already set up!" & vbCrLf & vbCrLf & _
           "Just click any date button on the slides during slideshow (F5)" & vbCrLf & _
           "Remarks will appear at the bottom-left of the chart.", _
           vbInformation, "Setup Complete"
End Sub

Public Sub SetupInteractiveSlide()
    MsgBox "Date buttons are already set up!" & vbCrLf & vbCrLf & _
           "Just click any date button on this slide during slideshow (F5)" & vbCrLf & _
           "Remarks will appear at the bottom-left of the chart.", _
           vbInformation, "Setup Complete"
End Sub

' DATA STORAGE

Private Function GetRemarksData() As String
    Dim data As String
    data = ""
'''

    # Add remarks data lines
    for rem in remarks:
        date_esc = escape_vba_string(rem['date'])
        comm_esc = escape_vba_string(rem['commodity'])
        complex_esc = escape_vba_string(rem['complex'])
        remark_esc = escape_vba_string(rem['remark'])
        unit_esc = escape_vba_string(rem['unit'])

        line = f'    data = data & "{date_esc}|{comm_esc}|{complex_esc}|{remark_esc}|{rem["loss"]}|{unit_esc}§"\n'
        bas_code += line

    bas_code += '''    GetRemarksData = data
End Function

' DISPLAY LOGIC

Public Sub ShowDateRemarks(ByVal oShp As Object)
    Dim sld As Object
    Dim dtStr As String
    Dim comm As String
    Dim remarksText As String
    Dim popCard As Object

    On Error Resume Next
    Set sld = oShp.Parent
    dtStr = GetDateFromButton(oShp)
    comm = GetCommodityFromButton(oShp, sld)

    If Len(dtStr) = 0 Then Exit Sub

    remarksText = BuildRemarksText(dtStr, comm)
    Set popCard = GetOrCreatePopupCard(sld)

    With popCard.TextFrame
        .WordWrap = msoTrue
        .TextRange.Text = remarksText
    End With

    Call FormatPopupText(popCard)
    popCard.Visible = msoTrue
    popCard.ZOrder msoBringToFront
    On Error GoTo 0
End Sub

Private Function BuildRemarksText(ByVal targetDate As String, ByVal commodity As String) As String
    Dim allData As String
    Dim normDate As String
    Dim msg As String
    Dim foundAny As Boolean
    Dim records() As String
    Dim i As Long
    Dim fields() As String
    Dim lStr As String
    Dim prefix As String

    normDate = NormalizeDate(targetDate)
    msg = normDate & vbCrLf
    foundAny = False

    allData = GetRemarksData()
    records = Split(allData, "§")

    For i = 0 To UBound(records)
        If Len(Trim(records(i))) > 0 Then
            fields = Split(records(i), "|")

            If UBound(fields) >= 5 Then
                If UCase(fields(0)) = UCase(normDate) And UCase(fields(1)) = UCase(commodity) Then
                    If Len(fields(3)) > 0 Then
                        If CDbl(fields(4)) <> 0 Then
                            If UCase(commodity) = "GAS" Then
                                lStr = " (" & Format(CDbl(fields(4)), "+0.000;-0.000;0") & " " & fields(5) & ")"
                            Else
                                lStr = " (" & Format(CDbl(fields(4)), "+#,##0;-#,##0;0") & " " & fields(5) & ")"
                            End If
                        Else
                            lStr = ""
                        End If

                        If Len(fields(2)) > 0 And UCase(fields(2)) <> "NONE" And UCase(fields(2)) <> "MH TOTAL" Then
                            prefix = fields(2) & ": "
                        Else
                            prefix = ""
                        End If

                        msg = msg & ChrW(8226) & " " & prefix & fields(3) & lStr & vbCrLf
                        foundAny = True
                    End If
                End If
            End If
        End If
    Next i

    If Not foundAny Then
        msg = msg & ChrW(8226) & " No remarks recorded for this date" & vbCrLf
    End If

    BuildRemarksText = msg
End Function

Private Function GetDateFromButton(ByVal oShp As Object) As String
    Dim alt As String
    Dim parts() As String
    On Error Resume Next
    alt = oShp.AlternativeText
    If InStr(alt, "|") > 0 Then
        parts = Split(alt, "|")
        GetDateFromButton = Trim(parts(0))
    End If
    On Error GoTo 0
End Function

Private Function GetCommodityFromButton(ByVal oShp As Object, ByVal sld As Object) As String
    Dim alt As String
    Dim parts() As String
    On Error Resume Next
    alt = oShp.AlternativeText
    If InStr(alt, "|") > 0 Then
        parts = Split(alt, "|")
        If UBound(parts) >= 1 Then
            GetCommodityFromButton = Trim(parts(1))
            Exit Function
        End If
    End If
    On Error GoTo 0
    GetCommodityFromButton = DetectSlideCommodity(sld)
End Function

Private Function DetectSlideCommodity(ByVal sld As Object) As String
    Dim shp As Object
    Dim i As Long
    Dim txt As String
    On Error Resume Next
    For i = 1 To sld.Shapes.Count
        Set shp = sld.Shapes(i)
        If shp.HasChart Then
            txt = UCase(shp.Chart.ChartTitle.Text)
            If InStr(txt, "GAS") > 0 Then
                DetectSlideCommodity = "Gas"
                Exit Function
            ElseIf InStr(txt, "WATER") > 0 Or InStr(txt, "WI") > 0 Then
                DetectSlideCommodity = "WI"
                Exit Function
            End If
        End If
    Next i
    On Error GoTo 0
    DetectSlideCommodity = "Oil"
End Function

Private Function NormalizeDate(ByVal dateStr As String) As String
    Dim dt As Date
    On Error Resume Next
    dt = CDate(dateStr)
    NormalizeDate = Format(dt, "dd-mmm-yyyy")
    On Error GoTo 0
End Function

Private Function GetOrCreatePopupCard(ByVal sld As Object) As Object
    Dim popCard As Object
    Dim chartShp As Object
    Dim i As Long
    Dim chartL As Single
    Dim chartT As Single
    Dim chartW As Single
    Dim chartH As Single

    On Error Resume Next
    Set popCard = sld.Shapes("Remarks_Popup_Card")
    On Error GoTo 0

    If Not popCard Is Nothing Then
        Set GetOrCreatePopupCard = popCard
        Exit Function
    End If

    On Error Resume Next
    For i = 1 To sld.Shapes.Count
        If sld.Shapes(i).HasChart Then
            Set chartShp = sld.Shapes(i)
            Exit For
        End If
    Next i
    On Error GoTo 0

    If Not chartShp Is Nothing Then
        chartL = chartShp.Left
        chartT = chartShp.Top
        chartW = chartShp.Width
        chartH = chartShp.Height
    Else
        chartL = 100
        chartT = 100
        chartW = 600
        chartH = 400
    End If

    Set popCard = sld.Shapes.AddShape(msoShapeRoundedRectangle, chartL + 12, chartT + chartH - 120, 280, 100)

    With popCard
        .Name = "Remarks_Popup_Card"
        .Adjustments.Item(1) = 0.1
        .Fill.Solid
        .Fill.ForeColor.RGB = RGB(240, 240, 235)
        .Line.Color.RGB = RGB(150, 120, 90)
        .Line.Weight = 1.5

        With .TextFrame
            .WordWrap = msoTrue
            .MarginLeft = 6
            .MarginRight = 6
            .MarginTop = 5
            .MarginBottom = 5
            .VerticalAnchor = msoAnchorTop
        End With

        .ActionSettings(ppMouseClick).Action = ppActionRunMacro
        .ActionSettings(ppMouseClick).Run = "HideRemarksPopup"
    End With

    Set GetOrCreatePopupCard = popCard
End Function

Private Sub FormatPopupText(ByVal popCard As Object)
    Dim para As Object
    Dim i As Long
    Dim tr As Object

    On Error Resume Next
    Set tr = popCard.TextFrame.TextRange

    With tr.Font
        .Name = "Calibri"
        .Size = 9
        .Bold = msoFalse
    End With

    For i = 1 To popCard.TextFrame.Paragraphs.Count
        Set para = popCard.TextFrame.Paragraphs(i)
        If i = 1 Then
            para.Font.Bold = msoTrue
            para.Font.Size = 10
        Else
            With para.ParagraphFormat
                .Alignment = ppAlignLeft
                .LineSpacing = 1
                .SpaceBefore = 0
                .SpaceAfter = 0
            End With
        End If
    Next i
    On Error GoTo 0
End Sub

Public Sub HideRemarksPopup(Optional ByVal oShp As Object = Nothing)
    Dim sld As Object
    Dim popCard As Object
    On Error Resume Next
    If oShp Is Nothing Then
        Set sld = ActivePresentation.SlideShowWindow.View.Slide
    Else
        Set sld = oShp.Parent
    End If
    If Not sld Is Nothing Then
        Set popCard = sld.Shapes("Remarks_Popup_Card")
        If Not popCard Is Nothing Then popCard.Visible = msoFalse
    End If
    On Error GoTo 0
End Sub
'''

    # Write to file
    with open(output_path, 'w', encoding='utf-8') as f:
        f.write(bas_code)

    return True


def main():
    print("=" * 70)
    print("GENERATE VBA MACRO WITH EMBEDDED REMARKS")
    print("=" * 70)
    print()

    base_path = Path("/Users/amanmishra/Library/CloudStorage/OneDrive-OilandNaturalGasCorporationLimited/MH_Dashboard")

    excel_file = base_path / "100_Days_Actual_vs_Projection.xlsx"
    output_file = base_path / "Remarks_Embedded_Macro.bas"

    # Check Excel file
    if not excel_file.exists():
        print(f"❌ Excel file not found: {excel_file}")
        return False

    print(f"📖 Reading from: {excel_file.name}")
    print()

    # Load remarks
    print("📖 Loading remarks from Excel...")
    remarks = load_remarks_from_excel(excel_file)

    if not remarks:
        print("❌ No remarks loaded.")
        return False

    print(f"✅ Loaded {len(remarks)} remarks")
    print()

    # Generate .bas file
    print(f"📝 Generating VBA macro...")
    if generate_bas_file(remarks, output_file):
        print(f"✅ Generated: {output_file.name}")
        print()
        print("=" * 70)
        print("✅ SUCCESS!")
        print("=" * 70)
        print()
        print(f"📊 {len(remarks)} remarks embedded in VBA code")
        print()
        print("Next steps:")
        print("1. Open PowerPoint (Presentation1.pptm)")
        print("2. Press Alt+F11 (or Fn+Option+F11 on Mac)")
        print("3. File → Import File → Select 'Remarks_Embedded_Macro.bas'")
        print("4. Close VBA editor")
        print("5. Press Alt+F8 to see: SetupAllSlides, SetupInteractiveSlide, HideRemarksPopup")
        print("6. Click any date button in slideshow (F5)")
        print("   → Remarks appear at bottom-left!")
        print()
        return True
    else:
        print("❌ Failed to generate macro.")
        return False


if __name__ == "__main__":
    import sys
    success = main()
    sys.exit(0 if success else 1)
