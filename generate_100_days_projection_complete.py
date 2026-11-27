#!/usr/bin/env python3
"""
Generate 100 Days Actual vs Projection Report (Complete Version)
Creates Excel file with Oil, Gas, and WI production data comparing actuals vs projections.
Output includes visible Remarks_Data sheet (editable) and Interactive VBA macro file.

Output:
    - 100_Days_Actual_vs_Projection.xlsx  (in MH_Dashboard folder)
    - Interactive_Charts_Macro_Dynamic.bas (in MH_Dashboard folder)

Usage:
    python3 generate_100_days_projection_complete.py <start_date> <end_date>

Example:
    python3 generate_100_days_projection_complete.py 11-09-2026 01-10-2026
    python3 generate_100_days_projection_complete.py 2026-09-11 2026-10-01
"""

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.chart import LineChart, BarChart, Reference
from openpyxl.chart.shapes import GraphicalProperties
from openpyxl.drawing.line import LineProperties
from datetime import datetime, timedelta
from pathlib import Path
from collections import defaultdict
import sys
import re
import zipfile
import io
import xml.etree.ElementTree as ET


class ReportGenerator:
    def __init__(self, base_path, start_date, end_date):
        self.base_path = Path(base_path)
        self.start_date = start_date
        self.end_date = end_date

        # Source files
        self.projection_file = self.base_path / "Outlook of OIL, Gas and WI 100 AND 365 DAYS_MH Asset (Final ).xlsx"
        self.production_file = self.base_path / "MH_Production_Dashboard.xlsx"

        # Output to MH_Dashboard root (NOT the Weekly subfolder)
        self.output_file = self.base_path / "100_Days_Actual_vs_Projection.xlsx"

        # Storage
        self.weekly_projections = []
        self.monthly_wi_projections = []
        self.daily_actuals = {}
        self.daily_loss_records = defaultdict(lambda: {'oil': [], 'gas': [], 'wi': []})

    # ------------------------------------------------------------------
    # DATA LOADING
    # ------------------------------------------------------------------

    def load_projection_data(self):
        """Load projection data from Outlook file"""
        print("Loading projection data...")

        wb = openpyxl.load_workbook(self.projection_file, data_only=True)

        # Oil & Gas weekly projections
        oil_sheet = wb['OIL 100 Days(without HP-B SD)']
        for r in range(12, 100):
            week_num = oil_sheet.cell(row=r, column=1).value
            week_start = oil_sheet.cell(row=r, column=2).value
            week_end_raw = oil_sheet.cell(row=r, column=3).value
            proj_oil = oil_sheet.cell(row=r, column=10).value   # Column J
            proj_gas = oil_sheet.cell(row=r, column=16).value   # Column P

            if not week_start or not isinstance(week_start, datetime):
                break

            if isinstance(week_end_raw, str):
                try:
                    week_end = datetime.strptime(week_end_raw.strip(), "%d %b %y")
                except Exception:
                    week_end = week_start + timedelta(days=6)
            elif isinstance(week_end_raw, datetime):
                week_end = week_end_raw
            else:
                week_end = week_start + timedelta(days=6)

            try:
                proj_oil_val = float(proj_oil) if proj_oil is not None else 0.0
            except (ValueError, TypeError):
                proj_oil_val = 0.0

            try:
                proj_gas_val = float(proj_gas) if proj_gas is not None else 0.0
            except (ValueError, TypeError):
                proj_gas_val = 0.0

            self.weekly_projections.append({
                'week': week_num,
                'start': week_start,
                'end': week_end,
                'proj_oil': proj_oil_val,
                'proj_gas': proj_gas_val
            })

        # WI monthly projections
        wi_sheet = wb['WI 100 Days & 365 Days Oulook']
        for r in range(9, 50):
            month_num = wi_sheet.cell(row=r, column=3).value    # Column C
            month_start = wi_sheet.cell(row=r, column=4).value  # Column D
            month_end = wi_sheet.cell(row=r, column=5).value    # Column E
            proj_wi = wi_sheet.cell(row=r, column=14).value     # Column N

            if not month_start or not isinstance(month_start, datetime):
                break

            try:
                proj_wi_val = float(proj_wi) if proj_wi is not None else 0.0
            except (ValueError, TypeError):
                proj_wi_val = 0.0

            self.monthly_wi_projections.append({
                'month': month_num,
                'start': month_start,
                'end': month_end,
                'proj_wi': proj_wi_val
            })

        wb.close()
        print(f"  Loaded {len(self.weekly_projections)} weekly projections (Oil & Gas)")
        print(f"  Loaded {len(self.monthly_wi_projections)} monthly projections (WI)")

    def load_actual_data(self):
        """Load actual daily production data from MH_Production_Dashboard"""
        print("Loading actual production data...")

        wb = openpyxl.load_workbook(self.production_file, data_only=True)
        daily_sheet = wb['Daily Production (MH)']

        for r in range(2, daily_sheet.max_row + 1):
            date_val = daily_sheet.cell(row=r, column=1).value
            if date_val and isinstance(date_val, datetime):
                oil = daily_sheet.cell(row=r, column=12).value  # Column L
                gas = daily_sheet.cell(row=r, column=21).value  # Column U
                wi = daily_sheet.cell(row=r, column=27).value   # Column AA

                self.daily_actuals[date_val.date()] = {
                    'oil': oil if oil else 0,
                    'gas': gas if gas else 0,
                    'wi': wi if wi else 0
                }

        wb.close()
        print(f"  Loaded {len(self.daily_actuals)} days of actual production data")

    def load_remarks_data(self):
        """Load daily complex-level reasons and remarks from Daily Comparison (MH)"""
        print("Loading daily comparison remarks...")
        try:
            wb = openpyxl.load_workbook(self.production_file, data_only=True)
            if 'Daily Comparison (MH)' not in wb.sheetnames:
                wb.close()
                print("  Warning: 'Daily Comparison (MH)' sheet not found")
                return

            comp_sheet = wb['Daily Comparison (MH)']
            for r in range(2, comp_sheet.max_row + 1):
                date_val = comp_sheet.cell(row=r, column=1).value
                if not date_val or not isinstance(date_val, datetime):
                    continue

                d_key = date_val.date()
                metric = str(comp_sheet.cell(row=r, column=2).value or '').strip()
                complex_name = str(comp_sheet.cell(row=r, column=3).value or '').strip()

                if complex_name in ('MH Total', 'MH Asset Total'):
                    continue

                reason = comp_sheet.cell(row=r, column=15).value
                short_rem = comp_sheet.cell(row=r, column=16).value
                oil_loss = comp_sheet.cell(row=r, column=17).value
                gas_loss = comp_sheet.cell(row=r, column=18).value
                wi_loss = comp_sheet.cell(row=r, column=19).value

                txt = short_rem or reason
                if not txt:
                    continue

                if metric == 'Oil':
                    self.daily_loss_records[d_key]['oil'].append({
                        'complex': complex_name,
                        'text': str(txt).strip(),
                        'loss': oil_loss if isinstance(oil_loss, (int, float)) else 0
                    })
                elif metric == 'Gas':
                    self.daily_loss_records[d_key]['gas'].append({
                        'complex': complex_name,
                        'text': str(txt).strip(),
                        'loss': gas_loss if isinstance(gas_loss, (int, float)) else 0
                    })
                else:
                    is_wi = ('wi' in metric.lower() or 'water' in metric.lower())
                    is_wi_by_loss = isinstance(wi_loss, (int, float)) and wi_loss != 0
                    if is_wi or is_wi_by_loss:
                        self.daily_loss_records[d_key]['wi'].append({
                            'complex': complex_name,
                            'text': str(txt).strip(),
                            'loss': wi_loss if isinstance(wi_loss, (int, float)) else 0
                        })

            wb.close()
            print(f"  Loaded daily loss records across {len(self.daily_loss_records)} dates")
        except Exception as e:
            print(f"  Warning: Could not load remarks: {e}")

    # ------------------------------------------------------------------
    # REMARK TEXT CLEANING
    # ------------------------------------------------------------------

    def _clean_remark_text(self, text):
        """Clean remark text — remove dates, timestamps, technical parameters"""
        if not text:
            return ""

        s = str(text).strip()
        s = re.sub(r'^[•\-\*\s]+', '', s)

        # Remove dates
        s = re.sub(r'\b(?:on\s+)?\d{1,2}[./-]\d{1,2}[./-]\d{2,4}\b', '', s, flags=re.IGNORECASE)
        s = re.sub(r'\(\s*\d{1,2}[./-]\d{1,2}[./-]\d{2,4}\s*\)', '', s)

        # Remove timestamps
        s = re.sub(r'\b(?:at|@|from|between)\s+\d{1,2}[:.]?\d{2}\s*(?:hrs|hours|am|pm)?(?:\s*(?:to|and|-)\s*\d{1,2}[:.]?\d{2}\s*(?:hrs|hours|am|pm)?)?\b', '', s, flags=re.IGNORECASE)
        s = re.sub(r'\b\d{1,2}[:.]\d{2}\s*(?:hrs|hours|am|pm)\b', '', s, flags=re.IGNORECASE)
        s = re.sub(r'@\s*\d{4}\s*hrs\b', '', s, flags=re.IGNORECASE)

        # Remove well testing parameter blocks
        s = re.sub(r'(?:Total\s+liquid|W/C|OIL|GIP|FTHP|FTHT|FLP|THP)\s*[:=-]\s*[^,;]+', '', s, flags=re.IGNORECASE)
        s = re.sub(r'\bWell\s+testing\s+is\s+carried\s+out\b', '', s, flags=re.IGNORECASE)
        s = re.sub(r'\bWell\s+testing\s+is\b', '', s, flags=re.IGNORECASE)

        # Remove loss tags in parentheses
        s = re.sub(r'\s*\([^)]*(?:loss|gain|bopd|bbld|bbls|mmscm|mwpd|bwpd|total)[^)]*\)', '', s, flags=re.IGNORECASE)
        s = re.sub(r'\s*\[[^\]]*\]', '', s)
        s = re.sub(r'\(from\s+[^)]+\)', '', s, flags=re.IGNORECASE)

        # Standardize common phrases
        s = re.sub(r'\bDevelopment completed\b', 'Dev completed', s, flags=re.IGNORECASE)
        s = re.sub(r'\b(?:loss\s+)?due to(?:\s+a)?\b', '', s, flags=re.IGNORECASE)
        s = re.sub(r'\bshortfall due to\b', '', s, flags=re.IGNORECASE)
        s = re.sub(r'\bstopped due to\b', 'stopped for', s, flags=re.IGNORECASE)

        # Abbreviations
        s = s.replace('Platform', 'Plat').replace('platform', 'plat')
        s = s.replace('Pipeline', 'Line').replace('pipeline', 'line')
        s = s.replace('Shutdown', 'SD').replace('shutdown', 'SD')
        s = s.replace('closed', 'clsd').replace('Closed', 'clsd')
        s = s.replace('production', 'prod').replace('Production', 'Prod')
        s = s.replace('workover job', 'WOJ')

        # Clean leftovers
        s = re.sub(r'\(\s*\)', '', s)
        s = re.sub(r'\[\s*\]', '', s)
        s = re.sub(r'[,;:.\s]+$', '', s)
        s = ' '.join(s.split()).strip(' ,;:-()[]')

        if len(s) > 75:
            s = s[:72].rsplit(' ', 1)[0] + '...'

        return s

    def _sanitize_for_xml(self, text):
        """Remove characters illegal in XML 1.0"""
        if not text:
            return ""
        # XML 1.0 legal chars: #x9 | #xA | #xD | [#x20-#xD7FF] | [#xE000-#xFFFD] | [#x10000-#x10FFFF]
        return re.sub(
            r'[^\x09\x0A\x0D\x20-퟿-�]',
            '',
            str(text)
        )

    # ------------------------------------------------------------------
    # PROJECTION LOOKUP
    # ------------------------------------------------------------------

    def get_projection_for_date(self, date, commodity):
        """Get projection for a specific date and commodity"""
        if commodity == 'wi':
            for month in self.monthly_wi_projections:
                if month['start'].date() <= date <= month['end'].date():
                    return month['proj_wi']
            return 0
        else:
            for week in self.weekly_projections:
                if week['start'].date() <= date <= week['end'].date():
                    return week[f'proj_{commodity}']
            return 0

    def get_daily_loss_bullets(self, date_key, commodity_key, shortfall, max_bullets=3):
        """Consolidate daily loss/shortfall remarks into bullet points"""
        records = self.daily_loss_records.get(date_key, {}).get(commodity_key, [])
        if not records:
            return []

        expanded = []
        for rec in records:
            for line in rec['text'].split('\n'):
                if line.strip():
                    expanded.append({
                        'complex': rec['complex'],
                        'text': line.strip(),
                        'loss': rec['loss']
                    })

        # Sort: gains first (positive loss), then losses (negative)
        expanded.sort(key=lambda x: (
            0 if x['loss'] > 0 else (1 if x['loss'] < 0 else 2),
            -x['loss'] if x['loss'] > 0 else (x['loss'] if x['loss'] < 0 else 0)
        ))

        bullets = []
        seen = set()
        for rec in expanded:
            cleaned = self._clean_remark_text(rec['text'])
            if not cleaned:
                continue

            if commodity_key == 'gas':
                loss_str = f"{rec['loss']:+.3f}"
            else:
                loss_str = f"{rec['loss']:+.0f}"

            prefix = f"{rec['complex']}: " if rec['complex'] and rec['complex'].lower() not in ('none', 'nil', '-', '0', 'nan', 'mh total', 'mh asset total') else ""
            bullet = f"• {prefix}{cleaned} ({loss_str})"

            if bullet not in seen:
                seen.add(bullet)
                bullets.append(bullet)
                if len(bullets) >= max_bullets:
                    break

        return bullets

    # ------------------------------------------------------------------
    # REPORT GENERATION
    # ------------------------------------------------------------------

    def generate_report(self):
        """Main entry point — load data and create report"""
        print("\n" + "=" * 80)
        print("  100 DAYS ACTUAL VS PROJECTION REPORT (COMPLETE)")
        print("=" * 80 + "\n")

        self.load_projection_data()
        self.load_actual_data()
        self.load_remarks_data()

        # Build daily date list
        daily_dates = []
        current = self.start_date
        while current <= self.end_date:
            daily_dates.append(current)
            current += timedelta(days=1)

        print(f"\nDate range: {self.start_date.strftime('%d-%b-%Y')} to {self.end_date.strftime('%d-%b-%Y')}")
        print(f"Total days: {len(daily_dates)}")

        # ---- Create workbook ----
        # IMPORTANT: Excel requires at least one visible sheet at all times.
        # We create Oil first (as the default active sheet), then Gas, WI,
        # and finally Remarks_Data. All sheets stay visible.
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Oil"

        print("\nGenerating Oil sheet...")
        self._populate_commodity_sheet_from_ws(wb, ws, daily_dates, 'oil', 'Oil Production', 'BOPD')

        print("Generating Gas sheet...")
        self._populate_commodity_sheet(wb, daily_dates, 'gas', 'Gas Production', 'MMSCMD')

        print("Generating WI sheet...")
        self._populate_commodity_sheet(wb, daily_dates, 'wi', 'Water Injection', 'BOPD')

        print("Generating Remarks_Data sheet...")
        self._create_remarks_sheet(wb, daily_dates)

        # Save workbook
        wb.save(self.output_file)
        print(f"\n  ✓ Saved: {self.output_file.name}")

        # Inject bar chart labels (bullet-point data labels in OOXML)
        try:
            self._inject_bar_chart_labels(self.output_file, daily_dates)
            print("  ✓ Injected bar chart data labels")
        except Exception as e:
            print(f"  ⚠ Label injection notice: {e}")

        # Generate VBA macro files
        self._generate_vba_file(daily_dates)
        print(f"  ✓ Created: Interactive_Charts_Macro_Dynamic.bas (Excel)")

        self._generate_ppt_vba_file(daily_dates)
        print(f"  ✓ Created: PowerPoint_Interactive_Remarks.bas (PowerPoint Slide Show)")

        print("\n" + "=" * 80)
        print(f"  ✓ COMPLETED!")
        print(f"  ✓ Excel: {self.output_file}")
        print(f"  ✓ Excel VBA: {self.base_path / 'Interactive_Charts_Macro_Dynamic.bas'}")
        print(f"  ✓ PPT VBA:   {self.base_path / 'PowerPoint_Interactive_Remarks.bas'}")
        print("=" * 80 + "\n")

        return self.output_file

    # ------------------------------------------------------------------
    # SHEET POPULATION
    # ------------------------------------------------------------------

    def _populate_commodity_sheet_from_ws(self, wb, ws, daily_dates, commodity_key, title, unit):
        """Populate an already-existing worksheet (used for the default active sheet)"""
        self._fill_commodity_sheet(ws, daily_dates, commodity_key, title, unit)

    def _populate_commodity_sheet(self, wb, daily_dates, commodity_key, title, unit):
        """Create a new sheet and populate it"""
        sheet_name = {
            'oil': 'Oil',
            'gas': 'Gas',
            'wi': 'WI'
        }.get(commodity_key, title)

        ws = wb.create_sheet(sheet_name)
        self._fill_commodity_sheet(ws, daily_dates, commodity_key, title, unit)

    def _fill_commodity_sheet(self, sheet, daily_dates, commodity_key, title, unit):
        """Fill a commodity sheet with data, formatting, and charts"""

        # Title row
        sheet.merge_cells('A1:E1')
        title_cell = sheet['A1']
        title_cell.value = f"{title} — Actual vs Projected"
        title_cell.font = Font(size=14, bold=True, color="FFFFFF")
        title_cell.fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
        title_cell.alignment = Alignment(horizontal="center", vertical="center")

        # Headers
        if commodity_key == 'oil':
            headers = ['Date', 'Actual Oil (BOPD)', 'Projected Oil (BOPD)', 'Shortfall/Gain (BOPD)', 'Deviation (%)']
        elif commodity_key == 'gas':
            headers = ['Date', 'Actual Gas (MMSCMD)', 'Projected Gas (MMSCMD)', 'Shortfall/Gain (MMSCMD)', 'Deviation (%)']
        else:
            headers = ['Date', f'Actual {title} ({unit})', f'Projected {title} ({unit})', f'Shortfall/Gain ({unit})', 'Deviation (%)']

        header_row = 2
        for col, header in enumerate(headers, start=1):
            cell = sheet.cell(row=header_row, column=col)
            cell.value = header
            cell.font = Font(bold=True, color="FFFFFF")
            cell.fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
            cell.alignment = Alignment(horizontal="center", vertical="center")
            cell.border = Border(
                left=Side(style='thin'), right=Side(style='thin'),
                top=Side(style='thin'), bottom=Side(style='thin')
            )

        # Data rows
        data_start_row = 3
        for i, date in enumerate(daily_dates):
            row = data_start_row + i
            date_key = date.date()

            # Date
            cell = sheet.cell(row=row, column=1)
            cell.value = date.strftime('%d-%b-%Y')
            cell.alignment = Alignment(horizontal="center")

            # Actual
            actual_val = self.daily_actuals.get(date_key, {}).get(commodity_key, 0)
            cell = sheet.cell(row=row, column=2)
            if commodity_key == 'gas':
                cell.value = round(actual_val, 3) if actual_val else 0
                cell.number_format = '#,##0.000'
            else:
                cell.value = round(actual_val) if actual_val else 0
                cell.number_format = '#,##0'

            # Projected
            projection = self.get_projection_for_date(date_key, commodity_key)
            cell = sheet.cell(row=row, column=3)
            if commodity_key == 'gas':
                cell.value = round(projection, 3)
                cell.number_format = '#,##0.000'
            else:
                cell.value = round(projection)
                cell.number_format = '#,##0'

            # Shortfall/Gain
            shortfall = actual_val - projection
            cell = sheet.cell(row=row, column=4)
            if commodity_key == 'gas':
                cell.value = round(shortfall, 3)
                cell.number_format = '#,##0.000'
            else:
                cell.value = round(shortfall)
                cell.number_format = '#,##0'

            if shortfall < 0:
                cell.fill = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
                cell.font = Font(color="9C0006")
            else:
                cell.fill = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
                cell.font = Font(color="006100")

            # Deviation %
            deviation_pct = (shortfall / projection * 100) if projection != 0 else 0
            cell = sheet.cell(row=row, column=5)
            cell.value = round(deviation_pct, 2)
            cell.number_format = '0.00"%"'

            # Borders
            for col in range(1, 6):
                cell = sheet.cell(row=row, column=col)
                cell.border = Border(
                    left=Side(style='thin', color='D3D3D3'),
                    right=Side(style='thin', color='D3D3D3'),
                    top=Side(style='thin', color='D3D3D3'),
                    bottom=Side(style='thin', color='D3D3D3')
                )

        # Summary row
        summary_row = data_start_row + len(daily_dates)
        sheet.cell(row=summary_row, column=1).value = "Average"
        sheet.cell(row=summary_row, column=1).font = Font(bold=True)
        sheet.cell(row=summary_row, column=2).value = f"=AVERAGE(B{data_start_row}:B{summary_row-1})"
        sheet.cell(row=summary_row, column=3).value = f"=AVERAGE(C{data_start_row}:C{summary_row-1})"
        sheet.cell(row=summary_row, column=4).value = f"=AVERAGE(D{data_start_row}:D{summary_row-1})"
        sheet.cell(row=summary_row, column=5).value = f"=AVERAGE(E{data_start_row}:E{summary_row-1})"

        for col in range(1, 6):
            cell = sheet.cell(row=summary_row, column=col)
            cell.font = Font(bold=True)
            cell.fill = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")
            if col >= 2:
                cell.number_format = '#,##0.000' if commodity_key == 'gas' else '#,##0'
                if col == 5:
                    cell.number_format = '0.00"%"'

        # Column widths
        sheet.column_dimensions['A'].width = 15
        sheet.column_dimensions['B'].width = 18
        sheet.column_dimensions['C'].width = 18
        sheet.column_dimensions['D'].width = 20
        sheet.column_dimensions['E'].width = 15

        # Charts
        self._create_combo_chart(sheet, data_start_row, summary_row - 1, title, commodity_key)

    # ------------------------------------------------------------------
    # REMARKS SHEET (VISIBLE — user can edit manually)
    # ------------------------------------------------------------------

    def _create_remarks_sheet(self, wb, daily_dates):
        """Create visible Remarks_Data sheet with Date, Commodity, Complex, Remark, Loss/Gain, Unit"""
        remarks_sheet = wb.create_sheet("Remarks_Data")
        remarks_sheet.sheet_state = 'visible'  # Explicitly visible

        # Title
        remarks_sheet.merge_cells('A1:F1')
        title_cell = remarks_sheet['A1']
        title_cell.value = "Daily Remarks — Loss/Gain Analysis"
        title_cell.font = Font(size=14, bold=True, color="FFFFFF")
        title_cell.fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
        title_cell.alignment = Alignment(horizontal="center", vertical="center")

        # Headers
        headers = ['Date', 'Commodity', 'Complex', 'Remark', 'Loss/Gain', 'Unit']
        for col, header in enumerate(headers, start=1):
            cell = remarks_sheet.cell(row=2, column=col)
            cell.value = header
            cell.font = Font(bold=True, color="FFFFFF")
            cell.fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
            cell.alignment = Alignment(horizontal="center", vertical="center")
            cell.border = Border(
                left=Side(style='thin'), right=Side(style='thin'),
                top=Side(style='thin'), bottom=Side(style='thin')
            )

        row = 3
        for date in daily_dates:
            date_key = date.date()
            date_str = date.strftime('%d-%b-%Y')

            for commodity_key, commodity_label, unit_label in [
                ('oil', 'Oil', 'BOPD'),
                ('gas', 'Gas', 'MMSCMD'),
                ('wi', 'WI', 'BOPD')
            ]:
                for rec in self.daily_loss_records.get(date_key, {}).get(commodity_key, []):
                    remarks_sheet.cell(row=row, column=1).value = date_str
                    remarks_sheet.cell(row=row, column=1).alignment = Alignment(horizontal="center")

                    remarks_sheet.cell(row=row, column=2).value = commodity_label
                    remarks_sheet.cell(row=row, column=2).alignment = Alignment(horizontal="center")

                    remarks_sheet.cell(row=row, column=3).value = self._sanitize_for_xml(rec['complex'])
                    remarks_sheet.cell(row=row, column=3).alignment = Alignment(horizontal="center")

                    cleaned = self._clean_remark_text(rec['text'])
                    remarks_sheet.cell(row=row, column=4).value = self._sanitize_for_xml(cleaned)
                    remarks_sheet.cell(row=row, column=4).alignment = Alignment(horizontal="left", wrap_text=True)

                    loss_cell = remarks_sheet.cell(row=row, column=5)
                    loss_cell.value = rec['loss']
                    loss_cell.number_format = '#,##0.000' if commodity_key == 'gas' else '#,##0'
                    loss_cell.alignment = Alignment(horizontal="right")

                    if rec['loss'] < 0:
                        loss_cell.fill = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
                        loss_cell.font = Font(color="9C0006")
                    elif rec['loss'] > 0:
                        loss_cell.fill = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
                        loss_cell.font = Font(color="006100")

                    remarks_sheet.cell(row=row, column=6).value = unit_label
                    remarks_sheet.cell(row=row, column=6).alignment = Alignment(horizontal="center")

                    for c in range(1, 7):
                        remarks_sheet.cell(row=row, column=c).border = Border(
                            left=Side(style='thin', color='D3D3D3'),
                            right=Side(style='thin', color='D3D3D3'),
                            top=Side(style='thin', color='D3D3D3'),
                            bottom=Side(style='thin', color='D3D3D3')
                        )

                    row += 1

        # Column widths
        remarks_sheet.column_dimensions['A'].width = 15
        remarks_sheet.column_dimensions['B'].width = 12
        remarks_sheet.column_dimensions['C'].width = 15
        remarks_sheet.column_dimensions['D'].width = 70
        remarks_sheet.column_dimensions['E'].width = 12
        remarks_sheet.column_dimensions['F'].width = 10

        # Freeze header
        remarks_sheet.freeze_panes = 'A3'

        print(f"  Remarks_Data: {row - 3} rows written (sheet is VISIBLE)")

    # ------------------------------------------------------------------
    # CHARTS
    # ------------------------------------------------------------------

    def _create_combo_chart(self, sheet, data_start_row, data_end_row, title, commodity_key):
        """Create line chart (Actual vs Projected) and bar chart (Shortfall/Gain)"""

        if commodity_key == 'gas':
            y_label = "Gas Production (MMSCMD)"
        elif commodity_key == 'wi':
            y_label = "Water Injection (BOPD)"
        else:
            y_label = "Oil Production (BOPD)"

        dates = Reference(sheet, min_col=1, min_row=data_start_row, max_row=data_end_row)
        actual_data = Reference(sheet, min_col=2, min_row=data_start_row - 1, max_row=data_end_row)
        projected_data = Reference(sheet, min_col=3, min_row=data_start_row - 1, max_row=data_end_row)
        shortfall_data = Reference(sheet, min_col=4, min_row=data_start_row - 1, max_row=data_end_row)

        # --- LINE CHART ---
        line_chart = LineChart()
        line_chart.title = f"{title} — Actual vs Projected"
        line_chart.style = 2
        line_chart.height = 12
        line_chart.width = 24

        line_chart.legend.legendPos = "b"
        line_chart.legend.overlay = False

        line_chart.add_data(actual_data, titles_from_data=True)
        line_chart.add_data(projected_data, titles_from_data=True)
        line_chart.set_categories(dates)

        line_chart.y_axis.delete = False
        line_chart.y_axis.title = y_label
        line_chart.y_axis.numFmt = '#,##0' if commodity_key != 'gas' else '#,##0.000'
        line_chart.y_axis.tickLblPos = "nextTo"
        line_chart.y_axis.majorTickMark = "out"
        line_chart.y_axis.majorGridlines = None
        line_chart.y_axis.crosses = "autoZero"

        line_chart.x_axis.delete = False
        line_chart.x_axis.title = "Date"
        line_chart.x_axis.tickLblPos = "low"
        line_chart.x_axis.majorTickMark = "out"
        line_chart.x_axis.numFmt = "dd-mmm"

        # Actual — solid blue
        s0 = line_chart.series[0]
        s0.smooth = True
        s0.graphicalProperties.line.solidFill = "0070C0"
        s0.graphicalProperties.line.width = 28000
        s0.marker.symbol = "circle"
        s0.marker.size = 7
        s0.marker.graphicalProperties.solidFill = "0070C0"
        s0.marker.graphicalProperties.line.solidFill = "0070C0"

        # Projected — dashed red
        s1 = line_chart.series[1]
        s1.smooth = True
        s1.graphicalProperties.line.solidFill = "FF0000"
        s1.graphicalProperties.line.width = 28000
        s1.graphicalProperties.line.dashStyle = "dash"
        s1.marker.symbol = "square"
        s1.marker.size = 7
        s1.marker.graphicalProperties.solidFill = "FF0000"
        s1.marker.graphicalProperties.line.solidFill = "FF0000"

        # Plot area border
        if line_chart.plot_area.graphicalProperties is None:
            line_chart.plot_area.graphicalProperties = GraphicalProperties()
        if line_chart.plot_area.graphicalProperties.line is None:
            line_chart.plot_area.graphicalProperties.line = LineProperties()
        line_chart.plot_area.graphicalProperties.line.width = 19050
        line_chart.plot_area.graphicalProperties.line.solidFill = "000000"

        sheet.add_chart(line_chart, "G2")

        # --- BAR CHART ---
        bar_chart = BarChart()
        bar_chart.type = "col"
        bar_chart.title = f"{title} — Daily Shortfall / Gain"
        bar_chart.style = 2
        bar_chart.height = 12
        bar_chart.width = 24
        bar_chart.overlap = 100

        bar_chart.legend.legendPos = "b"
        bar_chart.legend.overlay = False

        bar_chart.add_data(shortfall_data, titles_from_data=True)
        bar_chart.set_categories(dates)

        bar_chart.y_axis.delete = False
        bar_chart.y_axis.title = f"Shortfall / Gain ({'MMSCMD' if commodity_key == 'gas' else 'BOPD'})"
        bar_chart.y_axis.numFmt = '#,##0' if commodity_key != 'gas' else '#,##0.000'
        bar_chart.y_axis.tickLblPos = "nextTo"
        bar_chart.y_axis.majorTickMark = "out"
        bar_chart.y_axis.majorGridlines = None
        bar_chart.y_axis.crosses = "autoZero"

        bar_chart.x_axis.delete = False
        bar_chart.x_axis.title = "Date"
        bar_chart.x_axis.tickLblPos = "low"
        bar_chart.x_axis.majorTickMark = "out"
        bar_chart.x_axis.numFmt = "dd-mmm"

        # Orange bars
        bar_chart.series[0].graphicalProperties.solidFill = "FFA500"
        bar_chart.series[0].graphicalProperties.line.solidFill = "E07B00"

        if bar_chart.plot_area.graphicalProperties is None:
            bar_chart.plot_area.graphicalProperties = GraphicalProperties()
        if bar_chart.plot_area.graphicalProperties.line is None:
            bar_chart.plot_area.graphicalProperties.line = LineProperties()
        bar_chart.plot_area.graphicalProperties.line.width = 19050
        bar_chart.plot_area.graphicalProperties.line.solidFill = "000000"

        sheet.add_chart(bar_chart, "G22")

    # ------------------------------------------------------------------
    # BAR CHART LABEL INJECTION (OOXML)
    # ------------------------------------------------------------------

    def _inject_bar_chart_labels(self, filepath, daily_dates):
        """Inject custom DrawingML bullet-point data labels onto Bar Charts"""
        labels_by_comm = {'oil': [], 'gas': [], 'wi': []}

        for d in daily_dates:
            d_key = d.date()
            act = self.daily_actuals.get(d_key, {})

            for comm, unit_fmt, fmt_fn in [
                ('oil', 'BOPD', lambda v: f"{v:+,.0f}"),
                ('gas', 'MMSCMD', lambda v: f"{v:+.3f}"),
                ('wi', 'BOPD', lambda v: f"{v:+,.0f}")
            ]:
                act_val = act.get(comm, 0)
                proj_val = self.get_projection_for_date(d_key, comm)
                sf = act_val - proj_val
                bullets = self.get_daily_loss_bullets(d_key, comm, sf)
                val_str = f"{fmt_fn(sf)} {unit_fmt}" if sf != 0 else f"0 {unit_fmt}"
                labels_by_comm[comm].append({'val': val_str, 'bullets': bullets, 'sf': sf})

        temp_buf = io.BytesIO()
        with zipfile.ZipFile(filepath, 'r') as zin:
            with zipfile.ZipFile(temp_buf, 'w', zipfile.ZIP_DEFLATED) as zout:
                for item in zin.infolist():
                    content = zin.read(item.filename)
                    if item.filename.startswith('xl/charts/chart') and item.filename.endswith('.xml'):
                        try:
                            root = ET.fromstring(content)
                            ET.register_namespace('c', 'http://schemas.openxmlformats.org/drawingml/2006/chart')
                            ET.register_namespace('a', 'http://schemas.openxmlformats.org/drawingml/2006/main')
                            ET.register_namespace('r', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships')

                            bar_chart = root.find('.//{http://schemas.openxmlformats.org/drawingml/2006/chart}barChart')
                            if bar_chart is not None:
                                content_str = content.decode('utf-8', errors='ignore')
                                if 'Gas' in content_str or 'MMSCMD' in content_str:
                                    comm_key = 'gas'
                                elif 'Water' in content_str or 'WI' in content_str:
                                    comm_key = 'wi'
                                else:
                                    comm_key = 'oil'

                                labels_list = labels_by_comm.get(comm_key, [])
                                ser = bar_chart.find('{http://schemas.openxmlformats.org/drawingml/2006/chart}ser')
                                if ser is not None and labels_list:
                                    existing = ser.find('{http://schemas.openxmlformats.org/drawingml/2006/chart}dLbls')
                                    if existing is not None:
                                        ser.remove(existing)

                                    c_ns = 'http://schemas.openxmlformats.org/drawingml/2006/chart'
                                    a_ns = 'http://schemas.openxmlformats.org/drawingml/2006/main'

                                    dLbls = ET.Element(f'{{{c_ns}}}dLbls')
                                    pos = ET.SubElement(dLbls, f'{{{c_ns}}}dLblPos')
                                    pos.set('val', 'outEnd')

                                    for flag in ['showLegendKey', 'showVal', 'showCatName', 'showSerName', 'showPercent', 'showBubbleSize']:
                                        el = ET.SubElement(dLbls, f'{{{c_ns}}}{flag}')
                                        el.set('val', '0')

                                    for idx, info in enumerate(labels_list):
                                        dLbl = ET.SubElement(dLbls, f'{{{c_ns}}}dLbl')
                                        idx_el = ET.SubElement(dLbl, f'{{{c_ns}}}idx')
                                        idx_el.set('val', str(idx))

                                        tx = ET.SubElement(dLbl, f'{{{c_ns}}}tx')
                                        rich = ET.SubElement(tx, f'{{{c_ns}}}rich')

                                        bodyPr = ET.SubElement(rich, f'{{{a_ns}}}bodyPr')
                                        bodyPr.set('wrap', 'square')
                                        bodyPr.set('lIns', '20000')
                                        bodyPr.set('rIns', '20000')
                                        bodyPr.set('tIns', '10000')
                                        bodyPr.set('bIns', '10000')

                                        ET.SubElement(rich, f'{{{a_ns}}}lstStyle')

                                        # Value line
                                        p1 = ET.SubElement(rich, f'{{{a_ns}}}p')
                                        pPr1 = ET.SubElement(p1, f'{{{a_ns}}}pPr')
                                        pPr1.set('algn', 'ctr')
                                        r1 = ET.SubElement(p1, f'{{{a_ns}}}r')
                                        rPr1 = ET.SubElement(r1, f'{{{a_ns}}}rPr')
                                        rPr1.set('sz', '750')
                                        rPr1.set('b', '1')
                                        clr1 = ET.SubElement(rPr1, f'{{{a_ns}}}solidFill')
                                        srgb1 = ET.SubElement(clr1, f'{{{a_ns}}}srgbClr')
                                        srgb1.set('val', '9C0006' if info['sf'] < 0 else ('006100' if info['sf'] > 0 else '000000'))
                                        t1 = ET.SubElement(r1, f'{{{a_ns}}}t')
                                        t1.text = self._sanitize_for_xml(info['val'])

                                        # Bullet lines
                                        for bullet_text in info.get('bullets', []):
                                            p_b = ET.SubElement(rich, f'{{{a_ns}}}p')
                                            pPr_b = ET.SubElement(p_b, f'{{{a_ns}}}pPr')
                                            pPr_b.set('algn', 'ctr')
                                            r_b = ET.SubElement(p_b, f'{{{a_ns}}}r')
                                            rPr_b = ET.SubElement(r_b, f'{{{a_ns}}}rPr')
                                            rPr_b.set('sz', '550')
                                            clr_b = ET.SubElement(rPr_b, f'{{{a_ns}}}solidFill')
                                            srgb_b = ET.SubElement(clr_b, f'{{{a_ns}}}srgbClr')
                                            srgb_b.set('val', '000000')
                                            t_b = ET.SubElement(r_b, f'{{{a_ns}}}t')
                                            t_b.text = self._sanitize_for_xml(bullet_text)

                                    ser.append(dLbls)

                            content = ET.tostring(root, encoding='utf-8', xml_declaration=True)
                        except Exception:
                            pass
                    zout.writestr(item, content)

        with open(filepath, 'wb') as f:
            f.write(temp_buf.getvalue())

    # ------------------------------------------------------------------
    # VBA MACRO GENERATION
    # ------------------------------------------------------------------

    def _clean_vba_literal(self, text):
        """Sanitize text for clean, error-free VBA string literals"""
        if not text:
            return ""
        s = str(text)
        # Normalize quotes and dashes
        s = s.replace('“', '"').replace('”', '"')
        s = s.replace('‘', "'").replace('’', "'")
        s = s.replace('–', '-').replace('—', '-')
        s = s.replace('\n', ' ').replace('\r', ' ').replace('\t', ' ')
        # Keep ASCII characters only to avoid VBA code page syntax errors
        s = s.encode('ascii', 'ignore').decode('ascii')
        s = ' '.join(s.split())
        if len(s) > 115:
            s = s[:112] + '...'
        # Escape double quotes for VBA
        s = s.replace('"', '""')
        return s

    def _build_full_vba_code(self, daily_dates, file_type="PPT"):
        """Build universal, robust VBA macro code supporting both Excel and PowerPoint"""
        sample_date = daily_dates[0].strftime('%d-%b-%Y') if daily_dates else '11-Sep-2026'

        date_list_vba = []
        add_rec_lines = []
        total_records = 0

        for d in daily_dates:
            d_key = d.date()
            d_str = d.strftime('%d-%b-%Y')
            date_list_vba.append(d_str)

            for comm_key, comm_label, unit_label in [
                ('oil', 'Oil', 'BOPD'),
                ('gas', 'Gas', 'MMSCMD'),
                ('wi', 'WI', 'BOPD')
            ]:
                for rec in self.daily_loss_records.get(d_key, {}).get(comm_key, []):
                    comp = self._clean_vba_literal(rec.get('complex') or '')
                    cleaned = self._clean_vba_literal(self._clean_remark_text(rec.get('text') or ''))
                    loss = rec.get('loss', 0)
                    try:
                        loss_val = f"{float(loss):.3f}"
                    except Exception:
                        loss_val = "0.0"

                    if cleaned:
                        add_rec_lines.append(
                            f'    AddRec "{d_str}", "{comm_label}", "{comp}", "{cleaned}", {loss_val}#, "{unit_label}"'
                        )
                        total_records += 1

        add_records_vba_str = "\n".join(add_rec_lines)
        date_array_items = ', '.join([f'"{d}"' for d in date_list_vba])
        alloc_size = max(total_records + 50, 100)

        header_title = "POWERPOINT & EXCEL INTERACTIVE REMARKS MACRO"
        desc_line = "Compatible with PowerPoint & Excel on Windows & macOS (Office 365 / 2021 / 2019)"

        vba_code = f"""Option Explicit

' ==============================================================================
' {header_title}
' {desc_line}
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
            If UBound(m_Dates) >= LBound(m_Dates) Then dateValue = CStr(m_Dates(LBound(m_Dates))) Else dateValue = "{sample_date}"
        Else
            dateValue = "{sample_date}"
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
    msg = BuildSingleCommodityRemarks("{sample_date}", "Oil")
    MsgBox msg, vbInformation, "Oil Remarks Preview — {sample_date}"
End Sub

Public Sub TestGasRemarks()
    Dim msg As String
    msg = BuildSingleCommodityRemarks("{sample_date}", "Gas")
    MsgBox msg, vbInformation, "Gas Remarks Preview — {sample_date}"
End Sub

Public Sub TestWIRemarks()
    Dim msg As String
    msg = BuildSingleCommodityRemarks("{sample_date}", "WI")
    MsgBox msg, vbInformation, "WI Remarks Preview — {sample_date}"
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

    m_Dates = Array({date_array_items})
    m_Count = 0
    ReDim m_Remarks(0 To {alloc_size})

{add_records_vba_str}

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
"""
        return vba_code

    def _generate_vba_file(self, daily_dates):
        """Generate Interactive_Charts_Macro_Dynamic.bas"""
        vba_code = self._build_full_vba_code(daily_dates, file_type="Excel")
        vba_file = self.base_path / "Interactive_Charts_Macro_Dynamic.bas"
        with open(vba_file, 'w', encoding='utf-8') as f:
            f.write(vba_code)

    def _generate_ppt_vba_file(self, daily_dates):
        """Generate PowerPoint_Interactive_Remarks.bas"""
        ppt_vba_code = self._build_full_vba_code(daily_dates, file_type="PPT")
        ppt_vba_file = self.base_path / "PowerPoint_Interactive_Remarks.bas"
        with open(ppt_vba_file, 'w', encoding='utf-8') as f:
            f.write(ppt_vba_code)


# ======================================================================
# CLI
# ======================================================================

def parse_date(date_str):
    """Parse date string in various formats"""
    formats = [
        '%d-%m-%Y',    # 11-09-2026
        '%Y-%m-%d',    # 2026-09-11
        '%d/%m/%Y',    # 11/09/2026
        '%Y/%m/%d',    # 2026/09/11
        '%d.%m.%Y',    # 11.09.2026
        '%Y.%m.%d',    # 2026.09.11
    ]
    for fmt in formats:
        try:
            return datetime.strptime(date_str, fmt)
        except ValueError:
            continue
    raise ValueError(f"Could not parse date: {date_str}. Use DD-MM-YYYY or YYYY-MM-DD.")


def main():
    if len(sys.argv) != 3:
        print("\n" + "=" * 80)
        print("  100 Days Actual vs Projection Report (Complete)")
        print("=" * 80)
        print("\nUsage:")
        print("  python3 generate_100_days_projection_complete.py <start_date> <end_date>")
        print("\nDate formats:")
        print("  DD-MM-YYYY  (e.g. 11-09-2026)")
        print("  YYYY-MM-DD  (e.g. 2026-09-11)")
        print("\nExample:")
        print("  python3 generate_100_days_projection_complete.py 11-09-2026 01-10-2026")
        print("\nOutput:")
        print("  100_Days_Actual_vs_Projection.xlsx   (in MH_Dashboard folder)")
        print("  Interactive_Charts_Macro_Dynamic.bas  (in MH_Dashboard folder)")
        print("=" * 80 + "\n")
        return 1

    base_path = "/Users/amanmishra/Library/CloudStorage/OneDrive-OilandNaturalGasCorporationLimited/MH_Dashboard"

    try:
        start_date = parse_date(sys.argv[1])
        end_date = parse_date(sys.argv[2])

        if end_date < start_date:
            print("❌ ERROR: End date must be after start date!")
            return 1

        generator = ReportGenerator(base_path, start_date, end_date)
        filepath = generator.generate_report()

        print(f"📄 Report saved: {filepath.name}")
        return 0

    except ValueError as e:
        print(f"\n❌ DATE ERROR: {e}")
        return 1
    except Exception as e:
        print(f"\n❌ ERROR: {e}", file=sys.stderr)
        import traceback
        traceback.print_exc()
        return 1


if __name__ == "__main__":
    sys.exit(main())
