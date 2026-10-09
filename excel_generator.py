import os
import re
import json
import glob
from datetime import datetime
from typing import Dict, Any, Optional

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.utils import get_column_letter
import pandas as pd


class ExcelDashboardGenerator:
    """
    Builds a styled, interactive Excel Workbook for job application tracking.
    Features:
    - Native openpyxl hyperlinks (handles long UTM tracking URLs safely)
    - Persistent Application Status & Notes memory via applications.json
    - Automatic sync from the single most recent tracker file (newest-wins precedence)
    - Graceful fallback on PermissionError if yesterday's file is open in Excel
    - Dual-output generation: timestamped archive in leads/ and stable Tracker_Latest.xlsx
    - Tab 2 Run Summary audit dashboard
    - Dropdown Data Validation for Application Status
    - Clean professional theme with freeze panes & auto-filters
    - Graceful fallback if file is currently open in Excel
    """

    def __init__(self, config: Dict[str, Any]):
        self.output_config = config.get("output", {})
        self.filename = self.output_config.get("filename", "Job_Applications_Tracker.xlsx")
        self.sheet_name = self.output_config.get("sheet_name", "Active Job Leads")
        self.status_options = self.output_config.get("status_options", [
            "Not Applied", "Applied", "Referral Requested", 
            "Screening Call", "Interviewing", "Assessment", "Rejected", "Offer Received"
        ])
        self.leads_dir = self.output_config.get("leads_dir", "leads")
        self.stable_filename = self.output_config.get("stable_filename", "Tracker_Latest.xlsx")
        self.applications_file = self.output_config.get("applications_file", "applications.json")
        self.applications_ledger = self._sync_and_load_applications()

    def _sync_and_load_applications(self) -> Dict[str, Dict[str, str]]:
        """
        Loads existing applications.json memory as source of truth.
        Syncs newly recorded statuses/notes from the single newest previous workbook.
        Precedence:
          1. Newest timestamped file in leads/ (Job_Leads_*.xlsx)
          2. Fallback to stable_filename or legacy filename
        Safely catches PermissionError if a file is currently open in Excel.
        Matches by header names ('Application Status', 'Referral / Notes', 'Apply Action').
        """
        ledger = {}
        if os.path.exists(self.applications_file):
            try:
                with open(self.applications_file, "r", encoding="utf-8") as f:
                    data = json.load(f)
                    if isinstance(data, dict):
                        ledger = data
            except Exception as e:
                print(f"[!] Warning: Could not read '{self.applications_file}': {e}")

        # Locate single newest historical file
        latest_file: Optional[str] = None
        if os.path.isdir(self.leads_dir):
            pattern = os.path.join(self.leads_dir, "Job_Leads_*.xlsx")
            lead_files = glob.glob(pattern)
            if lead_files:
                lead_files.sort(reverse=True)
                latest_file = lead_files[0]

        if not latest_file:
            for candidate in [self.stable_filename, self.filename]:
                if candidate and os.path.exists(candidate):
                    latest_file = candidate
                    break

        if latest_file and os.path.exists(latest_file):
            try:
                wb = openpyxl.load_workbook(latest_file, data_only=True)
                if self.sheet_name in wb.sheetnames:
                    ws = wb[self.sheet_name]
                else:
                    ws = wb.active

                headers = [str(cell.value or "").strip() for cell in ws[1]]
                header_indices = {}
                for idx, h in enumerate(headers, 1):
                    h_lower = h.lower()
                    if "application status" in h_lower:
                        header_indices["status"] = idx
                    elif "referral / notes" in h_lower or "notes" in h_lower:
                        header_indices["notes"] = idx
                    elif "apply action" in h_lower or "apply" in h_lower:
                        header_indices["apply"] = idx
                    elif "company" in h_lower:
                        header_indices["company"] = idx
                    elif "title" in h_lower:
                        header_indices["title"] = idx

                status_col = header_indices.get("status")
                notes_col = header_indices.get("notes")
                apply_col = header_indices.get("apply")
                company_col = header_indices.get("company")
                title_col = header_indices.get("title")

                if status_col or notes_col:
                    updates_count = 0
                    for row in range(2, ws.max_row + 1):
                        status_val = str(ws.cell(row=row, column=status_col).value or "").strip() if status_col else ""
                        notes_val = str(ws.cell(row=row, column=notes_col).value or "").strip() if notes_col else ""

                        if (status_val in ["", "Not Applied", "None"]) and not notes_val:
                            continue

                        apply_cell = ws.cell(row=row, column=apply_col) if apply_col else None
                        url_target = ""
                        if apply_cell:
                            if apply_cell.hyperlink and apply_cell.hyperlink.target:
                                url_target = apply_cell.hyperlink.target
                            elif apply_cell.value and str(apply_cell.value).startswith("http"):
                                url_target = str(apply_cell.value).strip()

                        norm_url = re.sub(r'\?.*$', '', url_target).strip().lower() if url_target else ""
                        company_val = str(ws.cell(row=row, column=company_col).value or "").strip().lower() if company_col else ""
                        title_val = str(ws.cell(row=row, column=title_col).value or "").strip().lower() if title_col else ""

                        key = norm_url if norm_url.startswith("http") else (f"{company_val}::{title_val}" if company_val and title_val else "")
                        if key:
                            ledger[key] = {
                                "status": status_val if status_val and status_val != "None" else "Not Applied",
                                "notes": notes_val,
                                "last_updated": datetime.now().strftime("%Y-%m-%d")
                            }
                            updates_count += 1

                    if updates_count > 0:
                        try:
                            parent_dir = os.path.dirname(os.path.abspath(self.applications_file))
                            os.makedirs(parent_dir, exist_ok=True)
                            with open(self.applications_file, "w", encoding="utf-8") as f:
                                json.dump(ledger, f, indent=2)
                        except Exception as e:
                            print(f"[!] Warning: Could not save '{self.applications_file}': {e}")

            except PermissionError:
                print(f"[!] Warning: Tracker file '{latest_file}' is currently locked in Excel.")
                print(f"    Proceeding safely with existing '{self.applications_file}' memory.")
            except Exception as e:
                print(f"[!] Warning: Could not sync from '{latest_file}': {e}")

        return ledger

    def generate(self, jobs_df: pd.DataFrame, output_path: str = None, run_stats: dict = None) -> str:
        """
        Creates and styles the Excel spreadsheet from the enriched DataFrame.
        Produces:
          - Tab 1: Active Job Leads (with status/notes carried over from applications.json)
          - Tab 2: Run Summary (detailed audit statistics)
          - Dual outputs: timestamped archive in leads/ and stable Tracker_Latest.xlsx in root.
        """
        save_stable = False
        if output_path is not None:
            primary_target = output_path
        else:
            os.makedirs(self.leads_dir, exist_ok=True)
            now_str = datetime.now().strftime("%Y-%m-%d_%H%M")
            primary_target = os.path.join(self.leads_dir, f"Job_Leads_{now_str}.xlsx")
            save_stable = True

        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = self.sheet_name
        ws.views.sheetView[0].showGridLines = True

        # Define Columns for Tab 1
        headers = [
            "Sl. No",
            "Lead Status",
            "Date Posted",
            "Company Name",
            "Job Title / Role",
            "Main Skills Required",
            "Experience Level",
            "Location / Mode",
            "Platform",
            "Apply Action",
            "Application Status",
            "Referral / Notes"
        ]

        ws.append(headers)

        # Theme Colors & Styles
        header_fill = PatternFill(start_color="0F172A", end_color="0F172A", fill_type="solid")  # Dark Slate
        header_font = Font(name="Segoe UI", size=11, bold=True, color="FFFFFF")
        
        row_alt_fill = PatternFill(start_color="F8FAFC", end_color="F8FAFC", fill_type="solid") # Soft cool grey
        row_white_fill = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")
        
        apply_btn_fill = PatternFill(start_color="2563EB", end_color="2563EB", fill_type="solid") # Vibrant Blue
        apply_btn_font = Font(name="Segoe UI", size=10, bold=True, color="FFFFFF", underline="single")
        
        skill_font = Font(name="Segoe UI Semibold", size=10, color="047857") # Emerald Green
        standard_font = Font(name="Segoe UI", size=10, color="1E293B")
        company_font = Font(name="Segoe UI", size=10, bold=True, color="0F172A")
        applied_font = Font(name="Segoe UI Semibold", size=10, color="2563EB")
        
        thin_border = Border(
            left=Side(style='thin', color='E2E8F0'),
            right=Side(style='thin', color='E2E8F0'),
            top=Side(style='thin', color='E2E8F0'),
            bottom=Side(style='thin', color='E2E8F0')
        )

        # Style Header Row
        ws.row_dimensions[1].height = 28
        for col_idx, header in enumerate(headers, 1):
            cell = ws.cell(row=1, column=col_idx)
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            cell.border = thin_border

        if jobs_df is None or jobs_df.empty:
            ws.append([1, "N/A", "N/A", "No Jobs Found", "Try adjusting search query or location", "-", "-", "-", "-", "No Link", "Not Applied", ""])
        else:
            # Populate Job Rows
            for idx, row in jobs_df.iterrows():
                row_idx = ws.max_row + 1
                ws.row_dimensions[row_idx].height = 24
                
                is_even = (row_idx % 2 == 0)
                current_fill = row_alt_fill if is_even else row_white_fill

                lead_status = str(row.get("lead_status", "🟢 NEW"))
                date_posted = str(row.get("formatted_date", "Recent"))
                company = str(row.get("company", "Hiring Company"))
                title = str(row.get("title", "Job Role"))
                skills = str(row.get("extracted_skills", "Data Analysis, Problem Solving"))
                exp = str(row.get("extracted_experience", "Not Specified"))
                work_mode = str(row.get("work_mode", row.get("location", "Hyderabad, India")))
                platform = str(row.get("site", "Web")).capitalize()
                apply_url = str(row.get("final_apply_url", ""))

                # Memory lookup: check URL or company::title in applications_ledger
                norm_url = re.sub(r'\?.*$', '', apply_url).strip().lower() if apply_url else ""
                fallback_key = f"{company.strip().lower()}::{title.strip().lower()}"
                saved_entry = self.applications_ledger.get(norm_url) or self.applications_ledger.get(fallback_key)
                
                if saved_entry:
                    status_val = saved_entry.get("status", "Not Applied")
                    notes_val = saved_entry.get("notes", "")
                else:
                    status_val = "Not Applied"
                    notes_val = ""

                # Populate cells
                c1 = ws.cell(row=row_idx, column=1, value=idx + 1)
                c2 = ws.cell(row=row_idx, column=2, value=lead_status)
                c3 = ws.cell(row=row_idx, column=3, value=date_posted)
                c4 = ws.cell(row=row_idx, column=4, value=company)
                c5 = ws.cell(row=row_idx, column=5, value=title)
                c6 = ws.cell(row=row_idx, column=6, value=skills)
                c7 = ws.cell(row=row_idx, column=7, value=exp)
                c8 = ws.cell(row=row_idx, column=8, value=work_mode)
                c9 = ws.cell(row=row_idx, column=9, value=platform)
                c10 = ws.cell(row=row_idx, column=10)
                c11 = ws.cell(row=row_idx, column=11, value=status_val)
                c12 = ws.cell(row=row_idx, column=12, value=notes_val)

                for c in [c1, c2, c3, c4, c5, c6, c7, c8, c9, c11, c12]:
                    c.fill = current_fill
                    c.border = thin_border
                    c.alignment = Alignment(vertical="center")

                c1.alignment = Alignment(horizontal="center", vertical="center")
                if "NEW" in lead_status:
                    c2.font = Font(name="Segoe UI", size=10, bold=True, color="047857")
                else:
                    c2.font = Font(name="Segoe UI", size=10, color="64748B")
                c2.alignment = Alignment(horizontal="center", vertical="center")
                c3.alignment = Alignment(horizontal="center", vertical="center")
                c4.font = company_font
                c5.font = standard_font
                c6.font = skill_font
                c7.alignment = Alignment(horizontal="center", vertical="center")
                c9.alignment = Alignment(horizontal="center", vertical="center")
                c11.alignment = Alignment(horizontal="center", vertical="center")
                if status_val and status_val != "Not Applied":
                    c11.font = applied_font

                # Native openpyxl Hyperlink
                c10.border = thin_border
                if apply_url and apply_url.startswith("http"):
                    c10.value = "Apply Now"
                    c10.hyperlink = apply_url
                    c10.font = apply_btn_font
                    c10.fill = apply_btn_fill
                    c10.alignment = Alignment(horizontal="center", vertical="center")
                else:
                    c10.value = "Link Unavailable"
                    c10.font = standard_font
                    c10.fill = current_fill
                    c10.alignment = Alignment(horizontal="center", vertical="center")

        # Add Status Dropdown Data Validation (Column K)
        if ws.max_row >= 2:
            status_list_str = f'"{",".join(self.status_options)}"'
            status_dv = DataValidation(
                type="list",
                formula1=status_list_str,
                allow_blank=True,
                showErrorMessage=True
            )
            status_dv.error = "Please select a valid application status"
            status_dv.errorTitle = "Invalid Status"
            ws.add_data_validation(status_dv)
            status_dv.add(f"K2:K{ws.max_row}")

        # Enable Auto-filter across all columns
        last_col_letter = get_column_letter(len(headers))
        ws.auto_filter.ref = f"A1:{last_col_letter}{ws.max_row}"

        # Freeze Header Row
        ws.freeze_panes = "A2"

        # Dynamic Column Widths for Tab 1
        col_widths = {
            1: 8,   # Sl No
            2: 15,  # Lead Status
            3: 14,  # Date Posted
            4: 26,  # Company Name
            5: 34,  # Job Title
            6: 35,  # Main Skills Required
            7: 22,  # Experience Level
            8: 22,  # Location / Mode
            9: 15,  # Platform
            10: 18, # Apply Action
            11: 20, # Application Status
            12: 28  # Notes
        }

        for col_idx, width in col_widths.items():
            col_letter = get_column_letter(col_idx)
            ws.column_dimensions[col_letter].width = width

        # Build Tab 2: Run Summary
        ws_summary = wb.create_sheet(title="Run Summary")
        ws_summary.views.sheetView[0].showGridLines = True
        ws_summary.append(["Audit Metric", "Value"])
        ws_summary.row_dimensions[1].height = 26

        for col_idx in [1, 2]:
            cell = ws_summary.cell(row=1, column=col_idx)
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(horizontal="center", vertical="center")
            cell.border = thin_border

        stats = run_stats or {}
        mode_val = str(stats.get("mode", "frugal")).upper()
        total_leads_val = len(jobs_df) if jobs_df is not None else 0
        fc_cands = stats.get("fc_candidates", 0)
        credits_val = stats.get("credits_used", 0)
        sr_dropped = stats.get("dropped_senior_title", 0)
        ceil_dropped = stats.get("dropped_experience_ceiling", 0)
        ts_val = stats.get("timestamp", datetime.now().strftime("%Y-%m-%d %H:%M"))

        summary_rows = [
            ("Run Timestamp", ts_val),
            ("Scraping Mode", mode_val),
            ("Total Raw Candidates Gathered", fc_cands + total_leads_val),
            ("Senior Roles Filtered", sr_dropped),
            ("Experience Ceiling Filtered (> 2 YOE)", ceil_dropped),
            ("Final Verified Leads Retained", total_leads_val),
            ("Firecrawl Credits Consumed", credits_val)
        ]

        for r_idx, (metric, val) in enumerate(summary_rows, 2):
            ws_summary.row_dimensions[r_idx].height = 22
            c_m = ws_summary.cell(row=r_idx, column=1, value=metric)
            c_v = ws_summary.cell(row=r_idx, column=2, value=val)

            c_fill = row_alt_fill if (r_idx % 2 == 0) else row_white_fill
            c_m.fill = c_fill
            c_m.font = standard_font
            c_m.border = thin_border
            c_m.alignment = Alignment(vertical="center", indent=1)

            c_v.fill = c_fill
            c_v.font = Font(name="Segoe UI Semibold", size=10, color="0F172A")
            c_v.border = thin_border
            c_v.alignment = Alignment(horizontal="center", vertical="center")

        ws_summary.column_dimensions["A"].width = 40
        ws_summary.column_dimensions["B"].width = 24

        saved_file = self._safe_save(wb, primary_target)
        if save_stable:
            self._safe_save(wb, self.stable_filename)

        print(f"\n[+] Excel Workbook successfully created: {os.path.abspath(saved_file)}")
        return saved_file

    def _safe_save(self, wb, output_path: str) -> str:
        """Saves workbook safely, falling back if the target file is locked by Excel."""
        try:
            wb.save(output_path)
            return output_path
        except PermissionError:
            base, ext = os.path.splitext(output_path)
            fallback_path = f"{base}_Latest{ext}"
            print(f"\n[!] Warning: '{output_path}' is currently open in Excel.")
            print(f"    Saving clean copy to: '{fallback_path}' instead.")
            wb.save(fallback_path)
            return fallback_path
