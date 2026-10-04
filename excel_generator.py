import os
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.utils import get_column_letter
import pandas as pd
from typing import Dict, Any

class ExcelDashboardGenerator:
    """
    Builds a styled, interactive Excel Workbook for job application tracking.
    Features:
    - Native openpyxl hyperlinks (handles long UTM tracking URLs safely)
    - Dropdown Data Validation for Application Status
    - Highlighted skill sections and experience levels
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

    def generate(self, jobs_df: pd.DataFrame, output_path: str = None) -> str:
        """
        Creates and styles the Excel spreadsheet from the enriched DataFrame.
        """
        if output_path is None:
            output_path = self.filename

        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = self.sheet_name
        ws.views.sheetView[0].showGridLines = True

        # Define Columns
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
            self._safe_save(wb, output_path)
            return output_path

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
            c11 = ws.cell(row=row_idx, column=11, value="Not Applied")
            c12 = ws.cell(row=row_idx, column=12, value="")

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
                showDropDown=True
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

        # Dynamic Column Widths
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

        saved_file = self._safe_save(wb, output_path)
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
