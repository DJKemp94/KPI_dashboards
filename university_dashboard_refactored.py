import pandas as pd
import json
import tkinter as tk
from tkinter import filedialog, messagebox
import os


class UniversityDashboardGenerator:
    def __init__(self):
        self.university_data = None
        self.university_history_data = None
        self.kpi_definitions = {
            "Written Arrangements Complete": {"percentage_col": "% of Written Arrangements Complete", "number_col": "Number of Arrangements", "completed_col": "Number of Arrangements Completed"},
            "Risk Assessments in Register up to date": {"percentage_col": "% Risk Assessments on Register up-to-date", "number_col": "Number of Risk Assessments on Register", "completed_col": None},
            "H&S Induction Completion": {"percentage_col": "% of Staff Completed UoN H&S Induction", "number_col": "Number of Staff", "completed_col": None},
            "Fire Training Completion": {"percentage_col": "% of Staff Completed UoN Fire Training", "number_col": "Number of Staff", "completed_col": None},
            "Fire Drills Completed": {"percentage_col": "% of Fire Drills Carried out", "number_col": "Number of Buildings Allocated for Fire Drills to be undertaken", "completed_col": None},
            "PEEPS in Place": {"percentage_col": "% of PEEPS in Place, Reviewed and Controlled", "number_col": "Number of PEEPS Identified", "completed_col": None},
            "PEEPS Drilled": {"percentage_col": "% of PEEPS that are tested/drilled", "number_col": "Number of PEEPS Identified", "completed_col": None},
            "Assets without A&B Defects": {"percentage_col": "% of Assets without active A and B defects", "number_col": "Number of BU Owned Assets", "completed_col": None},
            "Assets Inspected by Allianz": {"percentage_col": "% of Assets seen to by Allianz", "number_col": "Number of BU Owned Assets", "completed_col": None},
            "Accidents and Incidents Investigated": {"percentage_col": "% of Incidents + Near Missed Investigated", "number_col": "Total Incidents (Accidents + Near Misses)", "completed_col": None},
            "Inspections Carried Out": {"percentage_col": "% of Inspections Carried out against Monitoring Schedule", "number_col": "Number of Inspections on Monitoring Schedule", "completed_col": None},
            "Leadership Walkarounds": {"percentage_col": "% of Leadership Walkarounds Carried out", "number_col": "Number of Leadership walkarounds on Monitoring Schedule", "completed_col": None},
            "Risk Assessment Coverage": {"percentage_col": "Percentage Coverage of Risk Assessments", "number_col": None, "completed_col": None},
            "Training Matrix Coverage": {"percentage_col": "% of Training identified in Matrix that is accessible", "number_col": None, "completed_col": None},
            "Staff Training Requirements": {"percentage_col": "% of Staff who are in date with all training requirements", "number_col": "Number of Staff", "completed_col": None}
        }
        self.kpi_metadata = {
            "Written Arrangements Complete": {
                "question": "Arrangements completed, reviewed, and approved out of arrangements relevant to the department.",
                "denominator_label": "arrangements",
                "na_message": "No written arrangements available/generated for this period"
            },
            "Risk Assessments in Register up to date": {
                "question": "Risk assessments updated in the last 2 years out of all risk assessments listed on the register.",
                "denominator_label": "risk assessments",
                "na_message": "No risk assessments in register this period or no risk assessment register in place"
            },
            "H&S Induction Completion": {
                "question": "Staff in-date for UoN H&S induction training (last 3 years).",
                "denominator_label": "staff"
            },
            "Fire Training Completion": {
                "question": "Staff in-date for UoN fire training (last 3 years).",
                "denominator_label": "staff"
            },
            "Fire Drills Completed": {
                "question": "Fire drills carried out in period out of buildings allocated to undertake a fire drill.",
                "denominator_label": "buildings",
                "na_message": "No scheduled drills this period"
            },
            "PEEPS in Place": {
                "question": "PEEPs in place (reviewed, communicated, controls in place) out of PEEPs required.",
                "denominator_label": "PEEPs",
                "na_message": "No PEEPs identified as required"
            },
            "PEEPS Drilled": {
                "question": "PEEPs tested/drilled in the period out of PEEPs required.",
                "denominator_label": "PEEPs",
                "na_message": "No PEEPs identified as required"
            },
            "Assets without A&B Defects": {
                "question": "BU-owned assets without unresolved A/B defects.",
                "denominator_label": "assets",
                "na_message": "No department-owned assets that require inspection"
            },
            "Assets Inspected by Allianz": {
                "question": "BU-owned assets inspected by Allianz (not overdue / plant available).",
                "denominator_label": "assets",
                "na_message": "No department-owned assets that require inspection"
            },
            "Accidents and Incidents Investigated": {
                "question": "Investigations completed out of total incidents and near misses reported in the period.",
                "denominator_label": "incidents/near misses",
                "na_message": "No incidents reported in period"
            },
            "Inspections Carried Out": {
                "question": "Inspections carried out out of inspections on the monitoring schedule.",
                "denominator_label": "inspections",
                "na_message": "No inspections scheduled in this period"
            },
            "Leadership Walkarounds": {
                "question": "Leadership walkarounds completed out of walkarounds on the monitoring schedule.",
                "denominator_label": "walkarounds",
                "na_message": "No walkarounds scheduled in period"
            },
            "Risk Assessment Coverage": {
                "question": "Percentage coverage of risk assessment across department/PS.",
                "denominator_label": None
            },
            "Training Matrix Coverage": {
                "question": "Training in the matrix that is accessible to staff who need it.",
                "denominator_label": None,
                "na_message": "No training matrix available/generated for this period",
                "zero_score_message": "No training matrix available/generated for this period"
            },
            "Staff Training Requirements": {
                "question": "Staff in-date with all required training.",
                "denominator_label": "staff"
            }
        }
        # Tooltips reuse the percentage column descriptions
        self.kpi_tooltips = {kpi: kpi_def["percentage_col"] for kpi, kpi_def in self.kpi_definitions.items()}
        self.tooltip_data = None

    def select_file_and_output(self):
        root = tk.Tk()
        root.withdraw()
        excel_file = filedialog.askopenfilename(
            title="Select Health & Safety Excel File",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
        )
        if not excel_file:
            root.destroy()
            return None, None
        output_dir = filedialog.askdirectory(title="Select Directory for University Report...")
        root.destroy()
        return excel_file, output_dir

    def load_excel_data(self, file_path):
        try:
            xl_file = pd.ExcelFile(file_path)
            if 'University_Summary' not in xl_file.sheet_names:
                raise ValueError("University_Summary sheet not found")
            self.university_data = pd.read_excel(file_path, sheet_name='University_Summary')
            if 'University_Summary_History' in xl_file.sheet_names:
                self.university_history_data = pd.read_excel(file_path, sheet_name='University_Summary_History')
            # Load tooltips from optional sheet "Question Tooltips"
            if 'Question Tooltips' in xl_file.sheet_names:
                tooltip_df = pd.read_excel(file_path, sheet_name='Question Tooltips')
                tooltip_mapping = {}
                for col_idx in range(len(tooltip_df.columns)):
                    col_name, tooltip_text = tooltip_df.iloc[0, col_idx], tooltip_df.iloc[1, col_idx]
                    if pd.notna(col_name) and pd.notna(tooltip_text):
                        tooltip_mapping[str(col_name)] = str(tooltip_text)
                self.tooltip_data = {kpi: tooltip_mapping.get(col, None) for kpi, col in self.kpi_tooltips.items()}
            else:
                self.tooltip_data = {kpi: None for kpi in self.kpi_definitions.keys()}
        except Exception as e:
            raise RuntimeError(f"Failed to load Excel data: {e}")

    def _build_university_history(self):
        """Build KPI -> [{date, percentage, number, applicable}] from University_Summary_History."""
        empty = {k: [] for k in self.kpi_definitions.keys()}
        if self.university_history_data is None or self.university_history_data.empty:
            return empty

        df = self.university_history_data.copy()
        if 'Faculty' in df.columns:
            df = df[df['Faculty'].astype(str).str.strip().str.lower().eq('university')].copy()
        if df.empty:
            return empty

        if 'Date' in df.columns:
            df['_parsed_date'] = pd.to_datetime(df['Date'], dayfirst=True, errors='coerce')
            df = df.sort_values('_parsed_date')
        else:
            df['_parsed_date'] = pd.NaT

        result = {}
        for kpi_name, kpi_def in self.kpi_definitions.items():
            series = []
            for _, row in df.iterrows():
                val = row.get(kpi_def["percentage_col"], None)
                num = row.get(kpi_def["number_col"], None) if kpi_def["number_col"] else None
                pct = self._safe_float(val) if not self._is_empty(val) else None
                num_val = self._safe_float(num) if not self._is_empty(num) else None
                applicable = self._is_kpi_applicable(kpi_name, num_val)
                date_label = row.get('Date', '')
                series.append({
                    "date": str(date_label) if pd.notna(date_label) else "",
                    "percentage": pct if applicable else None,
                    "raw_percentage": pct,
                    "number": num_val,
                    "applicable": applicable
                })
            result[kpi_name] = series
        return result

    def _is_empty(self, value):
        return pd.isna(value) or value == '/' or value == ''

    def _safe_float(self, value):
        try:
            return float(value)
        except (ValueError, TypeError):
            return None

    def _is_kpi_applicable(self, kpi_name, number):
        number_col = self.kpi_definitions.get(kpi_name, {}).get("number_col")
        if not number_col or number is None:
            return True
        return abs(float(number)) > 1e-9

    def _format_display(self, kpi_name, percentage, number):
        applicable = self._is_kpi_applicable(kpi_name, number)
        metadata = self.kpi_metadata.get(kpi_name, {})
        denominator_label = metadata.get("denominator_label")
        if not applicable:
            custom_na_message = metadata.get("na_message")
            if custom_na_message:
                return custom_na_message
            if denominator_label:
                return f"Not applicable (0 {denominator_label} in scope)"
            return "Not applicable this period"
        if percentage is None:
            return "No return submitted"
        zero_score_message = metadata.get("zero_score_message")
        if zero_score_message and abs(float(percentage)) < 1e-9:
            return zero_score_message
        if number is None:
            return f"{percentage:.1f}%"
        if denominator_label:
            return f"{percentage:.1f}% ({int(number)} {denominator_label} in scope)"
        return f"{percentage:.1f}% ({int(number)} in scope)"

    def _extract_university_kpis(self):
        if self.university_data is None or self.university_data.empty:
            raise ValueError("No university data available")

        # Pick a row that represents overall university totals.
        df = self.university_data
        row = None
        # Prefer a row labeled as University if present
        for col in ["Entity", "Level", "Faculty", "Name"]:
            if col in df.columns:
                match = df[df[col].astype(str).str.strip().str.lower().eq("university")]
                if not match.empty:
                    row = match.iloc[0]
                    break
        if row is None:
            row = df.iloc[0]

        kpis = {}
        history = self._build_university_history()
        for kpi_name, kpi_def in self.kpi_definitions.items():
            perc_val = row.get(kpi_def["percentage_col"], None)
            num_val = row.get(kpi_def["number_col"], None) if kpi_def["number_col"] else None
            perc_val = self._safe_float(perc_val) if not self._is_empty(perc_val) else None
            num_val = self._safe_float(num_val) if not self._is_empty(num_val) else None
            kpis[kpi_name] = {
                "percentage": perc_val,
                "number": num_val,
                "applicable": self._is_kpi_applicable(kpi_name, num_val),
                "display_text": self._format_display(kpi_name, perc_val, num_val),
                "history": history.get(kpi_name, [])
            }
        return {"name": "University", "kpis": kpis}

    def _performance_class(self, percentage, applicable=True):
        if not applicable:
            return "performance-not-applicable"
        if percentage is None:
            return "performance-no-data"
        if percentage >= 90:
            return "performance-excellent"
        if percentage >= 75:
            return "performance-good"
        if percentage >= 50:
            return "performance-warning"
        return "performance-poor"

    def create_university_html_dashboard(self, uni_data, output_path):
        template_path = os.path.join(os.path.dirname(__file__), 'university_dashboard_template.html')
        if not os.path.exists(template_path):
            raise FileNotFoundError(f"University dashboard template not found: {template_path}")

        tooltip_json = json.dumps(self.tooltip_data or {}, indent=2, default=str)
        kpi_meta_json = json.dumps(self.kpi_metadata, default=str)
        cards_html = self._render_kpi_cards(uni_data)

        with open(template_path, 'r', encoding='utf-8') as f:
            html_content = f.read()

        replacements = {
            '__TOOLTIP_JSON__': tooltip_json,
            '__KPI_META_JSON__': kpi_meta_json,
            '__UNIVERSITY_CARDS__': cards_html,
        }
        for token, value in replacements.items():
            html_content = html_content.replace(token, value)

        with open(output_path, 'w', encoding='utf-8') as f:
            f.write(html_content)

    def _render_kpi_cards(self, uni_data):
        cards = []
        items = list(uni_data["kpis"].items())
        # Sort highest to lowest; None at end
        items.sort(key=lambda kv: (kv[1].get("percentage") is None, -(kv[1].get("percentage") or 0)))
        for kpi_name, kpi in items:
            pct = kpi.get("percentage")
            val = kpi.get("display_text", "/")
            applicable = kpi.get("applicable", True)
            perf_cls = self._performance_class(pct, applicable)
            safe_key = kpi_name.replace('"', "&quot;")
            safe_kpi_name = kpi_name.replace('"', '&quot;')
            history = kpi.get("history", [])
            history_json = json.dumps(history, default=str).replace('"', '&quot;')
            out_of_range = False
            if pct is not None:
                try:
                    out_of_range = float(pct) < 0 or float(pct) > 100
                except (ValueError, TypeError):
                    out_of_range = False
            history_out_of_range = any(
                isinstance(point, dict)
                and point.get("raw_percentage") is not None
                and self._safe_float(point.get("raw_percentage")) is not None
                and (self._safe_float(point.get("raw_percentage")) < 0 or self._safe_float(point.get("raw_percentage")) > 100)
                for point in history
            )
            warning_badge = '<span class="data-warning" title="Current or historic source percentage outside 0-100">!</span>' if (out_of_range or history_out_of_range) else ''
            question_text = self.kpi_metadata.get(kpi_name, {}).get("question", "")
            card = f'''<div class="kpi-card">
                <div class="kpi-label-wrap">
                    <div class="kpi-title">{kpi_name}
                        <span class="tooltip-trigger" data-kpi="{safe_key}" title="Info" onmouseenter="handleTooltipShow(event)" onmouseleave="handleTooltipHide()">&#9432;</span>
                        {warning_badge}
                    </div>
                    <div class="kpi-question">{question_text}</div>
                </div>
                <span class="kpi-value {perf_cls}">{val}</span>
                <span style="height:1px"></span>
                <div class="trend-wrap"><canvas class="trend-canvas" data-kpi-name="{safe_kpi_name}" data-history="{history_json}"></canvas></div>
            </div>'''
            cards.append(card)
        return "\n".join(cards)

    def run(self):
        # If a default test file exists, prefer it for convenience; otherwise prompt
        default_file = 'test.xlsx'
        if os.path.exists(default_file):
            excel_file = default_file
            output_dir = os.getcwd()
        else:
            excel_file, output_dir = self.select_file_and_output()
            if not excel_file or not output_dir:
                print("Operation cancelled.")
                return

        try:
            self.load_excel_data(excel_file)
            uni_data = self._extract_university_kpis()
        except Exception as e:
            messagebox.showerror("Error", str(e))
            print(f"Error: {e}")
            return

        output_path = os.path.join(output_dir, "University_KPI_Report.html")

        try:
            self.create_university_html_dashboard(uni_data, output_path)
            print("\n✅ University Report generated successfully!")
            print(f"📁 Saved to: {output_path}")
            messagebox.showinfo("Success", f"University Report generated successfully!\n\nSaved to: {output_path}\n\nOpen the HTML file in your web browser to view the dashboard.")
        except Exception as e:
            print(f"\n❌ Failed to generate report: {e}")
            messagebox.showerror("Error", f"Failed to generate report: {e}")


if __name__ == "__main__":
    generator = UniversityDashboardGenerator()
    generator.run()
