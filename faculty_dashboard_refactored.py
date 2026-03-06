import pandas as pd
import json
import tkinter as tk
from tkinter import filedialog, messagebox
import os

class FacultyDashboardGenerator:
    def __init__(self):
        self.university_data = None
        self.faculty_data = None
        self.school_data = None
        self.university_history_data = None
        self.faculty_history_data = None
        self.school_history_data = None
        self.tooltip_data = None
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
        self.faculty_school_mapping = {"CLAS": "Arts", "English": "Arts", "Humanities": "Arts", "Engineering": "Engineering", "Estates": "Estates", "Finance and Infrastructure": "Finance and Infrastructure", "HR": "HR", "BDI": "Medicine & Health Sciences", "Life Sciences": "Medicine & Health Sciences", "Medicine": "Medicine & Health Sciences", "Veterinary Medicine and Science": "Medicine & Health Sciences", "Clinical Skills": "Medicine & Health Sciences", "Health Sciences": "Medicine & Health Sciences", "Libraries": "Registrars", "Sport": "Registrars", "BSU": "Registrars", "Student & Public Facing Services": "Registrars", "Computer Sciences": "Science", "Biosciences": "Science", "Chemistry": "Science", "Mathematical Sciences": "Science", "Pharmacy": "Science", "Physics and Astronomy": "Science", "Psychology": "Science", "Economics": "Social Sciences", "Business School": "Social Sciences", "Education": "Social Sciences", "Law": "Social Sciences", "Politics and IR": "Social Sciences", "Rights Lab": "Social Sciences", "Sociology": "Social Sciences", "Geography": "Social Sciences", "Faculty Office": "Social Sciences"}
        # KPIs where more can be done than was scheduled (e.g. extra inspections beyond the plan)
        self.exceedable_kpis = {
            "Inspections Carried Out", "Fire Drills Completed",
            "Leadership Walkarounds", "PEEPS Drilled"
        }
        self.kpi_tooltips = {kpi: kpi_def["percentage_col"] for kpi, kpi_def in self.kpi_definitions.items()}

    def select_file_and_output(self):
        root = tk.Tk()
        root.withdraw()
        excel_file = filedialog.askopenfilename(title="Select Health & Safety Excel File", filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")])
        if not excel_file:
            root.destroy()
            return None, None
        output_dir = filedialog.askdirectory(title="Select Directory for Faculty Reports...")
        root.destroy()
        return excel_file, output_dir

    def load_excel_data(self, file_path):
        try:
            xl_file = pd.ExcelFile(file_path)
            required_sheets = ['University_Summary', 'Faculty_Summary', 'School_Raw_Data']
            for sheet in required_sheets:
                if sheet not in xl_file.sheet_names:
                    raise ValueError(f"{sheet} sheet not found")
            
            self.university_data = pd.read_excel(file_path, sheet_name='University_Summary')
            self.faculty_data = pd.read_excel(file_path, sheet_name='Faculty_Summary')
            self.school_data = pd.read_excel(file_path, sheet_name='School_Raw_Data')
            self.school_data['Faculty'] = self.school_data['School'].map(self.faculty_school_mapping)
            if 'University_Summary_History' in xl_file.sheet_names:
                self.university_history_data = pd.read_excel(file_path, sheet_name='University_Summary_History')
            if 'Faculty_Summary_History' in xl_file.sheet_names:
                self.faculty_history_data = pd.read_excel(file_path, sheet_name='Faculty_Summary_History')
            if 'School_Raw_Data_History' in xl_file.sheet_names:
                self.school_history_data = pd.read_excel(file_path, sheet_name='School_Raw_Data_History')
                self.school_history_data['Faculty'] = self.school_history_data['School'].map(self.faculty_school_mapping)
            
            if 'Question Tooltips' in xl_file.sheet_names:
                tooltip_df = pd.read_excel(file_path, sheet_name='Question Tooltips')
                tooltip_mapping = {}
                for col_idx in range(len(tooltip_df.columns)):
                    col_name, tooltip_text = tooltip_df.iloc[0, col_idx], tooltip_df.iloc[1, col_idx]
                    if pd.notna(col_name) and pd.notna(tooltip_text):
                        tooltip_mapping[str(col_name)] = str(tooltip_text)
                self.tooltip_data = {kpi: tooltip_mapping.get(col, None) for kpi, col in self.kpi_tooltips.items()}
                print(f"   ✓ Loaded {len([x for x in self.tooltip_data.values() if x])} KPI tooltips")
            else:
                print("Warning: Question Tooltips sheet not found - tooltips will not be available")
            return True
        except Exception as e:
            messagebox.showerror("Error", f"Could not load Excel file: {str(e)}")
            return False

    def _build_kpi_history(self, history_df, entity_col, entity_name):
        """Build KPI -> [{date, percentage, number, applicable}] history series for one entity."""
        empty = {k: [] for k in self.kpi_definitions.keys()}
        if history_df is None or history_df.empty or entity_col not in history_df.columns:
            return empty

        scoped = history_df[history_df[entity_col].astype(str).str.strip() == str(entity_name).strip()].copy()
        if scoped.empty:
            return empty

        if 'Date' in scoped.columns:
            scoped['_parsed_date'] = pd.to_datetime(scoped['Date'], dayfirst=True, errors='coerce')
            scoped = scoped.sort_values('_parsed_date')
        else:
            scoped['_parsed_date'] = pd.NaT

        result = {}
        for kpi_name, kpi_def in self.kpi_definitions.items():
            series = []
            for _, row in scoped.iterrows():
                row_missing_return = self._is_no_return_row(row)
                perc = row.get(kpi_def["percentage_col"], None)
                num = row.get(kpi_def["number_col"], None) if kpi_def["number_col"] else None
                perc_val = self._safe_float(perc) if not self._is_empty(perc) else None
                num_val = self._safe_float(num) if not self._is_empty(num) else None
                if row_missing_return:
                    perc_val = None
                    num_val = None
                # For count-based KPIs, a missing denominator means no return was submitted.
                # Only apply this when the denominator column exists in this sheet/row.
                if kpi_def["number_col"] and kpi_def["number_col"] in row.index and num_val is None:
                    perc_val = None
                applicable = self._is_kpi_applicable(kpi_name, num_val)
                exceeded = False
                if kpi_name in self.exceedable_kpis and kpi_def["number_col"] and num_val is not None and perc_val is not None:
                    if not applicable and perc_val > 0:
                        exceeded = True
                    elif perc_val > 100 and num_val > 0:
                        exceeded = True
                if 'Date' in scoped.columns:
                    date_label = row.get('Date', '')
                else:
                    date_label = ''
                series.append({
                    "date": str(date_label) if pd.notna(date_label) else "",
                    "percentage": perc_val if (applicable or exceeded) else None,
                    "raw_percentage": perc_val,
                    "number": num_val,
                    "applicable": applicable if not exceeded else True,
                    "exceeded": exceeded
                })
            result[kpi_name] = series
        return result

    def extract_and_process_data(self):
        dashboard_data = {"university": {}, "faculties": {}, "schools": {}}
        
        # Process university data
        university_kpis = {}
        university_history = self._build_kpi_history(self.university_history_data, 'Faculty', 'University')
        if self.university_data is not None and len(self.university_data) > 0:
            uni_row = self.university_data.iloc[0]
            dashboard_data["university"] = self._extract_kpi_data(uni_row, "University")
            for kpi_name, kpi_def in self.kpi_definitions.items():
                perc_val = uni_row.get(kpi_def["percentage_col"], None)
                university_kpis[kpi_name] = self._safe_float(perc_val) if not self._is_empty(perc_val) else None

        # Process faculty data
        if self.faculty_data is not None:
            for _, faculty_row in self.faculty_data.iterrows():
                faculty_name = faculty_row['Faculty']
                faculty_data = self._extract_kpi_data(faculty_row, faculty_name)
                faculty_history = self._build_kpi_history(self.faculty_history_data, 'Faculty', faculty_name)
                
                # Add university comparison
                for kpi_name in faculty_data["kpis"]:
                    fac_perc = faculty_data["kpis"][kpi_name]["percentage"]
                    uni_perc = university_kpis.get(kpi_name, None)
                    fac_applicable = faculty_data["kpis"][kpi_name].get("applicable", True)
                    uni_number_col = self.kpi_definitions[kpi_name]["number_col"]
                    uni_number = uni_row.get(uni_number_col, None) if 'uni_row' in locals() and uni_number_col else None
                    uni_number_val = self._safe_float(uni_number) if not self._is_empty(uni_number) else None
                    uni_applicable = self._is_kpi_applicable(kpi_name, uni_number_val)
                    faculty_data["kpis"][kpi_name]["university_percentage"] = uni_perc
                    faculty_data["kpis"][kpi_name]["university_applicable"] = uni_applicable
                    faculty_data["kpis"][kpi_name]["history"] = faculty_history.get(kpi_name, [])
                    faculty_data["kpis"][kpi_name]["university_history"] = university_history.get(kpi_name, [])
                    
                    if fac_applicable and uni_applicable and fac_perc is not None and uni_perc is not None:
                        if fac_perc > uni_perc:
                            faculty_data["kpis"][kpi_name]["comparison"] = "above"
                        elif fac_perc < uni_perc:
                            faculty_data["kpis"][kpi_name]["comparison"] = "below"
                        else:
                            faculty_data["kpis"][kpi_name]["comparison"] = "equal"
                    else:
                        faculty_data["kpis"][kpi_name]["comparison"] = "unknown"
                
                dashboard_data["faculties"][faculty_name] = faculty_data
                
                # Add schools for this faculty
                faculty_schools = self.school_data[self.school_data['Faculty'] == faculty_name] if self.school_data is not None else pd.DataFrame()
                dashboard_data["faculties"][faculty_name]['schools'] = {}
                for _, school_row in faculty_schools.iterrows():
                    school_name = school_row['School']
                    school_data = self._extract_kpi_data(school_row, school_name)
                    school_history = self._build_kpi_history(self.school_history_data, 'School', school_name)
                    for kpi_name in school_data["kpis"]:
                        school_data["kpis"][kpi_name]["history"] = school_history.get(kpi_name, [])
                    dashboard_data["faculties"][faculty_name]['schools'][school_name] = school_data

        return dashboard_data

    def _extract_kpi_data(self, data_row, entity_name):
        kpi_data = {"name": entity_name, "kpis": {}}
        row_missing_return = self._is_no_return_row(data_row)
        for kpi_name, kpi_def in self.kpi_definitions.items():
            perc_val = data_row.get(kpi_def["percentage_col"], None)
            num_val = data_row.get(kpi_def["number_col"], None) if kpi_def["number_col"] else None
            
            perc_val = self._safe_float(perc_val) if not self._is_empty(perc_val) else None
            num_val = self._safe_float(num_val) if not self._is_empty(num_val) else None
            if row_missing_return:
                perc_val = None
                num_val = None
            # For count-based KPIs, a missing denominator means no return was submitted.
            # Only apply this when the denominator column exists in this sheet/row.
            if kpi_def["number_col"] and kpi_def["number_col"] in data_row.index and num_val is None:
                perc_val = None
            
            # Detect exceeded-schedule cases (e.g. 1 inspection done, 0 scheduled)
            applicable = self._is_kpi_applicable(kpi_name, num_val)
            completed_count = None
            exceeded = False
            if kpi_name in self.exceedable_kpis and kpi_def["number_col"] and num_val is not None and perc_val is not None:
                completed_count = round(perc_val * num_val / 100) if num_val > 0 else None
                if not applicable and perc_val > 0:
                    # 0 scheduled but percentage > 0 means work was done beyond schedule
                    # We can't calculate completed from percentage alone when denominator is 0,
                    # so we just know at least some were done
                    exceeded = True
                    completed_count = None  # unknown exact count
                elif perc_val > 100 and num_val > 0:
                    exceeded = True
                    completed_count = round(perc_val * num_val / 100)

            kpi_data["kpis"][kpi_name] = {
                "percentage": perc_val,
                "number": num_val,
                "applicable": applicable if not exceeded else True,
                "exceeded": exceeded,
                "completed_count": completed_count,
                "display_text": self._format_display(kpi_name, perc_val, num_val, exceeded, completed_count)
            }
        return kpi_data

    def _is_empty(self, value):
        if pd.isna(value):
            return True
        if isinstance(value, str):
            cleaned = value.strip()
            return cleaned in {'', '/', '-', '–', '—'}
        return False

    def _safe_float(self, value):
        try:
            return float(value)
        except (ValueError, TypeError):
            return None

    def _is_no_return_row(self, row):
        """
        Treat a row as 'no return' when it looks like an untouched template row:
        - at least one explicit missing marker is present ('/', '-', blank), and
        - all KPI denominator/percentage values are either missing or zero.
        """
        denominator_cols = [cfg["number_col"] for cfg in self.kpi_definitions.values() if cfg.get("number_col")]
        percentage_cols = [cfg["percentage_col"] for cfg in self.kpi_definitions.values() if cfg.get("percentage_col")]
        if not denominator_cols and not percentage_cols:
            return False

        has_explicit_missing = False
        for col in denominator_cols + percentage_cols:
            value = row.get(col, None)
            if self._is_empty(value):
                has_explicit_missing = True
                continue
            numeric = self._safe_float(value)
            if numeric is None:
                has_explicit_missing = True
                continue
            if abs(float(numeric)) > 1e-9:
                return False
        return has_explicit_missing

    def _is_kpi_applicable(self, kpi_name, number):
        number_col = self.kpi_definitions.get(kpi_name, {}).get("number_col")
        if not number_col or number is None:
            return True
        return abs(float(number)) > 1e-9

    def _format_display(self, kpi_name, percentage, number, exceeded=False, completed_count=None):
        applicable = self._is_kpi_applicable(kpi_name, number)
        metadata = self.kpi_metadata.get(kpi_name, {})
        denominator_label = metadata.get("denominator_label")
        if exceeded:
            label = denominator_label or "items"
            if number is not None and abs(float(number)) < 1e-9:
                # 0 scheduled but work was done
                return f"Exceeded (0 {label} scheduled)"
            elif completed_count is not None:
                return f"{percentage:.1f}% ({completed_count} carried out, {int(number)} {label} scheduled)"
            else:
                return f"{percentage:.1f}% ({int(number)} {label} in scope)"
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

    def create_faculty_html_dashboard(self, dashboard_data, faculty_name, output_path):
        faculty_data = dashboard_data["faculties"].get(faculty_name, {})
        if not faculty_data:
            raise ValueError(f"No data found for faculty: {faculty_name}")

        template_path = os.path.join(os.path.dirname(__file__), 'faculty_dashboard_template.html')
        if not os.path.exists(template_path):
            raise FileNotFoundError(f"Faculty dashboard template not found: {template_path}")

        faculty_json = json.dumps(faculty_data, indent=2, default=str)
        tooltip_json = json.dumps(self.tooltip_data or {}, indent=2, default=str)
        kpi_meta_json = json.dumps(self.kpi_metadata, default=str)

        with open(template_path, 'r', encoding='utf-8') as f:
            html_content = f.read()

        replacements = {
            '__FACULTY_NAME__': faculty_name,
            '__FACULTY_JSON__': faculty_json,
            '__TOOLTIP_JSON__': tooltip_json,
            '__KPI_META_JSON__': kpi_meta_json,
        }
        for token, value in replacements.items():
            html_content = html_content.replace(token, value)

        with open(output_path, 'w', encoding='utf-8') as f:
            f.write(html_content)

    def run(self):
        print("Faculty Health & Safety Dashboard Generator")
        print("=" * 50)

        print("1. Select Excel file and output directory...")
        excel_file, output_dir = self.select_file_and_output()
        
        if not excel_file or not output_dir:
            print("No file or directory selected. Exiting.")
            return

        if not os.path.exists(excel_file):
            print(f"Error: File '{excel_file}' does not exist.")
            return

        print("2. Loading Excel data...")
        if not self.load_excel_data(excel_file):
            return

        print(f"   ✓ University: {len(self.university_data) if self.university_data is not None else 0} rows")
        print(f"   ✓ Faculty: {len(self.faculty_data) if self.faculty_data is not None else 0} rows") 
        print(f"   ✓ School: {len(self.school_data) if self.school_data is not None else 0} rows")

        print("3. Processing KPI data...")
        dashboard_data = self.extract_and_process_data()

        print("4. Creating faculty-specific HTML dashboards...")
        
        if not dashboard_data["faculties"]:
            print("❌ Error: No faculty data available to generate reports.")
            messagebox.showerror("Error", "No faculty data available to generate reports.")
            return
        
        generated_files = []
        
        for faculty_name in dashboard_data["faculties"].keys():
            safe_filename = "".join(c for c in faculty_name if c.isalnum() or c in (' ', '-', '_')).rstrip().replace(' ', '_')
            output_file = os.path.join(output_dir, f"{safe_filename}_Faculty_Report.html")
            
            try:
                self.create_faculty_html_dashboard(dashboard_data, faculty_name, output_file)
                generated_files.append(output_file)
                print(f"   ✓ Generated: {faculty_name} → {os.path.basename(output_file)}")
            except Exception as e:
                print(f"   ❌ Failed to generate {faculty_name}: {str(e)}")
        
        if generated_files:
            print(f"\n✅ Faculty Reports generated successfully!")
            print(f"📁 Saved to directory: {output_dir}")
            print(f"📊 Generated {len(generated_files)} faculty reports")
            
            file_list = "\\n".join([f"• {os.path.basename(f)}" for f in generated_files])
            messagebox.showinfo("Success", f"Faculty Reports generated successfully!\\n\\nSaved to: {output_dir}\\n\\nGenerated files:\\n{file_list}\\n\\nOpen any HTML file in your web browser to view the faculty dashboard.")
        else:
            print(f"\n❌ No faculty reports were generated successfully.")
            messagebox.showerror("Error", "No faculty reports were generated successfully.")

if __name__ == "__main__":
    generator = FacultyDashboardGenerator()
    generator.run()
