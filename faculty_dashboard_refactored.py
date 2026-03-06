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
        self.faculty_school_mapping = {"CLAS": "Arts", "English": "Arts", "Humanities": "Arts", "Engineering": "Engineering", "Estates": "Estates", "Finance and Infrastructure": "Finance and Infrastructure", "HR": "HR", "BDI": "Medicine & Health Sciences", "Life Sciences": "Medicine & Health Sciences", "Medicine": "Medicine & Health Sciences", "Veterinary Medicine and Science": "Medicine & Health Sciences", "Clinical Skills": "Medicine & Health Sciences", "Health Sciences": "Medicine & Health Sciences", "Libraries": "Registrars", "Sport": "Registrars", "BSU": "Registrars", "Student & Public Facing Services": "Registrars", "Computer Sciences": "Science", "Biosciences": "Science", "Chemistry": "Science", "Mathematical Sciences": "Science", "Pharmacy": "Science", "Physics and Astronomy": "Science", "Psychology": "Science", "Economics": "Social Sciences", "Business School": "Social Sciences", "Education": "Social Sciences", "Law": "Social Sciences", "Politics and IR": "Social Sciences", "Rights Lab": "Social Sciences", "Sociology": "Social Sciences", "Geography": "Social Sciences", "Faculty Office": "Social Sciences"}
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

    def extract_and_process_data(self):
        dashboard_data = {"university": {}, "faculties": {}, "schools": {}}
        
        # Process university data
        university_kpis = {}
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
                
                # Add university comparison
                for kpi_name in faculty_data["kpis"]:
                    fac_perc = faculty_data["kpis"][kpi_name]["percentage"]
                    uni_perc = university_kpis.get(kpi_name, None)
                    faculty_data["kpis"][kpi_name]["university_percentage"] = uni_perc
                    
                    if fac_perc is not None and uni_perc is not None:
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
                    dashboard_data["faculties"][faculty_name]['schools'][school_name] = self._extract_kpi_data(school_row, school_name)

        return dashboard_data

    def _extract_kpi_data(self, data_row, entity_name):
        kpi_data = {"name": entity_name, "kpis": {}}
        for kpi_name, kpi_def in self.kpi_definitions.items():
            perc_val = data_row.get(kpi_def["percentage_col"], None)
            num_val = data_row.get(kpi_def["number_col"], None) if kpi_def["number_col"] else None
            
            perc_val = self._safe_float(perc_val) if not self._is_empty(perc_val) else None
            num_val = self._safe_float(num_val) if not self._is_empty(num_val) else None
            
            kpi_data["kpis"][kpi_name] = {
                "percentage": perc_val,
                "number": num_val,
                "display_text": self._format_display(perc_val, num_val)
            }
        return kpi_data

    def _is_empty(self, value):
        return pd.isna(value) or value == '/' or value == ''

    def _safe_float(self, value):
        try:
            return float(value)
        except (ValueError, TypeError):
            return None

    def _format_display(self, percentage, number):
        if percentage is None:
            return "/"
        if number is None:
            return f"{percentage:.1f}%"
        return f"{percentage:.1f}% ({int(number)})"

    def create_faculty_html_dashboard(self, dashboard_data, faculty_name, output_path):
        faculty_data = dashboard_data["faculties"].get(faculty_name, {})
        if not faculty_data:
            raise ValueError(f"No data found for faculty: {faculty_name}")
        
        faculty_json = json.dumps(faculty_data, indent=2, default=str)
        tooltip_json = json.dumps(self.tooltip_data or {}, indent=2, default=str)
        
        # Minified HTML with compressed CSS and JavaScript
        html_content = f'''<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>{faculty_name} Faculty - Health & Safety Dashboard</title>
<script src="https://cdnjs.cloudflare.com/ajax/libs/Chart.js/3.9.1/chart.min.js"></script>
<style>
*{{margin:0;padding:0;box-sizing:border-box}}
body{{font-family:'Segoe UI',Tahoma,Geneva,Verdana,sans-serif;background:linear-gradient(135deg,#f8fafc 0%,#e2e8f0 100%);min-height:100vh;color:#333;font-size:15px}}
.container{{max-width:1400px;margin:0 auto;padding:20px}}
.header{{text-align:center;color:#1e293b;margin-bottom:30px;background:linear-gradient(135deg,#667eea 0%,#764ba2 100%);padding:40px;border-radius:20px;box-shadow:0 8px 32px rgba(0,0,0,0.12);position:relative;overflow:hidden}}
.header::before{{content:'';position:absolute;top:0;left:0;right:0;bottom:0;background:linear-gradient(45deg,rgba(255,255,255,0.1) 0%,transparent 50%,rgba(255,255,255,0.05) 100%);pointer-events:none}}
.header h1{{font-size:2.5rem;margin-bottom:12px;font-weight:800;color:white;text-shadow:0 2px 4px rgba(0,0,0,0.1);position:relative;z-index:1}}
.header p{{color:rgba(255,255,255,0.9);font-size:1.1rem;font-weight:500;position:relative;z-index:1}}
.faculty-overview{{background:white;border-radius:20px;padding:35px;margin-bottom:30px;box-shadow:0 8px 32px rgba(0,0,0,0.1);border:1px solid #e2e8f0;position:relative;overflow:hidden}}
.faculty-overview::before{{content:'';position:absolute;top:0;left:0;right:0;height:5px;background:linear-gradient(90deg,#667eea,#764ba2,#667eea);background-size:200% 100%;animation:shimmer 3s ease-in-out infinite}}
@keyframes shimmer{{0%,100%{{background-position:200% 0}}50%{{background-position:-200% 0}}}}
.overview-title{{font-size:1.6rem;font-weight:800;background:linear-gradient(135deg,#667eea,#764ba2);-webkit-background-clip:text;-webkit-text-fill-color:transparent;background-clip:text;margin-bottom:25px;text-align:center}}
.controls-container{{text-align:center;margin-bottom:25px}}
.view-controls{{margin-bottom:15px}}
.view-button{{background:#f1f5f9;color:#64748b;border:1px solid #e2e8f0;padding:10px 20px;border-radius:20px;font-size:15px;font-weight:500;cursor:pointer;transition:all 0.3s ease;margin:0 5px}}
.view-button:hover{{background:#e2e8f0;color:#334155}}
.view-button.active{{background:#667eea;color:white;font-weight:600}}
.sort-controls{{text-align:center}}
.sort-button{{background:#f1f5f9;color:#64748b;border:1px solid #e2e8f0;padding:10px 20px;border-radius:20px;font-size:15px;font-weight:500;cursor:pointer;transition:all 0.3s ease;margin:0 5px}}
.sort-button:hover{{background:#e2e8f0;color:#334155}}
.sort-button.active{{background:#3b82f6;color:white;font-weight:600}}
.kpi-grid{{display:grid;grid-template-columns:repeat(3,1fr);gap:25px;margin-bottom:30px}}
.kpi-card{{background:white;border-radius:16px;padding:20px;box-shadow:0 4px 20px rgba(0,0,0,0.08);border:1px solid #e2e8f0;transition:all 0.4s cubic-bezier(0.4,0,0.2,1);cursor:pointer;position:relative;overflow:hidden;align-self:start}}
.kpi-card::before{{content:'';position:absolute;top:0;left:0;right:0;height:4px;background:linear-gradient(90deg,#3b82f6,#8b5cf6,#06b6d4);opacity:0;transition:opacity 0.3s ease}}
.kpi-card:hover{{transform:translateY(-4px);box-shadow:0 12px 40px rgba(0,0,0,0.15);border-color:#3b82f6}}
.kpi-card:hover::before{{opacity:1}}
.kpi-card.expanded{{box-shadow:0 12px 35px rgba(0,0,0,0.15)}}
.university-comparison{{background:#f8fafc;border:1px solid #e2e8f0;border-radius:10px;padding:10px 12px;margin-bottom:12px;display:flex;align-items:center;justify-content:center;gap:12px}}
.comparison-text{{color:#64748b;font-weight:500;font-size:13px;margin:0}}
.comparison-values{{display:flex;align-items:center;gap:10px}}
.faculty-value{{font-weight:700;color:#1e293b;font-size:16px}}
.university-value{{color:#64748b;font-size:16px;font-weight:500}}
.comparison-indicator{{font-size:16px;font-weight:700;margin-left:6px;padding:6px;border-radius:50%;display:inline-flex;align-items:center;justify-content:center;width:28px;height:28px;box-shadow:0 2px 6px rgba(0,0,0,0.12)}}
.comparison-above{{background:linear-gradient(135deg,#3b82f6,#1d4ed8);color:white}}
.comparison-below{{background:linear-gradient(135deg,#ef4444,#dc2626);color:white}}
.comparison-equal{{background:linear-gradient(135deg,#6b7280,#4b5563);color:white}}
.kpi-title{{font-size:16px;font-weight:600;color:#374151;line-height:1.4;margin-bottom:15px;text-align:center}}
.kpi-value{{font-size:24px;font-weight:700;color:white;padding:14px 20px;border-radius:16px;text-align:center;box-shadow:0 4px 16px rgba(0,0,0,0.15);margin:0 auto 15px;display:block;width:fit-content;min-width:100px}}
.expand-icon{{position:absolute;bottom:13px;left:50%;transform:translateX(-50%);font-size:15px;color:#9ca3af;transition:all 0.3s ease;cursor:pointer}}
.kpi-card.expanded .expand-icon{{transform:translateX(-50%) rotate(180deg);color:#3b82f6}}
.compressed-view .kpi-grid{{display:flex;flex-wrap:wrap;gap:8px;align-items:flex-start}}
.compressed-view .kpi-card{{display:grid;grid-template-columns:7fr 1.5fr 1.5fr auto;column-gap:12px;align-items:center;padding:6px 12px;margin-bottom:4px;cursor:default;transform:none;border-radius:8px;min-height:40px;position:relative;align-self:flex-start;flex:1 1 calc(50% - 4px);box-sizing:border-box}}
.compressed-view .kpi-card .school-breakdown-content{{grid-column:1/-1;width:100%}}
.compressed-view .kpi-card.expanded{{margin-bottom:8px}}
.compressed-view .kpi-card:hover{{transform:none;box-shadow:0 4px 20px rgba(0,0,0,0.08)}}
.compressed-view .kpi-card::before{{display:none}}
.compressed-view .kpi-title{{font-size:14px;font-weight:600;margin:0;line-height:1.3;align-self:center;display:flex;align-items:center}}
.compressed-view .kpi-value{{font-size:14px;padding:3px 8px;margin:0;border-radius:6px;justify-self:center;text-align:center;display:inline-flex;align-items:center;justify-content:center;line-height:1.2;min-height:38px;max-height:38px;overflow:hidden}}
.compressed-view .university-comparison{{background:transparent;border:none;padding:0;margin:0;flex:none;width:auto;justify-self:end;display:flex;justify-content:flex-end;align-items:center;align-self:center}}
.compressed-view .comparison-text{{display:none}}
.compressed-view .comparison-values{{margin:0;gap:8px;justify-content:flex-end;align-items:center}}
.compressed-view .university-value{{font-size:14px;color:#64748b;font-weight:500}}
.compressed-view .comparison-indicator{{font-size:14px;width:24px;height:24px;margin-left:0}}
.compressed-view .expand-icon{{display:block;position:static;transform:none;margin-left:0;font-size:12px;color:#9ca3af;z-index:1;pointer-events:none;justify-self:end;align-self:center;grid-column:4;grid-row:1}}.compressed-view .kpi-card.expanded .expand-icon{{transform:rotate(180deg);color:#3b82f6}}
.compressed-view .tooltip-trigger{{margin-left:6px}}
.school-breakdown-content{{max-height:0;overflow:hidden;transition:max-height 0.3s ease,opacity 0.3s ease;opacity:0;margin-top:0}}
.kpi-card.expanded .school-breakdown-content{{max-height:10000px;opacity:1;margin-top:18px;padding-top:18px;border-top:1px solid #e5e7eb}}
.school-breakdown-title{{font-size:15px;font-weight:600;color:#6b7280;margin-bottom:14px}}
.school-bar-inline{{display:flex;align-items:center;margin-bottom:14px;padding:14px;background:#f8fafc;border-radius:8px;border:1px solid #e2e8f0;flex-direction:column;align-items:flex-start}}
.school-bar-header{{display:flex;justify-content:space-between;width:100%;margin-bottom:9px}}
.school-bar-name-inline{{font-size:15px;font-weight:600;color:#374151}}
.school-count{{font-size:13px;color:#6b7280}}
.school-bar-container{{width:100%;position:relative}}
.school-percentage-bar{{height:22px;border-radius:11px;background:#e5e7eb;position:relative;overflow:hidden;display:flex;align-items:center;justify-content:flex-start}}
.school-percentage-bar.zero-count-bar{{border:1px dashed #cbd5f5;background:#f8fafc}}
.school-percentage-bar.no-data-bar{{background:#6b7280}}
.school-percentage-fill{{height:100%;border-radius:11px;transition:width 0.3s ease;display:block}}
.school-percentage-label{{position:absolute;inset:0;display:flex;align-items:center;justify-content:center;font-size:13px;font-weight:600;color:#1f2937;pointer-events:none;text-align:center;padding:0 8px;white-space:nowrap;z-index:1}}
.school-percentage-label.label-on-fill{{color:white}}
.school-percentage-label.label-no-data{{color:white}}
.school-percentage-label.label-zero-count{{color:#475569;font-style:italic}}
.performance-excellent{{background:linear-gradient(135deg,#3b82f6,#1d4ed8);box-shadow:0 4px 15px rgba(59,130,246,0.3)}}
.performance-good{{background:linear-gradient(135deg,#10b981,#059669);box-shadow:0 4px 15px rgba(16,185,129,0.3)}}
.performance-warning{{background:linear-gradient(135deg,#f59e0b,#d97706);box-shadow:0 4px 15px rgba(245,158,11,0.3)}}
.performance-poor{{background:linear-gradient(135deg,#ef4444,#dc2626);box-shadow:0 4px 15px rgba(239,68,68,0.3)}}
.performance-no-data{{background:linear-gradient(135deg,#6b7280,#4b5563);box-shadow:0 4px 15px rgba(107,114,128,0.2)}}
.tooltip-trigger{{display:inline-flex;align-items:center;cursor:help;margin-left:8px;font-size:16px;color:#6b7280;transition:color 0.2s ease}}
.tooltip-trigger:hover{{color:#3b82f6}}
.tooltip-content{{visibility:hidden;opacity:0;position:fixed;z-index:9999;background-color:#1f2937;color:white;padding:16px;border-radius:8px;font-size:14px;line-height:1.5;max-width:320px;width:max-content;box-shadow:0 10px 25px rgba(0,0,0,0.3);transition:opacity 0.3s,visibility 0.3s;pointer-events:none}}
@media (max-width:1200px){{.kpi-grid{{grid-template-columns:repeat(2,1fr)}}}}
@media (max-width:768px){{.kpi-grid{{grid-template-columns:1fr}}.container{{padding:10px}}.header h1{{font-size:2rem}}.kpi-card{{padding:20px}}.kpi-title{{font-size:16px}}.kpi-value{{font-size:20px;padding:6px 12px;min-width:70px}}.view-button,.sort-button{{padding:8px 16px;font-size:14px;margin:0 3px}}.compressed-view .kpi-grid{{display:block}}.compressed-view .kpi-card{{flex-direction:column;align-items:center;justify-content:center;padding:14px;position:relative;flex-basis:100%;width:100%}}.compressed-view .kpi-value{{margin-bottom:8px;align-self:center;text-align:center}}.compressed-view .university-comparison{{align-self:center;justify-content:center}}}}
</style>
</head>
<body>
<div class="container">
<div class="header">
<h1> {faculty_name} Faculty</h1>
<p>Health & Safety KPI Overview & School Breakdown</p>
</div>
<div class="faculty-overview">
<div class="overview-title">📊 Faculty KPI Overview</div>
<p id="instructions" style="text-align:center;color:#64748b;margin-bottom:20px;font-size:15px">Click any KPI card to view school performance breakdown</p>
<div class="controls-container">
<div class="view-controls">
<button class="view-button active" id="detailedView" onclick="toggleViewMode('detailed')">Detailed View</button>
<button class="view-button" id="compressedView" onclick="toggleViewMode('compressed')">Compressed View</button>
</div>
<div class="sort-controls">
<button class="sort-button active" id="sortHighLow" onclick="sortFacultyCards('highLow')">Highest First</button>
<button class="sort-button" id="sortLowHigh" onclick="sortFacultyCards('lowHigh')">Lowest First</button>
</div>
</div>
<div class="kpi-grid" id="facultyKpis"></div>
</div>
</div>
<script>
const facultyData={faculty_json};const facultyName="{faculty_name}";const tooltipData={tooltip_json};let currentSortOrder='highLow';let currentViewMode='detailed';let globalTooltip=null;
function createTooltip(kpiName,tooltipText){{if(!tooltipText)return '';return `<span class="tooltip-trigger" data-tooltip="${{encodeURIComponent(tooltipText)}}" data-kpi="${{kpiName}}"><svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><circle cx="12" cy="12" r="10"></circle><path d="M9,9h0a3,3,0,0,1,6,0c0,2-3,3-3,3"></path><path d="m12,17h0"></path></svg></span>`;}}
function createGlobalTooltip(){{if(globalTooltip)return;globalTooltip=document.createElement('div');globalTooltip.className='tooltip-content';globalTooltip.style.visibility='hidden';globalTooltip.style.opacity='0';document.body.appendChild(globalTooltip);}}
function positionTooltips(){{createGlobalTooltip();document.querySelectorAll('.tooltip-trigger').forEach(trigger=>{{trigger.removeEventListener('mouseenter',handleTooltipShow);trigger.removeEventListener('mouseleave',handleTooltipHide);trigger.addEventListener('mouseenter',handleTooltipShow);trigger.addEventListener('mouseleave',handleTooltipHide);}});}}
function handleTooltipShow(event){{const trigger=event.currentTarget;const tooltipText=decodeURIComponent(trigger.getAttribute('data-tooltip'));if(!tooltipText||!globalTooltip)return;showTooltip(trigger,tooltipText);}}
function handleTooltipHide(event){{hideTooltip();}}
function showTooltip(trigger,tooltipText){{if(!globalTooltip)return;globalTooltip.innerHTML=tooltipText;const triggerRect=trigger.getBoundingClientRect();const viewportWidth=window.innerWidth;const viewportHeight=window.innerHeight;globalTooltip.style.visibility='visible';globalTooltip.style.opacity='0';globalTooltip.style.left='0px';globalTooltip.style.top='0px';const tooltipRect=globalTooltip.getBoundingClientRect();const triggerCenterX=triggerRect.left+(triggerRect.width/2);const triggerTopY=triggerRect.top;let left=triggerCenterX-(tooltipRect.width/2);let top=triggerTopY-tooltipRect.height-12;if(left<10){{left=10;}}else if(left+tooltipRect.width>viewportWidth-10){{left=viewportWidth-tooltipRect.width-10;}}if(top<10){{top=triggerRect.bottom+12;}}globalTooltip.style.left=left+'px';globalTooltip.style.top=top+'px';globalTooltip.style.visibility='visible';globalTooltip.style.opacity='1';}}
function hideTooltip(){{if(globalTooltip){{globalTooltip.style.visibility='hidden';globalTooltip.style.opacity='0';}}}}
function getPerformanceClass(percentage){{if(percentage===null||percentage===undefined)return 'performance-no-data';if(percentage>=95)return 'performance-excellent';if(percentage>=80)return 'performance-good';if(percentage>=60)return 'performance-warning';return 'performance-poor';}}
function formatValue(percentage,number){{if(percentage===null||percentage===undefined)return '/';if(number===null||number===undefined)return `${{percentage.toFixed(1)}}%`;return `${{percentage.toFixed(1)}}%`;}}
function formatDisplayTextWithCount(percentage,number){{if(percentage===null||percentage===undefined)return '/';if(number===null||number===undefined)return `${{percentage.toFixed(1)}}%`;return `${{percentage.toFixed(1)}}% (${{Math.round(number)}})`;}}
function createUniversityComparison(kpiData){{const facultyPercentage=kpiData.percentage;const universityPercentage=kpiData.university_percentage;const comparison=kpiData.comparison;if(facultyPercentage===null||universityPercentage===null){{return '';}}let indicatorText='';let indicatorClass='';switch(comparison){{case 'above':indicatorText='▲';indicatorClass='comparison-above';break;case 'below':indicatorText='▼';indicatorClass='comparison-below';break;case 'equal':indicatorText='=';indicatorClass='comparison-equal';break;default:return '';}}return `<div class="university-comparison"><div class="comparison-text">vs University:</div><div class="comparison-values"><span class="university-value">${{universityPercentage.toFixed(1)}}%</span><span class="comparison-indicator ${{indicatorClass}}">${{indicatorText}}</span></div></div>`;}}
function toggleCard(cardElement,kpiName){{cardElement.classList.toggle('expanded');}}
function sortFacultyCards(order){{currentSortOrder=order;document.getElementById('sortHighLow').classList.toggle('active',order==='highLow');document.getElementById('sortLowHigh').classList.toggle('active',order==='lowHigh');initializeFacultyDashboard();setTimeout(()=>{{positionTooltips();}},100);}}
function toggleViewMode(mode){{currentViewMode=mode;document.getElementById('detailedView').classList.toggle('active',mode==='detailed');document.getElementById('compressedView').classList.toggle('active',mode==='compressed');const container=document.querySelector('.faculty-overview');const instructions=document.getElementById('instructions');if(mode==='compressed'){{container.classList.add('compressed-view');instructions.textContent='Compact view showing faculty and university averages for all KPIs';}}else{{container.classList.remove('compressed-view');instructions.textContent='Click any KPI card to view school performance breakdown';}}initializeFacultyDashboard();setTimeout(()=>{{positionTooltips();}},100);}}
function initializeFacultyDashboard(){{const facultyKpisContainer=document.getElementById('facultyKpis');facultyKpisContainer.innerHTML='';if(facultyData.kpis){{let kpiEntries=Object.entries(facultyData.kpis);if(currentSortOrder==='highLow'){{kpiEntries.sort((a,b)=>{{const aPercentage=a[1].percentage!==null?a[1].percentage:-1;const bPercentage=b[1].percentage!==null?b[1].percentage:-1;return bPercentage-aPercentage;}});}}else{{kpiEntries.sort((a,b)=>{{const aPercentage=a[1].percentage!==null?a[1].percentage:999;const bPercentage=b[1].percentage!==null?b[1].percentage:999;return aPercentage-bPercentage;}});}}for(const[kpiName,kpiData]of kpiEntries){{const card=document.createElement('div');card.className='kpi-card';card.id=`card-${{kpiName.replace(/[^a-zA-Z0-9]/g,'-')}}`;const performanceClass=getPerformanceClass(kpiData.percentage);let schoolBreakdownHtml='';const hasSchools=facultyData.schools&&Object.keys(facultyData.schools).length>0;if(hasSchools){{let schoolEntries=Object.entries(facultyData.schools).filter(([schoolName,schoolData])=>schoolData.kpis&&schoolData.kpis[kpiName]);if(currentSortOrder==='highLow'){{schoolEntries.sort((a,b)=>{{const aPerc=a[1].kpis[kpiName].percentage!==null?a[1].kpis[kpiName].percentage:-1;const bPerc=b[1].kpis[kpiName].percentage!==null?b[1].kpis[kpiName].percentage:-1;return bPerc-aPerc;}});}}else{{schoolEntries.sort((a,b)=>{{const aPerc=a[1].kpis[kpiName].percentage!==null?a[1].kpis[kpiName].percentage:999;const bPerc=b[1].kpis[kpiName].percentage!==null?b[1].kpis[kpiName].percentage:999;return aPerc-bPerc;}});}}if(schoolEntries.length>0){{schoolBreakdownHtml=`<div class="school-breakdown-content"><div class="school-breakdown-title">School Performance (${{schoolEntries.length}} schools):</div>`;for(const[schoolName,schoolData]of schoolEntries){{
    const schoolKpi=schoolData.kpis[kpiName];
    const percentage=schoolKpi.percentage;
    const number=schoolKpi.number;
    let barWidth=0;
    let barColor='#6b7280';
    const labelClasses=['school-percentage-label'];
    const barClasses=['school-percentage-bar'];
    let labelText='';
    const hasNumber=(number!==null&&number!==undefined&&number!=='/'&&!isNaN(parseFloat(number)));
    const countValue=hasNumber?parseFloat(number):null;
    const hasZeroCount=countValue===0;
    if(hasZeroCount){{
        barClasses.push('zero-count-bar');
    }}
    if(percentage===null||percentage===undefined||(typeof percentage==='string'&&percentage.trim()==='/')){{
        labelText='No Return Provided';
        labelClasses.push('label-no-data');
        barClasses.push('no-data-bar');
    }}else{{
        const numericPercentage=parseFloat(percentage);
        if(!isNaN(numericPercentage)){{
            const boundedPercentage=Math.max(0,Math.min(100,numericPercentage));
            if(hasZeroCount){{
                labelText='0 items returned';
                labelClasses.push('label-zero-count');
            }}else{{
                barWidth=boundedPercentage;
                const performanceClass=getPerformanceClass(numericPercentage);
                switch(performanceClass){{
                    case 'performance-excellent':
                        barColor='#3b82f6';
                        break;
                    case 'performance-good':
                        barColor='#10b981';
                        break;
                    case 'performance-warning':
                        barColor='#f59e0b';
                        break;
                    case 'performance-poor':
                        barColor='#ef4444';
                        break;
                    default:
                        barColor='#6b7280';
                }}
                labelText=`${{numericPercentage.toFixed(1)}}%`;
                if(boundedPercentage>=60){{
                    labelClasses.push('label-on-fill');
                }}
            }}
        }}else{{
            labelText='No Return Provided';
            labelClasses.push('label-no-data');
            barClasses.push('no-data-bar');
        }}
    }}
    if(labelText===''){{
        labelText='No Return Provided';
        if(!labelClasses.includes('label-no-data')){{
            labelClasses.push('label-no-data');
        }}
        if(!barClasses.includes('no-data-bar')){{
            barClasses.push('no-data-bar');
        }}
    }}
    const labelClass=labelClasses.join(' ');
    const barClass=barClasses.join(' ');
    const countDisplay=countValue!==null?`<div class="school-count">${{Math.round(countValue)}} items</div>`:'';
    const fillMarkup=!hasZeroCount&&barWidth>0?`<div class="school-percentage-fill" style="width:${{barWidth}}%;background:${{barColor}};"></div>`:'';
    schoolBreakdownHtml+=`<div class="school-bar-inline"><div class="school-bar-header"><div class="school-bar-name-inline">${{schoolName}}</div>${{countDisplay}}</div><div class="school-bar-container"><div class="${{barClass}}">${{fillMarkup}}<div class="${{labelClass}}">${{labelText}}</div></div></div></div>`;
}}
schoolBreakdownHtml+='</div>';
}}
}}const hasExpandableContent=schoolBreakdownHtml!=='';const expandIcon=hasExpandableContent?'<div class=\"expand-icon\">▼</div>':'';const tooltipHtml=createTooltip(kpiName,tooltipData[kpiName]);const universityComparisonHtml=createUniversityComparison(kpiData);if(currentViewMode==='compressed'){{card.innerHTML=`<div class="kpi-title">${{kpiName}} ${{tooltipHtml}}</div><div class="kpi-value ${{performanceClass}}">${{formatDisplayTextWithCount(kpiData.percentage,kpiData.number)}}</div>${{universityComparisonHtml}}${{schoolBreakdownHtml}}${{expandIcon}}`;if(hasExpandableContent){{card.onclick=()=>toggleCard(card,kpiName);}}else{{card.style.cursor='default';card.onclick=null;}}}}else{{card.onclick=()=>toggleCard(card,kpiName);card.innerHTML=`<div class="kpi-header"><div class="kpi-title">${{kpiName}} ${{tooltipHtml}}</div><div class="kpi-value ${{performanceClass}}">${{formatDisplayTextWithCount(kpiData.percentage,kpiData.number)}}</div></div>${{universityComparisonHtml}}${{schoolBreakdownHtml}}${{expandIcon}}`;if(!hasExpandableContent){{card.style.cursor='default';card.onclick=null;}}}}facultyKpisContainer.appendChild(card);}}}}}}
initializeFacultyDashboard();setTimeout(()=>{{positionTooltips();}},100);window.addEventListener('resize',()=>{{setTimeout(positionTooltips,100);}});
</script>
</body>
</html>'''

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
