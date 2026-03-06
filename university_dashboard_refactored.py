import pandas as pd
import json
import tkinter as tk
from tkinter import filedialog, messagebox
import os


class UniversityDashboardGenerator:
    def __init__(self):
        self.university_data = None
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
        for kpi_name, kpi_def in self.kpi_definitions.items():
            perc_val = row.get(kpi_def["percentage_col"], None)
            num_val = row.get(kpi_def["number_col"], None) if kpi_def["number_col"] else None
            perc_val = self._safe_float(perc_val) if not self._is_empty(perc_val) else None
            num_val = self._safe_float(num_val) if not self._is_empty(num_val) else None
            kpis[kpi_name] = {
                "percentage": perc_val,
                "number": num_val,
                "display_text": self._format_display(perc_val, num_val)
            }
        return {"name": "University", "kpis": kpis}

    def _performance_class(self, percentage):
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
        tooltip_json = json.dumps(self.tooltip_data or {}, indent=2, default=str)

        # HTML style/theme mirrors faculty_dashboard_refactored.py (simplified; no controls)
        # Cards show percentage and optional number in a colored pill reflecting performance
        html_content = f'''<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>University - Health & Safety Dashboard</title>
<style>
*{{margin:0;padding:0;box-sizing:border-box}}
body{{font-family:'Segoe UI',Tahoma,Geneva,Verdana,sans-serif;min-height:100vh;color:#333;font-size:15px}}
.container{{max-width:1200px;margin:0 auto;padding:20px}}
.header{{text-align:center;color:#1e293b;margin-bottom:30px;background:linear-gradient(135deg,#667eea 0%,#764ba2 100%);padding:40px;border-radius:20px;box-shadow:0 8px 32px rgba(0,0,0,0.12);position:relative;overflow:hidden}}
.header::before{{content:'';position:absolute;top:0;left:0;right:0;bottom:0;background:linear-gradient(45deg,rgba(255,255,255,0.1) 0%,transparent 50%,rgba(255,255,255,0.05) 100%);pointer-events:none}}
.header h1{{font-size:2.3rem;margin-bottom:10px;font-weight:800;color:white;text-shadow:0 2px 4px rgba(0,0,0,0.1);position:relative;z-index:1}}
.header p{{color:rgba(255,255,255,0.9);font-size:1rem;font-weight:500;position:relative;z-index:1}}
.kpi-grid{{display:grid;grid-template-columns:repeat(2,1fr);gap:12px;margin-bottom:30px}}
.kpi-card{{background:white;border-radius:8px;padding:8px 12px;box-shadow:0 4px 20px rgba(0,0,0,0.08);border:1px solid #e2e8f0;position:relative;overflow:hidden;display:grid;grid-template-columns:7fr 3fr auto;column-gap:12px;align-items:center}}
.kpi-card::before{{content:'';position:absolute;top:0;left:0;right:0;height:4px}}
.kpi-title{{font-size:14px;font-weight:600;color:#374151;line-height:1.3;display:flex;align-items:center;margin:0}}
.kpi-value{{font-size:14px;font-weight:700;color:white;padding:6px 10px;border-radius:6px;text-align:center;box-shadow:0 4px 16px rgba(0,0,0,0.15);margin:0;display:inline-flex;align-items:center;justify-self:center;min-height:28px}}
.tooltip-trigger{{display:inline-flex;align-items:center;cursor:help;margin-left:6px;font-size:14px;color:#6b7280;transition:color 0.2s ease}}
.tooltip-trigger:hover{{color:#3b82f6}}
.tooltip-content{{visibility:hidden;opacity:0;position:fixed;z-index:9999;background-color:#1f2937;color:white;padding:12px 14px;border-radius:8px;font-size:14px;line-height:1.4;max-width:320px;width:max-content;box-shadow:0 10px 25px rgba(0,0,0,0.3);transition:opacity 0.2s,visibility 0.2s;pointer-events:none}}
.performance-excellent{{background:linear-gradient(135deg,#10b981,#059669);box-shadow:0 4px 15px rgba(16,185,129,0.3)}}
.performance-good{{background:linear-gradient(135deg,#3b82f6,#1d4ed8);box-shadow:0 4px 15px rgba(59,130,246,0.3)}}
.performance-warning{{background:linear-gradient(135deg,#f59e0b,#d97706);box-shadow:0 4px 15px rgba(245,158,11,0.3)}}
.performance-poor{{background:linear-gradient(135deg,#ef4444,#dc2626);box-shadow:0 4px 15px rgba(239,68,68,0.3)}}
.performance-no-data{{background:linear-gradient(135deg,#6b7280,#4b5563);box-shadow:0 4px 15px rgba(107,114,128,0.2)}}
@media (max-width:1200px){{.kpi-grid{{grid-template-columns:repeat(2,1fr)}}}}
@media (max-width:768px){{.kpi-grid{{grid-template-columns:1fr}}}}
</style>
</head>
<body>
  <div class="container">
    <div class="header">
      <h1>University - Health & Safety Dashboard</h1>
      <p>Overall KPI performance across the University</p>
    </div>

    <div class="kpi-grid compressed-view">
      <!-- KPI Cards inserted here -->
      {self._render_kpi_cards(uni_data)}
    </div>
  </div>

  <div id="tooltip" class="tooltip-content"></div>
  <script>
    const tooltipData = {tooltip_json};
    let globalTooltip = null;
    function createGlobalTooltip(){{
      if(globalTooltip) return;
      globalTooltip = document.getElementById('tooltip');
    }}
    function handleTooltipShow(e){{
      const trigger = e.currentTarget;
      const key = trigger.getAttribute('data-kpi');
      const text = tooltipData[key] || '';
      showTooltip(trigger, text);
    }}
    function handleTooltipHide(){{
      if(!globalTooltip) return;
      globalTooltip.style.visibility='hidden';
      globalTooltip.style.opacity='0';
    }}
    function showTooltip(trigger, text){{
      createGlobalTooltip();
      if(!globalTooltip) return;
      globalTooltip.textContent = text;
      const rect = trigger.getBoundingClientRect();
      const vw = window.innerWidth;
      globalTooltip.style.visibility='visible';
      globalTooltip.style.opacity='0';
      globalTooltip.style.left='0px';
      globalTooltip.style.top='0px';
      const tipRect = globalTooltip.getBoundingClientRect();
      let left = rect.left + (rect.width/2) - (tipRect.width/2);
      let top = rect.top - tipRect.height - 12;
      if(left < 10) left = 10;
      if(left + tipRect.width > vw - 10) left = vw - tipRect.width - 10;
      if(top < 10) top = rect.bottom + 12;
      globalTooltip.style.left = left + 'px';
      globalTooltip.style.top = top + 'px';
      globalTooltip.style.visibility='visible';
      globalTooltip.style.opacity='1';
    }}
  </script>
</body>
</html>
'''

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
            perf_cls = self._performance_class(pct)
            safe_key = kpi_name.replace('"', "&quot;")
            card = f'''<div class="kpi-card">
                <div class="kpi-title">{kpi_name}
                    <span class="tooltip-trigger" data-kpi="{safe_key}" title="Info" onmouseenter="handleTooltipShow(event)" onmouseleave="handleTooltipHide()">&#9432;</span>
                </div>
                <span class="kpi-value {perf_cls}">{val}</span>
                <span style="height:1px"></span>
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
