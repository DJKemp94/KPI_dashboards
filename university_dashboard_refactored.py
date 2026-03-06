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
                "denominator_label": "arrangements"
            },
            "Risk Assessments in Register up to date": {
                "question": "Risk assessments updated in the last 2 years out of all risk assessments listed on the register.",
                "denominator_label": "risk assessments"
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
                "denominator_label": "buildings"
            },
            "PEEPS in Place": {
                "question": "PEEPs in place (reviewed, communicated, controls in place) out of PEEPs required.",
                "denominator_label": "PEEPs"
            },
            "PEEPS Drilled": {
                "question": "PEEPs tested/drilled in the period out of PEEPs required.",
                "denominator_label": "PEEPs"
            },
            "Assets without A&B Defects": {
                "question": "BU-owned assets without unresolved A/B defects.",
                "denominator_label": "assets"
            },
            "Assets Inspected by Allianz": {
                "question": "BU-owned assets inspected by Allianz (not overdue / plant available).",
                "denominator_label": "assets"
            },
            "Accidents and Incidents Investigated": {
                "question": "Investigations completed out of total incidents and near misses reported in the period.",
                "denominator_label": "incidents/near misses"
            },
            "Inspections Carried Out": {
                "question": "Inspections carried out out of inspections on the monitoring schedule.",
                "denominator_label": "inspections"
            },
            "Leadership Walkarounds": {
                "question": "Leadership walkarounds completed out of walkarounds on the monitoring schedule.",
                "denominator_label": "walkarounds"
            },
            "Risk Assessment Coverage": {
                "question": "Percentage coverage of risk assessment across department/PS.",
                "denominator_label": None
            },
            "Training Matrix Coverage": {
                "question": "Training in the matrix that is accessible to staff who need it.",
                "denominator_label": None
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
        """Build KPI -> [{date, percentage}] from University_Summary_History."""
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
                pct = self._safe_float(val) if not self._is_empty(val) else None
                date_label = row.get('Date', '')
                series.append({
                    "date": str(date_label) if pd.notna(date_label) else "",
                    "percentage": pct
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

    def _format_display(self, kpi_name, percentage, number):
        if percentage is None:
            return "No return submitted"
        if number is None:
            return f"{percentage:.1f}%"
        denominator_label = self.kpi_metadata.get(kpi_name, {}).get("denominator_label")
        if denominator_label:
            return f"{percentage:.1f}% (base: {int(number)} {denominator_label})"
        return f"{percentage:.1f}% (base: {int(number)})"

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
                "display_text": self._format_display(kpi_name, perc_val, num_val),
                "history": history.get(kpi_name, [])
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
        uni_json = json.dumps(uni_data or {}, indent=2, default=str)

        # HTML style/theme mirrors faculty_dashboard_refactored.py (simplified; no controls)
        # Cards show percentage and optional number in a colored pill reflecting performance
        html_content = f'''<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>University KPI Dashboard</title>
<script src="https://cdnjs.cloudflare.com/ajax/libs/Chart.js/3.9.1/chart.min.js"></script>
<style>
*{{margin:0;padding:0;box-sizing:border-box}}
body{{font-family:'Source Sans 3','Segoe UI',Tahoma,Geneva,Verdana,sans-serif;background:#f5f7fb;min-height:100vh;color:#1f2937;font-size:15px}}
.container{{max-width:1360px;margin:0 auto;padding:24px}}
.header{{text-align:left;color:#0f172a;margin-bottom:18px;background:linear-gradient(145deg,#1e3a5f 0%,#2c5a87 100%);padding:28px;border-radius:14px;box-shadow:0 10px 24px rgba(15,23,42,0.18)}}
.header h1{{font-size:2rem;margin-bottom:6px;font-weight:700;color:#f8fafc}}
.header p{{color:#dbeafe;font-size:1rem;font-weight:500}}
.kpi-grid{{display:grid;grid-template-columns:repeat(2,minmax(0,1fr));gap:16px;margin-bottom:20px}}
.kpi-card{{background:#fff;border-radius:12px;padding:14px 16px;box-shadow:0 2px 10px rgba(15,23,42,0.06);border:1px solid #dbe2ea;position:relative;overflow:hidden;display:grid;grid-template-columns:minmax(0,1fr) auto;column-gap:14px;row-gap:8px;align-items:start}}
.kpi-card::before{{content:'';position:absolute;top:0;left:0;right:0;height:4px}}
.kpi-title{{font-size:15px;font-weight:700;color:#0f172a;line-height:1.25;display:flex;align-items:center;gap:4px;margin:0}}
.kpi-label-wrap{{display:flex;flex-direction:column;gap:4px}}
.kpi-question{{font-size:12px;color:#475569;line-height:1.35}}
.kpi-value{{font-size:14px;font-weight:700;color:white;padding:7px 11px;border-radius:8px;text-align:center;box-shadow:none;margin:0;display:inline-flex;align-items:center;justify-self:end;min-height:30px}}
.trend-wrap{{grid-column:1/-1;height:94px;background:#f8fafc;border:1px solid #e2e8f0;border-radius:10px;padding:8px}}
.trend-canvas{{width:100%!important;height:76px!important}}
.tooltip-trigger{{display:inline-flex;align-items:center;cursor:help;margin-left:6px;font-size:14px;color:#6b7280;transition:color 0.2s ease}}
.tooltip-trigger:hover{{color:#3b82f6}}
.data-warning{{display:inline-flex;align-items:center;justify-content:center;width:18px;height:18px;border-radius:50%;background:#dc2626;color:#fff;font-size:12px;font-weight:700;margin-left:6px;vertical-align:middle;cursor:default}}
.tooltip-content{{visibility:hidden;opacity:0;position:fixed;z-index:9999;background-color:#1f2937;color:white;padding:12px 14px;border-radius:8px;font-size:14px;line-height:1.4;max-width:320px;width:max-content;box-shadow:0 10px 25px rgba(0,0,0,0.3);transition:opacity 0.2s,visibility 0.2s;pointer-events:none}}
.performance-excellent{{background:linear-gradient(135deg,#10b981,#059669);box-shadow:0 4px 15px rgba(16,185,129,0.3)}}
.performance-good{{background:linear-gradient(135deg,#3b82f6,#1d4ed8);box-shadow:0 4px 15px rgba(59,130,246,0.3)}}
.performance-warning{{background:linear-gradient(135deg,#f59e0b,#d97706);box-shadow:0 4px 15px rgba(245,158,11,0.3)}}
.performance-poor{{background:linear-gradient(135deg,#ef4444,#dc2626);box-shadow:0 4px 15px rgba(239,68,68,0.3)}}
.performance-no-data{{background:linear-gradient(135deg,#6b7280,#4b5563);box-shadow:0 4px 15px rgba(107,114,128,0.2)}}
.guide-panel{{background:#ffffff;border:1px solid #dbe2ea;border-radius:12px;padding:12px 14px;margin-bottom:16px;box-shadow:0 2px 8px rgba(15,23,42,0.04)}}
.guide-title{{font-size:14px;font-weight:700;color:#0f172a;margin-bottom:8px}}
.guide-line{{font-size:13px;color:#334155;line-height:1.45}}
@media (max-width:1100px){{.kpi-grid{{grid-template-columns:1fr}}}}
@media (max-width:768px){{.container{{padding:12px}}.header{{padding:18px}}.header h1{{font-size:1.55rem}}.kpi-card{{padding:12px}}}}
</style>
</head>
<body>
  <div class="container">
    <div class="header">
      <h1>University KPI Dashboard</h1>
      <p>Overall KPI performance across the University</p>
    </div>
    <div class="guide-panel">
      <div class="guide-title">How to read this</div>
      <div class="guide-line">Each score is a percentage for the KPI. Where shown, <strong>base</strong> is the denominator used for that percentage.</div>
      <div class="guide-line"><span class="data-warning">!</span> means a source percentage is outside 0-100 and should be reviewed with the submitting department.</div>
      <div class="guide-line">Trend charts show reporting periods in date order.</div>
    </div>

    <div class="kpi-grid compressed-view">
      <!-- KPI Cards inserted here -->
      {self._render_kpi_cards(uni_data)}
    </div>
  </div>

  <div id="tooltip" class="tooltip-content"></div>
  <script>
    const tooltipData = {tooltip_json};
    const uniData = {uni_json};
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
    function buildTrendSeries(history){{
      if(!Array.isArray(history)) return null;
      const cleaned = history.filter(p => p && p.date && typeof p.percentage === 'number');
      if(cleaned.length < 2) return null;
      return {{
        labels: cleaned.map(p => p.date),
        data: cleaned.map(p => p.percentage)
      }};
    }}
    function renderTrendCharts(){{
      document.querySelectorAll('.trend-canvas').forEach(canvas => {{
        const historyRaw = canvas.getAttribute('data-history') || '[]';
        let history = [];
        try {{ history = JSON.parse(historyRaw); }} catch (e) {{ history = []; }}
        const series = buildTrendSeries(history);
        if(!series) {{
          const wrap = canvas.closest('.trend-wrap');
          if(wrap) wrap.style.display='none';
          return;
        }}
        const ctx = canvas.getContext('2d');
        new Chart(ctx, {{
          type: 'line',
          data: {{
            labels: series.labels,
            datasets: [{{
              label: 'University',
              data: series.data,
              borderColor: '#2563eb',
              backgroundColor: 'rgba(37,99,235,0.12)',
              fill: true,
              tension: 0.25,
              pointRadius: 3
            }}]
          }},
          options: {{
            responsive: true,
            maintainAspectRatio: false,
            plugins: {{ legend: {{ display: false }}, tooltip: {{ enabled: true }} }},
            scales: {{
              x: {{ ticks: {{ maxRotation: 0, autoSkip: true, maxTicksLimit: 4 }}, grid: {{ display: false }} }},
              y: {{ min: 0, max: 100, ticks: {{ callback: v => `${{v}}%` }}, grid: {{ color: '#e5e7eb' }} }}
            }}
          }}
        }});
      }});
    }}
    renderTrendCharts();
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
            history_json = json.dumps(kpi.get("history", []), default=str).replace('"', '&quot;')
            out_of_range = False
            if pct is not None:
                try:
                    out_of_range = float(pct) < 0 or float(pct) > 100
                except (ValueError, TypeError):
                    out_of_range = False
            warning_badge = '<span class="data-warning" title="Out-of-range percentage in source data">!</span>' if out_of_range else ''
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
                <div class="trend-wrap"><canvas class="trend-canvas" data-history="{history_json}"></canvas></div>
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
