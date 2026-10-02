import pandas as pd
import json
import tkinter as tk
from tkinter import filedialog, messagebox
import os
from pathlib import Path


class UniversityDashboardGenerator:
    def __init__(self):
        self.university_data = None
        self.faculty_data = None
        self.school_data = None
        self.tooltip_data = None
        self.kpi_definitions = self._define_kpis()
        self.faculty_school_mapping = self._define_faculty_school_mapping()
        self.kpi_tooltips = self._define_kpi_tooltips()

    def _define_kpis(self):
        """Define the 13 KPI metrics with their column mappings"""
        return {
            "Written Arrangements Complete": {
                "percentage_col": "% of Written Arrangements Complete",
                "number_col": "Number of Arrangements",
                "completed_col": "Number of Arrangements Completed"
            },
            "Risk Assessments in Register up to date": {
                "percentage_col": "% Risk Assessments on Register up-to-date",
                "number_col": "Number of Risk Assessments on Register",
                "completed_col": None
            },
            "H&S Induction Completion": {
                "percentage_col": "% of Staff Completed UoN H&S Induction",
                "number_col": "Number of Staff",
                "completed_col": None
            },
            "Fire Training Completion": {
                "percentage_col": "% of Staff Completed UoN Fire Training", 
                "number_col": "Number of Staff",
                "completed_col": None
            },
            "Fire Drills Completed": {
                "percentage_col": "% of Fire Drills Carried out",
                "number_col": "Number of Buildings Allocated for Fire Drills to be undertaken",
                "completed_col": None
            },
            "PEEPS in Place": {
                "percentage_col": "% of PEEPS in Place, Reviewed and Controlled",
                "number_col": "Number of PEEPS Identified",
                "completed_col": None
            },
            "PEEPS Drilled": {
                "percentage_col": "% of PEEPS that are tested/drilled",
                "number_col": "Number of PEEPS Identified",
                "completed_col": None
            },
            "Assets without A&B Defects": {
                "percentage_col": "% of Assets without active A and B defects",
                "number_col": "Number of BU Owned Assets",
                "completed_col": None
            },
            "Assets Inspected by Allianz": {
                "percentage_col": "% of Assets seen to by Allianz",
                "number_col": "Number of BU Owned Assets", 
                "completed_col": None
            },
            "Accidents and Incidents Investigated": {
                "percentage_col": "% of Incidents + Near Missed Investigated",
                "number_col": "Total Incidents (Accidents + Near Misses)",
                "completed_col": None
            },
            "Inspections Carried Out": {
                "percentage_col": "% of Inspections Carried out against Monitoring Schedule",
                "number_col": "Number of Inspections on Monitoring Schedule",
                "completed_col": None
            },
            "Leadership Walkarounds": {
                "percentage_col": "% of Leadership Walkarounds Carried out",
                "number_col": "Number of Leadership walkarounds on Monitoring Schedule",
                "completed_col": None
            },
            "Risk Assessment Coverage": {
                "percentage_col": "Percentage Coverage of Risk Assessments",
                "number_col": None,
                "completed_col": None
            },
            "Training Matrix Coverage": {
                "percentage_col": "% of Training identified in Matrix that is accessible",
                "number_col": None,
                "completed_col": None
            },
            "Staff Training Requirements": {
                "percentage_col": "% of Staff who are in date with all training requirements",
                "number_col": "Number of Staff",
                "completed_col": None
            }
        }

    def _define_faculty_school_mapping(self):
        """Define the exact mapping between schools and faculties as provided"""
        return {
            # Arts Faculty
            "CLAS": "Arts",
            "English": "Arts",
            "Humanities": "Arts",

            # Engineering Faculty  
            "Engineering": "Engineering",

            # Estates Faculty
            "Estates": "Estates",

            # Finance and Infrastructure Faculty
            "Finance and Infrastructure": "Finance and Infrastructure",

            # HR Faculty
            "HR": "HR",

            # Medicine & Health Sciences Faculty
            "BDI": "Medicine & Health Sciences",
            "Life Sciences": "Medicine & Health Sciences",
            "Medicine": "Medicine & Health Sciences",
            "Veterinary Medicine and Science": "Medicine & Health Sciences",
            "Clinical Skills": "Medicine & Health Sciences",
            "Health Sciences": "Medicine & Health Sciences",

            # Registrars Faculty
            "Libraries": "Registrars",
            "Sport": "Registrars",
            "BSU": "Registrars",
            "Student & Public Facing Services": "Registrars",

            # Science Faculty
            "Computer Sciences": "Science",
            "Biosciences": "Science",
            "Chemistry": "Science",
            "Mathematical Sciences": "Science",
            "Pharmacy": "Science",
            "Physics and Astronomy": "Science",
            "Psychology": "Science",

            # Social Sciences Faculty
            "Economics": "Social Sciences",
            "Business School": "Social Sciences",
            "Education": "Social Sciences",
            "Law": "Social Sciences",
            "Politics and IR": "Social Sciences",
            "Rights Lab": "Social Sciences",
            "Sociology": "Social Sciences",
            "Geography": "Social Sciences",
            "Faculty Office": "Social Sciences"
        }

    def _define_kpi_tooltips(self):
        """Define the mapping between KPI names and their tooltip column names"""
        return {
            "Written Arrangements Complete": "% of Written Arrangements Complete",
            "Risk Assessments in Register up to date": "% Risk Assessments on Register up-to-date", 
            "H&S Induction Completion": "% of Staff Completed UoN H&S Induction",
            "Fire Training Completion": "% of Staff Completed UoN Fire Training",
            "Fire Drills Completed": "% of Fire Drills Carried out",
            "PEEPS in Place": "% of PEEPS in Place, Reviewed and Controlled",
            "PEEPS Drilled": "% of PEEPS that are tested0drilled",  # Note: uses '0' instead of '/'
            "Assets without A&B Defects": "% of Assets without active A and B defects",
            "Assets Inspected by Allianz": "% of Assets seen to by Allianz",
            "Accidents and Incidents Investigated": "% of Incidents + Near Missed Investigated",
            "Inspections Carried Out": "% of Inspections Carried out against Monitoring Schedule",
            "Leadership Walkarounds": "% of Leadership Walkarounds Carried out",
            "Risk Assessment Coverage": "Percentage Coverage of Risk Assessments",
            "Training Matrix Coverage": "% of Training identified in Matrix that is accessible",
            "Staff Training Requirements": "% of Staff who are in date with all training requirements"
        }

    def select_excel_file(self):
        """Open file dialog to select Excel file"""
        root = tk.Tk()
        root.withdraw()

        file_path = filedialog.askopenfilename(
            title="Select Health & Safety Excel File",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
        )

        root.destroy()
        return file_path if file_path else None

    def select_output_location(self):
        """Open file dialog to select output location for HTML dashboard"""
        root = tk.Tk()
        root.withdraw()

        file_path = filedialog.asksaveasfilename(
            title="Save University Dashboard HTML As...",
            defaultextension=".html",
            filetypes=[("HTML files", "*.html"), ("All files", "*.*")]
        )

        root.destroy()
        return file_path if file_path else None

    def load_excel_data(self, file_path):
        """Load data from the three sheets in Excel file"""
        try:
            xl_file = pd.ExcelFile(file_path)
            
            # Load university data (Table1)
            if 'University_Summary' in xl_file.sheet_names:
                self.university_data = pd.read_excel(file_path, sheet_name='University_Summary')
            else:
                raise ValueError("University_Summary sheet not found")
            
            # Load faculty data (Table2)  
            if 'Faculty_Summary' in xl_file.sheet_names:
                self.faculty_data = pd.read_excel(file_path, sheet_name='Faculty_Summary')
            else:
                raise ValueError("Faculty_Summary sheet not found")
            
            # Load school data (Table3)
            if 'School_Raw_Data' in xl_file.sheet_names:
                self.school_data = pd.read_excel(file_path, sheet_name='School_Raw_Data')
                # Add faculty mapping to school data
                self.school_data['Faculty'] = self.school_data['School'].map(self.faculty_school_mapping)
            else:
                raise ValueError("School_Raw_Data sheet not found")
            
            # Load tooltip data if available
            if 'Question Tooltips' in xl_file.sheet_names:
                self.tooltip_data = self._load_tooltip_data(file_path)
            else:
                print("Warning: Question Tooltips sheet not found - tooltips will not be available")
                
            return True
            
        except Exception as e:
            messagebox.showerror("Error", f"Could not load Excel file: {str(e)}")
            return False

    def _load_tooltip_data(self, file_path):
        """Load and process tooltip data from Excel file"""
        try:
            tooltip_df = pd.read_excel(file_path, sheet_name='Question Tooltips')
            
            # Extract column names (row 0) and tooltips (row 1)
            tooltip_mapping = {}
            
            for col_idx in range(len(tooltip_df.columns)):
                column_name = tooltip_df.iloc[0, col_idx]
                tooltip_text = tooltip_df.iloc[1, col_idx]
                
                if pd.notna(column_name) and pd.notna(tooltip_text):
                    tooltip_mapping[str(column_name)] = str(tooltip_text)
            
            # Create KPI to tooltip mapping using the predefined mapping
            kpi_tooltip_mapping = {}
            for kpi_name, column_name in self.kpi_tooltips.items():
                if column_name in tooltip_mapping:
                    kpi_tooltip_mapping[kpi_name] = tooltip_mapping[column_name]
            
            print(f"   ✓ Loaded {len(kpi_tooltip_mapping)} KPI tooltips")
            return kpi_tooltip_mapping
            
        except Exception as e:
            print(f"   ⚠ Error loading tooltip data: {str(e)}")
            return {}

    def process_kpi_data(self):
        """Process all KPI data into the format needed for dashboard"""
        dashboard_data = {
            "university": {},
            "faculties": {},
            "schools": {}
        }

        # Process University level data
        if self.university_data is not None and len(self.university_data) > 0:
            uni_row = self.university_data.iloc[0]
            dashboard_data["university"] = self._extract_kpi_data(uni_row, "University")

        # Process Faculty level data
        if self.faculty_data is not None:
            for _, faculty_row in self.faculty_data.iterrows():
                faculty_name = faculty_row['Faculty']
                dashboard_data["faculties"][faculty_name] = self._extract_kpi_data(faculty_row, faculty_name)
                
                # Add schools for this faculty
                faculty_schools = self.school_data[self.school_data['Faculty'] == faculty_name] if self.school_data is not None else pd.DataFrame()
                dashboard_data["faculties"][faculty_name]['schools'] = {}
                
                for _, school_row in faculty_schools.iterrows():
                    school_name = school_row['School']
                    dashboard_data["faculties"][faculty_name]['schools'][school_name] = self._extract_kpi_data(school_row, school_name)

        return dashboard_data

    def _extract_kpi_data(self, data_row, entity_name):
        """Extract KPI data from a data row"""
        kpi_data = {
            "name": entity_name,
            "kpis": {}
        }

        for kpi_name, kpi_def in self.kpi_definitions.items():
            percentage_col = kpi_def["percentage_col"]
            number_col = kpi_def["number_col"]
            
            # Get percentage value
            percentage_value = data_row.get(percentage_col, None)
            if pd.isna(percentage_value) or percentage_value == '/' or percentage_value == '':
                percentage_value = None
            else:
                # Try to convert to numeric
                try:
                    percentage_value = float(percentage_value)
                except (ValueError, TypeError):
                    percentage_value = None
            
            # Get number value  
            number_value = data_row.get(number_col, None) if number_col else None
            if pd.isna(number_value) or number_value == '/' or number_value == '':
                number_value = None
            else:
                # Try to convert to numeric
                try:
                    number_value = float(number_value)
                except (ValueError, TypeError):
                    number_value = None
                
            kpi_data["kpis"][kpi_name] = {
                "percentage": percentage_value,
                "number": number_value,
                "display_text": self._format_kpi_display(percentage_value, number_value)
            }

        return kpi_data

    def _format_kpi_display(self, percentage, number):
        """Format KPI for display"""
        if percentage is None or pd.isna(percentage):
            return "/"
        
        # Ensure percentage is numeric
        try:
            percentage = float(percentage)
        except (ValueError, TypeError):
            return "/"
        
        if number is None or pd.isna(number):
            return f"{percentage:.1f}%"
        else:
            try:
                number = float(number)
                return f"{percentage:.1f}% ({int(number)})"
            except (ValueError, TypeError):
                return f"{percentage:.1f}%"

    def create_html_dashboard(self, dashboard_data, output_path):
        """Create the interactive HTML dashboard for University level"""
        
        # Convert data to JSON for JavaScript
        data_json = json.dumps(dashboard_data, indent=2, default=str)
        
        html_content = f"""<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>University Health & Safety KPI Dashboard</title>
    <script src="https://cdnjs.cloudflare.com/ajax/libs/Chart.js/3.9.1/chart.min.js"></script>
    <style>
        * {{
            margin: 0;
            padding: 0;
            box-sizing: border-box;
        }}

        body {{
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
            background: linear-gradient(135deg, #f8fafc 0%, #e2e8f0 100%);
            min-height: 100vh;
            color: #333;
        }}

        .container {{
            max-width: 1400px;
            margin: 0 auto;
            padding: 20px;
        }}

        .header {{
            text-align: center;
            color: #1e293b;
            margin-bottom: 30px;
            background: white;
            padding: 30px;
            border-radius: 15px;
            box-shadow: 0 4px 20px rgba(0, 0, 0, 0.08);
            border: 1px solid #e2e8f0;
        }}

        .sort-controls {{
            text-align: center;
            margin-bottom: 20px;
        }}

        .sort-button {{
            background: #f1f5f9;
            color: #64748b;
            border: 1px solid #e2e8f0;
            padding: 8px 16px;
            border-radius: 20px;
            font-size: 14px;
            font-weight: 500;
            cursor: pointer;
            transition: all 0.3s ease;
            margin: 0 5px;
        }}

        .sort-button:hover {{
            background: #e2e8f0;
            color: #334155;
        }}

        .sort-button.active {{
            background: #3b82f6;
            color: white;
            font-weight: 600;
        }}

        .header h1 {{
            font-size: 2.5rem;
            margin-bottom: 10px;
            font-weight: 700;
        }}

        .level-indicator {{
            display: inline-block;
            background: white;
            color: #64748b;
            padding: 15px 25px;
            border-radius: 25px;
            font-size: 16px;
            font-weight: 600;
            margin-bottom: 25px;
            box-shadow: 0 4px 20px rgba(0, 0, 0, 0.08);
            border: 1px solid #e2e8f0;
        }}

        .dashboard-grid {{
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(320px, 1fr));
            gap: 20px;
            margin-bottom: 30px;
        }}

        .kpi-card {{
            background: white;
            border-radius: 15px;
            padding: 25px;
            box-shadow: 0 4px 20px rgba(0, 0, 0, 0.08);
            cursor: pointer;
            transition: all 0.3s ease;
            border: 1px solid #e2e8f0;
            position: relative;
            overflow: hidden;
        }}

        .kpi-card::before {{
            content: '';
            position: absolute;
            top: 0;
            left: 0;
            right: 0;
            height: 4px;
            background: #3b82f6;
            transform: scaleX(0);
            transition: transform 0.3s ease;
        }}

        .kpi-card:hover {{
            transform: translateY(-8px);
            box-shadow: 0 15px 40px rgba(0,0,0,0.15);
        }}

        .kpi-card:hover::before {{
            transform: scaleX(1);
        }}

        .kpi-header {{
            display: flex;
            justify-content: space-between;
            align-items: flex-start;
            margin-bottom: 20px;
        }}

        .kpi-title {{
            font-size: 16px;
            font-weight: 600;
            color: #2c3e50;
            line-height: 1.3;
            flex: 1;
            margin-right: 15px;
        }}

        .kpi-value {{
            font-size: 20px;
            font-weight: 700;
            color: white;
            padding: 8px 16px;
            border-radius: 20px;
            min-width: 80px;
            text-align: center;
            box-shadow: 0 4px 15px rgba(0,0,0,0.2);
        }}

        .kpi-context {{
            margin-top: 15px;
            color: #64748b;
            font-size: 14px;
        }}

        .performance-excellent {{ background: linear-gradient(135deg, #10b981, #059669); }}
        .performance-good {{ background: linear-gradient(135deg, #3b82f6, #1d4ed8); }}  
        .performance-warning {{ background: linear-gradient(135deg, #f59e0b, #d97706); }}
        .performance-poor {{ background: linear-gradient(135deg, #ef4444, #dc2626); }}
        .performance-no-data {{ background: linear-gradient(135deg, #6b7280, #4b5563); }}

        .no-data {{
            text-align: center;
            padding: 60px 20px;
            color: #64748b;
            font-size: 18px;
            background: rgba(255, 255, 255, 0.9);
            border-radius: 15px;
            backdrop-filter: blur(10px);
        }}

        @media (max-width: 768px) {{
            .dashboard-grid {{
                grid-template-columns: 1fr;
            }}
            
            .container {{
                padding: 10px;
            }}
            
            .header h1 {{
                font-size: 2rem;
            }}
        }}
    </style>
</head>
<body>
    <div class="container">
        <div class="header">
            <h1>🏛️ University Health & Safety KPI Dashboard</h1>
            <p>Comprehensive university-wide performance monitoring</p>
        </div>

        <div class="level-indicator" id="levelIndicator">University Level: All KPI Overview</div>

        <div class="sort-controls" id="sortControls">
            <button class="sort-button active" id="sortHighLow" onclick="sortKpis('highLow')">Highest First</button>
            <button class="sort-button" id="sortLowHigh" onclick="sortKpis('lowHigh')">Lowest First</button>
        </div>

        <div class="dashboard-grid" id="dashboardContainer">
            <!-- Dynamic content will be populated here -->
        </div>
    </div>

    <script>
        const dashboardData = {data_json};
        let currentSortOrder = 'highLow'; // Default sort order

        function getPerformanceClass(percentage) {{
            if (percentage === null || percentage === undefined) return 'performance-no-data';
            if (percentage >= 95) return 'performance-excellent';
            if (percentage >= 80) return 'performance-good';
            if (percentage >= 60) return 'performance-warning';
            return 'performance-poor';
        }}

        function formatValue(percentage, number) {{
            if (percentage === null || percentage === undefined) return '/';
            if (number === null || number === undefined) return `${{percentage.toFixed(1)}}%`;
            return `${{percentage.toFixed(1)}}%`;
        }}

        function createKpiContextDisplay(percentage, number) {{
            if (percentage === null || percentage === undefined) {{
                return '<div class="kpi-context">No data available</div>';
            }}
            
            if (number !== null && number !== undefined) {{
                return `<div class="kpi-context">Total items: ${{Math.round(number)}}</div>`;
            }} else {{
                return '<div class="kpi-context">Percentage metric</div>';
            }}
        }}

        function showUniversityView() {{
            const container = document.getElementById('dashboardContainer');
            container.innerHTML = '';

            if (!dashboardData.university || !dashboardData.university.kpis) {{
                container.innerHTML = '<div class="no-data">No university data available</div>';
                return;
            }}

            // Create KPI cards with sorting
            let kpiEntries = Object.entries(dashboardData.university.kpis);
            
            // Sort based on current sort order
            if (currentSortOrder === 'highLow') {{
                kpiEntries.sort((a, b) => {{
                    const aPercentage = a[1].percentage !== null ? a[1].percentage : -1;
                    const bPercentage = b[1].percentage !== null ? b[1].percentage : -1;
                    return bPercentage - aPercentage;
                }});
                document.getElementById('sortHighLow').classList.add('active');
                document.getElementById('sortLowHigh').classList.remove('active');
            }} else {{
                kpiEntries.sort((a, b) => {{
                    const aPercentage = a[1].percentage !== null ? a[1].percentage : 999;
                    const bPercentage = b[1].percentage !== null ? b[1].percentage : 999;
                    return aPercentage - bPercentage;
                }});
                document.getElementById('sortLowHigh').classList.add('active');
                document.getElementById('sortHighLow').classList.remove('active');
            }}

            // Create KPI cards
            for (const [kpiName, kpiData] of kpiEntries) {{
                const card = document.createElement('div');
                card.className = 'kpi-card';
                
                const performanceClass = getPerformanceClass(kpiData.percentage);
                
                card.innerHTML = `
                    <div class="kpi-header">
                        <div class="kpi-title">${{kpiName}}</div>
                        <div class="kpi-value ${{performanceClass}}">
                            ${{formatValue(kpiData.percentage, kpiData.number)}}
                        </div>
                    </div>
                    ${{createKpiContextDisplay(kpiData.percentage, kpiData.number)}}
                `;
                
                container.appendChild(card);
            }}
        }}

        function sortKpis(order) {{
            currentSortOrder = order;
            showUniversityView();
        }}

        // Initialize with university view
        showUniversityView();
    </script>
</body>
</html>"""

        with open(output_path, 'w', encoding='utf-8') as f:
            f.write(html_content)

    def run(self):
        """Main execution function"""
        print("University Health & Safety Dashboard Generator")
        print("=" * 50)

        # Step 1: Select Excel file
        print("1. Select Excel file containing health & safety data...")
        excel_file = self.select_excel_file()
        
        if not excel_file:
            print("No file selected. Exiting.")
            return

        if not os.path.exists(excel_file):
            print(f"Error: File '{excel_file}' does not exist.")
            return

        # Step 2: Load data
        print("2. Loading Excel data...")
        if not self.load_excel_data(excel_file):
            return

        print(f"   ✓ University data: {len(self.university_data) if self.university_data is not None else 0} rows")
        print(f"   ✓ Faculty data: {len(self.faculty_data) if self.faculty_data is not None else 0} rows") 
        print(f"   ✓ School data: {len(self.school_data) if self.school_data is not None else 0} rows")

        # Step 3: Select output location
        print("3. Select output location for University HTML dashboard...")
        output_file = self.select_output_location()
        
        if not output_file:
            print("No output location selected. Exiting.")
            return

        print("4. Processing KPI data...")
        dashboard_data = self.process_kpi_data()

        print("5. Creating interactive HTML dashboard...")
        self.create_html_dashboard(dashboard_data, output_file)

        print(f"\n✅ University Dashboard generated successfully!")
        print(f"📊 Saved to: {output_file}")
        print(f"\n📋 Dashboard features:")
        print(f"   • 13 interactive KPI cards at university level")
        print(f"   • Visual performance indicators (color-coded)")
        print(f"   • Sortable KPI display (highest/lowest first)")
        print(f"   • Handles missing data with '/' indicators")

        messagebox.showinfo("Success", f"University Dashboard generated successfully!\n\nSaved to: {output_file}\n\nOpen the HTML file in your web browser to view the interactive dashboard.")


if __name__ == "__main__":
    generator = UniversityDashboardGenerator()
    generator.run()