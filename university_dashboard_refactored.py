import json
import os
import tkinter as tk
from tkinter import filedialog, messagebox

from faculty_dashboard_refactored import FacultyDashboardGenerator


class UniversityDashboardGenerator(FacultyDashboardGenerator):
    """Generate the university dashboard using the same KPI model as faculty pages."""

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
        output_dir = filedialog.askdirectory(title="Select Directory for University Dashboard...")
        root.destroy()
        return excel_file, output_dir

    def _extract_university_dashboard_data(self):
        dashboard_data = self.extract_and_process_data()
        university_data = dashboard_data.get("university", {})
        if not university_data:
            raise ValueError("No university data available")

        university_history = self._build_kpi_history(self.university_history_data, 'Faculty', 'University')
        for kpi_name, kpi_data in university_data.get("kpis", {}).items():
            kpi_data["history"] = university_history.get(kpi_name, [])
            kpi_data["is_university"] = True

        university_data["schools"] = dashboard_data.get("faculties", {})
        return university_data

    def create_university_html_dashboard(self, university_data, output_path):
        template_path = os.path.join(os.path.dirname(__file__), 'university_dashboard_template.html')
        if not os.path.exists(template_path):
            raise FileNotFoundError(f"University dashboard template not found: {template_path}")

        university_json = json.dumps(university_data, indent=2, default=str)
        tooltip_json = json.dumps(self.tooltip_data or {}, indent=2, default=str)
        kpi_meta_json = json.dumps(self.kpi_metadata, default=str)

        with open(template_path, 'r', encoding='utf-8') as f:
            html_content = f.read()

        replacements = {
            '__FACULTY_NAME__': 'University',
            '__FACULTY_JSON__': university_json,
            '__TOOLTIP_JSON__': tooltip_json,
            '__KPI_META_JSON__': kpi_meta_json,
        }
        for token, value in replacements.items():
            html_content = html_content.replace(token, value)

        with open(output_path, 'w', encoding='utf-8') as f:
            f.write(html_content)

    def run(self):
        print("University Health & Safety Dashboard Generator")
        print("=" * 50)

        excel_file, output_dir = self.select_file_and_output()
        if not excel_file or not output_dir:
            print("No file or directory selected. Exiting.")
            return

        if not os.path.exists(excel_file):
            print(f"Error: File '{excel_file}' does not exist.")
            return

        print("1. Loading Excel data...")
        if not self.load_excel_data(excel_file):
            return

        print("2. Processing KPI data...")
        university_data = self._extract_university_dashboard_data()

        output_path = os.path.join(output_dir, "index.html")
        print("3. Creating university HTML dashboard...")

        try:
            self.create_university_html_dashboard(university_data, output_path)
            print("\nUniversity Dashboard generated successfully!")
            print(f"Saved to: {output_path}")
            messagebox.showinfo(
                "Success",
                f"University Dashboard generated successfully!\n\nSaved to: {output_path}\n\nOpen index.html in your web browser to view the dashboard."
            )
        except Exception as e:
            print(f"\nFailed to generate report: {e}")
            messagebox.showerror("Error", f"Failed to generate report: {e}")


if __name__ == "__main__":
    generator = UniversityDashboardGenerator()
    generator.run()
