import customtkinter as ctk
from tkinter import messagebox, filedialog
import tkinter.ttk as ttk
import pandas as pd
import os
import json
from datetime import datetime, timedelta
import threading
import time
import webbrowser
import urllib.parse
from dateutil import parser
import tkinter as tk
import html
import re

ctk.set_appearance_mode("System")  # "Dark", "Light", or "System"
ctk.set_default_color_theme("blue")

class AuditManagerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Audit Issue Management System")
        self.root.geometry("1200x800")
        self.root.minsize(900, 700)
        self.selected_issue = None
        self.team_filter = None
        self.type_filter = None
        
        # Configuration
        self.excel_file = 'test_issues.xlsx'
        self.config_file = 'config.json'
        self.email_template_file = 'email_template.html'
        self.dashboard_columns_file = 'dashboard_columns.json'
        self.dashboard_display_columns = []  # Will be set after loading data
        
        # Email templates
        self.email_templates_file = 'email_templates.json'
        self.email_templates = self.load_email_templates()
        # Reminder intervals (make available everywhere)
        self.reminder_intervals = [
            ("7 days before", 7),
            ("30 days before", 30),
            ("3 months before", 90),
            ("6 months before", 180)
        ]
        # Initialize files
        self.init_files()
        
        # Load data
        self.load_data()
        self.load_dashboard_columns_config()
        
        # Create UI
        self.create_ui()
        
        # Start reminder scheduler
        self.start_reminder_scheduler()

    def init_files(self):
        """Initialize configuration files"""
        # Create config file if it doesn't exist
        if not os.path.exists(self.config_file):
            config = {
                "reminder_intervals": [
                    {"days_before": 30, "enabled": True},
                    {"days_before": 14, "enabled": True},
                    {"days_before": 7, "enabled": True},
                    {"days_before": 3, "enabled": True},
                    {"days_before": 1, "enabled": True},
                    {"days_before": 0, "enabled": True}
                ]
            }
            with open(self.config_file, 'w') as f:
                json.dump(config, f, indent=2)
        
        # Create email template if it doesn't exist
        if not os.path.exists(self.email_template_file):
            template = """<!DOCTYPE html>
<html>
<head>
    <style>
        body { font-family: Arial, sans-serif; margin: 20px; }
        .header { background-color: #f0f0f0; padding: 15px; border-radius: 5px; }
        .content { margin: 20px 0; }
        .footer { color: #666; font-size: 12px; }
        .urgent { color: #d32f2f; font-weight: bold; }
    </style>
</head>
<body>
    <div class=\"header\">
        <h2>Audit Issue Resolution Required</h2>
    </div>
    <div class=\"content\">
        <p><strong>Issue ID:</strong> {{ISSUE_ID}}</p>
        <p><strong>Description:</strong> {{DESCRIPTION}}</p>
        <p><strong>Priority:</strong> {{PRIORITY}}</p>
        <p><strong>Status:</strong> {{STATUS}}</p>
        <p><strong>Resolution Due Date:</strong> <span class=\"urgent\">{{RESOLUTION_DATE}}</span></p>
        <p><strong>Days Remaining:</strong> <span class=\"urgent\">{{DAYS_REMAINING}}</span></p>
        <p><strong>Team:</strong> {{TEAM}}</p>
        <p><strong>Created Date:</strong> {{CREATED_DATE}}</p>
        <p><strong>Reminder Count:</strong> {{REMINDER_COUNT}}</p>
    </div>
    <div class=\"content\">
        <p>Please review and resolve this audit issue by the specified resolution date. If you have any questions, please contact the audit team.</p>
        <p>This is reminder #{{REMINDER_COUNT}} of this issue.</p>
    </div>
    <div class=\"footer\">
        <p>This is an automated reminder from the Audit Management System.</p>
        <p>Generated on: {{CURRENT_DATE}}</p>
    </div>
</body>
</html>"""
            with open(self.email_template_file, 'w') as f:
                f.write(template)

    def load_data(self):
        """Load data from Excel file"""
        try:
            if not os.path.exists(self.excel_file):
                messagebox.showerror("Error", f"Excel file '{self.excel_file}' not found. Please ensure the file exists in the same directory as this application.")
                self.df = pd.DataFrame()
                return
            self.df = pd.read_excel(self.excel_file)
            # Check if required columns exist, add them if missing
            required_columns = {
                'Issue ID': [],
                'Issue Description': [],
                'Team Name': [],
                'Management Resolution Date': [],
                'Reporting Month': [],
                'Responsible Business Segment': [],
                'Issue Summary': [],
                'Impact and Likelihood': [],
                'Recommendation': [],
                'Previous Month Rationale': [],
                'Line of Business': [],
                'Reporting Categories': []
            }
            for col in required_columns:
                if col not in self.df.columns:
                    print(f"File is missing column: {col}")
            if self.df.empty:
                print("Excel file is empty. Ready to add new issues.")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load Excel file: {str(e)}")
            self.df = pd.DataFrame()

    # Use dateutil.parser for robust date parsing
    def parse_any_date(self, date_str):
        try:
            return parser.parse(str(date_str)).date()
        except Exception:
            return None

    def save_data(self):
        try:
            self.df.to_excel(self.excel_file, index=False)
        except Exception as e:
            messagebox.showerror("Error", f"Failed to save Excel file: {str(e)}")

    def create_ui(self):
        self.tabview = ctk.CTkTabview(self.root)
        self.tabview.pack(fill="both", expand=True, padx=10, pady=10)
        self.dashboard_tab = self.tabview.add("Dashboard")
        self.issues_tab = self.tabview.add("Manage Issues")
        self.email_tab = self.tabview.add("Email Management")
        self.settings_tab = self.tabview.add("Settings")
        self.create_dashboard_tab()
        self.create_issues_tab()
        self.create_email_tab()
        self.create_settings_tab()

    def create_dashboard_tab(self):
        frame = self.dashboard_tab
        for widget in frame.winfo_children():
            widget.destroy()
        # Use a fixed, hardcoded set of columns for the dashboard
        self.dashboard_display_columns = [
            'Issue ID', 'Issue Description', 'Team Name', 'Management Resolution Date',
            'Reporting Month', 'Responsible Business Segment', 'Issue Summary',
            'Impact and Likelihood', 'Recommendation', 'Previous Month Rationale',
            'Line of Business', 'Reporting Categories'
        ]
        # --- Audit Issue Dashboard title and stats (title stays at top) ---
        ctk.CTkLabel(frame, text="Audit Issue Dashboard", font=("Arial", 20, "bold")).pack(pady=(2, 0))
        stats_frame = ctk.CTkFrame(frame, fg_color="transparent")
        stats_frame.pack(fill='x', padx=20, pady=(0, 2))
        total_issues = len(self.df)
        stats = [
            ("Total", total_issues, "#3b82f6"),
        ]
        for i, (label, value, color) in enumerate(stats):
            stat = ctk.CTkFrame(stats_frame, fg_color="transparent")
            stat.grid(row=0, column=i, padx=4, pady=0, sticky='w')
            ctk.CTkLabel(stat, text=label, font=("Arial", 10)).pack(side='left', padx=(0, 2))
            ctk.CTkLabel(stat, text=str(value), font=("Arial", 12, "bold"), text_color=color).pack(side='left')
        stats_frame.grid_columnconfigure((0,), weight=0)
        # --- ISSUES TABLE (reduced height) ---
        table_frame = ctk.CTkFrame(frame)
        table_frame.pack(fill='both', expand=False, padx=20, pady=(5, 10))
        ctk.CTkLabel(table_frame, text="Audit Issues", font=("Arial", 14, "bold")).pack(anchor='w', pady=(0, 2))
        dashboard_cols = [col for col in self.dashboard_display_columns if col in self.df.columns]
        self.dashboard_tree = ttk.Treeview(table_frame, columns=dashboard_cols, show='headings', height=10)  # Reduced height
        for col in dashboard_cols:
            self.dashboard_tree.heading(col, text=col, anchor='w')
        self.dashboard_tree.pack(fill='x', expand=False, pady=5)
        self.dashboard_tree.bind('<<TreeviewSelect>>', self.on_dashboard_select)
        self.dashboard_tree_cols = dashboard_cols
        self.sort_state = {col: None for col in dashboard_cols}  # None, 'asc', 'desc'
        # Populate the table with Excel data
        for _, row in self.df.iterrows():
            values = [row.get(col, '') for col in self.dashboard_tree_cols]
            self.dashboard_tree.insert('', 'end', values=values)
        # Set selected_issue to the first row if available, and render preview
        if len(self.df) > 0:
            self.selected_issue = self.df.iloc[0].to_dict()
            self.render_email_preview(self.selected_issue)
        else:
            self.selected_issue = None
            # Render the HTML template with dummy data (not issue data)
            self.render_email_preview({col: f"[{col}]" for col in self.dashboard_tree_cols})

        # --- EMAIL PREVIEW PANEL ---
        preview_frame = ctk.CTkFrame(frame)
        preview_frame.pack(fill='x', padx=20, pady=(0, 10))
        ctk.CTkLabel(preview_frame, text="Email Preview", font=("Arial", 14, "bold")).pack(anchor='w')
        # Template and reminder interval in same row
        top_row = ctk.CTkFrame(preview_frame, fg_color="transparent")
        top_row.pack(anchor='w', pady=(0, 5), fill='x')
        ctk.CTkLabel(top_row, text="Select Template:").pack(side='left', padx=(0, 5))
        self.dashboard_template_names = list(self.email_templates.keys())
        self.dashboard_template_select_var = ctk.StringVar(value=self.dashboard_template_names[0] if self.dashboard_template_names else "")
        self.dashboard_template_select_box = ctk.CTkComboBox(top_row, values=self.dashboard_template_names, variable=self.dashboard_template_select_var, width=160, command=self.on_dashboard_template_select)
        self.dashboard_template_select_box.pack(side='left', padx=(0, 10))
        ctk.CTkLabel(top_row, text="Reminder Interval:").pack(side='left', padx=(0, 5))
        self.reminder_interval_var = ctk.StringVar(value="7 days before")
        interval_options = [label for label, _ in self.reminder_intervals]
        self.reminder_interval_box = ctk.CTkComboBox(top_row, values=interval_options, variable=self.reminder_interval_var, width=140, command=self.on_reminder_interval_change)
        self.reminder_interval_box.pack(side='left', padx=(0, 10))
        # Only one editable email preview box in the dashboard
        self.email_preview_box = tk.Text(preview_frame, height=20, font=("Consolas", 12))
        self.email_preview_box.pack(fill='x', expand=True, pady=5)
        # Remove any creation or packing of self.email_template_box in the dashboard tab

    def create_issues_tab(self):
        frame = self.issues_tab
        for widget in frame.winfo_children():
            widget.destroy()
        header_frame = ctk.CTkFrame(frame)
        header_frame.pack(fill='x', padx=20, pady=20)
        ctk.CTkLabel(header_frame, text="Audit Issues Management", font=("Arial", 18, "bold")).pack(side='left')
        button_frame = ctk.CTkFrame(header_frame, fg_color="transparent")
        button_frame.pack(side='right')
        ctk.CTkButton(button_frame, text="Add New Issue", command=self.show_add_issue_dialog).pack(side='left', padx=5)
        ctk.CTkButton(button_frame, text="Edit Selected", command=self.edit_selected_issue).pack(side='left', padx=5)
        ctk.CTkButton(button_frame, text="Delete Selected", command=self.delete_selected_issue).pack(side='left', padx=5)
        ctk.CTkButton(button_frame, text="Send Reminder", command=self.send_reminder_to_selected).pack(side='left', padx=5)
        table_frame = ctk.CTkFrame(frame)
        table_frame.pack(fill='both', expand=True, padx=20, pady=10)
        ctk.CTkLabel(table_frame, text="All Issues (All Columns)", font=("Arial", 14, "bold")).pack(anchor='w')
        all_columns = list(self.df.columns)
        self.issues_tree = ttk.Treeview(table_frame, columns=all_columns, show='headings', height=16)
        for col in all_columns:
            self.issues_tree.heading(col, text=col)
            self.issues_tree.column(col, width=150)
        self.issues_tree.pack(fill='both', expand=True, pady=10)
        self.update_issues_table()

    def create_email_tab(self):
        frame = self.email_tab
        for widget in frame.winfo_children():
            widget.destroy()
        ctk.CTkLabel(frame, text="Email Template Management", font=("Arial", 18, "bold")).pack(pady=20)
        # Template selection and management
        top_frame = ctk.CTkFrame(frame)
        top_frame.pack(fill='x', padx=20, pady=5)
        ctk.CTkLabel(top_frame, text="Select Template:").pack(side='left', padx=(0, 5))
        self.template_names = list(self.email_templates.keys())
        self.template_select_var = ctk.StringVar(value=self.template_names[0] if self.template_names else "")
        self.template_select_box = ctk.CTkComboBox(top_frame, values=self.template_names, variable=self.template_select_var, width=200, command=self.on_template_select)
        self.template_select_box.pack(side='left', padx=(0, 10))
        ctk.CTkButton(top_frame, text="New Template", command=self.new_email_template).pack(side='left', padx=5)
        ctk.CTkButton(top_frame, text="Delete Template", command=self.delete_email_template).pack(side='left', padx=5)
        # Association with column value
        assoc_frame = ctk.CTkFrame(frame, fg_color="transparent")
        assoc_frame.pack(fill='x', padx=20, pady=2)
        ctk.CTkLabel(assoc_frame, text="Associate with column value (optional):").pack(side='left', padx=(0, 5))
        self.assoc_column_var = ctk.StringVar(value="")
        assoc_columns = [col for col in self.df.columns if col not in ("", None)]
        self.assoc_column_box = ctk.CTkComboBox(assoc_frame, values=["None"] + assoc_columns, variable=self.assoc_column_var, width=150)
        self.assoc_column_box.pack(side='left', padx=(0, 10))
        self.assoc_value_var = ctk.StringVar(value="")
        self.assoc_value_box = ctk.CTkEntry(assoc_frame, textvariable=self.assoc_value_var, width=150)
        self.assoc_value_box.pack(side='left', padx=(0, 10))
        # Only one template editing box and Save button
        preview_frame = ctk.CTkFrame(frame)
        preview_frame.pack(fill='x', padx=20, pady=10)
        ctk.CTkLabel(preview_frame, text="Edit & Format Template", font=("Arial", 14, "bold")).pack(anchor='w')
        self.email_template_box = tk.Text(preview_frame, height=14, font=("Arial", 12), wrap='word')
        self.email_template_box.pack(fill='x', expand=True, pady=5)
        # Save Template button directly below the editing box
        ctk.CTkButton(preview_frame, text="Save Template", command=self.save_email_template).pack(anchor='w', pady=(5, 0))
        # Remove any unused toolbar or extra frames above the edit area

    def load_email_templates(self):
        import os, json
        default_template_content = """<!DOCTYPE html>\n<html>\n<head>\n    <style>\n        body { font-family: Arial, sans-serif; margin: 20px; }\n        .header { background-color: #f0f0f0; padding: 15px; border-radius: 5px; }\n        .content { margin: 20px 0; }\n        .footer { color: #666; font-size: 12px; }\n        .urgent { color: #d32f2f; font-weight: bold; }\n    </style>\n</head>\n<body>\n    <div class=\"header\">\n        <h2>Audit Issue Resolution Required</h2>\n    </div>\n    <div class=\"content\">\n        <p><strong>Issue ID:</strong> {{ISSUE_ID}}</p>\n        <p><strong>Description:</strong> {{DESCRIPTION}}</p>\n        <p><strong>Resolution Due Date:</strong> <span class=\"urgent\">{{RESOLUTION_DATE}}</span></p>\n        <p><strong>Days Remaining:</strong> <span class=\"urgent\">{{DAYS_REMAINING}}</span></p>\n        <p><strong>Team:</strong> {{TEAM}}</p>\n    </div>\n    <div class=\"content\">\n        <p>Please review and resolve this audit issue by the specified resolution date. If you have any questions, please contact the audit team.</p>\n    </div>\n    <div class=\"footer\">\n        <p>This is an automated reminder from the Audit Management System.</p>\n        <p>Generated on: {{CURRENT_DATE}}</p>\n    </div>\n</body>\n</html>"""
        if not os.path.exists(self.email_templates_file):
            templates = {"Default": {"content": default_template_content, "assoc_column": "", "assoc_value": ""}}
            with open(self.email_templates_file, 'w', encoding='utf-8') as f:
                json.dump(templates, f, indent=2)
            return templates
        with open(self.email_templates_file, 'r', encoding='utf-8') as f:
            templates = json.load(f)
        # Ensure 'Default' template always exists
        if "Default" not in templates:
            templates["Default"] = {"content": default_template_content, "assoc_column": "", "assoc_value": ""}
            with open(self.email_templates_file, 'w', encoding='utf-8') as f:
                json.dump(templates, f, indent=2)
        return templates

    def save_email_templates(self):
        print(f"[DEBUG] Saving to: {self.email_templates_file}")
        with open(self.email_templates_file, 'w', encoding='utf-8') as f:
            json.dump(self.email_templates, f, indent=2)

    def on_template_select(self, *args):
        self.selected_template_name = self.template_select_var.get()
        self.load_selected_template()
        self.render_live_template()

    def load_selected_template(self):
        name = self.template_select_var.get()
        if name and name in self.email_templates:
            template = self.email_templates[name]
            self.email_template_box.delete('1.0', 'end')
            self.email_template_box.insert('1.0', template.get('content', ''))
            self.assoc_column_var.set(template.get('assoc_column', 'None'))
            self.assoc_value_var.set(template.get('assoc_value', ''))
        else:
            self.email_template_box.delete('1.0', 'end')
            self.assoc_column_var.set('None')
            self.assoc_value_var.set('')
        self.render_live_template()

    def save_email_template(self):
        name = self.template_select_var.get()
        if not name:
            messagebox.showerror("Error", "Please enter a template name.")
            return
        # Save plain text and tag ranges
        content = self.email_template_box.get('1.0', 'end').strip()
        tags = self.get_text_tags(self.email_template_box)
        assoc_column = self.assoc_column_var.get() if self.assoc_column_var.get() != 'None' else ''
        assoc_value = self.assoc_value_var.get().strip()
        self.email_templates[name] = {
            'content': content,
            'assoc_column': assoc_column,
            'assoc_value': assoc_value
        }
        self.save_email_templates()
        messagebox.showinfo("Saved", f"Template '{name}' saved.")
        self.template_names = list(self.email_templates.keys())
        self.template_select_box.configure(values=self.template_names)
        if hasattr(self, 'dashboard_template_select_box'):
            self.dashboard_template_names = list(self.email_templates.keys())
            self.dashboard_template_select_box.configure(values=self.dashboard_template_names)
            if self.dashboard_template_select_var.get() == name:
                self.render_email_preview(self.selected_issue if hasattr(self, 'selected_issue') and self.selected_issue else {})

    def new_email_template(self):
        import tkinter.simpledialog
        name = tkinter.simpledialog.askstring("New Template", "Enter template name:")
        if not name:
            return
        if name in self.email_templates:
            messagebox.showerror("Error", "A template with this name already exists.")
            return
        self.email_templates[name] = {'content': '', 'tags': {}, 'assoc_column': '', 'assoc_value': ''}
        self.save_email_templates()
        self.template_names = list(self.email_templates.keys())
        self.template_select_box.configure(values=self.template_names)
        self.template_select_var.set(name)
        self.load_selected_template()
        self.render_live_template()

    def delete_email_template(self):
        name = self.template_select_var.get()
        if not name or name not in self.email_templates:
            return
        if messagebox.askyesno("Delete Template", f"Delete template '{name}'?"):
            del self.email_templates[name]
            self.save_email_templates()
            self.template_names = list(self.email_templates.keys())
            self.template_select_box.configure(values=self.template_names)
            if self.template_names:
                self.template_select_var.set(self.template_names[0])
                self.load_selected_template()
                self.render_live_template()
            else:
                self.template_select_var.set("")
                self.email_template_box.delete('1.0', 'end')
                # If all templates are deleted, add a minimal placeholder
                placeholder = "<html><body><p>[No template available]</p></body></html>"
                self.email_templates["Default"] = {"content": placeholder, "assoc_column": "", "assoc_value": ""}
                self.save_email_templates()
                self.template_names = list(self.email_templates.keys())
                self.template_select_box.configure(values=self.template_names)
                self.template_select_var.set("Default")
                self.load_selected_template()
                self.render_live_template()

    def render_live_template(self):
        # Render the current template as plain text with dummy/example data, but keep the box editable
        template = self.email_template_box.get('1.0', 'end')
        example_issue = {col: f"Example {col}" for col in self.df.columns}
        preview = template
        for col in self.df.columns:
            preview = preview.replace(f'{{{{{col}}}}}', str(example_issue.get(col, f'[{col}]')))
        preview = preview.replace('{{CURRENT_DATE}}', datetime.now().strftime('%Y-%m-%d'))
        import re
        text_preview = re.sub('<[^<]+?>', '', preview)
        # Replace the box content only if it differs from the rendered preview
        current_content = self.email_template_box.get('1.0', 'end')
        if current_content.strip() != text_preview.strip():
            self.email_template_box.delete('1.0', 'end')
            self.email_template_box.insert('1.0', text_preview)

    def create_settings_tab(self):
        frame = self.settings_tab
        for widget in frame.winfo_children():
            widget.destroy()
        ctk.CTkLabel(frame, text="System Settings", font=("Arial", 18, "bold")).pack(pady=20)
        # (Removed Configure Dashboard Columns section from settings tab)
        email_settings_frame = ctk.CTkFrame(frame)
        email_settings_frame.pack(fill='x', padx=20, pady=10)
        info_text = """Corporate Email Setup:\n\nThis application will use your corporate email settings automatically.\nNo manual SMTP configuration is required.\n\nTo send emails:\n1. Ensure you're logged into your corporate email on this machine\n2. The system will use your default email application\n3. Emails will be sent through your corporate email system\n\nNote: If you encounter permission issues, contact your IT department."""
        ctk.CTkLabel(email_settings_frame, text=info_text, justify='left', font=("Arial", 12)).pack(anchor='w', pady=10)
        ctk.CTkButton(email_settings_frame, text="Test Email Configuration", command=self.test_email_config).pack(pady=10)
        # Reminder interval options removed from settings tab

    def show_add_issue_dialog(self):
        """Show dialog to add a new issue"""
        dialog = ctk.CTkToplevel(self.root)
        dialog.title("Add New Audit Issue")
        dialog.geometry("500x900")
        dialog.transient(self.root)
        dialog.grab_set()
        
        # Form fields
        ctk.CTkLabel(dialog, text="Add New Audit Issue", font=("Arial", 16, "bold")).pack(pady=20)
        
        # Reporting Month
        ctk.CTkLabel(dialog, text="Reporting Month (MM-YYYY):").pack(anchor='w', padx=20)
        reporting_month_entry = ctk.CTkEntry(dialog, width=60)
        reporting_month_entry.pack(fill='x', padx=20, pady=5)
        # Responsible Business Segment
        ctk.CTkLabel(dialog, text="Responsible Business Segment:").pack(anchor='w', padx=20)
        business_segment_entry = ctk.CTkEntry(dialog, width=60)
        business_segment_entry.pack(fill='x', padx=20, pady=5)
        # Issue Summary
        ctk.CTkLabel(dialog, text="Issue Summary:").pack(anchor='w', padx=20)
        issue_summary_entry = ctk.CTkEntry(dialog, width=60)
        issue_summary_entry.pack(fill='x', padx=20, pady=5)
        # Issue Description
        ctk.CTkLabel(dialog, text="Issue Description:").pack(anchor='w', padx=20)
        description_entry = ctk.CTkEntry(dialog, width=60)
        description_entry.pack(fill='x', padx=20, pady=5)
        # Team Name
        ctk.CTkLabel(dialog, text="Team Name:").pack(anchor='w', padx=20)
        team_entry = ctk.CTkEntry(dialog, width=60)
        team_entry.pack(fill='x', padx=20, pady=5)
        # Management Resolution Date
        ctk.CTkLabel(dialog, text="Management Resolution Date (YYYY-MM-DD):").pack(anchor='w', padx=20)
        date_entry = ctk.CTkEntry(dialog, width=60)
        date_entry.pack(fill='x', padx=20, pady=5)
        # Impact and Likelihood
        ctk.CTkLabel(dialog, text="Impact and Likelihood:").pack(anchor='w', padx=20)
        impact_entry = ctk.CTkEntry(dialog, width=60)
        impact_entry.pack(fill='x', padx=20, pady=5)
        # Recommendation
        ctk.CTkLabel(dialog, text="Recommendation:").pack(anchor='w', padx=20)
        recommendation_entry = ctk.CTkEntry(dialog, width=60)
        recommendation_entry.pack(fill='x', padx=20, pady=5)
        # Previous Month Rationale
        ctk.CTkLabel(dialog, text="Previous Month Rationale:").pack(anchor='w', padx=20)
        rationale_entry = ctk.CTkEntry(dialog, width=60)
        rationale_entry.pack(fill='x', padx=20, pady=5)
        # Line of Business
        ctk.CTkLabel(dialog, text="Line of Business:").pack(anchor='w', padx=20)
        lob_entry = ctk.CTkEntry(dialog, width=60)
        lob_entry.pack(fill='x', padx=20, pady=5)
        # Reporting Categories
        ctk.CTkLabel(dialog, text="Reporting Categories:").pack(anchor='w', padx=20)
        reporting_cat_entry = ctk.CTkEntry(dialog, width=60)
        reporting_cat_entry.pack(fill='x', padx=20, pady=5)
        
        def save_issue():
            # Validate fields
            if not all([
                reporting_month_entry.get(), business_segment_entry.get(), issue_summary_entry.get(),
                description_entry.get(), team_entry.get(), date_entry.get(), impact_entry.get(),
                recommendation_entry.get(), rationale_entry.get(), lob_entry.get(), reporting_cat_entry.get()
            ]):
                messagebox.showerror("Error", "All fields are required")
                return
            
            # Validate date
            try:
                datetime.strptime(date_entry.get(), '%Y-%m-%d')
            except ValueError:
                messagebox.showerror("Error", "Invalid date format. Use YYYY-MM-DD")
                return
            
            # Generate ID
            new_id = f"AUDIT-{len(self.df) + 1:04d}" if not self.df.empty else "AUDIT-0001"
            
            # Create new issue
            new_issue = {
                'Issue ID': new_id,
                'Reporting Month': reporting_month_entry.get(),
                'Responsible Business Segment': business_segment_entry.get(),
                'Issue Summary': issue_summary_entry.get(),
                'Issue Description': description_entry.get(),
                'Team Name': team_entry.get(),
                'Management Resolution Date': date_entry.get(),
                'Impact and Likelihood': impact_entry.get(),
                'Recommendation': recommendation_entry.get(),
                'Previous Month Rationale': rationale_entry.get(),
                'Line of Business': lob_entry.get(),
                'Reporting Categories': reporting_cat_entry.get()
            }
            
            # Add to dataframe
            self.df = pd.concat([self.df, pd.DataFrame([new_issue])], ignore_index=True)
            self.save_data()
            
            # Update UI
            self.update_issues_table()
            self.update_recent_issues()
            self.update_issue_combo()
            
            dialog.destroy()
            messagebox.showinfo("Success", "Issue added successfully!")
        
        # Buttons
        button_frame = ctk.CTkFrame(dialog)
        button_frame.pack(fill='x', padx=20, pady=20)
        
        ctk.CTkButton(button_frame, text="Save", command=save_issue).pack(side='right', padx=5)
        ctk.CTkButton(button_frame, text="Cancel", command=dialog.destroy).pack(side='right', padx=5)
    
    def update_issues_table(self):
        """Update the issues table with current data"""
        for item in self.issues_tree.get_children():
            self.issues_tree.delete(item)
        for _, row in self.df.iterrows():
            values = [row[col] if col in row else '' for col in self.issues_tree['columns']]
            self.issues_tree.insert('', 'end', values=values)
    
    def update_recent_issues(self):
        """Update the recent issues table"""
        for item in self.recent_tree.get_children():
            self.recent_tree.delete(item)
        recent_data = self.df.tail(8) if not self.df.empty else pd.DataFrame()
        for _, row in recent_data.iterrows():
            values = [row[col] if col in row else '' for col in self.recent_tree['columns']]
            self.recent_tree.insert('', 'end', values=values)
    
    def update_issue_combo(self):
        """Update the issue combo box"""
        if not self.df.empty:
            issue_list = [f"{row['Issue ID']} - {row['Issue Description'][:50]}..." for _, row in self.df.iterrows()]
            self.issue_combo['values'] = issue_list
            if issue_list:
                self.issue_combo.set(issue_list[0])
    
    def load_email_template(self):
        """Load the email template from file"""
        try:
            with open(self.email_template_file, 'r') as f:
                template = f.read()
                self.email_template_box.delete(1.0, ctk.END)
                self.email_template_box.insert(1.0, template)
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load email template: {str(e)}")
    
    def reset_email_template(self):
        """Reset email template to default"""
        if messagebox.askyesno("Confirm", "Are you sure you want to reset the email template to default?"):
            self.init_files()
            self.load_email_template()
            messagebox.showinfo("Success", "Email template reset to default!")
    
    def load_settings(self):
        """Load settings from config file"""
        try:
            with open(self.config_file, 'r') as f:
                config = json.load(f)
            print("Settings loaded successfully")
        except Exception as e:
            print(f"Failed to load settings: {e}")
    
    def save_settings(self):
        """Save settings to config file"""
        try:
            with open(self.config_file, 'r') as f:
                config = json.load(f)
            
            with open(self.config_file, 'w') as f:
                json.dump(config, f, indent=2)
            
            messagebox.showinfo("Success", "Settings saved successfully!")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to save settings: {str(e)}")
    
    def send_reminder_email(self):
        """Send reminder email for selected issue"""
        if not self.issue_var.get():
            messagebox.showerror("Error", "Please select an issue")
            return
        
        # Get selected issue
        issue_id = self.issue_var.get().split(' - ')[0]
        issue = self.df[self.df['Issue ID'] == issue_id].iloc[0]
        
        # Send email
        success = self.send_email(issue)
        
        if success:
            # Update reminder count
            self.df.loc[self.df['Issue ID'] == issue_id, 'Last_Reminder'] = datetime.now().strftime('%Y-%m-%d')
            self.df.loc[self.df['Issue ID'] == issue_id, 'Reminder_Count'] = issue['Reminder_Count'] + 1
            self.save_data()
            self.update_issues_table()
            messagebox.showinfo("Success", "Reminder email sent successfully!")
        else:
            messagebox.showerror("Error", "Failed to send reminder email")
    
    def send_email(self, issue):
        """Send email for a specific issue using default email application"""
        try:
            # Load template
            with open(self.email_template_file, 'r') as f:
                template = f.read()
            
            # Calculate days remaining
            mgmt_date = self.parse_any_date(issue['Management Resolution Date'])
            if mgmt_date:
                days_remaining = (mgmt_date - datetime.now().date()).days
            else:
                days_remaining = ''
            
            # Replace template variables
            template = template.replace('{{ISSUE_ID}}', issue['Issue ID'])
            template = template.replace('{{DESCRIPTION}}', issue['Issue Description'])
            template = template.replace('{{RESOLUTION_DATE}}', issue['Management Resolution Date'])
            template = template.replace('{{DAYS_REMAINING}}', str(days_remaining))
            template = template.replace('{{TEAM}}', issue['Team Name'])
            template = template.replace('{{CURRENT_DATE}}', datetime.now().strftime('%Y-%m-%d'))
            
            # Create email content
            subject = f"Audit Issue Reminder: {issue['Issue ID']} - {issue['Issue Description'][:50]}"
            body = f"""
Audit Issue Resolution Required

Issue ID: {issue['Issue ID']}
Description: {issue['Issue Description']}
Resolution Due Date: {issue['Management Resolution Date']}
Days Remaining: {days_remaining}
Team: {issue['Team Name']}

Please review and resolve this audit issue by the specified resolution date. 
If you have any questions, please contact the audit team.

This is reminder #{issue['Reminder_Count'] + 1} of this issue.

This is an automated reminder from the Audit Management System.
Generated on: {datetime.now().strftime('%Y-%m-%d')}
            """
            
            # Use default email application
            mailto_link = f"mailto:{issue['Team_Email']}?subject={urllib.parse.quote(subject)}&body={urllib.parse.quote(body)}"
            
            # Open default email application
            webbrowser.open(mailto_link)
            
            print(f"Email application opened for issue {issue['Issue ID']}")
            print(f"To: {issue['Team_Email']}")
            print(f"Subject: {subject}")
            
            return True
        except Exception as e:
            print(f"Error opening email application: {e}")
            return False
    
    def test_email_config(self):
        """Test email configuration by opening default email application"""
        try:
            test_subject = "Test Email - Audit Management System"
            test_body = "This is a test email to verify your email configuration is working properly."
            
            mailto_link = f"mailto:test@example.com?subject={urllib.parse.quote(test_subject)}&body={urllib.parse.quote(test_body)}"
            webbrowser.open(mailto_link)
            
            messagebox.showinfo("Test Email", "Your default email application should have opened with a test email. If it didn't open, please check your email configuration.")
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to test email configuration: {str(e)}")
    
    def start_reminder_scheduler(self):
        """Start the reminder scheduler in a separate thread"""
        def scheduler():
            while True:
                try:
                    self.check_and_send_reminders()
                    time.sleep(3600)  # Check every hour
                except Exception as e:
                    print(f"Scheduler error: {e}")
                    time.sleep(3600)
        
        thread = threading.Thread(target=scheduler, daemon=True)
        thread.start()
    
    def check_and_send_reminders(self):
        """Check for issues that need reminders and send them"""
        try:
            today = datetime.now().date()
            
            for _, issue in self.df.iterrows():
                mgmt_date = self.parse_any_date(issue.get('Management Resolution Date', ''))
                if not mgmt_date:
                    continue
                
                days_until_due = (mgmt_date - today).days
                
                # Check if reminder should be sent
                should_send = False
                for label, days_before in self.reminder_intervals:
                    if label == self.reminder_interval_var.get() and days_until_due == days_before:
                        should_send = True
                        break
                
                if should_send:
                    # Check if reminder was already sent today
                    last_reminder = issue['Last_Reminder']
                    if last_reminder and last_reminder == today.strftime('%Y-%m-%d'):
                        continue
                    
                    # Send reminder
                    self.send_email(issue)
                    
                    # Update reminder count
                    self.df.loc[self.df['Issue ID'] == issue['Issue ID'], 'Last_Reminder'] = today.strftime('%Y-%m-%d')
                    self.df.loc[self.df['Issue ID'] == issue['Issue ID'], 'Reminder_Count'] = issue['Reminder_Count'] + 1
                    self.save_data()
                    
        except Exception as e:
            print(f"Error checking reminders: {e}")
    
    def edit_selected_issue(self):
        """Edit the selected issue"""
        selection = self.issues_tree.selection()
        if not selection:
            messagebox.showwarning("Warning", "Please select an issue to edit")
            return
        
        # Get selected item
        item = self.issues_tree.item(selection[0])
        values = item['values']
        
        # Create edit dialog
        self.show_edit_issue_dialog(values)
    
    def show_edit_issue_dialog(self, values):
        """Show dialog to edit an issue"""
        dialog = ctk.CTkToplevel(self.root)
        dialog.title("Edit Audit Issue")
        dialog.geometry("500x600")
        dialog.transient(self.root)
        dialog.grab_set()
        
        # Form fields
        ctk.CTkLabel(dialog, text="Edit Audit Issue", font=("Arial", 16, "bold")).pack(pady=20)
        
        # Issue ID (read-only)
        ctk.CTkLabel(dialog, text="Issue ID:").pack(anchor='w', padx=20)
        id_label = ctk.CTkLabel(dialog, text=values[0], font=("Arial", 12, "bold"))
        id_label.pack(anchor='w', padx=20, pady=5)
        
        # Description
        ctk.CTkLabel(dialog, text="Issue Description:").pack(anchor='w', padx=20)
        description_entry = ctk.CTkEntry(dialog, width=600)
        description_entry.insert(0, values[1])
        description_entry.pack(fill='x', padx=20, pady=5)
        
        # Team
        ctk.CTkLabel(dialog, text="Team Name:").pack(anchor='w', padx=20)
        team_entry = ctk.CTkEntry(dialog, width=600)
        team_entry.insert(0, values[2])
        team_entry.pack(fill='x', padx=20, pady=5)
        
        # Management Resolution Date
        ctk.CTkLabel(dialog, text="Management Resolution Date (YYYY-MM-DD):").pack(anchor='w', padx=20)
        date_entry = ctk.CTkEntry(dialog, width=600)
        date_entry.insert(0, values[7] if values[7] != 'Not set' else '')
        date_entry.pack(fill='x', padx=20, pady=5)
        
        def save_changes():
            # Validate fields
            if not all([description_entry.get(), team_entry.get()]):
                messagebox.showerror("Error", "Description and Team are required")
                return
            
            # Validate date if provided
            if date_entry.get():
                try:
                    datetime.strptime(date_entry.get(), '%Y-%m-%d')
                except ValueError:
                    messagebox.showerror("Error", "Invalid date format. Use YYYY-MM-DD")
                    return
            
            # Update issue
            issue_id = values[0]
            self.df.loc[self.df['Issue ID'] == issue_id, 'Issue Description'] = description_entry.get()
            self.df.loc[self.df['Issue ID'] == issue_id, 'Team Name'] = team_entry.get()
            self.df.loc[self.df['Issue ID'] == issue_id, 'Management Resolution Date'] = date_entry.get() if date_entry.get() else ''
            
            self.save_data()
            self.update_issues_table()
            self.update_recent_issues()
            self.update_issue_combo()
            
            dialog.destroy()
            messagebox.showinfo("Success", "Issue updated successfully!")
        
        # Buttons
        button_frame = ctk.CTkFrame(dialog)
        button_frame.pack(fill='x', padx=20, pady=20)
        
        ctk.CTkButton(button_frame, text="Save Changes", command=save_changes).pack(side='right', padx=5)
        ctk.CTkButton(button_frame, text="Cancel", command=dialog.destroy).pack(side='right', padx=5)
    
    def delete_selected_issue(self):
        """Delete the selected issue"""
        selection = self.issues_tree.selection()
        if not selection:
            messagebox.showwarning("Warning", "Please select an issue to delete")
            return
        
        # Confirm deletion
        if not messagebox.askyesno("Confirm Delete", "Are you sure you want to delete this issue?"):
            return
        
        # Get selected item
        item = self.issues_tree.item(selection[0])
        values = item['values']
        issue_id = values[0]
        
        # Remove from dataframe
        self.df = self.df[self.df['Issue ID'] != issue_id]
        self.save_data()
        
        # Update UI
        self.update_issues_table()
        self.update_recent_issues()
        self.update_issue_combo()
        
        messagebox.showinfo("Success", "Issue deleted successfully!")
    
    def send_reminder_to_selected(self):
        """Send reminder to the selected issue"""
        selection = self.issues_tree.selection()
        if not selection:
            messagebox.showwarning("Warning", "Please select an issue to send reminder")
            return
        
        # Get selected item
        item = self.issues_tree.item(selection[0])
        values = item['values']
        issue_id = values[0]
        
        # Get issue data
        issue = self.df[self.df['Issue ID'] == issue_id].iloc[0]
        
        # Send email
        success = self.send_email(issue)
        
        if success:
            # Update reminder count
            self.df.loc[self.df['Issue ID'] == issue_id, 'Last_Reminder'] = datetime.now().strftime('%Y-%m-%d')
            self.df.loc[self.df['Issue ID'] == issue_id, 'Reminder_Count'] = issue['Reminder_Count'] + 1
            self.save_data()
            self.update_issues_table()
            messagebox.showinfo("Success", "Reminder email sent successfully!")
        else:
            messagebox.showerror("Error", "Failed to send reminder email")
    
    def send_all_overdue_reminders(self):
        """Send reminders for all overdue issues"""
        if not messagebox.askyesno("Confirm", "Send reminders for all overdue issues?"):
            return
        
        overdue_issues = self.df[
            (self.df['Management Resolution Date'] != '') & 
            (pd.to_datetime(self.df['Management Resolution Date']) < datetime.now())
        ]
        
        if overdue_issues.empty:
            messagebox.showinfo("Info", "No overdue issues found")
            return
        
        sent_count = 0
        for _, issue in overdue_issues.iterrows():
            if self.send_email(issue):
                sent_count += 1
                # Update reminder count
                self.df.loc[self.df['Issue ID'] == issue['Issue ID'], 'Last_Reminder'] = datetime.now().strftime('%Y-%m-%d')
                self.df.loc[self.df['Issue ID'] == issue['Issue ID'], 'Reminder_Count'] = issue['Reminder_Count'] + 1
        
        self.save_data()
        self.update_issues_table()
        messagebox.showinfo("Success", f"Sent {sent_count} overdue reminders!")
    
    def send_weekly_reminders(self):
        """Send reminders for issues due this week"""
        if not messagebox.askyesno("Confirm", "Send reminders for issues due this week?"):
            return
        
        today = datetime.now().date()
        week_end = today + timedelta(days=7)
        
        weekly_issues = self.df[
            (self.df['Management Resolution Date'] != '') & 
            (pd.to_datetime(self.df['Management Resolution Date']) <= week_end) &
            (pd.to_datetime(self.df['Management Resolution Date']) >= today)
        ]
        
        if weekly_issues.empty:
            messagebox.showinfo("Info", "No issues due this week")
            return
        
        sent_count = 0
        for _, issue in weekly_issues.iterrows():
            if self.send_email(issue):
                sent_count += 1
                # Update reminder count
                self.df.loc[self.df['Issue ID'] == issue['Issue ID'], 'Last_Reminder'] = datetime.now().strftime('%Y-%m-%d')
                self.df.loc[self.df['Issue ID'] == issue['Issue ID'], 'Reminder_Count'] = issue['Reminder_Count'] + 1
        
        self.save_data()
        self.update_issues_table()
        messagebox.showinfo("Success", f"Sent {sent_count} weekly reminders!")

    def on_dashboard_select(self, event):
        selected = self.dashboard_tree.selection()
        if not selected:
            self.selected_issue = None
            if hasattr(self, 'email_preview_box'):
                self.email_preview_box.delete('1.0', 'end')
            return
        item = self.dashboard_tree.item(selected[0])
        values = item['values']
        columns = self.dashboard_tree_cols
        issue = {col: values[i] for i, col in enumerate(columns)}
        # Find the full issue row in self.df (to get all fields)
        if 'Issue ID' in issue and 'Issue ID' in self.df.columns:
            df_row = self.df[self.df['Issue ID'] == issue['Issue ID']]
        else:
            df_row = pd.DataFrame()
        if not df_row.empty:
            issue = df_row.iloc[0].to_dict()
        self.selected_issue = issue
        # Auto-select template if associated
        for name, template in self.email_templates.items():
            assoc_col = template.get('assoc_column', '')
            assoc_val = template.get('assoc_value', '')
            if assoc_col and assoc_val and issue.get(assoc_col, '') == assoc_val:
                self.dashboard_template_select_var.set(name)
                self.load_selected_template()
                break
        self.render_email_preview(issue)

    def on_reminder_interval_change(self, *args):
        if self.selected_issue:
            self.render_email_preview(self.selected_issue)

    def on_dashboard_template_select(self, *args):
        # Always reload the latest template content and render the preview
        self.render_email_preview(self.selected_issue if self.selected_issue else {})

    def render_email_preview(self, issue):
        if not hasattr(self, 'email_preview_box'):
            return
        template = None
        if hasattr(self, 'dashboard_template_select_var') and self.dashboard_template_select_var.get():
            name = self.dashboard_template_select_var.get()
            if name in self.email_templates:
                template = self.email_templates[name]['content']
        if template is None:
            try:
                with open(self.email_template_file, 'r') as f:
                    template = f.read()
            except Exception as e:
                self.email_preview_box.delete('1.0', 'end')
                return
        # Calculate reminder date and days remaining
        interval_label = self.reminder_interval_var.get() if hasattr(self, 'reminder_interval_var') else "7 days before"
        interval_days = 7
        for label, days in self.reminder_intervals:
            if label == interval_label:
                interval_days = days
                break
        try:
            mgmt_date_str = str(issue.get('Management Resolution Date', ''))
            mgmt_date = self.parse_any_date(mgmt_date_str)
            if mgmt_date:
                reminder_date = mgmt_date - timedelta(days=interval_days)
                days_remaining = (reminder_date - datetime.now().date()).days
                reminder_date_str = reminder_date.strftime('%Y-%m-%d')
            else:
                reminder_date_str = ''
                days_remaining = ''
        except Exception:
            reminder_date_str = ''
            days_remaining = ''
        preview = template
        preview = preview.replace('{{ISSUE_ID}}', str(issue.get('Issue ID', '')))
        preview = preview.replace('{{DESCRIPTION}}', str(issue.get('Issue Description', '')))
        preview = preview.replace('{{RESOLUTION_DATE}}', str(issue.get('Management Resolution Date', '')))
        preview = preview.replace('{{DAYS_REMAINING}}', str(days_remaining))
        preview = preview.replace('{{REMINDER_DATE}}', str(reminder_date_str))
        preview = preview.replace('{{TEAM}}', str(issue.get('Team Name', '')))
        preview = preview.replace('{{CURRENT_DATE}}', datetime.now().strftime('%Y-%m-%d'))
        preview = preview.replace('{{REMINDER_COUNT}}', str(issue.get('Reminder_Count', '')))
        self.email_preview_box.delete('1.0', 'end')
        self.email_preview_box.insert('1.0', preview)
        # The user can now edit/format the preview as desired before sending

    def send_dashboard_email(self):
        if not hasattr(self, 'email_preview_box'):
            return
        if not self.selected_issue:
            messagebox.showwarning("No Issue Selected", "Please select an issue from the table.")
            return
        # Use the current content of the preview box as the email body
        subject = f"Audit Issue Reminder: {self.selected_issue.get('Issue ID', '')} - {self.selected_issue.get('Issue Description', '')[:50]}"
        body = self.email_preview_box.get('1.0', 'end').strip()
        mailto_link = f"mailto:?subject={urllib.parse.quote(subject)}&body={urllib.parse.quote(body)}"
        webbrowser.open(mailto_link)

    def copy_email_preview(self):
        if not hasattr(self, 'email_preview_box'):
            return
        self.root.clipboard_clear()
        self.root.clipboard_append(self.email_preview_box.get('1.0', 'end').strip())
        messagebox.showinfo("Copied", "Email preview copied to clipboard.")

    def load_dashboard_columns_config(self):
        import os, json
        if os.path.exists(self.dashboard_columns_file):
            with open(self.dashboard_columns_file, 'r', encoding='utf-8') as f:
                cols = json.load(f)
            # Only keep columns that exist in the DataFrame
            self.dashboard_display_columns = [col for col in cols if col in self.df.columns]
        else:
            self.dashboard_display_columns = [
                'Reporting Month', 'Issue ID', 'Responsible Business Segment', 'Team Name', 'Issue Title', 'Issue Summary',
                'Issue Description', 'Impact and Likelihood', 'Recommendation', 'Action Plan Due Date',
                'Management Resolution Date', '1B Contact', 'Previous Month Rationale', 'Line of Business',
                'Reporting Categories', 'Root Cause Category Selected', 'Root Cause Sub Category Selected', 'Risk Group'
            ]
            self.dashboard_display_columns = [col for col in self.dashboard_display_columns if col in self.df.columns]

    def save_dashboard_columns_config(self):
        import json
        with open(self.dashboard_columns_file, 'w', encoding='utf-8') as f:
            json.dump(self.dashboard_display_columns, f, indent=2)

    def open_configure_columns_popup(self):
        import tkinter as tk
        popup = tk.Toplevel(self.root)
        popup.title("Configure Dashboard Columns")
        popup.geometry("350x600")
        popup.transient(self.root)
        popup.grab_set()
        required_columns = [
            'Issue ID', 'Issue Description', 'Team Name', 'Management Resolution Date',
            'Reporting Month', 'Responsible Business Segment', 'Issue Summary',
            'Impact and Likelihood', 'Recommendation', 'Previous Month Rationale',
            'Line of Business', 'Reporting Categories'
        ]
        all_columns = list(self.df.columns)
        for col in required_columns:
            if col not in all_columns:
                all_columns.append(col)
        var_dict = {}
        frame = tk.Frame(popup)
        frame.pack(fill='both', expand=True, padx=10, pady=10)
        for col in all_columns:
            var = tk.BooleanVar(value=(col in self.dashboard_display_columns))
            cb = tk.Checkbutton(frame, text=str(col), variable=var, anchor='w')
            cb.pack(fill='x', anchor='w')
            var_dict[col] = var
        def on_save():
            self.dashboard_display_columns = [col for col, var in var_dict.items() if var.get()]
            self.save_dashboard_columns_config()
            popup.destroy()
            self.create_dashboard_tab()
            self.create_email_tab()
            return  # Do not call any other update methods here
        tk.Button(popup, text="Save", command=on_save).pack(pady=10)

    def create_settings_tab(self):
        frame = self.settings_tab
        for widget in frame.winfo_children():
            widget.destroy()
        ctk.CTkLabel(frame, text="System Settings", font=("Arial", 18, "bold")).pack(pady=20)
        # (Removed Configure Dashboard Columns section from settings tab)
        email_settings_frame = ctk.CTkFrame(frame)
        email_settings_frame.pack(fill='x', padx=20, pady=10)
        info_text = """Corporate Email Setup:\n\nThis application will use your corporate email settings automatically.\nNo manual SMTP configuration is required.\n\nTo send emails:\n1. Ensure you're logged into your corporate email on this machine\n2. The system will use your default email application\n3. Emails will be sent through your corporate email system\n\nNote: If you encounter permission issues, contact your IT department."""
        ctk.CTkLabel(email_settings_frame, text=info_text, justify='left', font=("Arial", 12)).pack(anchor='w', pady=10)
        ctk.CTkButton(email_settings_frame, text="Test Email Configuration", command=self.test_email_config).pack(pady=10)
        # Reminder interval options removed from settings tab

    def apply_text_tag(self, tag):
        try:
            start = self.email_template_box.index("sel.first")
            end = self.email_template_box.index("sel.last")
            self.email_template_box.tag_add(tag, start, end)
        except Exception:
            pass  # No selection

    def apply_dashboard_text_tag(self, tag):
        if not hasattr(self, 'email_preview_box'):
            return
        try:
            start = self.email_preview_box.index("sel.first")
            end = self.email_preview_box.index("sel.last")
            self.email_preview_box.tag_add(tag, start, end)
        except Exception:
            pass

    def get_text_tags(self, text_widget):
        # Get all tag ranges for bold, italic, underline
        tags = {}
        for tag in ('bold', 'italic', 'underline'):
            tag_ranges = text_widget.tag_ranges(tag)
            tags[tag] = [(str(tag_ranges[i]), str(tag_ranges[i+1])) for i in range(0, len(tag_ranges), 2)]
        return tags

    def apply_text_tags(self, text_widget, tags):
        # Apply tag ranges to a Text widget
        for tag, ranges in tags.items():
            for start, end in ranges:
                try:
                    text_widget.tag_add(tag, start, end)
                except Exception:
                    pass

# Main entry point
if __name__ == "__main__":
    root = ctk.CTk()
    app = AuditManagerApp(root)
    root.mainloop()
    