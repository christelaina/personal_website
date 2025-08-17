import customtkinter as ctk
from tkinter import messagebox, filedialog
import tkinter.ttk as ttk
import pandas as pd
import os
import json
from datetime import datetime, timedelta
from email_validator import validate_email, EmailNotValidError
import threading
import time
import webbrowser
import urllib.parse

ctk.set_appearance_mode("System")  # "Dark", "Light", or "System"
ctk.set_default_color_theme("green")

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
        self.excel_file = 'audit_issues.xlsx'
        self.config_file = 'config.json'
        self.email_template_file = 'email_template.html'
        
        # Initialize files
        self.init_files()
        
        # Load data
        self.load_data()
        
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
                'ID': [],
                'Description': [],
                'Team': [],
                'Team_Email': [],
                'Priority': [],
                'Status': [],
                'Created_Date': [],
                'Resolution_Date': [],
                'Last_Reminder': [],
                'Reminder_Count': []
            }
            
            for col in required_columns:
                if col not in self.df.columns:
                    print(f"File is missing column: {col}")
            
            # If dataframe is empty, initialize with sample data structure
            if self.df.empty:
                print("Excel file is empty. Ready to add new issues.")
                
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load Excel file: {str(e)}")
            self.df = pd.DataFrame()

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
        # Set dashboard_display_columns before any use
        self.dashboard_display_columns = [
            'Reporting Month', 'Issue ID', 'Responsible Business Segment', 'Team Name', 'Issue Title', 'Issue Summary',
            'Issue Description', 'Impact and Likelihood', 'Recommendation', 'Action Plan Due Date',
            'Management Resolution Date', '1B Contact', 'Previous Month Rationale', 'Line of Business',
            'Reporting Categories', 'Root Cause Category Selected', 'Root Cause Sub Category Selected', 'Risk Group'
        ]
        ctk.CTkLabel(frame, text="Audit Issue Dashboard", font=("Arial", 22, "bold")).pack(pady=(10, 2))
        stats_frame = ctk.CTkFrame(frame, fg_color="transparent")
        stats_frame.pack(fill='x', padx=20, pady=(0, 5))
        total_issues = len(self.df)
        open_issues = len(self.df[self.df['Status'] == 'Open']) if not self.df.empty and 'Status' in self.df.columns else 0
        resolved_issues = len(self.df[self.df['Status'] == 'Resolved']) if not self.df.empty and 'Status' in self.df.columns else 0
        stats = [
            ("Total", total_issues, "#3b82f6"),
            ("Open", open_issues, "#f59e0b"),
            ("Resolved", resolved_issues, "#10b981")
        ]
        for i, (label, value, color) in enumerate(stats):
            stat = ctk.CTkFrame(stats_frame, fg_color="transparent")
            stat.grid(row=0, column=i, padx=8, pady=0, sticky='w')
            ctk.CTkLabel(stat, text=label, font=("Arial", 12)).pack(side='left', padx=(0, 2))
            ctk.CTkLabel(stat, text=str(value), font=("Arial", 14, "bold"), text_color=color).pack(side='left')
        stats_frame.grid_columnconfigure((0, 1, 2), weight=0)

        # --- SEARCH, FILTER, SORT ---
        filter_frame = ctk.CTkFrame(frame)
        filter_frame.pack(fill='x', padx=20, pady=5)
        # Multi-column search bar
        ctk.CTkLabel(filter_frame, text="Search:").pack(side='left', padx=(0, 5))
        self.multi_search_var = ctk.StringVar()
        self.multi_search_entry = ctk.CTkEntry(filter_frame, textvariable=self.multi_search_var, width=220)
        self.multi_search_entry.pack(side='left', padx=(0, 20))
        self.multi_search_entry.bind('<KeyRelease>', lambda e: self.update_dashboard_table())
        # Multi-select filter
        ctk.CTkLabel(filter_frame, text="Filter by:").pack(side='left', padx=(0, 5))
        filter_columns = [col for col in self.dashboard_display_columns if col in self.df.columns]
        self.filter_column_var = ctk.StringVar(value=filter_columns[0] if filter_columns else "")
        self.filter_column_box = ctk.CTkComboBox(filter_frame, values=filter_columns, variable=self.filter_column_var, width=150, command=self.on_filter_column_change)
        self.filter_column_box.pack(side='left', padx=(0, 5))
        self.filter_multi_values = set()
        self.filter_multi_button = ctk.CTkButton(filter_frame, text="Select Values", command=self.open_multi_filter_popup)
        self.filter_multi_button.pack(side='left', padx=(0, 20))
        # Sort by column (clickable headers)
        ctk.CTkLabel(filter_frame, text="Sort by:").pack(side='left', padx=(0, 5))
        self.sort_column_var = ctk.StringVar(value=filter_columns[0] if filter_columns else "")
        self.sort_column_box = ctk.CTkComboBox(filter_frame, values=filter_columns, variable=self.sort_column_var, width=150, command=self.update_dashboard_table)
        self.sort_column_box.pack(side='left', padx=(0, 5))
        self.sort_order_var = ctk.StringVar(value="Ascending")
        self.sort_order_box = ctk.CTkComboBox(filter_frame, values=["Ascending", "Descending"], variable=self.sort_order_var, width=100, command=self.update_dashboard_table)
        self.sort_order_box.pack(side='left', padx=(0, 5))
        self.on_filter_column_change()  # Initialize filter values

        # --- ISSUES TABLE ---
        table_frame = ctk.CTkFrame(frame)
        table_frame.pack(fill='both', expand=True, padx=20, pady=(5, 10))
        ctk.CTkLabel(table_frame, text="Audit Issues", font=("Arial", 14, "bold")).pack(anchor='w')
        dashboard_cols = [col for col in self.dashboard_display_columns if col in self.df.columns]
        self.dashboard_tree = ttk.Treeview(table_frame, columns=dashboard_cols, show='headings', height=12)
        for col in dashboard_cols:
            self.dashboard_tree.heading(col, text=col, command=lambda c=col: self.on_column_header_click(c))
            self.dashboard_tree.column(col, width=180)
        self.dashboard_tree.pack(fill='both', expand=True, pady=5)
        self.dashboard_tree.bind('<<TreeviewSelect>>', self.on_dashboard_select)
        self.dashboard_tree_cols = dashboard_cols
        self.sort_state = {col: None for col in dashboard_cols}  # None, 'asc', 'desc'
        self.update_dashboard_table()

        # --- EMAIL PREVIEW PANEL ---
        preview_frame = ctk.CTkFrame(frame)
        preview_frame.pack(fill='x', padx=20, pady=(0, 10))
        ctk.CTkLabel(preview_frame, text="Email Preview", font=("Arial", 14, "bold")).pack(anchor='w')
        interval_frame = ctk.CTkFrame(preview_frame, fg_color="transparent")
        interval_frame.pack(anchor='w', pady=(0, 5))
        ctk.CTkLabel(interval_frame, text="Reminder Interval:").pack(side='left', padx=(0, 5))
        self.reminder_interval_var = ctk.StringVar(value="7 days before")
        self.reminder_intervals = [
            ("7 days before", 7),
            ("30 days before", 30),
            ("3 months before", 90),
            ("6 months before", 180)
        ]
        interval_options = [label for label, _ in self.reminder_intervals]
        self.reminder_interval_box = ctk.CTkComboBox(interval_frame, values=interval_options, variable=self.reminder_interval_var, width=160, command=self.on_reminder_interval_change)
        self.reminder_interval_box.pack(side='left', padx=(0, 10))
        self.email_preview_box = ctk.CTkTextbox(preview_frame, height=200, font=("Consolas", 12))
        self.email_preview_box.pack(fill='x', expand=True, pady=5)
        button_frame = ctk.CTkFrame(preview_frame, fg_color="transparent")
        button_frame.pack(anchor='e', pady=5)
        ctk.CTkButton(button_frame, text="Send Email", command=self.send_dashboard_email).pack(side='left', padx=5)
        ctk.CTkButton(button_frame, text="Copy to Clipboard", command=self.copy_email_preview).pack(side='left', padx=5)

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
        # Load templates from file
        self.email_templates_file = 'email_templates.json'
        self.email_templates = self.load_email_templates()
        self.selected_template_name = None
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
        # HTML/Text toggle
        toggle_frame = ctk.CTkFrame(frame, fg_color="transparent")
        toggle_frame.pack(fill='x', padx=20, pady=2)
        self.html_mode = ctk.BooleanVar(value=True)
        ctk.CTkCheckBox(toggle_frame, text="Edit as HTML", variable=self.html_mode, command=self.on_html_mode_toggle).pack(side='left')
        # Template editor
        editor_frame = ctk.CTkFrame(frame)
        editor_frame.pack(fill='both', expand=True, padx=20, pady=5)
        ctk.CTkLabel(editor_frame, text="Edit Template:").pack(anchor='w')
        self.template_text = ctk.CTkTextbox(editor_frame, height=250, width=900, font=("Consolas", 12))
        self.template_text.pack(fill='both', expand=True, pady=10)
        # Save button
        save_frame = ctk.CTkFrame(frame, fg_color="transparent")
        save_frame.pack(fill='x', padx=20, pady=2)
        ctk.CTkButton(save_frame, text="Save Template", command=self.save_email_template).pack(side='left', padx=5)
        # Live preview
        preview_frame = ctk.CTkFrame(frame)
        preview_frame.pack(fill='x', padx=20, pady=10)
        ctk.CTkLabel(preview_frame, text="Live Preview", font=("Arial", 14, "bold")).pack(anchor='w')
        self.email_template_preview_box = ctk.CTkTextbox(preview_frame, height=200, font=("Consolas", 12))
        self.email_template_preview_box.pack(fill='x', expand=True, pady=5)
        # Initial load
        self.load_selected_template()
        self.update_email_template_preview()

    def load_email_templates(self):
        import os, json
        if not os.path.exists(self.email_templates_file):
            return {}
        with open(self.email_templates_file, 'r', encoding='utf-8') as f:
            return json.load(f)

    def save_email_templates(self):
        with open(self.email_templates_file, 'w', encoding='utf-8') as f:
            json.dump(self.email_templates, f, indent=2)

    def on_template_select(self, *args):
        self.selected_template_name = self.template_select_var.get()
        self.load_selected_template()
        self.update_email_template_preview()

    def load_selected_template(self):
        name = self.template_select_var.get()
        if name and name in self.email_templates:
            template = self.email_templates[name]
            self.template_text.delete('1.0', 'end')
            self.template_text.insert('1.0', template.get('content', ''))
            self.assoc_column_var.set(template.get('assoc_column', 'None'))
            self.assoc_value_var.set(template.get('assoc_value', ''))
        else:
            self.template_text.delete('1.0', 'end')
            self.assoc_column_var.set('None')
            self.assoc_value_var.set('')

    def save_email_template(self):
        name = self.template_select_var.get()
        if not name:
            messagebox.showerror("Error", "Please enter a template name.")
            return
        content = self.template_text.get('1.0', 'end').strip()
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
        self.update_email_template_preview()

    def new_email_template(self):
        import tkinter.simpledialog
        name = tkinter.simpledialog.askstring("New Template", "Enter template name:")
        if not name:
            return
        if name in self.email_templates:
            messagebox.showerror("Error", "A template with this name already exists.")
            return
        self.email_templates[name] = {'content': '', 'assoc_column': '', 'assoc_value': ''}
        self.save_email_templates()
        self.template_names = list(self.email_templates.keys())
        self.template_select_box.configure(values=self.template_names)
        self.template_select_var.set(name)
        self.load_selected_template()
        self.update_email_template_preview()

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
            else:
                self.template_select_var.set("")
            self.load_selected_template()
            self.update_email_template_preview()

    def on_html_mode_toggle(self):
        # For now, just a toggle; could add HTML syntax highlighting in the future
        pass

    def update_email_template_preview(self):
        # Use the current template content and render with a dummy issue or the last selected issue
        content = self.template_text.get('1.0', 'end')
        issue = self.selected_issue if hasattr(self, 'selected_issue') and self.selected_issue else {col: f"[{col}]" for col in self.df.columns}
        preview = content
        for col in self.df.columns:
            preview = preview.replace(f'{{{{{col}}}}}', str(issue.get(col, f'[{col}]')))
        preview = preview.replace('{{CURRENT_DATE}}', datetime.now().strftime('%Y-%m-%d'))
        import re
        text_preview = re.sub('<[^<]+?>', '', preview)
        self.email_template_preview_box.delete('1.0', 'end')
        self.email_template_preview_box.insert('1.0', text_preview)

    def create_settings_tab(self):
        frame = self.settings_tab
        for widget in frame.winfo_children():
            widget.destroy()
        ctk.CTkLabel(frame, text="System Settings", font=("Arial", 18, "bold")).pack(pady=20)
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
        dialog.geometry("500x600")
        dialog.transient(self.root)
        dialog.grab_set()
        
        # Form fields
        ctk.CTkLabel(dialog, text="Add New Audit Issue", font=("Arial", 16, "bold")).pack(pady=20)
        
        # Description
        ctk.CTkLabel(dialog, text="Description:").pack(anchor='w', padx=20)
        description_entry = ctk.CTkEntry(dialog, width=600)
        description_entry.pack(fill='x', padx=20, pady=5)
        
        # Team
        ctk.CTkLabel(dialog, text="Team:").pack(anchor='w', padx=20)
        team_entry = ctk.CTkEntry(dialog, width=600)
        team_entry.pack(fill='x', padx=20, pady=5)
        
        # Team Email
        ctk.CTkLabel(dialog, text="Team Email:").pack(anchor='w', padx=20)
        email_entry = ctk.CTkEntry(dialog, width=600)
        email_entry.pack(fill='x', padx=20, pady=5)
        
        # Priority
        ctk.CTkLabel(dialog, text="Priority:").pack(anchor='w', padx=20)
        priority_var = ctk.StringVar(value="Medium")
        priority_combo = ctk.CTkComboBox(dialog, textvariable=priority_var, values=["High", "Medium", "Low"], state='readonly')
        priority_combo.pack(fill='x', padx=20, pady=5)
        
        # Resolution Date
        ctk.CTkLabel(dialog, text="Resolution Date (YYYY-MM-DD):").pack(anchor='w', padx=20)
        date_entry = ctk.CTkEntry(dialog, width=600)
        date_entry.pack(fill='x', padx=20, pady=5)
        
        def save_issue():
            # Validate fields
            if not all([description_entry.get(), team_entry.get(), email_entry.get(), date_entry.get()]):
                messagebox.showerror("Error", "All fields are required")
                return
            
            # Validate email
            try:
                validate_email(email_entry.get())
            except EmailNotValidError:
                messagebox.showerror("Error", "Invalid email address")
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
                'ID': new_id,
                'Description': description_entry.get(),
                'Team': team_entry.get(),
                'Team_Email': email_entry.get(),
                'Priority': priority_var.get(),
                'Status': 'Open',
                'Created_Date': datetime.now().strftime('%Y-%m-%d'),
                'Resolution_Date': date_entry.get(),
                'Last_Reminder': '',
                'Reminder_Count': 0
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
            issue_list = [f"{row['ID']} - {row['Description'][:50]}..." for _, row in self.df.iterrows()]
            self.issue_combo['values'] = issue_list
            if issue_list:
                self.issue_combo.set(issue_list[0])
    
    def load_email_template(self):
        """Load the email template from file"""
        try:
            with open(self.email_template_file, 'r') as f:
                template = f.read()
                self.template_text.delete(1.0, ctk.END)
                self.template_text.insert(1.0, template)
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load email template: {str(e)}")
    
    def save_email_template(self):
        """Save the email template to file"""
        try:
            template = self.template_text.get(1.0, ctk.END)
            with open(self.email_template_file, 'w') as f:
                f.write(template)
            messagebox.showinfo("Success", "Email template saved successfully!")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to save email template: {str(e)}")
    
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
        issue = self.df[self.df['ID'] == issue_id].iloc[0]
        
        # Send email
        success = self.send_email(issue)
        
        if success:
            # Update reminder count
            self.df.loc[self.df['ID'] == issue_id, 'Last_Reminder'] = datetime.now().strftime('%Y-%m-%d')
            self.df.loc[self.df['ID'] == issue_id, 'Reminder_Count'] = issue['Reminder_Count'] + 1
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
            resolution_date = datetime.strptime(issue['Resolution_Date'], '%Y-%m-%d')
            days_remaining = (resolution_date - datetime.now()).days
            
            # Replace template variables
            template = template.replace('{{ISSUE_ID}}', issue['ID'])
            template = template.replace('{{DESCRIPTION}}', issue['Description'])
            template = template.replace('{{PRIORITY}}', issue['Priority'])
            template = template.replace('{{STATUS}}', issue['Status'])
            template = template.replace('{{RESOLUTION_DATE}}', issue['Resolution_Date'])
            template = template.replace('{{DAYS_REMAINING}}', str(days_remaining))
            template = template.replace('{{TEAM}}', issue['Team'])
            template = template.replace('{{CREATED_DATE}}', issue['Created_Date'])
            template = template.replace('{{REMINDER_COUNT}}', str(issue['Reminder_Count'] + 1))
            template = template.replace('{{CURRENT_DATE}}', datetime.now().strftime('%Y-%m-%d'))
            
            # Create email content
            subject = f"Audit Issue Reminder: {issue['ID']} - {issue['Description'][:50]}"
            body = f"""
Audit Issue Resolution Required

Issue ID: {issue['ID']}
Description: {issue['Description']}
Priority: {issue['Priority']}
Status: {issue['Status']}
Resolution Due Date: {issue['Resolution_Date']}
Days Remaining: {days_remaining}
Team: {issue['Team']}
Created Date: {issue['Created_Date']}
Reminder Count: {issue['Reminder_Count'] + 1}

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
            
            print(f"Email application opened for issue {issue['ID']}")
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
            today = datetime.now()
            
            for _, issue in self.df.iterrows():
                if issue['Status'] != 'Open':
                    continue
                
                if not issue['Resolution_Date']:
                    continue
                
                resolution_date = datetime.strptime(issue['Resolution_Date'], '%Y-%m-%d')
                days_until_due = (resolution_date - today).days
                
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
                    self.df.loc[self.df['ID'] == issue['ID'], 'Last_Reminder'] = today.strftime('%Y-%m-%d')
                    self.df.loc[self.df['ID'] == issue['ID'], 'Reminder_Count'] = issue['Reminder_Count'] + 1
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
        ctk.CTkLabel(dialog, text="Description:").pack(anchor='w', padx=20)
        description_entry = ctk.CTkEntry(dialog, width=600)
        description_entry.insert(0, values[1])
        description_entry.pack(fill='x', padx=20, pady=5)
        
        # Team
        ctk.CTkLabel(dialog, text="Team:").pack(anchor='w', padx=20)
        team_entry = ctk.CTkEntry(dialog, width=600)
        team_entry.insert(0, values[2])
        team_entry.pack(fill='x', padx=20, pady=5)
        
        # Team Email
        ctk.CTkLabel(dialog, text="Team Email:").pack(anchor='w', padx=20)
        email_entry = ctk.CTkEntry(dialog, width=600)
        email_entry.insert(0, values[3])
        email_entry.pack(fill='x', padx=20, pady=5)
        
        # Priority
        ctk.CTkLabel(dialog, text="Priority:").pack(anchor='w', padx=20)
        priority_var = ctk.StringVar(value=values[4])
        priority_combo = ctk.CTkComboBox(dialog, textvariable=priority_var, values=["High", "Medium", "Low"], state='readonly')
        priority_combo.pack(fill='x', padx=20, pady=5)
        
        # Status
        ctk.CTkLabel(dialog, text="Status:").pack(anchor='w', padx=20)
        status_var = ctk.StringVar(value=values[5])
        status_combo = ctk.CTkComboBox(dialog, textvariable=status_var, values=["Open", "In Progress", "Resolved", "Closed"], state='readonly')
        status_combo.pack(fill='x', padx=20, pady=5)
        
        # Resolution Date
        ctk.CTkLabel(dialog, text="Resolution Date (YYYY-MM-DD):").pack(anchor='w', padx=20)
        date_entry = ctk.CTkEntry(dialog, width=600)
        date_entry.insert(0, values[7] if values[7] != 'Not set' else '')
        date_entry.pack(fill='x', padx=20, pady=5)
        
        def save_changes():
            # Validate fields
            if not all([description_entry.get(), team_entry.get(), email_entry.get()]):
                messagebox.showerror("Error", "Description, Team, and Email are required")
                return
            
            # Validate email
            try:
                validate_email(email_entry.get())
            except EmailNotValidError:
                messagebox.showerror("Error", "Invalid email address")
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
            self.df.loc[self.df['ID'] == issue_id, 'Description'] = description_entry.get()
            self.df.loc[self.df['ID'] == issue_id, 'Team'] = team_entry.get()
            self.df.loc[self.df['ID'] == issue_id, 'Team_Email'] = email_entry.get()
            self.df.loc[self.df['ID'] == issue_id, 'Priority'] = priority_var.get()
            self.df.loc[self.df['ID'] == issue_id, 'Status'] = status_var.get()
            self.df.loc[self.df['ID'] == issue_id, 'Resolution_Date'] = date_entry.get() if date_entry.get() else ''
            
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
        self.df = self.df[self.df['ID'] != issue_id]
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
        issue = self.df[self.df['ID'] == issue_id].iloc[0]
        
        # Send email
        success = self.send_email(issue)
        
        if success:
            # Update reminder count
            self.df.loc[self.df['ID'] == issue_id, 'Last_Reminder'] = datetime.now().strftime('%Y-%m-%d')
            self.df.loc[self.df['ID'] == issue_id, 'Reminder_Count'] = issue['Reminder_Count'] + 1
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
            (self.df['Status'] == 'Open') & 
            (self.df['Resolution_Date'] != '') & 
            (pd.to_datetime(self.df['Resolution_Date']) < datetime.now())
        ]
        
        if overdue_issues.empty:
            messagebox.showinfo("Info", "No overdue issues found")
            return
        
        sent_count = 0
        for _, issue in overdue_issues.iterrows():
            if self.send_email(issue):
                sent_count += 1
                # Update reminder count
                self.df.loc[self.df['ID'] == issue['ID'], 'Last_Reminder'] = datetime.now().strftime('%Y-%m-%d')
                self.df.loc[self.df['ID'] == issue['ID'], 'Reminder_Count'] = issue['Reminder_Count'] + 1
        
        self.save_data()
        self.update_issues_table()
        messagebox.showinfo("Success", f"Sent {sent_count} overdue reminders!")
    
    def send_weekly_reminders(self):
        """Send reminders for issues due this week"""
        if not messagebox.askyesno("Confirm", "Send reminders for issues due this week?"):
            return
        
        today = datetime.now()
        week_end = today + timedelta(days=7)
        
        weekly_issues = self.df[
            (self.df['Status'] == 'Open') & 
            (self.df['Resolution_Date'] != '') & 
            (pd.to_datetime(self.df['Resolution_Date']) <= week_end) &
            (pd.to_datetime(self.df['Resolution_Date']) >= today)
        ]
        
        if weekly_issues.empty:
            messagebox.showinfo("Info", "No issues due this week")
            return
        
        sent_count = 0
        for _, issue in weekly_issues.iterrows():
            if self.send_email(issue):
                sent_count += 1
                # Update reminder count
                self.df.loc[self.df['ID'] == issue['ID'], 'Last_Reminder'] = datetime.now().strftime('%Y-%m-%d')
                self.df.loc[self.df['ID'] == issue['ID'], 'Reminder_Count'] = issue['Reminder_Count'] + 1
        
        self.save_data()
        self.update_issues_table()
        messagebox.showinfo("Success", f"Sent {sent_count} weekly reminders!")

    def on_filter_column_change(self, *args):
        col = self.filter_column_var.get()
        if col and col in self.df.columns:
            values = sorted(list(self.df[col].dropna().unique()))
            self.filter_value_box.configure(values=["All"] + values)
            self.filter_value_var.set("All")
        self.update_dashboard_table()

    def open_multi_filter_popup(self):
        import tkinter as tk
        col = self.filter_column_var.get()
        if not col or col not in self.df.columns:
            return
        values = sorted(list(self.df[col].dropna().unique()))
        popup = tk.Toplevel(self.root)
        popup.title(f"Select values for {col}")
        popup.geometry("300x400")
        popup.transient(self.root)
        popup.grab_set()
        var_dict = {}
        frame = tk.Frame(popup)
        frame.pack(fill='both', expand=True, padx=10, pady=10)
        for v in values:
            var = tk.BooleanVar(value=(v in self.filter_multi_values))
            cb = tk.Checkbutton(frame, text=str(v), variable=var, anchor='w')
            cb.pack(fill='x', anchor='w')
            var_dict[v] = var
        def on_ok():
            self.filter_multi_values = set(v for v, var in var_dict.items() if var.get())
            popup.destroy()
            self.update_dashboard_table()
        tk.Button(popup, text="OK", command=on_ok).pack(pady=10)

    def on_column_header_click(self, col):
        # Toggle sort order for the column
        current = self.sort_state.get(col)
        for k in self.sort_state:
            self.sort_state[k] = None
        if current == 'asc':
            self.sort_state[col] = 'desc'
        else:
            self.sort_state[col] = 'asc'
        self.sort_column_var.set(col)
        self.sort_order_var.set('Ascending' if self.sort_state[col] == 'asc' else 'Descending')
        self.update_dashboard_table()

    def update_dashboard_table(self, *args):
        if not hasattr(self, 'dashboard_tree'):
            return
        # Multi-column search
        search = self.multi_search_var.get().strip().lower() if hasattr(self, 'multi_search_var') else ""
        filter_col = self.filter_column_var.get() if hasattr(self, 'filter_column_var') else None
        filter_vals = self.filter_multi_values if hasattr(self, 'filter_multi_values') else set()
        sort_col = self.sort_column_var.get() if hasattr(self, 'sort_column_var') else None
        sort_order = self.sort_order_var.get() if hasattr(self, 'sort_order_var') else "Ascending"
        df = self.df.copy()
        # Multi-column search
        if search:
            mask = pd.Series([False] * len(df))
            for col in self.dashboard_tree_cols:
                if col in df.columns:
                    mask = mask | df[col].astype(str).str.lower().str.contains(search)
            df = df[mask]
        # Multi-select filter
        if filter_col and filter_col in df.columns and filter_vals:
            df = df[df[filter_col].isin(filter_vals)]
        # Sort
        if sort_col and sort_col in df.columns:
            df = df.sort_values(by=sort_col, ascending=(sort_order == "Ascending"))
        for item in self.dashboard_tree.get_children():
            self.dashboard_tree.delete(item)
        for _, row in df.iterrows():
            values = [row.get(col, '') for col in self.dashboard_tree_cols]
            self.dashboard_tree.insert('', 'end', values=values)
        self.selected_issue = None
        if hasattr(self, 'email_preview_box'):
            self.email_preview_box.delete('1.0', 'end')

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
                self.template_select_var.set(name)
                self.load_selected_template()
                break
        self.render_email_preview(issue)

    def on_reminder_interval_change(self, *args):
        if self.selected_issue:
            self.render_email_preview(self.selected_issue)

    def render_email_preview(self, issue):
        try:
            with open(self.email_template_file, 'r') as f:
                template = f.read()
        except Exception as e:
            if hasattr(self, 'email_preview_box'):
                self.email_preview_box.delete('1.0', 'end')
                self.email_preview_box.insert('1.0', f"Error loading template: {e}")
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
            mgmt_date = datetime.strptime(mgmt_date_str, '%Y-%m-%d')
            reminder_date = mgmt_date - timedelta(days=interval_days)
            days_remaining = (reminder_date - datetime.now()).days
            reminder_date_str = reminder_date.strftime('%Y-%m-%d')
        except Exception:
            reminder_date_str = ''
            days_remaining = ''
        preview = template
        preview = preview.replace('{{ISSUE_ID}}', str(issue.get('ID', '')))
        preview = preview.replace('{{DESCRIPTION}}', str(issue.get('Description', '')))
        preview = preview.replace('{{PRIORITY}}', str(issue.get('Priority', '')))
        preview = preview.replace('{{STATUS}}', str(issue.get('Status', '')))
        preview = preview.replace('{{RESOLUTION_DATE}}', str(issue.get('Management Resolution Date', '')))
        preview = preview.replace('{{DAYS_REMAINING}}', str(days_remaining))
        preview = preview.replace('{{REMINDER_DATE}}', str(reminder_date_str))
        preview = preview.replace('{{TEAM}}', str(issue.get('Team', '')))
        preview = preview.replace('{{CREATED_DATE}}', str(issue.get('Created_Date', '')))
        preview = preview.replace('{{REMINDER_COUNT}}', str(issue.get('Reminder_Count', '')))
        preview = preview.replace('{{CURRENT_DATE}}', datetime.now().strftime('%Y-%m-%d'))
        import re
        text_preview = re.sub('<[^<]+?>', '', preview)
        if hasattr(self, 'email_preview_box'):
            self.email_preview_box.delete('1.0', 'end')
            self.email_preview_box.insert('1.0', text_preview)

    def send_dashboard_email(self):
        if not self.selected_issue:
            messagebox.showwarning("No Issue Selected", "Please select an issue from the table.")
            return
        # Use the same logic as send_email, but for selected_issue
        issue = self.selected_issue
        try:
            with open(self.email_template_file, 'r') as f:
                template = f.read()
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load email template: {e}")
            return
        try:
            resolution_date = datetime.strptime(str(issue.get('Resolution_Date', '')), '%Y-%m-%d')
            days_remaining = (resolution_date - datetime.now()).days
        except Exception:
            days_remaining = ''
        email_body = template
        email_body = email_body.replace('{{ISSUE_ID}}', str(issue.get('ID', '')))
        email_body = email_body.replace('{{DESCRIPTION}}', str(issue.get('Description', '')))
        email_body = email_body.replace('{{PRIORITY}}', str(issue.get('Priority', '')))
        email_body = email_body.replace('{{STATUS}}', str(issue.get('Status', '')))
        email_body = email_body.replace('{{RESOLUTION_DATE}}', str(issue.get('Resolution_Date', '')))
        email_body = email_body.replace('{{DAYS_REMAINING}}', str(days_remaining))
        email_body = email_body.replace('{{TEAM}}', str(issue.get('Team', '')))
        email_body = email_body.replace('{{CREATED_DATE}}', str(issue.get('Created_Date', '')))
        email_body = email_body.replace('{{REMINDER_COUNT}}', str(issue.get('Reminder_Count', '')))
        email_body = email_body.replace('{{CURRENT_DATE}}', datetime.now().strftime('%Y-%m-%d'))
        subject = f"Audit Issue Reminder: {issue.get('ID', '')} - {issue.get('Description', '')[:50]}"
        body = re.sub('<[^<]+?>', '', email_body)
        mailto_link = f"mailto:{issue.get('Team_Email', '')}?subject={urllib.parse.quote(subject)}&body={urllib.parse.quote(body)}"
        webbrowser.open(mailto_link)

    def copy_email_preview(self):
        if hasattr(self, 'email_preview_box'):
            self.root.clipboard_clear()
            self.root.clipboard_append(self.email_preview_box.get('1.0', 'end').strip())
            messagebox.showinfo("Copied", "Email preview copied to clipboard.")

# Main entry point
if __name__ == "__main__":
    root = ctk.CTk()
    app = AuditManagerApp(root)
    root.mainloop()
    