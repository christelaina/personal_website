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
        ctk.CTkLabel(frame, text="Audit Issue Dashboard", font=("Arial", 22, "bold")).pack(pady=20)
        stats_frame = ctk.CTkFrame(frame)
        stats_frame.pack(fill='x', padx=20, pady=10)
        
        # Calculate statistics
        total_issues = len(self.df)
        open_issues = len(self.df[self.df['Status'] == 'Open']) if not self.df.empty else 0
        resolved_issues = len(self.df[self.df['Status'] == 'Resolved']) if not self.df.empty else 0
        
        # Create stat boxes
        stats = [
            ("Total Issues", total_issues, "#3b82f6"),
            ("Open Issues", open_issues, "#f59e0b"),
            ("Resolved Issues", resolved_issues, "#10b981")
        ]
        
        for i, (label, value, color) in enumerate(stats):
            stat_frame = ctk.CTkFrame(stats_frame)
            stat_frame.grid(row=0, column=i, padx=10, pady=10, sticky='ew')
            ctk.CTkLabel(stat_frame, text=label, font=("Arial", 14)).pack()
            ctk.CTkLabel(stat_frame, text=str(value), font=("Arial", 28, "bold"), text_color=color).pack()
        stats_frame.grid_columnconfigure((0, 1, 2), weight=1)

        # --- FILTERS ---
        filter_frame = ctk.CTkFrame(frame)
        filter_frame.pack(fill='x', padx=20, pady=5)
        ctk.CTkLabel(filter_frame, text="Filter by Team:").pack(side='left', padx=(0, 5))
        teams = sorted(list(self.df['Team'].dropna().unique())) if not self.df.empty else []
        self.team_filter_var = ctk.StringVar(value="All")
        team_options = ["All"] + teams
        self.team_filter_box = ctk.CTkComboBox(filter_frame, values=team_options, variable=self.team_filter_var, width=150, command=self.update_dashboard_table)
        self.team_filter_box.pack(side='left', padx=(0, 20))

        # --- ISSUES TABLE ---
        table_frame = ctk.CTkFrame(frame)
        table_frame.pack(fill='both', expand=True, padx=20, pady=10)
        ctk.CTkLabel(table_frame, text="Audit Issues", font=("Arial", 14, "bold")).pack(anchor='w')
        self.dashboard_display_columns = ('ID', 'Description', 'Team', 'Priority', 'Status', 'Resolution_Date')
        self.dashboard_tree = ttk.Treeview(table_frame, columns=self.dashboard_display_columns, show='headings', height=12)
        for col in self.dashboard_display_columns:
            self.dashboard_tree.heading(col, text=col)
            self.dashboard_tree.column(col, width=150)
        self.dashboard_tree.pack(fill='both', expand=True, pady=10)
        self.dashboard_tree.bind('<<TreeviewSelect>>', self.on_dashboard_select)
        self.update_dashboard_table()

        # --- EMAIL PREVIEW PANEL ---
        preview_frame = ctk.CTkFrame(frame)
        preview_frame.pack(fill='x', padx=20, pady=10)
        ctk.CTkLabel(preview_frame, text="Email Preview", font=("Arial", 14, "bold")).pack(anchor='w')
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
        ctk.CTkLabel(table_frame, text="All Issues", font=("Arial", 14, "bold")).pack(anchor='w')
        columns = ('ID', 'Description', 'Team', 'Team_Email', 'Priority', 'Status', 'Created_Date', 'Resolution_Date', 'Reminder_Count')
        self.issues_tree = ttk.Treeview(table_frame, columns=columns, show='headings', height=16)
        for col in columns:
            self.issues_tree.heading(col, text=col)
            self.issues_tree.column(col, width=120)
        self.issues_tree.pack(fill='both', expand=True, pady=10)
        self.update_issues_table()

    def create_email_tab(self):
        frame = self.email_tab
        for widget in frame.winfo_children():
            widget.destroy()
        ctk.CTkLabel(frame, text="Email Template & Sending", font=("Arial", 18, "bold")).pack(pady=20)
        template_frame = ctk.CTkFrame(frame)
        template_frame.pack(fill='both', expand=True, padx=20, pady=10)
        ctk.CTkLabel(template_frame, text="Edit the email template below:").pack(anchor='w')
        self.template_text = ctk.CTkTextbox(template_frame, height=250, width=900, font=("Consolas", 12))
        self.template_text.pack(fill='both', expand=True, pady=10)
        
        # Load current template
        self.load_email_template()
        template_buttons = ctk.CTkFrame(template_frame, fg_color="transparent")
        template_buttons.pack(fill='x', pady=10)
        ctk.CTkButton(template_buttons, text="Save Template", command=self.save_email_template).pack(side='left', padx=5)
        ctk.CTkButton(template_buttons, text="Reset to Default", command=self.reset_email_template).pack(side='left', padx=5)
        sending_frame = ctk.CTkFrame(frame)
        sending_frame.pack(fill='x', padx=20, pady=10)
        issue_frame = ctk.CTkFrame(sending_frame, fg_color="transparent")
        issue_frame.pack(fill='x', pady=10)
        ctk.CTkLabel(issue_frame, text="Select Issue:").pack(side='left')
        self.issue_var = ctk.StringVar()
        self.issue_combo = ctk.CTkComboBox(issue_frame, variable=self.issue_var, width=400, state='readonly')
        self.issue_combo.pack(side='left', padx=10)
        
        # Update issue list
        self.update_issue_combo()
        ctk.CTkButton(sending_frame, text="Send Reminder Email", command=self.send_reminder_email).pack(pady=10)
        bulk_frame = ctk.CTkFrame(sending_frame, fg_color="transparent")
        bulk_frame.pack(fill='x', pady=10)
        ctk.CTkLabel(bulk_frame, text="Bulk Operations:").pack(side='left')
        ctk.CTkButton(bulk_frame, text="Send All Overdue Reminders", command=self.send_all_overdue_reminders).pack(side='left', padx=10)
        ctk.CTkButton(bulk_frame, text="Send Weekly Reminders", command=self.send_weekly_reminders).pack(side='left', padx=10)

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
        reminder_frame = ctk.CTkFrame(frame)
        reminder_frame.pack(fill='x', padx=20, pady=10)
        ctk.CTkLabel(reminder_frame, text="Configure when reminders should be sent:").pack(anchor='w')
        self.reminder_intervals = []
        intervals_data = [
            ("30 days before", 30),
            ("14 days before", 14),
            ("7 days before", 7),
            ("3 days before", 3),
            ("1 day before", 1),
            ("On due date", 0)
        ]
        
        for i, (label, days) in enumerate(intervals_data):
            var = ctk.BooleanVar(value=True)
            self.reminder_intervals.append((var, days))
            ctk.CTkCheckBox(reminder_frame, text=label, variable=var).pack(anchor='w', pady=2)

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
                for var, days_before in self.reminder_intervals:
                    if var.get() and days_until_due == days_before:
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

    def update_dashboard_table(self, *args):
        # Get filter values
        team = self.team_filter_var.get() if hasattr(self, 'team_filter_var') else "All"
        # Filter DataFrame
        df = self.df.copy()
        if team != "All":
            df = df[df['Team'] == team]
        # Update table
        for item in self.dashboard_tree.get_children():
            self.dashboard_tree.delete(item)
        for _, row in df.iterrows():
            values = [row.get(col, '') for col in self.dashboard_display_columns]
            self.dashboard_tree.insert('', 'end', values=values)
        # Clear preview if selection is not in filtered table
        self.selected_issue = None
        self.email_preview_box.delete('1.0', 'end')

    def on_dashboard_select(self, event):
        selected = self.dashboard_tree.selection()
        if not selected:
            self.selected_issue = None
            self.email_preview_box.delete('1.0', 'end')
            return
        item = self.dashboard_tree.item(selected[0])
        values = item['values']
        columns = self.dashboard_display_columns
        issue = {col: values[i] for i, col in enumerate(columns)}
        # Find the full issue row in self.df (to get all fields)
        df_row = self.df[self.df['ID'] == issue['ID']]
        if not df_row.empty:
            issue = df_row.iloc[0].to_dict()
        self.selected_issue = issue
        self.render_email_preview(issue)

    def render_email_preview(self, issue):
        # Load template
        try:
            with open(self.email_template_file, 'r') as f:
                template = f.read()
        except Exception as e:
            self.email_preview_box.delete('1.0', 'end')
            self.email_preview_box.insert('1.0', f"Error loading template: {e}")
            return
        # Calculate days remaining
        try:
            resolution_date = datetime.strptime(str(issue.get('Resolution_Date', '')), '%Y-%m-%d')
            days_remaining = (resolution_date - datetime.now()).days
        except Exception:
            days_remaining = ''
        # Replace template variables
        preview = template
        preview = preview.replace('{{ISSUE_ID}}', str(issue.get('ID', '')))
        preview = preview.replace('{{DESCRIPTION}}', str(issue.get('Description', '')))
        preview = preview.replace('{{PRIORITY}}', str(issue.get('Priority', '')))
        preview = preview.replace('{{STATUS}}', str(issue.get('Status', '')))
        preview = preview.replace('{{RESOLUTION_DATE}}', str(issue.get('Resolution_Date', '')))
        preview = preview.replace('{{DAYS_REMAINING}}', str(days_remaining))
        preview = preview.replace('{{TEAM}}', str(issue.get('Team', '')))
        preview = preview.replace('{{CREATED_DATE}}', str(issue.get('Created_Date', '')))
        preview = preview.replace('{{REMINDER_COUNT}}', str(issue.get('Reminder_Count', '')))
        preview = preview.replace('{{CURRENT_DATE}}', datetime.now().strftime('%Y-%m-%d'))
        # Show as plain text (strip HTML tags for preview)
        import re
        text_preview = re.sub('<[^<]+?>', '', preview)
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
        self.root.clipboard_clear()
        self.root.clipboard_append(self.email_preview_box.get('1.0', 'end').strip())
        messagebox.showinfo("Copied", "Email preview copied to clipboard.")

# Main entry point
if __name__ == "__main__":
    root = ctk.CTk()
    app = AuditManagerApp(root)
    root.mainloop()
    