import customtkinter as ctk
import tkinter.filedialog as fd
import tkinter.messagebox as mbox
import os
import json
import threading
import graph as gh

class OrgChartApp:
    def __init__(self):
        # Global settings variable
        self.app_settings = {'graphviz_path': 'C:/Program Files/Graphviz/bin'}
        
        # Load settings when the form starts
        self.load_settings()
        
        # Initialize the main application
        self.app = ctk.CTk()
        self.app.title('Organizational Chart')
        self.app.geometry('800x600')
        
        ctk.set_appearance_mode('system')
        ctk.set_default_color_theme('blue')
        
        # Track open windows
        self.open_windows = []
        
        # Setup the UI
        self.setup_ui()
        
        # Handle main window close
        self.app.protocol("WM_DELETE_WINDOW", self.on_closing)
    
    def load_settings(self):
        """Load settings from configuration file"""
        try:
            script_dir = os.path.dirname(os.path.abspath(__file__))
            settings_file = os.path.join(script_dir, 'settings.json')
            if os.path.exists(settings_file):
                with open(settings_file, 'r') as f:
                    loaded_settings = json.load(f)
                    self.app_settings.update(loaded_settings)
                    print("Settings loaded:", self.app_settings)
        except Exception as e:
            print(f"Failed to load settings: {e}")
    
    def get_graphviz_path(self):
        """Get the current Graphviz path from settings"""
        return self.app_settings.get('graphviz_path')
    
    def setup_ui(self):
        """Setup the main user interface"""
        # Create a frame for the title and settings button
        title_frame = ctk.CTkFrame(self.app)
        title_frame.pack(pady=20, fill="x", padx=20)
        
        # Title on the left
        title_label = ctk.CTkLabel(title_frame, text='Organizational Chart', font=("Arial", 20, "bold"))
        title_label.pack(side="left", pady=10)
        
        # Settings button on the right
        settings_button = ctk.CTkButton(title_frame, text="⚙️", width=40, height=40, command=self.open_settings)
        settings_button.pack(side="right", pady=10, padx=10)
        
        # File selection
        self.file_path_var = ctk.StringVar()
        
        file_frame = ctk.CTkFrame(self.app)
        file_frame.pack(pady=10)
        file_label = ctk.CTkLabel(file_frame, text="Excel File:")
        file_label.pack(side="left", padx=5)
        file_entry = ctk.CTkEntry(file_frame, textvariable=self.file_path_var, width=400)
        file_entry.pack(side="left", padx=5)
        file_button = ctk.CTkButton(file_frame, text="Browse", command=self.select_file)
        file_button.pack(side="left", padx=5)
        
        # Note for file selection
        self.file_note = ctk.CTkLabel(self.app, text="Ensure selected excel file is not open", 
                                     text_color="orange", font=("Arial", 10))
        self.file_note.pack(pady=2)
        
        # Initially hide the file note
        self.file_note.pack_forget()
        
        # Manager and Month/Year fields
        manager_name_label = ctk.CTkLabel(self.app, text='Manager Name')
        manager_name_label.pack(pady=5)
        self.manager_name_entry = ctk.CTkEntry(self.app)
        self.manager_name_entry.pack(pady=5)
        
        month_year_label = ctk.CTkLabel(self.app, text='Month and Year')
        month_year_label.pack(pady=5)
        self.month_year_entry = ctk.CTkEntry(self.app)
        self.month_year_entry.pack(pady=5)
        
        # Note for Past Data users
        self.past_data_note = ctk.CTkLabel(self.app, text="Ensure Month and Year input matches GOBS Monthly HC File sheet name", 
                                          text_color="gray", font=("Arial", 10))
        self.past_data_note.pack(pady=2)
        
        # Initially hide the note
        self.past_data_note.pack_forget()
        
        # Checkboxes for options
        self.show_location_var = ctk.BooleanVar()
        self.show_level_var = ctk.BooleanVar()
        self.past_data_var = ctk.BooleanVar()
        checkbox_frame = ctk.CTkFrame(self.app)
        checkbox_frame.pack(pady=10)
        show_location_cb = ctk.CTkCheckBox(checkbox_frame, text="Show Location", variable=self.show_location_var)
        show_location_cb.pack(side="left", padx=10)
        show_level_cb = ctk.CTkCheckBox(checkbox_frame, text="Show Level", variable=self.show_level_var)
        show_level_cb.pack(side="left", padx=10)
        past_data_cb = ctk.CTkCheckBox(checkbox_frame, text="Past Data", variable=self.past_data_var)
        past_data_cb.pack(side="left", padx=10)
        
        # Bind the Past Data checkbox to the change function
        past_data_cb.configure(command=self.on_past_data_change)
        
        # Generate button
        button = ctk.CTkButton(self.app, text='Generate Graph', command=self.button_function)
        button.pack(pady=20)
    
    def select_file(self):
        """Select Excel file and show/hide file note"""
        file_path = fd.askopenfilename(
            filetypes=[("Excel files", "*.xlsx")],
            title="Select Excel File"
        )
        if file_path:
            self.file_path_var.set(file_path)
            # Show the file note when a file is selected
            self.file_note.pack(pady=2)
        else:
            # Hide the file note when no file is selected
            self.file_note.pack_forget()
    
    def on_past_data_change(self):
        """Enable/disable Show Location checkbox based on Past Data state"""
        if self.past_data_var.get():
            # Find the checkbox widgets and disable them
            for widget in self.app.winfo_children():
                if isinstance(widget, ctk.CTkFrame):
                    for child in widget.winfo_children():
                        if isinstance(child, ctk.CTkCheckBox) and "Location" in child.cget("text"):
                            child.configure(state="disabled", text_color="gray", fg_color="gray")
                            break
            
            self.past_data_note.pack(pady=2)  # Show the note
            # Clear the file input since it's not needed for past data
            self.file_path_var.set("")
            # Hide the file note since no file is selected
            self.file_note.pack_forget()
        else:
            # Find the checkbox widgets and enable them
            for widget in self.app.winfo_children():
                if isinstance(widget, ctk.CTkFrame):
                    for child in widget.winfo_children():
                        if isinstance(child, ctk.CTkCheckBox) and "Location" in child.cget("text"):
                            child.configure(state="normal", text_color="white", fg_color="blue")
                            break
            
            self.past_data_note.pack_forget()  # Hide the note
            # Show file note if a file is currently selected
            if self.file_path_var.get():
                self.file_note.pack(pady=2)
    
    def open_settings(self):
        """Open settings dialog for additional configuration options"""
        # Check if settings window is already open
        for window in self.open_windows:
            if hasattr(window, 'title') and window.title() == "Settings":
                window.lift()  # Bring to front
                return
        
        settings_window = ctk.CTkToplevel(self.app)
        settings_window.title("Settings")
        settings_window.geometry("500x600")
        settings_window.resizable(False, False)
        settings_window.transient(self.app)
        settings_window.grab_set()
        
        # Add to open windows list
        self.open_windows.append(settings_window)
        
        # Center the settings window
        settings_window.update_idletasks()
        x = (settings_window.winfo_screenwidth() // 2) - (500 // 2)
        y = (settings_window.winfo_screenheight() // 2) - (600 // 2)
        settings_window.geometry(f"500x600+{x}+{y}")
        
        # Settings title
        settings_label = ctk.CTkLabel(settings_window, text="Settings", font=("Arial", 20, "bold"))
        settings_label.pack(pady=20)
        
        # Create scrollable frame for settings
        settings_frame = ctk.CTkScrollableFrame(settings_window, width=450, height=500)
        settings_frame.pack(pady=10, padx=20, fill="both", expand=True)
        
        # Graph Generation Settings Section
        graph_settings_label = ctk.CTkLabel(settings_frame, text="Graph Generation Settings", font=("Arial", 16, "bold"))
        graph_settings_label.pack(pady=10, anchor="w")
        
        # Graphviz path setting
        graphviz_frame = ctk.CTkFrame(settings_frame)
        graphviz_frame.pack(pady=10, fill="x")
        
        graphviz_label = ctk.CTkLabel(graphviz_frame, text="Graphviz Installation Path:", font=("Arial", 14, "bold"))
        graphviz_label.pack(pady=5, anchor="w")
        
        graphviz_path_frame = ctk.CTkFrame(graphviz_frame)
        graphviz_path_frame.pack(pady=5, fill="x")
        
        graphviz_path_var = ctk.StringVar(value=self.app_settings['graphviz_path'])
        graphviz_path_entry = ctk.CTkEntry(graphviz_path_frame, textvariable=graphviz_path_var, width=300)
        graphviz_path_entry.pack(side="left", padx=5, fill="x", expand=True)
        
        def browse_graphviz_path():
            """Browse for Graphviz installation directory"""
            path = fd.askdirectory(title="Select Graphviz Installation Directory")
            if path:
                graphviz_path_var.set(path)
        
        graphviz_browse_button = ctk.CTkButton(graphviz_path_frame, text="Browse", command=browse_graphviz_path, width=80)
        graphviz_browse_button.pack(side="right", padx=5)
        
        # Help text for Graphviz path
        graphviz_help = ctk.CTkLabel(graphviz_frame, text="Path to the directory containing 'dot.exe' (e.g., C:/Program Files/Graphviz/bin)", 
                                     text_color="gray", font=("Arial", 10))
        graphviz_help.pack(pady=2, anchor="w")
        
        # Output Settings Section
        output_settings_label = ctk.CTkLabel(settings_frame, text="Output Settings", font=("Arial", 16, "bold"))
        output_settings_label.pack(pady=10, anchor="w")
        
        output_frame = ctk.CTkFrame(settings_frame)
        output_frame.pack(pady=10, fill="x")
        
        # Buttons
        button_frame = ctk.CTkFrame(settings_window)
        button_frame.pack(pady=20)
        
        def save_settings():
            """Save settings to a configuration file"""
            settings = {
                'graphviz_path': graphviz_path_var.get()
            }
            
            # Update global settings
            self.app_settings.update(settings)
            
            # Save to JSON file
            try:
                script_dir = os.path.dirname(os.path.abspath(__file__))
                settings_file = os.path.join(script_dir, 'settings.json')
                with open(settings_file, 'w') as f:
                    json.dump(settings, f, indent=2)
                mbox.showinfo("Success", "Settings saved successfully!")
            except Exception as e:
                mbox.showerror("Error", f"Failed to save settings: {e}")
        
        def reset_settings():
            """Reset settings to defaults"""
            graphviz_path_var.set(self.app_settings['graphviz_path'])
        
        def close_settings():
            """Properly close the settings window"""
            try:
                if settings_window in self.open_windows:
                    self.open_windows.remove(settings_window)
                settings_window.grab_release()
                settings_window.destroy()
            except:
                pass
        
        save_button = ctk.CTkButton(button_frame, text="Save Settings", command=save_settings)
        save_button.pack(side="left", padx=10)
        
        reset_button = ctk.CTkButton(button_frame, text="Reset to Defaults", command=reset_settings)
        reset_button.pack(side="left", padx=10)
        
        close_button = ctk.CTkButton(button_frame, text="Close", command=close_settings)
        close_button.pack(side="left", padx=10)
        
        # Handle window close button (X)
        settings_window.protocol("WM_DELETE_WINDOW", close_settings)
    
    def show_login_dialog(self):
        """Show login dialog for past data access"""
        # Check if login window is already open
        for window in self.open_windows:
            if hasattr(window, 'title') and window.title() == "Login Required":
                window.lift()  # Bring to front
                return
        
        login_window = ctk.CTkToplevel(self.app)
        login_window.title("Login Required")
        login_window.geometry("400x300")
        login_window.resizable(False, False)
        login_window.transient(self.app)
        login_window.grab_set()
        
        # Add to open windows list
        self.open_windows.append(login_window)
        
        # Center the login window
        login_window.update_idletasks()
        x = (login_window.winfo_screenwidth() // 2) - (400 // 2)
        y = (login_window.winfo_screenheight() // 2) - (300 // 2)
        login_window.geometry(f"400x300+{x}+{y}")
        
        # Login form
        login_label = ctk.CTkLabel(login_window, text="Login Required for Past Data", font=("Arial", 16, "bold"))
        login_label.pack(pady=20)
        
        email_frame = ctk.CTkFrame(login_window)
        email_frame.pack(pady=10, padx=20, fill="x")
        email_label = ctk.CTkLabel(email_frame, text="Email:")
        email_label.pack(side="left", padx=5)
        email_entry = ctk.CTkEntry(email_frame, width=250)
        email_entry.pack(side="left", padx=5, fill="x", expand=True)
        
        password_frame = ctk.CTkFrame(login_window)
        password_frame.pack(pady=10, padx=20, fill="x")
        password_label = ctk.CTkLabel(password_frame, text="Password:")
        password_label.pack(side="left", padx=5)
        password_entry = ctk.CTkEntry(password_frame, show="*", width=250)
        password_entry.pack(side="left", padx=5, fill="x", expand=True)
        
        def validate_login():
            email = email_entry.get().strip()
            password = password_entry.get().strip()
            
            if not email or not password:
                mbox.showerror("Error", "Please enter both email and password.")
                return
            
            # Here you can add actual authentication logic
            # For now, we'll just check if both fields are filled
            if email and password:
                mbox.showinfo("Success", "Login successful! Proceeding with past data generation.")
                close_login()
                # Continue with the original generation logic, passing credentials
                self.generate_with_past_data(email, password)
            else:
                mbox.showerror("Error", "Invalid credentials. Please try again.")
        
        def close_login():
            """Properly close the login window"""
            try:
                if login_window in self.open_windows:
                    self.open_windows.remove(login_window)
                login_window.grab_release()
                login_window.destroy()
            except:
                pass
        
        # Buttons
        button_frame = ctk.CTkFrame(login_window)
        button_frame.pack(pady=20)
        
        login_button = ctk.CTkButton(button_frame, text="Login", command=validate_login)
        login_button.pack(side="left", padx=10)
        
        cancel_button = ctk.CTkButton(button_frame, text="Cancel", command=close_login)
        cancel_button.pack(side="left", padx=10)
        
        # Focus on email entry
        email_entry.focus()
        
        # Bind Enter key to login
        login_window.bind('<Return>', lambda event: validate_login())
        
        # Handle window close button (X)
        login_window.protocol("WM_DELETE_WINDOW", close_login)
    
    def generate_with_past_data(self, email, password):
        """Generate org chart with past data after successful login"""
        file_path = self.file_path_var.get()
        manager = self.manager_name_entry.get()
        month_year = self.month_year_entry.get()
        show_location = self.show_location_var.get()
        show_level = self.show_level_var.get()
        
        if not file_path or not os.path.exists(file_path):
            mbox.showerror("Error", "Please select a valid Excel file.")
            return
        if not manager or not month_year:
            mbox.showerror("Error", "Please enter both Manager Name and Month/Year.")
            return
        
        try:
            # Pass the selected file_path and credentials to load_data
            datasheet = gh.load_data(manager, month_year, file_path=file_path, email=email, password=password)
            df = gh.save_df(datasheet, manager, month_year)
            gh.generateGraph(df, manager, month_year, show_location=show_location, show_level=show_level)
            mbox.showinfo("Success", "Org chart with past data generated successfully!")
        except Exception as e:
            mbox.showerror("Error", f"Failed to generate org chart:\n{e}")
    
    def button_function(self):
        """Main button function for generating org chart"""
        file_path = self.file_path_var.get()
        manager = self.manager_name_entry.get()
        month_year = self.month_year_entry.get()
        show_location = self.show_location_var.get()
        show_level = self.show_level_var.get()
        
        # Check if past data is selected
        if self.past_data_var.get():
            # For past data, file is not required
            if not manager or not month_year:
                mbox.showerror("Error", "Please enter both Manager Name and Month/Year.")
                return
            self.show_login_dialog()
        else:
            # For regular generation, file is required
            if not file_path or not os.path.exists(file_path):
                mbox.showerror("Error", "Please select a valid Excel file.")
                return
            if not manager or not month_year:
                mbox.showerror("Error", "Please enter both Manager Name and Month/Year.")
                return
            
            try:
                # Pass the selected file_path to load_data
                datasheet = gh.load_data(manager, month_year, file_path=file_path)
                df = gh.save_df(datasheet, manager, month_year)
                gh.generateGraph(df, manager, month_year, show_location=show_location, show_level=show_level)
                mbox.showinfo("Success", "Org chart generated successfully!")
            except Exception as e:
                mbox.showerror("Error", f"Failed to generate org chart:\n{e}")
    
    def close_all_windows(self):
        """Close all open windows properly"""
        for window in self.open_windows[:]:  # Copy list to avoid modification during iteration
            try:
                if hasattr(window, 'grab_release'):
                    window.grab_release()
                window.destroy()
            except:
                pass
        self.open_windows.clear()
    
    def on_closing(self):
        """Handle application closing properly"""
        try:
            # Close all open windows first
            self.close_all_windows()
            
            # Quit the main application
            self.app.quit()
            self.app.destroy()
        except Exception as e:
            print(f"Error during cleanup: {e}")
            # Force quit if cleanup fails
            try:
                self.app.quit()
            except:
                pass
    
    def run(self):
        """Start the application"""
        try:
            self.app.mainloop()
        except Exception as e:
            print(f"Application error: {e}")
        finally:
            # Ensure cleanup happens
            self.close_all_windows()

# Create and run the application
if __name__ == "__main__":
    app = OrgChartApp()
    app.run()


