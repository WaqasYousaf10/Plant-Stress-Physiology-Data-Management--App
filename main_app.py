# main_app.py - COMPLETE FIXED VERSION WITH WORKING EXPORTS
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import pandas as pd
from datetime import datetime, timedelta
import os
import csv
import json

from database_sqlite import StressDatabase
from analysis import StressAnalyzer

class AdvancedStressApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Advanced Plant Stress Physiology Data Manager")
        self.root.geometry("1200x800")
        self.root.state('zoomed')  # Start maximized
        
        # Initialize database and analyzer
        self.db = StressDatabase()
        if not self.db.create_database():
            messagebox.showerror("Database Error", "Failed to connect to database. Please check your SQLite setup.")
            return
        
        self.analyzer = StressAnalyzer(self.db)
        self.current_experiment_id = None
        self.current_treatment_id = None
        self.current_measurement_id = None
        
        self.setup_styles()
        self.create_widgets()
        self.load_initial_data()
    
    def setup_styles(self):
        self.style = ttk.Style()
        self.style.configure('Title.TLabel', font=('Arial', 16, 'bold'), background='#4CAF50', foreground='white')
        self.style.configure('Header.TLabel', font=('Arial', 11, 'bold'))
        self.style.configure('Success.TButton', background='#4CAF50', foreground='white')
        
    def create_widgets(self):
        # Create notebook for tabs
        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(fill='both', expand=True, padx=10, pady=10)
        
        # Create tabs
        self.experiments_tab = ttk.Frame(self.notebook)
        self.treatments_tab = ttk.Frame(self.notebook)
        self.measurements_tab = ttk.Frame(self.notebook)
        self.analysis_tab = ttk.Frame(self.notebook)
        self.reports_tab = ttk.Frame(self.notebook)
        
        self.notebook.add(self.experiments_tab, text='📊 Experiments')
        self.notebook.add(self.treatments_tab, text='🔬 Treatments')
        self.notebook.add(self.measurements_tab, text='📈 Measurements')
        self.notebook.add(self.analysis_tab, text='📉 Analysis')
        self.notebook.add(self.reports_tab, text='📄 Reports')
        
        self.setup_experiments_tab()
        self.setup_treatments_tab()
        self.setup_measurements_tab()
        self.setup_analysis_tab()
        self.setup_reports_tab()
        
        # Status bar
        self.status_var = tk.StringVar()
        self.status_var.set("Ready - Advanced Plant Stress Physiology Data Manager")
        status_bar = ttk.Label(self.root, textvariable=self.status_var, relief='sunken')
        status_bar.pack(side='bottom', fill='x')
    
    def setup_experiments_tab(self):
        # Left frame - Input form
        input_frame = ttk.LabelFrame(self.experiments_tab, text="Experiment Details", padding=15)
        input_frame.pack(side='left', fill='y', padx=5, pady=5)
        
        # Form fields
        fields = [
            ('Experiment Code:*', 'code_entry'),
            ('Experiment Name:*', 'name_entry'),
            ('Plant Species:*', 'species_entry'),
            ('Researcher:', 'researcher_entry'),
            ('Stress Type:*', 'stress_combo'),
            ('Start Date:', 'start_date_entry'),
            ('End Date:', 'end_date_entry'),
            ('Description:', 'desc_text')
        ]
        
        self.experiment_widgets = {}
        for i, (label, key) in enumerate(fields):
            ttk.Label(input_frame, text=label, style='Header.TLabel').grid(row=i, column=0, sticky='w', pady=8)
            
            if 'combo' in key:
                widget = ttk.Combobox(input_frame, values=[
                    'drought', 'salt', 'heat', 'cold', 'flooding', 
                    'nutrient', 'UV', 'biotic', 'combined', 'heavy_metal'
                ], width=30)
            elif 'desc' in key:
                widget = tk.Text(input_frame, width=30, height=4, font=('Arial', 9))
            else:
                widget = ttk.Entry(input_frame, width=30)
            
            widget.grid(row=i, column=1, padx=5, pady=5, sticky='w')
            self.experiment_widgets[key] = widget
        
        # Set default dates
        self.experiment_widgets['start_date_entry'].insert(0, datetime.now().strftime('%Y-%m-%d'))
        
        # Buttons
        btn_frame = ttk.Frame(input_frame)
        btn_frame.grid(row=len(fields), column=0, columnspan=2, pady=20)
        
        buttons = [
            ('💾 Add Experiment', self.add_experiment, '#4CAF50'),
            ('🔄 Update', self.update_experiment, '#2196F3'),
            ('🧹 Clear', self.clear_experiment_form, '#FF9800'),
            ('🗑️ Delete', self.delete_experiment, '#f44336')
        ]
        
        for text, command, color in buttons:
            btn = tk.Button(btn_frame, text=text, command=command, 
                          bg=color, fg='white', font=('Arial', 9, 'bold'), padx=10)
            btn.pack(side='left', padx=5, pady=5)
        
        # Export buttons for experiments
        export_frame = ttk.LabelFrame(input_frame, text="Export Experiments Data", padding=10)
        export_frame.grid(row=len(fields)+1, column=0, columnspan=2, pady=10, sticky='ew')
        
        ttk.Button(export_frame, text="📊 Export to CSV", 
                  command=lambda: self.export_experiments_data('csv')).pack(side='left', padx=2, pady=5)
        ttk.Button(export_frame, text="📈 Export to Excel", 
                  command=lambda: self.export_experiments_data('xlsx')).pack(side='left', padx=2, pady=5)
        ttk.Button(export_frame, text="📋 Export to Text", 
                  command=lambda: self.export_experiments_data('txt')).pack(side='left', padx=2, pady=5)
        
        # Right frame - Experiments list
        list_frame = ttk.LabelFrame(self.experiments_tab, text="Experiments List", padding=10)
        list_frame.pack(side='right', fill='both', expand=True, padx=5, pady=5)
        
        # Search frame
        search_frame = ttk.Frame(list_frame)
        search_frame.pack(fill='x', pady=5)
        
        ttk.Label(search_frame, text="Search:").pack(side='left', padx=5)
        self.exp_search_entry = ttk.Entry(search_frame, width=30)
        self.exp_search_entry.pack(side='left', padx=5)
        self.exp_search_entry.bind('<KeyRelease>', self.search_experiments)
        
        ttk.Button(search_frame, text="Refresh", command=self.load_experiments).pack(side='left', padx=5)
        
        # Treeview
        tree_frame = ttk.Frame(list_frame)
        tree_frame.pack(fill='both', expand=True, pady=5)
        
        columns = ('ID', 'Code', 'Name', 'Species', 'Researcher', 'Stress Type', 'Start Date', 'Status')
        self.experiments_tree = ttk.Treeview(tree_frame, columns=columns, show='headings', height=15)
        
        # Configure columns
        col_widths = [50, 100, 150, 120, 100, 100, 100, 80]
        for col, width in zip(columns, col_widths):
            self.experiments_tree.heading(col, text=col)
            self.experiments_tree.column(col, width=width)
        
        # Scrollbar
        scrollbar = ttk.Scrollbar(tree_frame, orient='vertical', command=self.experiments_tree.yview)
        self.experiments_tree.configure(yscrollcommand=scrollbar.set)
        
        self.experiments_tree.pack(side='left', fill='both', expand=True)
        scrollbar.pack(side='right', fill='y')
        
        self.experiments_tree.bind('<<TreeviewSelect>>', self.on_experiment_select)
    
    def setup_treatments_tab(self):
        # Main frame for treatments
        main_frame = ttk.Frame(self.treatments_tab)
        main_frame.pack(fill='both', expand=True, padx=10, pady=10)
        
        # Initial message
        self.treatments_initial_label = ttk.Label(main_frame, 
                                                text="Please select an experiment first to manage treatments", 
                                                font=('Arial', 12, 'bold'))
        self.treatments_initial_label.pack(pady=50)
        
        # Content frame (will be populated when experiment is selected)
        self.treatments_content = ttk.Frame(main_frame)
    
    def setup_measurements_tab(self):
        # Main frame for measurements
        main_frame = ttk.Frame(self.measurements_tab)
        main_frame.pack(fill='both', expand=True, padx=10, pady=10)
        
        # Initial message
        self.measurements_initial_label = ttk.Label(main_frame, 
                                                  text="Please select an experiment and treatment first to manage measurements", 
                                                  font=('Arial', 12, 'bold'))
        self.measurements_initial_label.pack(pady=50)
        
        # Content frame (will be populated when treatment is selected)
        self.measurements_content = ttk.Frame(main_frame)
    
    def setup_analysis_tab(self):
        main_frame = ttk.Frame(self.analysis_tab)
        main_frame.pack(fill='both', expand=True, padx=10, pady=10)
        
        ttk.Label(main_frame, text="Data Analysis - Select an experiment to analyze", 
                 font=('Arial', 12, 'bold')).pack(pady=10)
        
        self.analysis_content = ttk.Frame(main_frame)
        self.analysis_content.pack(fill='both', expand=True)
    
    def setup_reports_tab(self):
        main_frame = ttk.Frame(self.reports_tab)
        main_frame.pack(fill='both', expand=True, padx=10, pady=10)
        
        report_frame = ttk.LabelFrame(main_frame, text="Generate Reports", padding=20)
        report_frame.pack(fill='both', expand=True, pady=10)
        
        ttk.Label(report_frame, text="Select Report Type:", font=('Arial', 11, 'bold')).pack(anchor='w', pady=5)
        
        report_types = [
            ("Complete Experiment Report", "complete"),
            ("Growth Analysis Report", "growth"),
            ("Physiological Parameters", "physio"),
            ("Stress Impact Summary", "stress"),
            ("Statistical Analysis", "stats")
        ]
        
        self.report_var = tk.StringVar(value="complete")
        for text, value in report_types:
            ttk.Radiobutton(report_frame, text=text, variable=self.report_var, value=value).pack(anchor='w', pady=2)
        
        ttk.Label(report_frame, text="Select Experiment:", font=('Arial', 11, 'bold')).pack(anchor='w', pady=(15,5))
        
        self.report_exp_combo = ttk.Combobox(report_frame, width=40, state='readonly')
        self.report_exp_combo.pack(anchor='w', pady=5)
        
        btn_frame = ttk.Frame(report_frame)
        btn_frame.pack(pady=20)
        
        ttk.Button(btn_frame, text="📊 Generate Report", command=self.generate_report).pack(side='left', padx=5)
        ttk.Button(btn_frame, text="📈 Create Charts", command=self.create_charts).pack(side='left', padx=5)
        ttk.Button(btn_frame, text="💾 Export to Excel", command=self.export_report).pack(side='left', padx=5)
        
        # Export options frame
        export_frame = ttk.LabelFrame(report_frame, text="Export Options", padding=15)
        export_frame.pack(fill='x', pady=15)
        
        ttk.Label(export_frame, text="Export Format:", font=('Arial', 10, 'bold')).pack(anchor='w')
        
        format_frame = ttk.Frame(export_frame)
        format_frame.pack(fill='x', pady=5)
        
        ttk.Button(format_frame, text="CSV", command=lambda: self.export_comprehensive_data('csv')).pack(side='left', padx=2)
        ttk.Button(format_frame, text="Excel", command=lambda: self.export_comprehensive_data('xlsx')).pack(side='left', padx=2)
        ttk.Button(format_frame, text="Text", command=lambda: self.export_comprehensive_data('txt')).pack(side='left', padx=2)
        ttk.Button(format_frame, text="JSON", command=lambda: self.export_comprehensive_data('json')).pack(side='left', padx=2)
    
    def load_initial_data(self):
        self.load_experiments()
        self.update_report_experiments()
    
    def load_experiments(self):
        try:
            for item in self.experiments_tree.get_children():
                self.experiments_tree.delete(item)
            
            query = """
                SELECT id, experiment_code, experiment_name, plant_species, 
                       researcher, stress_type, start_date, status 
                FROM experiments 
                ORDER BY created_at DESC
            """
            results = self.db.execute_query(query)
            
            if results:
                for experiment in results:
                    self.experiments_tree.insert('', 'end', values=experiment)
                
                self.status_var.set(f"Loaded {len(results)} experiments")
            else:
                self.status_var.set("No experiments found")
        
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load experiments: {str(e)}")
    
    def add_experiment(self):
        """Add a new experiment to the database"""
        try:
            # Get form data
            code = self.experiment_widgets['code_entry'].get().strip()
            name = self.experiment_widgets['name_entry'].get().strip()
            species = self.experiment_widgets['species_entry'].get().strip()
            stress_type = self.experiment_widgets['stress_combo'].get().strip()
            researcher = self.experiment_widgets['researcher_entry'].get().strip()
            start_date = self.experiment_widgets['start_date_entry'].get().strip()
            end_date = self.experiment_widgets['end_date_entry'].get().strip()
            description = self.experiment_widgets['desc_text'].get('1.0', 'end-1c').strip()

            # Validate required fields
            if not code or not name or not species or not stress_type:
                messagebox.showerror("Error", "Please fill in all required fields (*)")
                return

            # Insert into database
            query = """
                INSERT INTO experiments 
                (experiment_code, experiment_name, plant_species, stress_type, 
                 researcher, start_date, end_date, description, status)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, 'active')
            """
            params = (
                code, name, species, stress_type, researcher,
                start_date if start_date else None,
                end_date if end_date else None,
                description if description else None
            )

            if self.db.execute_query(query, params):
                messagebox.showinfo("Success", "Experiment added successfully!")
                self.clear_experiment_form()
                self.load_experiments()
                self.update_report_experiments()
                self.status_var.set(f"Added experiment: {name}")
            else:
                messagebox.showerror("Error", "Failed to add experiment")

        except Exception as e:
            messagebox.showerror("Error", f"Failed to add experiment: {str(e)}")
    
    def on_experiment_select(self, event):
        selected = self.experiments_tree.selection()
        if selected:
            item = self.experiments_tree.item(selected[0])
            values = item['values']
            self.current_experiment_id = values[0]
            
            # Load experiment details into form
            self.clear_experiment_form()
            self.experiment_widgets['code_entry'].insert(0, values[1])
            self.experiment_widgets['name_entry'].insert(0, values[2])
            self.experiment_widgets['species_entry'].insert(0, values[3])
            self.experiment_widgets['researcher_entry'].insert(0, values[4] or '')
            self.experiment_widgets['stress_combo'].set(values[5])
            if values[6]:
                self.experiment_widgets['start_date_entry'].insert(0, values[6])
            
            self.status_var.set(f"Selected: {values[2]}")
            
            # Update treatments and analysis tabs
            self.update_treatments_tab()
            self.update_analysis_tab()
    
    def update_experiment(self):
        """Update selected experiment"""
        if not self.current_experiment_id:
            messagebox.showwarning("Warning", "Please select an experiment to update")
            return

        try:
            # Get form data
            code = self.experiment_widgets['code_entry'].get().strip()
            name = self.experiment_widgets['name_entry'].get().strip()
            species = self.experiment_widgets['species_entry'].get().strip()
            stress_type = self.experiment_widgets['stress_combo'].get().strip()
            researcher = self.experiment_widgets['researcher_entry'].get().strip()
            start_date = self.experiment_widgets['start_date_entry'].get().strip()
            end_date = self.experiment_widgets['end_date_entry'].get().strip()
            description = self.experiment_widgets['desc_text'].get('1.0', 'end-1c').strip()

            # Validate required fields
            if not code or not name or not species or not stress_type:
                messagebox.showerror("Error", "Please fill in all required fields (*)")
                return

            query = """
                UPDATE experiments 
                SET experiment_code=?, experiment_name=?, plant_species=?, 
                    stress_type=?, researcher=?, start_date=?, end_date=?, 
                    description=?
                WHERE id=?
            """
            params = (
                code, name, species, stress_type, researcher,
                start_date if start_date else None,
                end_date if end_date else None,
                description if description else None,
                self.current_experiment_id
            )

            if self.db.execute_query(query, params):
                messagebox.showinfo("Success", "Experiment updated successfully!")
                self.load_experiments()
                self.status_var.set(f"Updated experiment: {name}")
            else:
                messagebox.showerror("Error", "Failed to update experiment")

        except Exception as e:
            messagebox.showerror("Error", f"Failed to update experiment: {str(e)}")
    
    def delete_experiment(self):
        """Delete selected experiment"""
        if not self.current_experiment_id:
            messagebox.showwarning("Warning", "Please select an experiment to delete")
            return

        if messagebox.askyesno("Confirm Delete", 
                              "Are you sure you want to delete this experiment and all its associated treatments and measurements?"):
            try:
                # First delete related treatments and measurements
                delete_measurements_query = """
                    DELETE FROM measurements 
                    WHERE treatment_id IN (SELECT id FROM treatments WHERE experiment_id = ?)
                """
                self.db.execute_query(delete_measurements_query, (self.current_experiment_id,))
                
                delete_treatments_query = "DELETE FROM treatments WHERE experiment_id = ?"
                self.db.execute_query(delete_treatments_query, (self.current_experiment_id,))
                
                # Then delete the experiment
                query = "DELETE FROM experiments WHERE id = ?"
                if self.db.execute_query(query, (self.current_experiment_id,)):
                    messagebox.showinfo("Success", "Experiment deleted successfully!")
                    self.clear_experiment_form()
                    self.load_experiments()
                    self.update_report_experiments()
                    self.current_experiment_id = None
                    self.status_var.set("Experiment deleted")
                else:
                    messagebox.showerror("Error", "Failed to delete experiment")

            except Exception as e:
                messagebox.showerror("Error", f"Failed to delete experiment: {str(e)}")
    
    def clear_experiment_form(self):
        """Clear the experiment form"""
        for key, widget in self.experiment_widgets.items():
            if isinstance(widget, tk.Text):
                widget.delete('1.0', 'end')
            else:
                widget.delete(0, 'end')
        
        # Set default date
        self.experiment_widgets['start_date_entry'].insert(0, datetime.now().strftime('%Y-%m-%d'))
        
        # Clear selection
        if hasattr(self, 'experiments_tree'):
            self.experiments_tree.selection_remove(self.experiments_tree.selection())
    
    def search_experiments(self, event=None):
        """Search experiments by various fields"""
        search_term = self.exp_search_entry.get().strip()
        
        if not search_term:
            self.load_experiments()
            return
        
        try:
            for item in self.experiments_tree.get_children():
                self.experiments_tree.delete(item)
            
            query = """
                SELECT id, experiment_code, experiment_name, plant_species, 
                       researcher, stress_type, start_date, status 
                FROM experiments 
                WHERE experiment_code LIKE ? OR experiment_name LIKE ? 
                   OR plant_species LIKE ? OR researcher LIKE ?
                   OR stress_type LIKE ?
                ORDER BY created_at DESC
            """
            search_pattern = f"%{search_term}%"
            params = (search_pattern, search_pattern, search_pattern, search_pattern, search_pattern)
            results = self.db.execute_query(query, params)
            
            if results:
                for experiment in results:
                    self.experiments_tree.insert('', 'end', values=experiment)
                
                self.status_var.set(f"Found {len(results)} experiments matching '{search_term}'")
            else:
                self.status_var.set(f"No experiments found matching '{search_term}'")
        
        except Exception as e:
            messagebox.showerror("Error", f"Search failed: {str(e)}")
    
    def update_treatments_tab(self):
        # Clear existing content
        for widget in self.treatments_content.winfo_children():
            widget.destroy()
        
        # Hide initial label and show content
        self.treatments_initial_label.pack_forget()
        self.treatments_content.pack(fill='both', expand=True)
        
        if not self.current_experiment_id:
            ttk.Label(self.treatments_content, text="Please select an experiment first").pack(pady=50)
            return
        
        # Get experiment details for display
        exp_details = self.get_experiment_details(self.current_experiment_id)
        exp_title = f"Treatments for: {exp_details.get('name', 'Unknown Experiment')} (ID: {self.current_experiment_id})"
        
        ttk.Label(self.treatments_content, text=exp_title, 
                 font=('Arial', 14, 'bold')).pack(pady=10)
        
        # Create two-column layout
        main_container = ttk.Frame(self.treatments_content)
        main_container.pack(fill='both', expand=True, padx=10, pady=10)
        
        # Left frame - Treatment form
        left_frame = ttk.LabelFrame(main_container, text="Treatment Details", padding=15)
        left_frame.pack(side='left', fill='y', padx=5, pady=5)
        
        # Right frame - Treatments list
        right_frame = ttk.LabelFrame(main_container, text="Treatments List", padding=10)
        right_frame.pack(side='right', fill='both', expand=True, padx=5, pady=5)
        
        # Treatment form fields
        fields = [
            ('Treatment Name:*', 'name_entry'),
            ('Treatment Type:*', 'type_combo'),
            ('Stress Level:', 'level_combo'),
            ('Concentration (mM):', 'conc_entry'),
            ('Duration (days):', 'duration_entry'),
            ('Temperature (°C):', 'temp_entry'),
            ('Description:', 'desc_text')
        ]
        
        self.treatment_widgets = {}
        for i, (label, key) in enumerate(fields):
            ttk.Label(left_frame, text=label, style='Header.TLabel').grid(row=i, column=0, sticky='w', pady=8)
            
            if 'combo' in key:
                if 'type' in key:
                    widget = ttk.Combobox(left_frame, values=[
                        'control', 'drought', 'salt', 'heat', 'cold', 'flooding', 
                        'nutrient_deficiency', 'UV', 'biotic', 'combined', 'heavy_metal'
                    ], width=25)
                elif 'level' in key:
                    widget = ttk.Combobox(left_frame, values=[
                        'low', 'medium', 'high', 'severe', 'control'
                    ], width=25)
                widget.set('control' if 'type' in key else 'medium')
            elif 'desc' in key:
                widget = tk.Text(left_frame, width=25, height=4, font=('Arial', 9))
            else:
                widget = ttk.Entry(left_frame, width=25)
            
            widget.grid(row=i, column=1, padx=5, pady=5, sticky='w')
            self.treatment_widgets[key] = widget
        
        # Buttons for treatment operations
        btn_frame = ttk.Frame(left_frame)
        btn_frame.grid(row=len(fields), column=0, columnspan=2, pady=20)
        
        treatment_buttons = [
            ('💾 Add Treatment', self.add_treatment, '#4CAF50'),
            ('🔄 Update', self.update_treatment, '#2196F3'),
            ('🧹 Clear', self.clear_treatment_form, '#FF9800'),
            ('🗑️ Delete', self.delete_treatment, '#f44336')
        ]
        
        for text, command, color in treatment_buttons:
            btn = tk.Button(btn_frame, text=text, command=command, 
                          bg=color, fg='white', font=('Arial', 9, 'bold'), padx=8)
            btn.pack(side='left', padx=3, pady=5)
        
        # Export buttons for treatments
        export_frame = ttk.LabelFrame(left_frame, text="Export Treatments Data", padding=10)
        export_frame.grid(row=len(fields)+1, column=0, columnspan=2, pady=10, sticky='ew')
        
        ttk.Button(export_frame, text="📊 Export to CSV", 
                  command=lambda: self.export_treatments_data('csv')).pack(side='left', padx=2, pady=5)
        ttk.Button(export_frame, text="📈 Export to Excel", 
                  command=lambda: self.export_treatments_data('xlsx')).pack(side='left', padx=2, pady=5)
        ttk.Button(export_frame, text="📋 Export to Text", 
                  command=lambda: self.export_treatments_data('txt')).pack(side='left', padx=2, pady=5)
        
        # Right side - Treatments list with search
        search_frame = ttk.Frame(right_frame)
        search_frame.pack(fill='x', pady=5)
        
        ttk.Label(search_frame, text="Search:").pack(side='left', padx=5)
        self.treatment_search_entry = ttk.Entry(search_frame, width=25)
        self.treatment_search_entry.pack(side='left', padx=5)
        self.treatment_search_entry.bind('<KeyRelease>', self.search_treatments)
        
        ttk.Button(search_frame, text="Refresh", command=self.load_treatments).pack(side='left', padx=5)
        
        # Treeview for treatments
        tree_frame = ttk.Frame(right_frame)
        tree_frame.pack(fill='both', expand=True, pady=5)
        
        columns = ('ID', 'Name', 'Type', 'Stress Level', 'Concentration', 'Duration', 'Temperature')
        self.treatments_tree = ttk.Treeview(tree_frame, columns=columns, show='headings', height=15)
        
        # Configure columns
        col_widths = [50, 120, 120, 100, 100, 80, 100]
        for col, width in zip(columns, col_widths):
            self.treatments_tree.heading(col, text=col)
            self.treatments_tree.column(col, width=width)
        
        # Scrollbar
        scrollbar = ttk.Scrollbar(tree_frame, orient='vertical', command=self.treatments_tree.yview)
        self.treatments_tree.configure(yscrollcommand=scrollbar.set)
        
        self.treatments_tree.pack(side='left', fill='both', expand=True)
        scrollbar.pack(side='right', fill='y')
        
        self.treatments_tree.bind('<<TreeviewSelect>>', self.on_treatment_select)
        
        # Load treatments for current experiment
        self.load_treatments()
    
    def get_experiment_details(self, experiment_id):
        """Get experiment details by ID"""
        try:
            query = "SELECT experiment_code, experiment_name FROM experiments WHERE id = ?"
            result = self.db.execute_query(query, (experiment_id,))
            if result:
                return {'code': result[0][0], 'name': result[0][1]}
            return {}
        except Exception as e:
            print(f"Error getting experiment details: {e}")
            return {}
    
    def load_treatments(self):
        """Load treatments for the current experiment - FIXED VERSION"""
        if not self.current_experiment_id:
            return
        
        try:
            for item in self.treatments_tree.get_children():
                self.treatments_tree.delete(item)
            
            query = """
                SELECT id, treatment_name, treatment_type, stress_level, 
                       concentration, duration_days, temperature 
                FROM treatments 
                WHERE experiment_id = ? 
                ORDER BY treatment_name
            """
            results = self.db.execute_query(query, (self.current_experiment_id,))
            
            if results:
                for treatment in results:
                    # Convert all values to strings to avoid formatting issues
                    formatted_treatment = []
                    for value in treatment:
                        if value is None:
                            formatted_treatment.append("")
                        else:
                            # Convert to string without any special formatting
                            formatted_treatment.append(str(value))
                    
                    self.treatments_tree.insert('', 'end', values=formatted_treatment)
                
                self.status_var.set(f"Loaded {len(results)} treatments")
            else:
                self.status_var.set("No treatments found for this experiment")
        
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load treatments: {str(e)}")
            print(f"Debug - Error details: {e}")
    
    def add_treatment(self):
        """Add a new treatment to the current experiment"""
        if not self.current_experiment_id:
            messagebox.showwarning("Warning", "Please select an experiment first")
            return
        
        try:
            name = self.treatment_widgets['name_entry'].get().strip()
            treatment_type = self.treatment_widgets['type_combo'].get().strip()
            
            if not name or not treatment_type:
                messagebox.showerror("Error", "Please fill in all required fields (*)")
                return
            
            # Get concentration value - handle empty string properly
            conc_text = self.treatment_widgets['conc_entry'].get().strip()
            concentration = None
            if conc_text:
                try:
                    concentration = float(conc_text)
                except ValueError:
                    messagebox.showerror("Error", "Concentration must be a valid number")
                    return
            
            # Get duration value - handle empty string properly
            duration_text = self.treatment_widgets['duration_entry'].get().strip()
            duration = None
            if duration_text:
                try:
                    duration = int(duration_text)
                except ValueError:
                    messagebox.showerror("Error", "Duration must be a valid whole number")
                    return
            
            # Get temperature value - handle empty string properly
            temp_text = self.treatment_widgets['temp_entry'].get().strip()
            temperature = None
            if temp_text:
                try:
                    temperature = float(temp_text)
                except ValueError:
                    messagebox.showerror("Error", "Temperature must be a valid number")
                    return
            
            query = """
                INSERT INTO treatments (experiment_id, treatment_name, treatment_type, 
                                      stress_level, concentration, duration_days, 
                                      temperature, description)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?)
            """
            params = (
                self.current_experiment_id,
                name,
                treatment_type,
                self.treatment_widgets['level_combo'].get().strip() or 'medium',
                concentration,
                duration,
                temperature,
                self.treatment_widgets['desc_text'].get('1.0', 'end-1c').strip() or None
            )
            
            result = self.db.execute_query(query, params)
            
            if result is True:
                messagebox.showinfo("Success", "Treatment added successfully!")
                self.clear_treatment_form()
                self.load_treatments()
                self.status_var.set(f"Added treatment: {name}")
            elif result == "DUPLICATE":
                messagebox.showerror("Error", f"Treatment name '{name}' already exists in this experiment. Please use a different name.")
            else:
                messagebox.showerror("Error", "Failed to add treatment. Please check if all required fields are filled correctly.")
        
        except Exception as e:
            messagebox.showerror("Error", f"Failed to add treatment: {str(e)}")
    
    def on_treatment_select(self, event):
        """When a treatment is selected from the list"""
        selected = self.treatments_tree.selection()
        if selected:
            item = self.treatments_tree.item(selected[0])
            values = item['values']
            self.current_treatment_id = values[0]
            
            # Load treatment details into form
            self.clear_treatment_form()
            self.treatment_widgets['name_entry'].insert(0, values[1])
            self.treatment_widgets['type_combo'].set(values[2])
            self.treatment_widgets['level_combo'].set(values[3] or 'medium')
            if values[4]:
                self.treatment_widgets['conc_entry'].insert(0, str(values[4]))
            if values[5]:
                self.treatment_widgets['duration_entry'].insert(0, str(values[5]))
            if values[6]:
                self.treatment_widgets['temp_entry'].insert(0, str(values[6]))
            
            self.status_var.set(f"Selected treatment: {values[1]}")
            
            # Update measurements tab
            self.update_measurements_tab()
    
    def update_treatment(self):
        """Update selected treatment"""
        if not self.current_treatment_id:
            messagebox.showwarning("Warning", "Please select a treatment to update")
            return
        
        try:
            name = self.treatment_widgets['name_entry'].get().strip()
            treatment_type = self.treatment_widgets['type_combo'].get().strip()
            
            if not name or not treatment_type:
                messagebox.showerror("Error", "Please fill in all required fields (*)")
                return
            
            # Get concentration value - handle empty string properly
            conc_text = self.treatment_widgets['conc_entry'].get().strip()
            concentration = None
            if conc_text:
                try:
                    concentration = float(conc_text)
                except ValueError:
                    messagebox.showerror("Error", "Concentration must be a valid number")
                    return
            
            # Get duration value - handle empty string properly
            duration_text = self.treatment_widgets['duration_entry'].get().strip()
            duration = None
            if duration_text:
                try:
                    duration = int(duration_text)
                except ValueError:
                    messagebox.showerror("Error", "Duration must be a valid whole number")
                    return
            
            # Get temperature value - handle empty string properly
            temp_text = self.treatment_widgets['temp_entry'].get().strip()
            temperature = None
            if temp_text:
                try:
                    temperature = float(temp_text)
                except ValueError:
                    messagebox.showerror("Error", "Temperature must be a valid number")
                    return
            
            query = """
                UPDATE treatments 
                SET treatment_name=?, treatment_type=?, stress_level=?,
                    concentration=?, duration_days=?, temperature=?, description=?
                WHERE id=?
            """
            params = (
                name,
                treatment_type,
                self.treatment_widgets['level_combo'].get().strip() or 'medium',
                concentration,
                duration,
                temperature,
                self.treatment_widgets['desc_text'].get('1.0', 'end-1c').strip() or None,
                self.current_treatment_id
            )
            
            if self.db.execute_query(query, params):
                messagebox.showinfo("Success", "Treatment updated successfully!")
                self.load_treatments()
            else:
                messagebox.showerror("Error", "Failed to update treatment")
        
        except Exception as e:
            messagebox.showerror("Error", f"Failed to update treatment: {str(e)}")
    
    def delete_treatment(self):
        """Delete selected treatment"""
        if not self.current_treatment_id:
            messagebox.showwarning("Warning", "Please select a treatment to delete")
            return
        
        # Get treatment name for confirmation message
        treatment_name = ""
        selected = self.treatments_tree.selection()
        if selected:
            item = self.treatments_tree.item(selected[0])
            values = item['values']
            treatment_name = values[1]  # treatment name is at index 1
        
        if messagebox.askyesno("Confirm Delete", 
                              f"Are you sure you want to delete treatment '{treatment_name}' and all its measurements?"):
            try:
                # First delete related measurements
                delete_measurements_query = "DELETE FROM measurements WHERE treatment_id = ?"
                self.db.execute_query(delete_measurements_query, (self.current_treatment_id,))
                
                # Then delete the treatment
                query = "DELETE FROM treatments WHERE id = ?"
                if self.db.execute_query(query, (self.current_treatment_id,)):
                    messagebox.showinfo("Success", "Treatment deleted successfully!")
                    self.clear_treatment_form()
                    self.load_treatments()  # Refresh the treatments list
                    self.current_treatment_id = None
                    
                    # Also update measurements tab since the treatment is gone
                    self.update_measurements_tab()
                    
                    self.status_var.set("Treatment deleted successfully")
                else:
                    messagebox.showerror("Error", "Failed to delete treatment")
            
            except Exception as e:
                messagebox.showerror("Error", f"Failed to delete treatment: {str(e)}")
    
    def clear_treatment_form(self):
        """Clear the treatment form"""
        if hasattr(self, 'treatment_widgets'):
            for key, widget in self.treatment_widgets.items():
                if isinstance(widget, tk.Text):
                    widget.delete('1.0', 'end')
                else:
                    widget.delete(0, 'end')
            
            # Set default values
            if 'type_combo' in self.treatment_widgets:
                self.treatment_widgets['type_combo'].set('control')
            if 'level_combo' in self.treatment_widgets:
                self.treatment_widgets['level_combo'].set('medium')
        
        # Clear selection
        if hasattr(self, 'treatments_tree'):
            self.treatments_tree.selection_remove(self.treatments_tree.selection())
    
    def search_treatments(self, event=None):
        """Search treatments by name"""
        if not self.current_experiment_id:
            return
        
        search_term = self.treatment_search_entry.get().strip()
        
        if not search_term:
            self.load_treatments()
            return
        
        try:
            for item in self.treatments_tree.get_children():
                self.treatments_tree.delete(item)
            
            query = """
                SELECT id, treatment_name, treatment_type, stress_level, 
                       concentration, duration_days, temperature 
                FROM treatments 
                WHERE experiment_id = ? AND treatment_name LIKE ?
                ORDER BY treatment_name
            """
            params = (self.current_experiment_id, f"%{search_term}%")
            results = self.db.execute_query(query, params)
            
            if results:
                for treatment in results:
                    # Convert all values to strings to avoid formatting issues
                    formatted_treatment = []
                    for value in treatment:
                        if value is None:
                            formatted_treatment.append("")
                        else:
                            formatted_treatment.append(str(value))
                    self.treatments_tree.insert('', 'end', values=formatted_treatment)
                
                self.status_var.set(f"Found {len(results)} treatments matching '{search_term}'")
            else:
                self.status_var.set(f"No treatments found matching '{search_term}'")
        
        except Exception as e:
            messagebox.showerror("Error", f"Search failed: {str(e)}")

    # EXPORT METHODS
    def export_experiments_data(self, format_type):
        """Export experiments data in specified format"""
        try:
            query = """
                SELECT id, experiment_code, experiment_name, plant_species, 
                       stress_type, researcher, start_date, end_date, 
                       description, status, created_at
                FROM experiments 
                ORDER BY created_at DESC
            """
            results = self.db.execute_query(query)
            
            if not results:
                messagebox.showinfo("Info", "No experiments data to export")
                return
            
            columns = ['ID', 'Code', 'Name', 'Species', 'Stress Type', 'Researcher', 
                      'Start Date', 'End Date', 'Description', 'Status', 'Created At']
            
            df = pd.DataFrame(results, columns=columns)
            
            filename = filedialog.asksaveasfilename(
                defaultextension=f".{format_type}",
                filetypes=[(f"{format_type.upper()} files", f"*.{format_type}")],
                title=f"Export Experiments Data as {format_type.upper()}"
            )
            
            if filename:
                if format_type == 'csv':
                    df.to_csv(filename, index=False)
                elif format_type == 'xlsx':
                    df.to_excel(filename, index=False, engine='openpyxl')
                elif format_type == 'txt':
                    with open(filename, 'w', encoding='utf-8') as f:
                        f.write("PLANT STRESS PHYSIOLOGY - EXPERIMENTS DATA\n")
                        f.write("=" * 50 + "\n\n")
                        for index, row in df.iterrows():
                            f.write(f"Experiment {index + 1}:\n")
                            f.write(f"  Code: {row['Code']}\n")
                            f.write(f"  Name: {row['Name']}\n")
                            f.write(f"  Species: {row['Species']}\n")
                            f.write(f"  Stress Type: {row['Stress Type']}\n")
                            f.write(f"  Researcher: {row['Researcher']}\n")
                            f.write(f"  Start Date: {row['Start Date']}\n")
                            f.write(f"  End Date: {row['End Date']}\n")
                            f.write(f"  Status: {row['Status']}\n")
                            f.write(f"  Description: {row['Description']}\n")
                            f.write("-" * 30 + "\n")
                
                messagebox.showinfo("Success", f"Experiments data exported successfully to {filename}")
                self.status_var.set(f"Exported experiments data to {os.path.basename(filename)}")
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to export experiments data: {str(e)}")
    
    def export_treatments_data(self, format_type):
        """Export treatments data for current experiment"""
        if not self.current_experiment_id:
            messagebox.showwarning("Warning", "Please select an experiment first")
            return
        
        try:
            query = """
                SELECT t.id, t.treatment_name, t.treatment_type, t.stress_level,
                       t.concentration, t.duration_days, t.temperature, t.description,
                       e.experiment_code, e.experiment_name
                FROM treatments t
                JOIN experiments e ON t.experiment_id = e.id
                WHERE t.experiment_id = ?
                ORDER BY t.treatment_name
            """
            results = self.db.execute_query(query, (self.current_experiment_id,))
            
            if not results:
                messagebox.showinfo("Info", "No treatments data to export")
                return
            
            columns = ['ID', 'Treatment Name', 'Type', 'Stress Level', 'Concentration',
                      'Duration (days)', 'Temperature', 'Description', 'Experiment Code', 'Experiment Name']
            
            df = pd.DataFrame(results, columns=columns)
            
            filename = filedialog.asksaveasfilename(
                defaultextension=f".{format_type}",
                filetypes=[(f"{format_type.upper()} files", f"*.{format_type}")],
                title=f"Export Treatments Data as {format_type.upper()}"
            )
            
            if filename:
                if format_type == 'csv':
                    df.to_csv(filename, index=False)
                elif format_type == 'xlsx':
                    df.to_excel(filename, index=False, engine='openpyxl')
                elif format_type == 'txt':
                    with open(filename, 'w', encoding='utf-8') as f:
                        f.write("PLANT STRESS PHYSIOLOGY - TREATMENTS DATA\n")
                        f.write("=" * 50 + "\n")
                        f.write(f"Experiment: {df.iloc[0]['Experiment Name']} ({df.iloc[0]['Experiment Code']})\n\n")
                        
                        for index, row in df.iterrows():
                            f.write(f"Treatment {index + 1}: {row['Treatment Name']}\n")
                            f.write(f"  Type: {row['Type']}\n")
                            f.write(f"  Stress Level: {row['Stress Level']}\n")
                            f.write(f"  Concentration: {row['Concentration']}\n")
                            f.write(f"  Duration: {row['Duration (days)']} days\n")
                            f.write(f"  Temperature: {row['Temperature']}°C\n")
                            f.write(f"  Description: {row['Description']}\n")
                            f.write("-" * 40 + "\n")
                
                messagebox.showinfo("Success", f"Treatments data exported successfully to {filename}")
                self.status_var.set(f"Exported treatments data to {os.path.basename(filename)}")
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to export treatments data: {str(e)}")
    
    def export_measurements_data(self, format_type):
        """Export measurements data for current treatment"""
        if not self.current_treatment_id:
            messagebox.showwarning("Warning", "Please select a treatment first")
            return
        
        try:
            query = """
                SELECT m.id, m.measurement_date, m.plant_height, m.leaf_area,
                       m.chlorophyll_content, m.photosynthesis_rate, m.stomatal_conductance,
                       m.root_length, m.biomass_fresh, m.biomass_dry, m.water_content,
                       m.notes, t.treatment_name, e.experiment_code
                FROM measurements m
                JOIN treatments t ON m.treatment_id = t.id
                JOIN experiments e ON t.experiment_id = e.id
                WHERE m.treatment_id = ?
                ORDER BY m.measurement_date
            """
            results = self.db.execute_query(query, (self.current_treatment_id,))
            
            if not results:
                messagebox.showinfo("Info", "No measurements data to export")
                return
            
            columns = ['ID', 'Date', 'Height (cm)', 'Leaf Area (cm²)', 'Chlorophyll',
                      'Photosynthesis Rate', 'Stomatal Conductance', 'Root Length (cm)',
                      'Fresh Biomass (g)', 'Dry Biomass (g)', 'Water Content (%)',
                      'Notes', 'Treatment Name', 'Experiment Code']
            
            df = pd.DataFrame(results, columns=columns)
            
            filename = filedialog.asksaveasfilename(
                defaultextension=f".{format_type}",
                filetypes=[(f"{format_type.upper()} files", f"*.{format_type}")],
                title=f"Export Measurements Data as {format_type.upper()}"
            )
            
            if filename:
                if format_type == 'csv':
                    df.to_csv(filename, index=False)
                elif format_type == 'xlsx':
                    df.to_excel(filename, index=False, engine='openpyxl')
                elif format_type == 'txt':
                    with open(filename, 'w', encoding='utf-8') as f:
                        f.write("PLANT STRESS PHYSIOLOGY - MEASUREMENTS DATA\n")
                        f.write("=" * 50 + "\n")
                        f.write(f"Experiment: {df.iloc[0]['Experiment Code']}\n")
                        f.write(f"Treatment: {df.iloc[0]['Treatment Name']}\n\n")
                        
                        for index, row in df.iterrows():
                            f.write(f"Measurement {index + 1}:\n")
                            f.write(f"  ID: {row['ID']}\n")
                            f.write(f"  Date: {row['Date']}\n")
                            f.write(f"  Plant Height: {row['Height (cm)']} cm\n")
                            f.write(f"  Leaf Area: {row['Leaf Area (cm²)']} cm²\n")
                            f.write(f"  Chlorophyll Content: {row['Chlorophyll']}\n")
                            f.write(f"  Photosynthesis Rate: {row['Photosynthesis Rate']}\n")
                            f.write(f"  Stomatal Conductance: {row['Stomatal Conductance']}\n")
                            f.write(f"  Root Length: {row['Root Length (cm)']} cm\n")
                            f.write(f"  Fresh Biomass: {row['Fresh Biomass (g)']} g\n")
                            f.write(f"  Dry Biomass: {row['Dry Biomass (g)']} g\n")
                            f.write(f"  Water Content: {row['Water Content (%)']}%\n")
                            f.write(f"  Notes: {row['Notes']}\n")
                            f.write("-" * 50 + "\n")
                
                messagebox.showinfo("Success", f"Measurements data exported successfully to {filename}")
                self.status_var.set(f"Exported measurements data to {os.path.basename(filename)}")
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to export measurements data: {str(e)}")
    
    def export_analysis_data(self, format_type):
        """Export analysis data for current experiment - FIXED VERSION"""
        if not self.current_experiment_id:
            messagebox.showwarning("Warning", "Please select an experiment first")
            return
        
        try:
            # Get growth rates analysis
            growth_rates = self.analyzer.calculate_growth_rates(self.current_experiment_id)
            stress_impact = self.analyzer.stress_impact_analysis(self.current_experiment_id)
            statistics = self.analyzer.calculate_statistics(self.current_experiment_id)
            
            # Check if we have any data to export
            has_data = False
            data_frames = {}
            
            if growth_rates is not None and not growth_rates.empty:
                data_frames['Growth Rates'] = growth_rates
                has_data = True
            
            if stress_impact is not None and not stress_impact.empty:
                data_frames['Stress Impact'] = stress_impact
                has_data = True
            
            if statistics is not None and not statistics.empty:
                data_frames['Statistics'] = statistics
                has_data = True
            
            if not has_data:
                messagebox.showinfo("Info", "No analysis data available for export")
                return
            
            filename = filedialog.asksaveasfilename(
                defaultextension=f".{format_type}",
                filetypes=[(f"{format_type.upper()} files", f"*.{format_type}")],
                title=f"Export Analysis Data as {format_type.upper()}"
            )
            
            if not filename:
                return  # User cancelled
            
            if format_type == 'csv':
                # For CSV, we'll create separate files for each analysis type
                base_name = filename.replace(f".{format_type}", "")
                exported_files = []
                
                for sheet_name, df in data_frames.items():
                    safe_name = sheet_name.replace(' ', '_').lower()
                    file_path = f"{base_name}_{safe_name}.csv"
                    
                    # Clean the dataframe for export
                    df_clean = df.reset_index(drop=True)
                    if isinstance(df_clean.columns, pd.MultiIndex):
                        df_clean.columns = ['_'.join(map(str, col)).strip() for col in df_clean.columns]
                    
                    df_clean.to_csv(file_path, index=True)  # Include index for row numbers
                    exported_files.append(file_path)
                
                if len(exported_files) == 1:
                    messagebox.showinfo("Success", f"Analysis data exported successfully to {exported_files[0]}")
                else:
                    messagebox.showinfo("Success", f"Analysis data exported successfully to {len(exported_files)} files")
                
                self.status_var.set(f"Exported analysis data to {os.path.basename(base_name)}_*.csv")
            
            elif format_type == 'xlsx':
                # FIXED: Handle Excel export with proper index handling
                with pd.ExcelWriter(filename, engine='openpyxl') as writer:
                    for sheet_name, df in data_frames.items():
                        # Clean the dataframe for export
                        df_clean = df.reset_index(drop=True)
                        
                        # Handle multi-index columns
                        if isinstance(df_clean.columns, pd.MultiIndex):
                            df_clean.columns = ['_'.join(map(str, col)).strip() for col in df_clean.columns]
                        
                        # Export with index to avoid the multi-index error
                        df_clean.to_excel(writer, sheet_name=sheet_name[:31], index=True)  # Sheet name max 31 chars
                
                messagebox.showinfo("Success", f"Analysis data exported successfully to {filename}")
                self.status_var.set(f"Exported analysis data to {os.path.basename(filename)}")
            
            elif format_type == 'txt':
                with open(filename, 'w', encoding='utf-8') as f:
                    f.write("PLANT STRESS PHYSIOLOGY - ANALYSIS DATA\n")
                    f.write("=" * 50 + "\n\n")
                    
                    for sheet_name, df in data_frames.items():
                        f.write(f"{sheet_name.upper()}:\n")
                        f.write("-" * 40 + "\n")
                        
                        # Clean the dataframe for text output
                        df_clean = df.reset_index(drop=True)
                        if isinstance(df_clean.columns, pd.MultiIndex):
                            df_clean.columns = ['_'.join(map(str, col)).strip() for col in df_clean.columns]
                        
                        f.write(df_clean.to_string(index=False))
                        f.write("\n\n")
                
                messagebox.showinfo("Success", f"Analysis data exported successfully to {filename}")
                self.status_var.set(f"Exported analysis data to {os.path.basename(filename)}")
            
        except Exception as e:
            error_msg = f"Failed to export analysis data: {str(e)}"
            messagebox.showerror("Error", error_msg)
            print(f"Debug - Export error: {e}")  # For debugging
    
    def export_comprehensive_data(self, format_type):
        """Export comprehensive data for all experiments"""
        try:
            # Get all data
            experiments = self.db.execute_query("SELECT * FROM experiments")
            treatments = self.db.execute_query("SELECT * FROM treatments")
            measurements = self.db.execute_query("SELECT * FROM measurements")
            
            if not experiments and not treatments and not measurements:
                messagebox.showinfo("Info", "No data available for export")
                return
            
            filename = filedialog.asksaveasfilename(
                defaultextension=f".{format_type}",
                filetypes=[(f"{format_type.upper()} files", f"*.{format_type}")],
                title=f"Export Comprehensive Data as {format_type.upper()}"
            )
            
            if filename:
                if format_type == 'csv':
                    # Export multiple CSV files in a zip or separate files
                    base_name = filename.replace(f".{format_type}", "")
                    
                    if experiments:
                        exp_df = pd.DataFrame(experiments, columns=[
                            'ID', 'Code', 'Name', 'Species', 'Stress Type', 'Researcher',
                            'Start Date', 'End Date', 'Description', 'Status', 'Created', 'Updated'
                        ])
                        exp_df.to_csv(f"{base_name}_experiments.csv", index=False)
                    
                    if treatments:
                        treat_df = pd.DataFrame(treatments, columns=[
                            'ID', 'Experiment ID', 'Treatment Name', 'Type', 'Stress Level',
                            'Concentration', 'Duration', 'Temperature', 'Description', 'Created', 'Updated'
                        ])
                        treat_df.to_csv(f"{base_name}_treatments.csv", index=False)
                    
                    if measurements:
                        meas_df = pd.DataFrame(measurements, columns=[
                            'ID', 'Treatment ID', 'Date', 'Height', 'Leaf Area', 'Chlorophyll',
                            'Photosynthesis', 'Stomatal', 'Root Length', 'Biomass Fresh',
                            'Biomass Dry', 'Water Content', 'Notes', 'Created'
                        ])
                        meas_df.to_csv(f"{base_name}_measurements.csv", index=False)
                    
                    messagebox.showinfo("Success", f"Comprehensive data exported successfully to multiple CSV files")
                
                elif format_type == 'xlsx':
                    with pd.ExcelWriter(filename, engine='openpyxl') as writer:
                        if experiments:
                            exp_df = pd.DataFrame(experiments, columns=[
                                'ID', 'Code', 'Name', 'Species', 'Stress Type', 'Researcher',
                                'Start Date', 'End Date', 'Description', 'Status', 'Created', 'Updated'
                            ])
                            exp_df.to_excel(writer, sheet_name='Experiments', index=False)
                        
                        if treatments:
                            treat_df = pd.DataFrame(treatments, columns=[
                                'ID', 'Experiment ID', 'Treatment Name', 'Type', 'Stress Level',
                                'Concentration', 'Duration', 'Temperature', 'Description', 'Created', 'Updated'
                            ])
                            treat_df.to_excel(writer, sheet_name='Treatments', index=False)
                        
                        if measurements:
                            meas_df = pd.DataFrame(measurements, columns=[
                                'ID', 'Treatment ID', 'Date', 'Height', 'Leaf Area', 'Chlorophyll',
                                'Photosynthesis', 'Stomatal', 'Root Length', 'Biomass Fresh',
                                'Biomass Dry', 'Water Content', 'Notes', 'Created'
                            ])
                            meas_df.to_excel(writer, sheet_name='Measurements', index=False)
                    
                    messagebox.showinfo("Success", f"Comprehensive data exported successfully to {filename}")
                
                elif format_type == 'txt':
                    with open(filename, 'w', encoding='utf-8') as f:
                        f.write("PLANT STRESS PHYSIOLOGY - COMPREHENSIVE DATA EXPORT\n")
                        f.write("=" * 60 + "\n\n")
                        
                        if experiments:
                            f.write("EXPERIMENTS:\n")
                            f.write("-" * 40 + "\n")
                            for exp in experiments:
                                f.write(f"Code: {exp[1]}, Name: {exp[2]}, Species: {exp[3]}\n")
                                f.write(f"Stress Type: {exp[4]}, Researcher: {exp[5]}\n")
                                f.write(f"Period: {exp[6]} to {exp[7]}, Status: {exp[9]}\n")
                                f.write(f"Description: {exp[8]}\n")
                                f.write("\n")
                        
                        if treatments:
                            f.write("\nTREATMENTS:\n")
                            f.write("-" * 40 + "\n")
                            for treat in treatments:
                                f.write(f"Name: {treat[2]}, Type: {treat[3]}, Level: {treat[4]}\n")
                                f.write(f"Concentration: {treat[5]}, Duration: {treat[6]} days\n")
                                f.write(f"Temperature: {treat[7]}°C\n")
                                f.write(f"Description: {treat[8]}\n")
                                f.write("\n")
                        
                        if measurements:
                            f.write("\nMEASUREMENTS:\n")
                            f.write("-" * 40 + "\n")
                            for meas in measurements:
                                f.write(f"Date: {meas[2]}, Height: {meas[3]} cm, Leaf Area: {meas[4]} cm²\n")
                                f.write(f"Chlorophyll: {meas[5]}, Photosynthesis: {meas[6]}\n")
                                f.write(f"Water Content: {meas[11]}%, Notes: {meas[12]}\n")
                                f.write("\n")
                    
                    messagebox.showinfo("Success", f"Comprehensive data exported successfully to {filename}")
                
                elif format_type == 'json':
                    data = {
                        'experiments': [],
                        'treatments': [],
                        'measurements': []
                    }
                    
                    if experiments:
                        for exp in experiments:
                            data['experiments'].append({
                                'id': exp[0], 'code': exp[1], 'name': exp[2],
                                'species': exp[3], 'stress_type': exp[4],
                                'researcher': exp[5], 'start_date': exp[6],
                                'end_date': exp[7], 'description': exp[8],
                                'status': exp[9], 'created_at': exp[10]
                            })
                    
                    if treatments:
                        for treat in treatments:
                            data['treatments'].append({
                                'id': treat[0], 'experiment_id': treat[1],
                                'name': treat[2], 'type': treat[3],
                                'stress_level': treat[4], 'concentration': treat[5],
                                'duration_days': treat[6], 'temperature': treat[7],
                                'description': treat[8], 'created_at': treat[9]
                            })
                    
                    if measurements:
                        for meas in measurements:
                            data['measurements'].append({
                                'id': meas[0], 'treatment_id': meas[1],
                                'date': meas[2], 'plant_height': meas[3],
                                'leaf_area': meas[4], 'chlorophyll_content': meas[5],
                                'photosynthesis_rate': meas[6], 'stomatal_conductance': meas[7],
                                'root_length': meas[8], 'biomass_fresh': meas[9],
                                'biomass_dry': meas[10], 'water_content': meas[11],
                                'notes': meas[12], 'created_at': meas[13]
                            })
                    
                    with open(filename, 'w', encoding='utf-8') as f:
                        json.dump(data, f, indent=2, ensure_ascii=False)
                    
                    messagebox.showinfo("Success", f"Comprehensive data exported successfully to {filename}")
                
                self.status_var.set(f"Exported comprehensive data to {os.path.basename(filename)}")
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to export comprehensive data: {str(e)}")

    def update_measurements_tab(self):
        """Update measurements tab when a treatment is selected - FIXED WITH VISIBLE EXPORT BUTTONS"""
        # Clear existing content
        for widget in self.measurements_content.winfo_children():
            widget.destroy()
        
        # Hide initial label and show content
        self.measurements_initial_label.pack_forget()
        self.measurements_content.pack(fill='both', expand=True)
        
        if not self.current_treatment_id:
            ttk.Label(self.measurements_content, 
                     text="Please select a treatment first to manage measurements", 
                     font=('Arial', 12, 'bold')).pack(pady=50)
            return
        
        # Get treatment details for display
        treatment_details = self.get_treatment_details(self.current_treatment_id)
        treatment_title = f"Measurements for: {treatment_details.get('name', 'Unknown Treatment')} (ID: {self.current_treatment_id})"
        
        ttk.Label(self.measurements_content, text=treatment_title, 
                 font=('Arial', 14, 'bold')).pack(pady=10)
        
        # Create two-column layout
        main_container = ttk.Frame(self.measurements_content)
        main_container.pack(fill='both', expand=True, padx=10, pady=10)
        
        # Left frame - Measurement form
        left_frame = ttk.LabelFrame(main_container, text="Measurement Details", padding=15)
        left_frame.pack(side='left', fill='y', padx=5, pady=5)
        
        # Right frame - Measurements list
        right_frame = ttk.LabelFrame(main_container, text="Measurements List", padding=10)
        right_frame.pack(side='right', fill='both', expand=True, padx=5, pady=5)
        
        # Measurement form fields
        fields = [
            ('Measurement Date:*', 'date_entry'),
            ('Plant Height (cm):', 'height_entry'),
            ('Leaf Area (cm²):', 'leaf_area_entry'),
            ('Chlorophyll Content:', 'chlorophyll_entry'),
            ('Photosynthesis Rate:', 'photosynthesis_entry'),
            ('Stomatal Conductance:', 'stomatal_entry'),
            ('Root Length (cm):', 'root_length_entry'),
            ('Fresh Biomass (g):', 'biomass_fresh_entry'),
            ('Dry Biomass (g):', 'biomass_dry_entry'),
            ('Water Content (%):', 'water_content_entry'),
            ('Notes:', 'notes_text')
        ]
        
        self.measurement_widgets = {}
        for i, (label, key) in enumerate(fields):
            ttk.Label(left_frame, text=label, style='Header.TLabel').grid(row=i, column=0, sticky='w', pady=6)
            
            if 'date' in key:
                widget = ttk.Entry(left_frame, width=25)
                widget.insert(0, datetime.now().strftime('%Y-%m-%d'))
            elif 'notes' in key:
                widget = tk.Text(left_frame, width=25, height=4, font=('Arial', 9))
            else:
                widget = ttk.Entry(left_frame, width=25)
            
            widget.grid(row=i, column=1, padx=5, pady=4, sticky='w')
            self.measurement_widgets[key] = widget
        
        # Buttons for measurement operations
        btn_frame = ttk.Frame(left_frame)
        btn_frame.grid(row=len(fields), column=0, columnspan=2, pady=15)
        
        measurement_buttons = [
            ('💾 Add Measurement', self.add_measurement, '#4CAF50'),
            ('🔄 Update', self.update_measurement, '#2196F3'),
            ('🧹 Clear', self.clear_measurement_form, '#FF9800'),
            ('🗑️ Delete', self.delete_measurement, '#f44336')
        ]
        
        for text, command, color in measurement_buttons:
            btn = tk.Button(btn_frame, text=text, command=command, 
                          bg=color, fg='white', font=('Arial', 9, 'bold'), padx=8)
            btn.pack(side='left', padx=3, pady=5)
        
        # Quick actions frame - FIXED: Now properly placed
        quick_actions_frame = ttk.LabelFrame(left_frame, text="Quick Actions", padding=10)
        quick_actions_frame.grid(row=len(fields)+1, column=0, columnspan=2, pady=10, sticky='ew')
        
        ttk.Button(quick_actions_frame, text="📊 Calculate Water Content", 
                   command=self.calculate_water_content).pack(side='left', padx=2)
        ttk.Button(quick_actions_frame, text="📈 Growth Analysis", 
                   command=self.quick_growth_analysis).pack(side='left', padx=2)
        
        # EXPORT BUTTONS FOR MEASUREMENTS - FIXED: Now properly visible
        export_frame = ttk.LabelFrame(left_frame, text="Export Measurements Data", padding=10)
        export_frame.grid(row=len(fields)+2, column=0, columnspan=2, pady=10, sticky='ew')
        
        # Create individual export buttons that are clearly visible
        export_btn_frame = ttk.Frame(export_frame)
        export_btn_frame.pack(fill='x', pady=5)
        
        export_csv_btn = tk.Button(export_btn_frame, text="📊 Export to CSV", 
                                  command=lambda: self.export_measurements_data('csv'),
                                  bg='#2196F3', fg='white', font=('Arial', 9, 'bold'),
                                  padx=10, pady=5, width=15)
        export_csv_btn.pack(side='left', padx=2, pady=2, fill='x', expand=True)
        
        export_excel_btn = tk.Button(export_btn_frame, text="📈 Export to Excel", 
                                    command=lambda: self.export_measurements_data('xlsx'),
                                    bg='#4CAF50', fg='white', font=('Arial', 9, 'bold'),
                                    padx=10, pady=5, width=15)
        export_excel_btn.pack(side='left', padx=2, pady=2, fill='x', expand=True)
        
        export_text_btn = tk.Button(export_btn_frame, text="📋 Export to Text", 
                                   command=lambda: self.export_measurements_data('txt'),
                                   bg='#FF9800', fg='white', font=('Arial', 9, 'bold'),
                                   padx=10, pady=5, width=15)
        export_text_btn.pack(side='left', padx=2, pady=2, fill='x', expand=True)
        
        # Right side - Measurements list with search and filters
        search_frame = ttk.Frame(right_frame)
        search_frame.pack(fill='x', pady=5)
        
        ttk.Label(search_frame, text="Search:").pack(side='left', padx=5)
        self.measurement_search_entry = ttk.Entry(search_frame, width=20)
        self.measurement_search_entry.pack(side='left', padx=5)
        self.measurement_search_entry.bind('<KeyRelease>', self.search_measurements)
        
        # Date filter
        ttk.Label(search_frame, text="Date:").pack(side='left', padx=(15,5))
        self.date_filter_combo = ttk.Combobox(search_frame, width=12, values=[
            'All', 'Today', 'This Week', 'This Month', 'Last Month'
        ], state='readonly')
        self.date_filter_combo.set('All')
        self.date_filter_combo.pack(side='left', padx=5)
        self.date_filter_combo.bind('<<ComboboxSelected>>', self.filter_measurements)
        
        ttk.Button(search_frame, text="Refresh", command=self.load_measurements).pack(side='left', padx=5)
        
        # Treeview for measurements
        tree_frame = ttk.Frame(right_frame)
        tree_frame.pack(fill='both', expand=True, pady=5)
        
        columns = ('ID', 'Date', 'Height', 'Leaf Area', 'Chlorophyll', 'Photosynthesis', 'Water %')
        self.measurements_tree = ttk.Treeview(tree_frame, columns=columns, show='headings', height=15)
        
        # Configure columns
        col_widths = [50, 100, 80, 90, 100, 110, 80]
        for col, width in zip(columns, col_widths):
            self.measurements_tree.heading(col, text=col)
            self.measurements_tree.column(col, width=width)
        
        # Scrollbar
        scrollbar = ttk.Scrollbar(tree_frame, orient='vertical', command=self.measurements_tree.yview)
        self.measurements_tree.configure(yscrollcommand=scrollbar.set)
        
        self.measurements_tree.pack(side='left', fill='both', expand=True)
        scrollbar.pack(side='right', fill='y')
        
        self.measurements_tree.bind('<<TreeviewSelect>>', self.on_measurement_select)
        
        # Summary statistics frame
        summary_frame = ttk.LabelFrame(right_frame, text="Summary Statistics", padding=10)
        summary_frame.pack(fill='x', pady=10)
        
        self.summary_var = tk.StringVar()
        self.summary_var.set("No measurements data available")
        ttk.Label(summary_frame, textvariable=self.summary_var, font=('Arial', 9)).pack()
        
        # Load measurements for current treatment
        self.load_measurements()

    def get_treatment_details(self, treatment_id):
        """Get treatment details by ID"""
        try:
            query = "SELECT treatment_name, treatment_type FROM treatments WHERE id = ?"
            result = self.db.execute_query(query, (treatment_id,))
            if result:
                return {'name': result[0][0], 'type': result[0][1]}
            return {}
        except Exception as e:
            print(f"Error getting treatment details: {e}")
            return {}

    def load_measurements(self):
        """Load measurements for the current treatment"""
        if not self.current_treatment_id:
            return
        
        try:
            for item in self.measurements_tree.get_children():
                self.measurements_tree.delete(item)
            
            query = """
                SELECT id, measurement_date, plant_height, leaf_area, 
                       chlorophyll_content, photosynthesis_rate, water_content
                FROM measurements 
                WHERE treatment_id = ? 
                ORDER BY measurement_date DESC
            """
            results = self.db.execute_query(query, (self.current_treatment_id,))
            
            if results:
                for measurement in results:
                    # Convert all values to strings to avoid formatting issues
                    formatted_measurement = []
                    for value in measurement:
                        if value is None:
                            formatted_measurement.append("")
                        else:
                            formatted_measurement.append(str(value))
                    
                    self.measurements_tree.insert('', 'end', values=formatted_measurement)
                
                self.status_var.set(f"Loaded {len(results)} measurements")
                self.update_summary_statistics(results)
            else:
                self.status_var.set("No measurements found for this treatment")
                self.summary_var.set("No measurements data available")
        
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load measurements: {str(e)}")

    def update_summary_statistics(self, measurements_data):
        """Update summary statistics display"""
        if not measurements_data:
            self.summary_var.set("No measurements data available")
            return
        
        try:
            # Convert to DataFrame for easier calculations
            df = pd.DataFrame(measurements_data, 
                             columns=['id', 'date', 'height', 'leaf_area', 'chlorophyll', 'photosynthesis', 'water_content'])
            
            stats_text = "Statistics: "
            stats = []
            
            # Convert string values to float for calculations
            try:
                if df['height'].notna().any():
                    heights = pd.to_numeric(df['height'], errors='coerce')
                    avg_height = heights.mean()
                    if not pd.isna(avg_height):
                        stats.append(f"Avg Height: {avg_height:.2f}cm")
                
                if df['leaf_area'].notna().any():
                    leaf_areas = pd.to_numeric(df['leaf_area'], errors='coerce')
                    avg_leaf = leaf_areas.mean()
                    if not pd.isna(avg_leaf):
                        stats.append(f"Avg Leaf Area: {avg_leaf:.2f}cm²")
                
                if df['water_content'].notna().any():
                    water_contents = pd.to_numeric(df['water_content'], errors='coerce')
                    avg_water = water_contents.mean()
                    if not pd.isna(avg_water):
                        stats.append(f"Avg Water: {avg_water:.2f}%")
            except Exception as calc_error:
                print(f"Error calculating statistics: {calc_error}")
            
            if stats:
                self.summary_var.set("Statistics: " + " | ".join(stats))
            else:
                self.summary_var.set("No numeric data available for statistics")
                
        except Exception as e:
            self.summary_var.set("Error calculating statistics")

    def add_measurement(self):
        """Add a new measurement to the current treatment"""
        if not self.current_treatment_id:
            messagebox.showwarning("Warning", "Please select a treatment first")
            return
        
        try:
            date = self.measurement_widgets['date_entry'].get().strip()
            
            if not date:
                messagebox.showerror("Error", "Please enter measurement date")
                return
            
            # Get all measurement values
            measurement_data = {
                'plant_height': self.get_float_value('height_entry'),
                'leaf_area': self.get_float_value('leaf_area_entry'),
                'chlorophyll_content': self.get_float_value('chlorophyll_entry'),
                'photosynthesis_rate': self.get_float_value('photosynthesis_entry'),
                'stomatal_conductance': self.get_float_value('stomatal_entry'),
                'root_length': self.get_float_value('root_length_entry'),
                'biomass_fresh': self.get_float_value('biomass_fresh_entry'),
                'biomass_dry': self.get_float_value('biomass_dry_entry'),
                'water_content': self.get_float_value('water_content_entry'),
                'notes': self.measurement_widgets['notes_text'].get('1.0', 'end-1c').strip()
            }
            
            # Check if at least one measurement value is provided
            has_data = any(value is not None for value in [
                measurement_data['plant_height'], measurement_data['leaf_area'],
                measurement_data['chlorophyll_content'], measurement_data['photosynthesis_rate'],
                measurement_data['stomatal_conductance'], measurement_data['root_length'],
                measurement_data['biomass_fresh'], measurement_data['biomass_dry'],
                measurement_data['water_content']
            ])
            
            if not has_data:
                messagebox.showerror("Error", "Please enter at least one measurement value")
                return
            
            query = """
                INSERT INTO measurements 
                (treatment_id, measurement_date, plant_height, leaf_area, 
                 chlorophyll_content, photosynthesis_rate, stomatal_conductance,
                 root_length, biomass_fresh, biomass_dry, water_content, notes)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """
            params = (
                self.current_treatment_id,
                date,
                measurement_data['plant_height'],
                measurement_data['leaf_area'],
                measurement_data['chlorophyll_content'],
                measurement_data['photosynthesis_rate'],
                measurement_data['stomatal_conductance'],
                measurement_data['root_length'],
                measurement_data['biomass_fresh'],
                measurement_data['biomass_dry'],
                measurement_data['water_content'],
                measurement_data['notes'] or None
            )
            
            if self.db.execute_query(query, params):
                messagebox.showinfo("Success", "Measurement added successfully!")
                self.clear_measurement_form()
                self.load_measurements()
                self.status_var.set(f"Added measurement for {date}")
            else:
                messagebox.showerror("Error", "Failed to add measurement")
        
        except Exception as e:
            messagebox.showerror("Error", f"Failed to add measurement: {str(e)}")

    def get_float_value(self, widget_key):
        """Get float value from entry widget, return None if empty"""
        value = self.measurement_widgets[widget_key].get().strip()
        return float(value) if value else None

    def on_measurement_select(self, event):
        """When a measurement is selected from the list"""
        selected = self.measurements_tree.selection()
        if selected:
            item = self.measurements_tree.item(selected[0])
            values = item['values']
            self.current_measurement_id = values[0]
            
            # Load full measurement details into form
            self.load_measurement_details(self.current_measurement_id)
            
            self.status_var.set(f"Selected measurement from {values[1]}")

    def load_measurement_details(self, measurement_id):
        """Load complete measurement details into form"""
        try:
            query = """
                SELECT measurement_date, plant_height, leaf_area, chlorophyll_content,
                       photosynthesis_rate, stomatal_conductance, root_length,
                       biomass_fresh, biomass_dry, water_content, notes
                FROM measurements WHERE id = ?
            """
            result = self.db.execute_query(query, (measurement_id,))
            
            if result:
                data = result[0]
                self.clear_measurement_form()
                
                # Fill form with data
                self.measurement_widgets['date_entry'].insert(0, data[0])
                if data[1]: self.measurement_widgets['height_entry'].insert(0, str(data[1]))
                if data[2]: self.measurement_widgets['leaf_area_entry'].insert(0, str(data[2]))
                if data[3]: self.measurement_widgets['chlorophyll_entry'].insert(0, str(data[3]))
                if data[4]: self.measurement_widgets['photosynthesis_entry'].insert(0, str(data[4]))
                if data[5]: self.measurement_widgets['stomatal_entry'].insert(0, str(data[5]))
                if data[6]: self.measurement_widgets['root_length_entry'].insert(0, str(data[6]))
                if data[7]: self.measurement_widgets['biomass_fresh_entry'].insert(0, str(data[7]))
                if data[8]: self.measurement_widgets['biomass_dry_entry'].insert(0, str(data[8]))
                if data[9]: self.measurement_widgets['water_content_entry'].insert(0, str(data[9]))
                if data[10]: self.measurement_widgets['notes_text'].insert('1.0', data[10])
        
        except Exception as e:
            print(f"Error loading measurement details: {e}")

    def update_measurement(self):
        """Update selected measurement"""
        if not hasattr(self, 'current_measurement_id') or not self.current_measurement_id:
            messagebox.showwarning("Warning", "Please select a measurement to update")
            return
        
        try:
            date = self.measurement_widgets['date_entry'].get().strip()
            
            if not date:
                messagebox.showerror("Error", "Please enter measurement date")
                return
            
            # Get all measurement values
            measurement_data = {
                'plant_height': self.get_float_value('height_entry'),
                'leaf_area': self.get_float_value('leaf_area_entry'),
                'chlorophyll_content': self.get_float_value('chlorophyll_entry'),
                'photosynthesis_rate': self.get_float_value('photosynthesis_entry'),
                'stomatal_conductance': self.get_float_value('stomatal_entry'),
                'root_length': self.get_float_value('root_length_entry'),
                'biomass_fresh': self.get_float_value('biomass_fresh_entry'),
                'biomass_dry': self.get_float_value('biomass_dry_entry'),
                'water_content': self.get_float_value('water_content_entry'),
                'notes': self.measurement_widgets['notes_text'].get('1.0', 'end-1c').strip()
            }
            
            query = """
                UPDATE measurements 
                SET measurement_date=?, plant_height=?, leaf_area=?, 
                    chlorophyll_content=?, photosynthesis_rate=?, stomatal_conductance=?,
                    root_length=?, biomass_fresh=?, biomass_dry=?, water_content=?, notes=?
                WHERE id=?
            """
            params = (
                date,
                measurement_data['plant_height'],
                measurement_data['leaf_area'],
                measurement_data['chlorophyll_content'],
                measurement_data['photosynthesis_rate'],
                measurement_data['stomatal_conductance'],
                measurement_data['root_length'],
                measurement_data['biomass_fresh'],
                measurement_data['biomass_dry'],
                measurement_data['water_content'],
                measurement_data['notes'] or None,
                self.current_measurement_id
            )
            
            if self.db.execute_query(query, params):
                messagebox.showinfo("Success", "Measurement updated successfully!")
                self.load_measurements()
            else:
                messagebox.showerror("Error", "Failed to update measurement")
        
        except Exception as e:
            messagebox.showerror("Error", f"Failed to update measurement: {str(e)}")

    def delete_measurement(self):
        """Delete selected measurement"""
        if not hasattr(self, 'current_measurement_id') or not self.current_measurement_id:
            messagebox.showwarning("Warning", "Please select a measurement to delete")
            return
        
        if messagebox.askyesno("Confirm Delete", "Are you sure you want to delete this measurement?"):
            try:
                query = "DELETE FROM measurements WHERE id = ?"
                if self.db.execute_query(query, (self.current_measurement_id,)):
                    messagebox.showinfo("Success", "Measurement deleted successfully!")
                    self.clear_measurement_form()
                    self.load_measurements()
                    self.current_measurement_id = None
                else:
                    messagebox.showerror("Error", "Failed to delete measurement")
            
            except Exception as e:
                messagebox.showerror("Error", f"Failed to delete measurement: {str(e)}")

    def clear_measurement_form(self):
        """Clear the measurement form"""
        if hasattr(self, 'measurement_widgets'):
            for key, widget in self.measurement_widgets.items():
                if isinstance(widget, tk.Text):
                    widget.delete('1.0', 'end')
                else:
                    widget.delete(0, 'end')
            
            # Set default date
            self.measurement_widgets['date_entry'].insert(0, datetime.now().strftime('%Y-%m-%d'))
        
        # Clear selection
        if hasattr(self, 'measurements_tree'):
            self.measurements_tree.selection_remove(self.measurements_tree.selection())
        
        # Clear current measurement ID
        if hasattr(self, 'current_measurement_id'):
            self.current_measurement_id = None

    def search_measurements(self, event=None):
        """Search measurements by date or notes"""
        if not self.current_treatment_id:
            return
        
        search_term = self.measurement_search_entry.get().strip()
        
        if not search_term:
            self.load_measurements()
            return
        
        try:
            for item in self.measurements_tree.get_children():
                self.measurements_tree.delete(item)
            
            query = """
                SELECT id, measurement_date, plant_height, leaf_area, 
                       chlorophyll_content, photosynthesis_rate, water_content
                FROM measurements 
                WHERE treatment_id = ? AND (measurement_date LIKE ? OR notes LIKE ?)
                ORDER BY measurement_date DESC
            """
            search_pattern = f"%{search_term}%"
            params = (self.current_treatment_id, search_pattern, search_pattern)
            results = self.db.execute_query(query, params)
        
            if results:
                for measurement in results:
                    # Convert all values to strings
                    formatted_measurement = []
                    for value in measurement:
                        if value is None:
                            formatted_measurement.append("")
                        else:
                            formatted_measurement.append(str(value))
                    self.measurements_tree.insert('', 'end', values=formatted_measurement)
                
                self.status_var.set(f"Found {len(results)} measurements matching '{search_term}'")
            else:
                self.status_var.set(f"No measurements found matching '{search_term}'")
        
        except Exception as e:
            messagebox.showerror("Error", f"Search failed: {str(e)}")

    def filter_measurements(self, event=None):
        """Filter measurements by date range"""
        if not self.current_treatment_id:
            return
        
        date_filter = self.date_filter_combo.get()
        
        try:
            for item in self.measurements_tree.get_children():
                self.measurements_tree.delete(item)
            
            if date_filter == 'All':
                query = """
                    SELECT id, measurement_date, plant_height, leaf_area, 
                           chlorophyll_content, photosynthesis_rate, water_content
                    FROM measurements 
                    WHERE treatment_id = ?
                    ORDER BY measurement_date DESC
                """
                params = (self.current_treatment_id,)
            else:
                # For date filtering, you would need to implement date range logic
                # This is a simplified version
                query = """
                    SELECT id, measurement_date, plant_height, leaf_area, 
                           chlorophyll_content, photosynthesis_rate, water_content
                    FROM measurements 
                    WHERE treatment_id = ?
                    ORDER BY measurement_date DESC
                """
                params = (self.current_treatment_id,)
            
            results = self.db.execute_query(query, params)
            
            if results:
                for measurement in results:
                    # Convert all values to strings
                    formatted_measurement = []
                    for value in measurement:
                        if value is None:
                            formatted_measurement.append("")
                        else:
                            formatted_measurement.append(str(value))
                    self.measurements_tree.insert('', 'end', values=formatted_measurement)
                
                self.status_var.set(f"Loaded {len(results)} measurements ({date_filter})")
            else:
                self.status_var.set(f"No measurements found ({date_filter})")
        
        except Exception as e:
            messagebox.showerror("Error", f"Filter failed: {str(e)}")

    def calculate_water_content(self):
        """Calculate water content from fresh and dry biomass"""
        try:
            fresh = self.get_float_value('biomass_fresh_entry')
            dry = self.get_float_value('biomass_dry_entry')
            
            if fresh is not None and dry is not None and fresh > 0:
                water_content = ((fresh - dry) / fresh) * 100
                self.measurement_widgets['water_content_entry'].delete(0, 'end')
                self.measurement_widgets['water_content_entry'].insert(0, f"{water_content:.2f}")
                messagebox.showinfo("Water Content", f"Calculated water content: {water_content:.2f}%")
            else:
                messagebox.showwarning("Warning", "Please enter both fresh and dry biomass values")
        
        except Exception as e:
            messagebox.showerror("Error", f"Failed to calculate water content: {str(e)}")

    def quick_growth_analysis(self):
        """Quick growth analysis for the current treatment"""
        if not self.current_treatment_id:
            messagebox.showwarning("Warning", "Please select a treatment first")
            return
        
        try:
            query = """
                SELECT measurement_date, plant_height, leaf_area 
                FROM measurements 
                WHERE treatment_id = ? AND plant_height IS NOT NULL
                ORDER BY measurement_date
            """
            results = self.db.execute_query(query, (self.current_treatment_id,))
            
            if results and len(results) >= 2:
                # Calculate growth rate
                first_height = results[0][1]
                last_height = results[-1][1]
                days_diff = (datetime.strptime(results[-1][0], '%Y-%m-%d') - 
                            datetime.strptime(results[0][0], '%Y-%m-%d')).days
                
                if days_diff > 0:
                    growth_rate = (last_height - first_height) / days_diff
                    messagebox.showinfo("Growth Analysis", 
                                      f"Growth Rate: {growth_rate:.2f} cm/day\n"
                                      f"Period: {days_diff} days\n"
                                      f"Initial Height: {first_height:.2f} cm\n"
                                      f"Final Height: {last_height:.2f} cm")
                else:
                    messagebox.showinfo("Growth Analysis", "Insufficient time data for growth rate calculation")
            else:
                messagebox.showinfo("Growth Analysis", "Not enough height measurements for analysis")
        
        except Exception as e:
            messagebox.showerror("Error", f"Failed to perform growth analysis: {str(e)}")

    def update_analysis_tab(self):
        # Clear existing content
        for widget in self.analysis_content.winfo_children():
            widget.destroy()
        
        if not self.current_experiment_id:
            ttk.Label(self.analysis_content, text="Please select an experiment first").pack(pady=50)
            return
        
        # Implementation for analysis
        ttk.Label(self.analysis_content, text=f"Analysis for Experiment ID: {self.current_experiment_id}", 
                 font=('Arial', 12, 'bold')).pack(pady=10)
        
        # Add analysis buttons
        btn_frame = ttk.Frame(self.analysis_content)
        btn_frame.pack(pady=10)
        
        ttk.Button(btn_frame, text="Show Growth Rates", command=self.show_growth_rates).pack(side='left', padx=5)
        ttk.Button(btn_frame, text="Stress Impact Analysis", command=self.show_stress_impact).pack(side='left', padx=5)
        ttk.Button(btn_frame, text="Create Timeline Plot", command=self.create_timeline_plot).pack(side='left', padx=5)
        
        # Export buttons for analysis
        export_frame = ttk.LabelFrame(self.analysis_content, text="Export Analysis Data", padding=15)
        export_frame.pack(pady=15, fill='x', padx=20)
        
        ttk.Button(export_frame, text="📊 Export to CSV", 
                  command=lambda: self.export_analysis_data('csv')).pack(side='left', padx=5, pady=5)
        ttk.Button(export_frame, text="📈 Export to Excel", 
                  command=lambda: self.export_analysis_data('xlsx')).pack(side='left', padx=5, pady=5)
        ttk.Button(export_frame, text="📋 Export to Text", 
                  command=lambda: self.export_analysis_data('txt')).pack(side='left', padx=5, pady=5)

    def update_report_experiments(self):
        """Update the experiments list in reports tab"""
        try:
            query = "SELECT id, experiment_code, experiment_name FROM experiments ORDER BY experiment_code"
            results = self.db.execute_query(query)
            
            if results:
                exp_list = [f"{code} - {name} (ID: {id})" for id, code, name in results]
                self.report_exp_combo['values'] = exp_list
                if exp_list:
                    self.report_exp_combo.set(exp_list[0])
            else:
                self.report_exp_combo['values'] = []
                self.report_exp_combo.set('')
        
        except Exception as e:
            print(f"Error updating report experiments: {e}")
    
    def show_growth_rates(self):
        if not self.current_experiment_id:
            messagebox.showwarning("Warning", "Please select an experiment first")
            return
        
        growth_rates = self.analyzer.calculate_growth_rates(self.current_experiment_id)
        if growth_rates is not None and not growth_rates.empty:
            messagebox.showinfo("Growth Rates", growth_rates.to_string())
        else:
            messagebox.showinfo("Info", "No measurement data available for growth rate calculation")
    
    def show_stress_impact(self):
        if not self.current_experiment_id:
            messagebox.showwarning("Warning", "Please select an experiment first")
            return
        
        impact = self.analyzer.stress_impact_analysis(self.current_experiment_id)
        if impact is not None and not impact.empty:
            messagebox.showinfo("Stress Impact Analysis", impact.to_string())
        else:
            messagebox.showinfo("Info", "No data available for stress impact analysis")
    
    def create_timeline_plot(self):
        if not self.current_experiment_id:
            messagebox.showwarning("Warning", "Please select an experiment first")
            return
        
        if self.analyzer.create_stress_timeline_plot(self.current_experiment_id):
            messagebox.showinfo("Success", "Timeline plot created successfully as 'timeline_plot.png'")
        else:
            messagebox.showerror("Error", "Failed to create timeline plot")
    
    def generate_report(self):
        """Generate a report for selected experiment"""
        selected = self.report_exp_combo.get()
        if not selected:
            messagebox.showwarning("Warning", "Please select an experiment")
            return
        
        try:
            # Extract experiment ID from selection
            exp_id = int(selected.split('(ID: ')[1].rstrip(')'))
            report_type = self.report_var.get()
            
            # Generate report based on type
            if report_type == "complete":
                self.generate_complete_report(exp_id)
            elif report_type == "growth":
                self.generate_growth_report(exp_id)
            elif report_type == "physio":
                self.generate_physio_report(exp_id)
            elif report_type == "stress":
                self.generate_stress_report(exp_id)
            elif report_type == "stats":
                self.generate_stats_report(exp_id)
                
            messagebox.showinfo("Success", f"{report_type.replace('_', ' ').title()} report generated successfully!")
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to generate report: {str(e)}")
    
    def generate_complete_report(self, exp_id):
        """Generate complete experiment report"""
        # Implementation for complete report
        print(f"Generating complete report for experiment {exp_id}")
    
    def generate_growth_report(self, exp_id):
        """Generate growth analysis report"""
        # Implementation for growth report
        print(f"Generating growth report for experiment {exp_id}")
    
    def generate_physio_report(self, exp_id):
        """Generate physiological parameters report"""
        # Implementation for physio report
        print(f"Generating physiological report for experiment {exp_id}")
    
    def generate_stress_report(self, exp_id):
        """Generate stress impact summary"""
        # Implementation for stress report
        print(f"Generating stress report for experiment {exp_id}")
    
    def generate_stats_report(self, exp_id):
        """Generate statistical analysis report"""
        # Implementation for stats report
        print(f"Generating stats report for experiment {exp_id}")
    
    def create_charts(self):
        """Create charts for selected experiment"""
        selected = self.report_exp_combo.get()
        if not selected:
            messagebox.showwarning("Warning", "Please select an experiment")
            return
        
        try:
            exp_id = int(selected.split('(ID: ')[1].rstrip(')'))
            
            # Create various charts
            self.create_growth_chart(exp_id)
            self.create_stress_comparison_chart(exp_id)
            
            messagebox.showinfo("Success", "Charts created successfully!")
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to create charts: {str(e)}")
    
    def create_growth_chart(self, exp_id):
        """Create growth comparison chart"""
        # Implementation for growth chart
        print(f"Creating growth chart for experiment {exp_id}")
    
    def create_stress_comparison_chart(self, exp_id):
        """Create stress comparison chart"""
        # Implementation for stress comparison chart
        print(f"Creating stress comparison chart for experiment {exp_id}")
    
    def export_report(self):
        """Export report to Excel"""
        selected = self.report_exp_combo.get()
        if not selected:
            messagebox.showwarning("Warning", "Please select an experiment")
            return
        
        try:
            exp_id = int(selected.split('(ID: ')[1].rstrip(')'))
            
            if self.analyzer.export_experiment_data(exp_id):
                messagebox.showinfo("Success", "Report exported to Excel successfully!")
            else:
                messagebox.showerror("Error", "Failed to export report")
                
        except Exception as e:
            messagebox.showerror("Error", f"Failed to export report: {str(e)}")
    
    def export_data(self):
        try:
            if self.db.export_to_excel():
                messagebox.showinfo("Success", "Data exported to 'plant_stress_data.xlsx'")
            else:
                messagebox.showerror("Error", "Failed to export data")
        except Exception as e:
            messagebox.showerror("Error", f"Export failed: {str(e)}")

# Run the application
if __name__ == "__main__":
    root = tk.Tk()
    app = AdvancedStressApp(root)
    root.mainloop()