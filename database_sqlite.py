# database_sqlite.py - UPDATED VERSION
import sqlite3
import pandas as pd
from datetime import datetime
import os

class StressDatabase:
    def __init__(self):
        self.db_file = "plant_stress.db"
        self.connection = None
        self.cursor = None
    
    def create_database(self):
        """Create SQLite database and tables"""
        try:
            self.connection = sqlite3.connect(self.db_file)
            self.cursor = self.connection.cursor()
            print(f"Using SQLite database: {self.db_file}")
            
            # Enable foreign keys
            self.cursor.execute("PRAGMA foreign_keys = ON")
            
            # Create experiments table
            experiments_table = """
            CREATE TABLE IF NOT EXISTS experiments (
                id INTEGER PRIMARY KEY,
                experiment_code TEXT UNIQUE NOT NULL,
                experiment_name TEXT NOT NULL,
                plant_species TEXT NOT NULL,
                stress_type TEXT NOT NULL,
                researcher TEXT,
                start_date TEXT,
                end_date TEXT,
                description TEXT,
                status TEXT DEFAULT 'active',
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
            )
            """
            self.cursor.execute(experiments_table)
            print("✅ Table 1 (experiments) created successfully")
            
            # Create treatments table
            treatments_table = """
            CREATE TABLE IF NOT EXISTS treatments (
                id INTEGER PRIMARY KEY,
                experiment_id INTEGER NOT NULL,
                treatment_name TEXT NOT NULL,
                treatment_type TEXT NOT NULL,
                stress_level TEXT,
                concentration REAL,
                duration_days INTEGER,
                temperature REAL,
                description TEXT,
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                FOREIGN KEY (experiment_id) REFERENCES experiments (id) ON DELETE CASCADE,
                UNIQUE(experiment_id, treatment_name)
            )
            """
            self.cursor.execute(treatments_table)
            print("✅ Table 2 (treatments) created successfully")
            
            # Create measurements table
            measurements_table = """
            CREATE TABLE IF NOT EXISTS measurements (
                id INTEGER PRIMARY KEY,
                treatment_id INTEGER NOT NULL,
                measurement_date TEXT NOT NULL,
                plant_height REAL,
                leaf_area REAL,
                chlorophyll_content REAL,
                photosynthesis_rate REAL,
                stomatal_conductance REAL,
                root_length REAL,
                biomass_fresh REAL,
                biomass_dry REAL,
                water_content REAL,
                notes TEXT,
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                FOREIGN KEY (treatment_id) REFERENCES treatments (id) ON DELETE CASCADE
            )
            """
            self.cursor.execute(measurements_table)
            print("✅ Table 3 (measurements) created successfully")
            
            self.connection.commit()
            print("✅ SQLite database connected successfully!")
            return True
            
        except sqlite3.Error as e:
            print(f"❌ Database error: {e}")
            return False
    
    def execute_query(self, query, params=None):
        """Execute a query and return results"""
        try:
            if params:
                self.cursor.execute(query, params)
            else:
                self.cursor.execute(query)
            
            # For SELECT queries, return results
            if query.strip().upper().startswith('SELECT'):
                return self.cursor.fetchall()
            else:
                # For INSERT, UPDATE, DELETE - commit changes
                self.connection.commit()
                return True
                
        except sqlite3.IntegrityError as e:
            if "UNIQUE constraint failed" in str(e):
                print(f"❌ Integrity error: Duplicate entry not allowed")
                return "DUPLICATE"
            else:
                print(f"❌ Integrity error: {e}")
                return False
        except sqlite3.Error as e:
            print(f"❌ Query execution error: {e}")
            print(f"Failed query: {query}")
            if params:
                print(f"With params: {params}")
            return False
    
    def export_to_excel(self):
        """Export all data to Excel file"""
        try:
            # Get experiments data
            experiments_df = pd.read_sql_query("SELECT * FROM experiments", self.connection)
            
            # Get treatments data
            treatments_df = pd.read_sql_query("SELECT * FROM treatments", self.connection)
            
            # Get measurements data
            measurements_df = pd.read_sql_query("SELECT * FROM measurements", self.connection)
            
            # Create Excel writer
            with pd.ExcelWriter('plant_stress_data.xlsx', engine='openpyxl') as writer:
                experiments_df.to_excel(writer, sheet_name='Experiments', index=False)
                treatments_df.to_excel(writer, sheet_name='Treatments', index=False)
                measurements_df.to_excel(writer, sheet_name='Measurements', index=False)
            
            print("✅ Data exported to plant_stress_data.xlsx")
            return True
            
        except Exception as e:
            print(f"❌ Export error: {e}")
            return False
    
    def close_connection(self):
        """Close database connection"""
        if self.connection:
            self.connection.close()
            print("✅ Database connection closed")