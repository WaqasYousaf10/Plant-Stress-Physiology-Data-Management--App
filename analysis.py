# analysis.py - UPDATED FOR SQLITE
import pandas as pd
import matplotlib.pyplot as plt
import numpy as np
from datetime import datetime

class StressAnalyzer:
    def __init__(self, database):
        self.db = database
    
    def calculate_growth_rates(self, experiment_id):
        """Calculate growth rates for all treatments in an experiment"""
        try:
            query = """
                SELECT t.treatment_name, m.measurement_date, m.plant_height, m.leaf_area, m.biomass_fresh
                FROM measurements m
                JOIN treatments t ON m.treatment_id = t.id
                WHERE t.experiment_id = ?
                ORDER BY t.treatment_name, m.measurement_date
            """
            results = self.db.execute_query(query, (experiment_id,))
            
            if results:
                df = pd.DataFrame(results, columns=['treatment_name', 'measurement_date', 'plant_height', 'leaf_area', 'biomass'])
                
                # Calculate growth rates
                growth_rates = df.groupby('treatment_name').agg({
                    'plant_height': ['mean', 'std'],
                    'leaf_area': ['mean', 'std'],
                    'biomass': ['mean', 'std']
                }).round(2)
                
                return growth_rates
            else:
                return None
                
        except Exception as e:
            print(f"Error calculating growth rates: {e}")
            return None
    
    def stress_impact_analysis(self, experiment_id):
        """Analyze stress impact by comparing treatments"""
        try:
            query = """
                SELECT t.treatment_name, t.treatment_type, t.stress_level,
                       AVG(m.plant_height) as avg_height,
                       AVG(m.leaf_area) as avg_leaf_area,
                       AVG(m.water_content) as avg_water_content,
                       COUNT(m.id) as measurement_count
                FROM treatments t
                LEFT JOIN measurements m ON t.id = m.treatment_id
                WHERE t.experiment_id = ?
                GROUP BY t.id, t.treatment_name, t.treatment_type, t.stress_level
                ORDER BY t.treatment_type, t.stress_level
            """
            results = self.db.execute_query(query, (experiment_id,))
            
            if results:
                df = pd.DataFrame(results, columns=[
                    'treatment_name', 'treatment_type', 'stress_level',
                    'avg_height', 'avg_leaf_area', 'avg_water_content', 'measurement_count'
                ])
                return df.round(2)
            else:
                return None
                
        except Exception as e:
            print(f"Error in stress impact analysis: {e}")
            return None
    
    def create_stress_timeline_plot(self, experiment_id):
        """Create a timeline plot showing stress development"""
        try:
            query = """
                SELECT t.treatment_name, m.measurement_date, m.plant_height, m.water_content
                FROM measurements m
                JOIN treatments t ON m.treatment_id = t.id
                WHERE t.experiment_id = ?
                ORDER BY m.measurement_date
            """
            results = self.db.execute_query(query, (experiment_id,))
            
            if results:
                df = pd.DataFrame(results, columns=['treatment_name', 'measurement_date', 'plant_height', 'water_content'])
                df['measurement_date'] = pd.to_datetime(df['measurement_date'])
                
                # Create plot
                fig, (ax1, ax2) = plt.subplots(2, 1, figsize=(12, 8))
                
                # Plot 1: Plant height over time
                for treatment in df['treatment_name'].unique():
                    treatment_data = df[df['treatment_name'] == treatment]
                    ax1.plot(treatment_data['measurement_date'], treatment_data['plant_height'], 
                            marker='o', label=treatment, linewidth=2)
                
                ax1.set_title('Plant Height Over Time')
                ax1.set_ylabel('Height (cm)')
                ax1.legend()
                ax1.grid(True, alpha=0.3)
                
                # Plot 2: Water content over time
                for treatment in df['treatment_name'].unique():
                    treatment_data = df[df['treatment_name'] == treatment]
                    ax2.plot(treatment_data['measurement_date'], treatment_data['water_content'], 
                            marker='s', label=treatment, linewidth=2)
                
                ax2.set_title('Water Content Over Time')
                ax2.set_ylabel('Water Content (%)')
                ax2.set_xlabel('Date')
                ax2.legend()
                ax2.grid(True, alpha=0.3)
                
                plt.tight_layout()
                plt.savefig('timeline_plot.png', dpi=300, bbox_inches='tight')
                plt.close()
                
                return True
            else:
                return False
                
        except Exception as e:
            print(f"Error creating timeline plot: {e}")
            return False
    
    def export_experiment_data(self, experiment_id):
        """Export experiment data to Excel"""
        try:
            # Get experiment details
            exp_query = "SELECT * FROM experiments WHERE id = ?"
            exp_data = self.db.execute_query(exp_query, (experiment_id,))
            
            # Get treatments data
            treatments_query = "SELECT * FROM treatments WHERE experiment_id = ?"
            treatments_data = self.db.execute_query(treatments_query, (experiment_id,))
            
            # Get measurements data
            measurements_query = """
                SELECT m.*, t.treatment_name 
                FROM measurements m
                JOIN treatments t ON m.treatment_id = t.id
                WHERE t.experiment_id = ?
            """
            measurements_data = self.db.execute_query(measurements_query, (experiment_id,))
            
            # Create DataFrames
            if exp_data:
                exp_df = pd.DataFrame(exp_data, columns=[
                    'id', 'experiment_code', 'experiment_name', 'plant_species', 
                    'stress_type', 'researcher', 'start_date', 'end_date', 
                    'description', 'status', 'created_at', 'updated_at'
                ])
            else:
                exp_df = pd.DataFrame()
            
            if treatments_data:
                treatments_df = pd.DataFrame(treatments_data, columns=[
                    'id', 'experiment_id', 'treatment_name', 'treatment_type',
                    'stress_level', 'concentration', 'duration_days', 'temperature',
                    'description', 'created_at', 'updated_at'
                ])
            else:
                treatments_df = pd.DataFrame()
            
            if measurements_data:
                measurements_df = pd.DataFrame(measurements_data, columns=[
                    'id', 'treatment_id', 'measurement_date', 'plant_height', 
                    'leaf_area', 'chlorophyll_content', 'photosynthesis_rate',
                    'stomatal_conductance', 'root_length', 'biomass_fresh',
                    'biomass_dry', 'water_content', 'notes', 'created_at', 'treatment_name'
                ])
            else:
                measurements_df = pd.DataFrame()
            
            # Create Excel file
            filename = f'experiment_{experiment_id}_data.xlsx'
            with pd.ExcelWriter(filename, engine='openpyxl') as writer:
                if not exp_df.empty:
                    exp_df.to_excel(writer, sheet_name='Experiment_Details', index=False)
                if not treatments_df.empty:
                    treatments_df.to_excel(writer, sheet_name='Treatments', index=False)
                if not measurements_df.empty:
                    measurements_df.to_excel(writer, sheet_name='Measurements', index=False)
            
            print(f"✅ Experiment data exported to {filename}")
            return True
            
        except Exception as e:
            print(f"❌ Export error: {e}")
            return False
    
    def calculate_statistics(self, experiment_id):
        """Calculate comprehensive statistics for an experiment"""
        try:
            query = """
                SELECT t.treatment_name, 
                       COUNT(m.id) as total_measurements,
                       AVG(m.plant_height) as avg_height,
                       STDDEV(m.plant_height) as std_height,
                       AVG(m.leaf_area) as avg_leaf_area,
                       AVG(m.water_content) as avg_water_content,
                       AVG(m.chlorophyll_content) as avg_chlorophyll,
                       MIN(m.measurement_date) as first_date,
                       MAX(m.measurement_date) as last_date
                FROM treatments t
                LEFT JOIN measurements m ON t.id = m.treatment_id
                WHERE t.experiment_id = ?
                GROUP BY t.id, t.treatment_name
            """
            results = self.db.execute_query(query, (experiment_id,))
            
            if results:
                df = pd.DataFrame(results, columns=[
                    'treatment_name', 'total_measurements', 'avg_height', 'std_height',
                    'avg_leaf_area', 'avg_water_content', 'avg_chlorophyll',
                    'first_date', 'last_date'
                ])
                return df.round(3)
            else:
                return None
                
        except Exception as e:
            print(f"Error calculating statistics: {e}")
            return None