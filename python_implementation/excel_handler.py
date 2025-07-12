"""
Excel handler for faculty scheduling.
This module provides functionality to read from and write to Excel files.
"""
import os
import pandas as pd
from typing import Dict, List, Tuple, Any


class ExcelHandler:
    """
    Handler for Excel file operations related to faculty scheduling.
    """

    @staticmethod
    def read_faculty_data(file_path: str) -> List[Dict[str, Any]]:
        """
        Read faculty data from an Excel file.
        
        Args:
            file_path: Path to the Excel file
            
        Returns:
            List of dictionaries containing faculty information including preferences
        """
        try:
            # Explicitly specify engine to avoid issues
            df = pd.read_excel(file_path, sheet_name='Faculty', engine='openpyxl')
            faculty_data = df.to_dict('records')
            
            # Process and handle faculty preferences
            for faculty in faculty_data:
                # Add default preferences if not present
                if 'day_preferences' not in faculty or pd.isna(faculty['day_preferences']):
                    faculty['day_preferences'] = "Monday,Tuesday,Wednesday,Thursday,Friday"
                    
                if 'time_preferences' not in faculty or pd.isna(faculty['time_preferences']):
                    faculty['time_preferences'] = "Morning,Afternoon"
                    
                if 'preference_weight' not in faculty or pd.isna(faculty['preference_weight']):
                    faculty['preference_weight'] = 1.0  # Default preference weight (0-5 scale)
                
                if 'consecutive_classes_preference' not in faculty or pd.isna(faculty['consecutive_classes_preference']):
                    faculty['consecutive_classes_preference'] = 0  # 0=no preference, -5=avoid, 5=prefer
                
                # Convert string preferences to lists
                if isinstance(faculty.get('day_preferences'), str):
                    faculty['day_preferences'] = [d.strip() for d in faculty['day_preferences'].split(',')]
                    
                if isinstance(faculty.get('time_preferences'), str):
                    faculty['time_preferences'] = [t.strip() for t in faculty['time_preferences'].split(',')]
                
                # Convert qualified_subjects to list if it's a string
                if 'qualified_subjects' in faculty and isinstance(faculty['qualified_subjects'], str):
                    faculty['qualified_subjects'] = [int(s.strip()) for s in faculty['qualified_subjects'].split(',') if s.strip()]
                
            return faculty_data
        except Exception as e:
            print(f"Error reading faculty data: {e}")
            return []
            
    @staticmethod
    def read_subject_data(file_path: str) -> List[Dict[str, Any]]:
        """
        Read subject data from an Excel file.
        
        Args:
            file_path: Path to the Excel file
            
        Returns:
            List of dictionaries containing subject information
        """
        try:
            # Explicitly specify engine to avoid issues
            df = pd.read_excel(file_path, sheet_name='Subjects', engine='openpyxl')
            return df.to_dict('records')
        except Exception as e:
            print(f"Error reading subject data: {e}")
            return []
            
    @staticmethod
    def read_classroom_data(file_path: str) -> List[Dict[str, Any]]:
        """
        Read classroom data from an Excel file.
        
        Args:
            file_path: Path to the Excel file
            
        Returns:
            List of dictionaries containing classroom information
        """
        try:
            # Explicitly specify engine to avoid issues
            df = pd.read_excel(file_path, sheet_name='Classrooms', engine='openpyxl')
            return df.to_dict('records')
        except Exception as e:
            print(f"Error reading classroom data: {e}")
            return []
            
    @staticmethod
    def read_timeslot_data(file_path: str) -> List[Dict[str, Any]]:
        """
        Read timeslot data from an Excel file.
        
        Args:
            file_path: Path to the Excel file
            
        Returns:
            List of dictionaries containing timeslot information
        """
        try:
            # Explicitly specify engine to avoid issues
            df = pd.read_excel(file_path, sheet_name='Timeslots', engine='openpyxl')
            return df.to_dict('records')
        except Exception as e:
            print(f"Error reading timeslot data: {e}")
            return []
    
    @staticmethod
    def write_schedule(file_path: str, schedule_data: List[Dict[str, Any]]) -> bool:
        """
        Write schedule data to an Excel file.
        
        Args:
            file_path: Path to the Excel file
            schedule_data: List of dictionaries containing schedule information
            
        Returns:
            bool: True if the write operation was successful, False otherwise
        """
        try:
            # Remove the file if it already exists
            if os.path.exists(file_path):
                os.remove(file_path)
                
            df = pd.DataFrame(schedule_data)
            # Explicitly specify engine to avoid issues
            with pd.ExcelWriter(file_path, engine='openpyxl') as writer:
                df.to_excel(writer, index=False)
            print(f"Schedule successfully written to {file_path}")
            return True
        except Exception as e:
            print(f"Error writing schedule data: {e}")
            return False
