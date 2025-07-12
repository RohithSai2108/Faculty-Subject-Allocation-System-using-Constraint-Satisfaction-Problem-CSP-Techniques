"""
Create a sample input Excel file for the Faculty Scheduler.
This module can create either a simple test case or a more complex one
to evaluate the scheduler's performance with larger datasets.
"""
import pandas as pd
import os
import sys
import time


def create_sample_input(file_path, complex_dataset=False):
    """
    Create a sample input Excel file for testing purposes
    
    Args:
        file_path: Path where the Excel file will be saved
        complex_dataset: If True, create a more challenging dataset with more
                         subjects, faculty, and constraints
    """
    
    # Create excel writer - make sure file doesn't already exist
    if os.path.exists(file_path):
        try:
            os.remove(file_path)
            # Small delay to ensure file system completes the deletion
            time.sleep(0.5)
        except Exception as e:
            print(f"Warning: Could not remove existing file: {e}")
    
    try:
        # Use a proper context manager to ensure file is closed properly
        with pd.ExcelWriter(file_path, engine='openpyxl', mode='w') as writer:
            # Faculty data
            if not complex_dataset:
                # Simple dataset with 4 faculty members
                faculty_data = [
                    {
                        'id': 1, 
                        'name': 'Dr. Smith', 
                        'qualified_subjects': [1, 2], 
                        'max_hours': 20,
                        'day_preferences': 'Monday,Tuesday,Wednesday', 
                        'time_preferences': 'Morning', 
                        'preference_weight': 3.0,
                        'consecutive_classes_preference': 2
                    },
                    {
                        'id': 2, 
                        'name': 'Prof. Johnson', 
                        'qualified_subjects': [2, 3, 4], 
                        'max_hours': 20,
                        'day_preferences': 'Tuesday,Thursday', 
                        'time_preferences': 'Afternoon', 
                        'preference_weight': 4.0,
                        'consecutive_classes_preference': -3
                    },
                    {
                        'id': 3, 
                        'name': 'Dr. Williams', 
                        'qualified_subjects': [1, 3, 5], 
                        'max_hours': 20,
                        'day_preferences': 'Monday,Wednesday,Friday', 
                        'time_preferences': 'Morning,Afternoon', 
                        'preference_weight': 2.0,
                        'consecutive_classes_preference': 0
                    },
                    {
                        'id': 4, 
                        'name': 'Prof. Davis', 
                        'qualified_subjects': [4, 5], 
                        'max_hours': 20,
                        'day_preferences': 'Thursday,Friday', 
                        'time_preferences': 'Morning', 
                        'preference_weight': 5.0,
                        'consecutive_classes_preference': 5
                    },
                ]
            else:
                # Complex dataset with 8 faculty members with varying specializations, hours, and preferences
                faculty_data = [
                    {
                        'id': 1, 
                        'name': 'Dr. Smith', 
                        'qualified_subjects': [1, 2, 3], 
                        'max_hours': 15,
                        'day_preferences': 'Monday,Tuesday', 
                        'time_preferences': 'Morning', 
                        'preference_weight': 4.0,
                        'consecutive_classes_preference': 3
                    },
                    {
                        'id': 2, 
                        'name': 'Prof. Johnson', 
                        'qualified_subjects': [2, 3, 4, 5], 
                        'max_hours': 18,
                        'day_preferences': 'Tuesday,Thursday', 
                        'time_preferences': 'Afternoon', 
                        'preference_weight': 3.5,
                        'consecutive_classes_preference': -4
                    },
                    {
                        'id': 3, 
                        'name': 'Dr. Williams', 
                        'qualified_subjects': [1, 3, 5, 7], 
                        'max_hours': 20,
                        'day_preferences': 'Monday,Wednesday,Friday', 
                        'time_preferences': 'Morning,Afternoon', 
                        'preference_weight': 2.0,
                        'consecutive_classes_preference': 0
                    },
                    {
                        'id': 4, 
                        'name': 'Prof. Davis', 
                        'qualified_subjects': [4, 5, 6], 
                        'max_hours': 16,
                        'day_preferences': 'Wednesday,Thursday', 
                        'time_preferences': 'Morning', 
                        'preference_weight': 4.5,
                        'consecutive_classes_preference': 5
                    },
                    {
                        'id': 5, 
                        'name': 'Dr. Brown', 
                        'qualified_subjects': [1, 5, 8, 9], 
                        'max_hours': 14,
                        'day_preferences': 'Monday,Thursday,Friday', 
                        'time_preferences': 'Afternoon', 
                        'preference_weight': 3.0,
                        'consecutive_classes_preference': -2
                    },
                    {
                        'id': 6, 
                        'name': 'Prof. Miller', 
                        'qualified_subjects': [6, 7, 8], 
                        'max_hours': 18,
                        'day_preferences': 'Tuesday,Wednesday', 
                        'time_preferences': 'Morning,Afternoon', 
                        'preference_weight': 2.5,
                        'consecutive_classes_preference': 1
                    },
                    {
                        'id': 7, 
                        'name': 'Dr. Wilson', 
                        'qualified_subjects': [2, 4, 9, 10], 
                        'max_hours': 16,
                        'day_preferences': 'Wednesday,Friday', 
                        'time_preferences': 'Morning', 
                        'preference_weight': 5.0,
                        'consecutive_classes_preference': 4
                    },
                    {
                        'id': 8, 
                        'name': 'Prof. Taylor', 
                        'qualified_subjects': [3, 7, 10], 
                        'max_hours': 12,
                        'day_preferences': 'Monday,Tuesday', 
                        'time_preferences': 'Afternoon', 
                        'preference_weight': 1.5,
                        'consecutive_classes_preference': -5
                    },
                ]
            faculty_df = pd.DataFrame(faculty_data)
            # Convert list to comma-separated string for Excel
            faculty_df['qualified_subjects'] = faculty_df['qualified_subjects'].apply(lambda x: ','.join(map(str, x)))
            faculty_df.to_excel(writer, sheet_name='Faculty', index=False)
            
            # Subject data
            if not complex_dataset:
                # Simple dataset with 5 subjects (2 with labs)
                subject_data = [
                    {'id': 1, 'name': 'Introduction to Programming', 'hours': 3, 'lab_hours': 2},
                    {'id': 2, 'name': 'Data Structures', 'hours': 4, 'lab_hours': 2},
                    {'id': 3, 'name': 'Artificial Intelligence', 'hours': 3, 'lab_hours': 0},
                    {'id': 4, 'name': 'Computer Networks', 'hours': 3, 'lab_hours': 0},
                    {'id': 5, 'name': 'Database Systems', 'hours': 4, 'lab_hours': 0},
                ]
            else:
                # Complex dataset with 10 subjects (4 with labs)
                subject_data = [
                    {'id': 1, 'name': 'Introduction to Programming', 'hours': 3, 'lab_hours': 2},
                    {'id': 2, 'name': 'Data Structures', 'hours': 4, 'lab_hours': 2},
                    {'id': 3, 'name': 'Artificial Intelligence', 'hours': 3, 'lab_hours': 0},
                    {'id': 4, 'name': 'Computer Networks', 'hours': 3, 'lab_hours': 0},
                    {'id': 5, 'name': 'Database Systems', 'hours': 4, 'lab_hours': 0},
                    {'id': 6, 'name': 'Operating Systems', 'hours': 4, 'lab_hours': 2},
                    {'id': 7, 'name': 'Web Development', 'hours': 3, 'lab_hours': 2},
                    {'id': 8, 'name': 'Software Engineering', 'hours': 4, 'lab_hours': 0},
                    {'id': 9, 'name': 'Computer Architecture', 'hours': 3, 'lab_hours': 0},
                    {'id': 10, 'name': 'Machine Learning', 'hours': 4, 'lab_hours': 0},
                ]
            subject_df = pd.DataFrame(subject_data)
            subject_df.to_excel(writer, sheet_name='Subjects', index=False)
            
            # Classroom data
            if not complex_dataset:
                # Simple dataset with 4 classrooms (2 with labs)
                classroom_data = [
                    {'id': 1, 'name': 'Room 101', 'has_lab': False},
                    {'id': 2, 'name': 'Room 102', 'has_lab': False},
                    {'id': 3, 'name': 'Lab 201', 'has_lab': True},
                    {'id': 4, 'name': 'Lab 202', 'has_lab': True},
                ]
            else:
                # Complex dataset with 8 classrooms (3 with labs)
                classroom_data = [
                    {'id': 1, 'name': 'Room 101', 'has_lab': False},
                    {'id': 2, 'name': 'Room 102', 'has_lab': False},
                    {'id': 3, 'name': 'Room 103', 'has_lab': False},
                    {'id': 4, 'name': 'Room 104', 'has_lab': False},
                    {'id': 5, 'name': 'Room 105', 'has_lab': False},
                    {'id': 6, 'name': 'Lab 201', 'has_lab': True},
                    {'id': 7, 'name': 'Lab 202', 'has_lab': True},
                    {'id': 8, 'name': 'Lab 203', 'has_lab': True},
                ]
            classroom_df = pd.DataFrame(classroom_data)
            classroom_df.to_excel(writer, sheet_name='Classrooms', index=False)
            
            # Timeslot data
            if not complex_dataset:
                # Simple dataset with 6 timeslots (3 days, 2 slots per day)
                timeslot_data = [
                    {'id': 1, 'day': 'Monday', 'time': '9:00-10:30', 'period': 'Morning'},
                    {'id': 2, 'day': 'Monday', 'time': '11:00-12:30', 'period': 'Morning'},
                    {'id': 3, 'day': 'Tuesday', 'time': '9:00-10:30', 'period': 'Morning'},
                    {'id': 4, 'day': 'Tuesday', 'time': '11:00-12:30', 'period': 'Morning'},
                    {'id': 5, 'day': 'Wednesday', 'time': '9:00-10:30', 'period': 'Morning'},
                    {'id': 6, 'day': 'Wednesday', 'time': '11:00-12:30', 'period': 'Morning'},
                ]
            else:
                # Complex dataset with 15 timeslots (5 days, 3 periods per day)
                timeslot_data = [
                    {'id': 1, 'day': 'Monday', 'time': '9:00-10:30', 'period': 'Morning'},
                    {'id': 2, 'day': 'Monday', 'time': '11:00-12:30', 'period': 'Morning'},
                    {'id': 3, 'day': 'Monday', 'time': '2:00-3:30', 'period': 'Afternoon'},
                    {'id': 4, 'day': 'Monday', 'time': '4:00-5:30', 'period': 'Afternoon'},
                    {'id': 5, 'day': 'Tuesday', 'time': '9:00-10:30', 'period': 'Morning'},
                    {'id': 6, 'day': 'Tuesday', 'time': '11:00-12:30', 'period': 'Morning'},
                    {'id': 7, 'day': 'Tuesday', 'time': '2:00-3:30', 'period': 'Afternoon'},
                    {'id': 8, 'day': 'Wednesday', 'time': '9:00-10:30', 'period': 'Morning'},
                    {'id': 9, 'day': 'Wednesday', 'time': '11:00-12:30', 'period': 'Morning'},
                    {'id': 10, 'day': 'Wednesday', 'time': '2:00-3:30', 'period': 'Afternoon'},
                    {'id': 11, 'day': 'Thursday', 'time': '9:00-10:30', 'period': 'Morning'},
                    {'id': 12, 'day': 'Thursday', 'time': '11:00-12:30', 'period': 'Morning'},
                    {'id': 13, 'day': 'Thursday', 'time': '2:00-3:30', 'period': 'Afternoon'},
                    {'id': 14, 'day': 'Friday', 'time': '9:00-10:30', 'period': 'Morning'},
                    {'id': 15, 'day': 'Friday', 'time': '2:00-3:30', 'period': 'Afternoon'},
                ]
            timeslot_df = pd.DataFrame(timeslot_data)
            timeslot_df.to_excel(writer, sheet_name='Timeslots', index=False)
            
            # No need to explicitly close the writer when using context manager
            
        # Additional safeguard: verify that the file exists and is not empty
        if not os.path.exists(file_path) or os.path.getsize(file_path) == 0:
            raise Exception(f"Failed to create file {file_path} or file is empty")
            
    except Exception as e:
        print(f"Error creating sample input file: {e}")
        return False
        
    return True


if __name__ == "__main__":
    # Check for command line arguments
    complex_dataset = '--complex' in sys.argv
    
    file_path = "sample_input.xlsx"
    if complex_dataset:
        print("Creating complex sample dataset...")
    else:
        print("Creating simple sample dataset...")
        
    success = create_sample_input(file_path, complex_dataset)
    if success:
        print(f"Full path: {os.path.abspath(file_path)}")
        print(f"File size: {os.path.getsize(file_path)} bytes")
    else:
        print("Failed to create sample input file")