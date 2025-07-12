"""
Main module for faculty scheduling application.
This module provides a command-line interface for the faculty scheduling system.

Usage:
    python main.py [--verbose] [--complex] [--generate-sample]
    
Options:
    --verbose         Enable verbose debug output
    --complex         Use a more complex dataset for testing
    --generate-sample Generate a sample input file and use it
"""
import os
import sys
from faculty_scheduler import FacultyScheduler
from create_sample_input import create_sample_input


def display_banner():
    """Display application banner"""
    print("=" * 80)
    print("Faculty Subject Allocation and Scheduling System")
    print("Constraint Satisfaction Problem Solver")
    print("=" * 80)
    print()


def get_input_file() -> str:
    """
    Get the input Excel file from the user.
    
    Returns:
        Path to the input Excel file
    """
    while True:
        file_path = input("Enter the path to the input Excel file: ").strip()
        if os.path.exists(file_path) and file_path.endswith(('.xlsx', '.xls')):
            return file_path
        print("Error: Invalid file path or not an Excel file. Please try again.")


def get_output_file() -> str:
    """
    Get the output Excel file from the user.
    
    Returns:
        Path to the output Excel file
    """
    while True:
        file_path = input("Enter the path for the output Excel file: ").strip()
        if file_path.endswith(('.xlsx', '.xls')):
            return file_path
        print("Error: Output file must have .xlsx or .xls extension. Please try again.")


def main():
    """Main function that runs the faculty scheduling application"""
    display_banner()
    
    # Parse command line arguments
    verbose = '--verbose' in sys.argv
    complex_dataset = '--complex' in sys.argv
    generate_sample = '--generate-sample' in sys.argv
    
    try:
        # Get input file or generate sample
        if generate_sample:
            print("Generating sample input file...")
            sample_file = "sample_input.xlsx"
            create_sample_input(sample_file, complex_dataset)
            input_file = sample_file
            print(f"Using generated sample file: {input_file}")
        else:
            input_file = get_input_file()
        
        # Create scheduler and load data
        scheduler = FacultyScheduler(verbose=verbose)
        if not scheduler.load_data(input_file):
            print("Error: Failed to load data from the input file.")
            return 1
            
        if verbose:
            print("Running in verbose mode - detailed debug information will be displayed")
            
        print("Data loaded successfully!")
        
        # In verbose mode, display detailed data info
        if verbose:
            print("\nFaculty:")
            for f in scheduler.faculty:
                qualified_str = f.get('qualified_subjects', '')
                if isinstance(qualified_str, str):
                    qualified_str = [int(s.strip()) for s in qualified_str.split(',') if s.strip()]
                print(f"ID: {f['id']}, Name: {f['name']}, Max Hours: {f.get('max_hours', 0)}, Qualified: {qualified_str}")
                
            print("\nSubjects:")
            for s in scheduler.subjects:
                print(f"ID: {s['id']}, Name: {s['name']}, Hours: {s.get('hours', 0)}, Lab Hours: {s.get('lab_hours', 0)}")
                
            print("\nClassrooms:")
            for c in scheduler.classrooms:
                print(f"ID: {c['id']}, Name: {c['name']}, Has Lab: {c.get('has_lab', False)}")
                
            print("\nTimeslots:")
            for t in scheduler.timeslots:
                print(f"ID: {t['id']}, Day: {t['day']}, Time: {t['time']}")
        else:
            # Show only summary in non-verbose mode
            print(f"Loaded {len(scheduler.faculty)} faculty, {len(scheduler.subjects)} subjects, " +
                  f"{len(scheduler.classrooms)} classrooms, and {len(scheduler.timeslots)} timeslots.")
        
        # Solve the scheduling problem
        print("\nSolving the faculty scheduling problem...")
        
        # Try direct construction approach first (faster)
        print("Trying direct construction approach first...")
        if scheduler.direct_solve():
            print("Solution found using direct construction!")
        else:
            # Fall back to CSP solver if direct construction fails
            print("\nFalling back to CSP solver...")
            if not scheduler.solve():
                print("Failed to find a solution that satisfies all constraints.")
                return 1
        
        print("Solution found!")
        
        # Display the schedule
        print("\nFaculty Schedule:")
        scheduler.display_schedule()
        
        # Get output file and save the schedule
        output_file = get_output_file()
        if scheduler.save_schedule(output_file):
            print(f"Schedule successfully saved to {output_file}")
        else:
            print("Error: Failed to save the schedule.")
            return 1
            
        return 0
    
    except KeyboardInterrupt:
        print("\nOperation cancelled by user.")
        return 1
    except Exception as e:
        print(f"Error: {e}")
        return 1


if __name__ == "__main__":
    sys.exit(main())
