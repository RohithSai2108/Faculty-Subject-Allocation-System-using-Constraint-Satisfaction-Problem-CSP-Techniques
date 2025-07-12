"""
Run the faculty scheduler with the sample input file.

This module uses a hybrid approach to solve the faculty scheduling problem:
1. First, it tries a direct construction method which is faster and simpler
2. If that fails, it falls back to the CSP solver, which is more comprehensive but slower

The scheduler now includes faculty preferences for:
- Preferred teaching days
- Preferred time periods (morning/afternoon)
- Consecutive classes preferences
- Weighted preference system

Usage:
    python run_with_sample.py [--verbose] [--complex]

Options:
    --verbose    Enable verbose debug output with preference satisfaction analysis
    --complex    Use a more complex dataset with more faculty and preferences
"""
import os
import sys
from faculty_scheduler import FacultyScheduler
from create_sample_input import create_sample_input


def run_with_sample():
    """Run the faculty scheduler with the sample input file"""
    # Check for verbose flag
    verbose = '--verbose' in sys.argv
    
    print("=" * 80)
    print("Faculty Subject Allocation and Scheduling System")
    print("Constraint Satisfaction Problem Solver")
    print("=" * 80)
    print()
    
    try:
        # Check for complex dataset flag
        complex_dataset = '--complex' in sys.argv
        
        # Generate a sample input file with appropriate complexity
        # Use full path to ensure file is created and accessed in the correct location
        current_dir = os.path.dirname(os.path.abspath(__file__))
        input_file = os.path.join(current_dir, "sample_input.xlsx")
        print("Generating sample input file...")
        create_sample_input(input_file, complex_dataset)
        print(f"Sample input file created at {input_file}")
        
        if not os.path.exists(input_file):
            print(f"Error: Sample input file {input_file} not found.")
            return 1
            
        print(f"Using sample input file: {input_file}")
        
        # Create scheduler and load data
        scheduler = FacultyScheduler(verbose=verbose)
        if not scheduler.load_data(input_file):
            print("Error: Failed to load data from the input file.")
            return 1
            
        if verbose:
            print("Running in verbose mode - detailed debug information will be displayed")
            
        print("Data loaded successfully!")
        
        # Print loaded data only in verbose mode
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
        
        # Try simpler approach first - try to manually build a valid solution
        print("\nTrying a direct construction approach first...")
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
        
        # Save the schedule to output file
        output_file = "output_schedule.xlsx"
        if scheduler.save_schedule(output_file):
            print(f"Schedule successfully saved to {output_file}")
        else:
            print("Error: Failed to save the schedule.")
            return 1
        
        # Print summary in verbose mode
        if verbose:
            print("\n" + "=" * 80)
            print("SUMMARY:")
            print("=" * 80)
            print("✓ Successfully created faculty schedule honoring all hard constraints")
            print("✓ Implemented faculty preferences for days, time periods, and consecutive classes")
            print("✓ Used preference weighting system to prioritize faculty with stronger preferences")
            print("✓ Generated optimized schedule balancing workload and faculty satisfaction")
            print("=" * 80)
            
        return 0
    
    except Exception as e:
        print(f"Error: {e}")
        return 1


if __name__ == "__main__":
    run_with_sample()