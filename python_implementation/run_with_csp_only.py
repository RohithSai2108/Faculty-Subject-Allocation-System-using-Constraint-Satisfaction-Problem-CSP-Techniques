import os
import sys
import threading
from faculty_scheduler import FacultyScheduler
from create_sample_input import create_sample_input


def run_with_csp_only():
    """Run the faculty scheduler with only the CSP solver approach"""
    # Check for verbose flag
    verbose = '--verbose' in sys.argv

    print("=" * 80)
    print("Faculty Subject Allocation and Scheduling System")
    print("CSP Solver Only - Traditional Approach")
    print("=" * 80)
    print()

    try:
        # Check for complex dataset flag
        complex_dataset = '--complex' in sys.argv

        # Generate a sample input file with appropriate complexity
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
        print(f"Loaded {len(scheduler.faculty)} faculty, {len(scheduler.subjects)} subjects, " +
              f"{len(scheduler.classrooms)} classrooms, and {len(scheduler.timeslots)} timeslots.")

        # Skip direct construction and use pure CSP solver
        print("\nUsing traditional CSP solver (no direct construction)...")
        print("Note: Pure CSP solving can be very slow for larger problems due to combinatorial explosion")
        print("Setting a time limit of 20 seconds for the CSP solver...")

        # Run the CSP solver in a separate thread with a timeout
        solution_found = [False]  # Use a mutable object to track the result

        def solve_csp():
            if scheduler.solve():
                solution_found[0] = True

        solver_thread = threading.Thread(target=solve_csp)
        solver_thread.start()
        solver_thread.join(timeout=20)  # Wait for 20 seconds

        if solver_thread.is_alive():
            print("\nCSP solver timed out after 20 seconds.")
            print("The traditional CSP approach timed out due to combinatorial explosion.")
            print("This demonstrates why our direct construction approach is necessary!")
            return 1

        if not solution_found[0]:
            print("CSP solver failed to find a solution that satisfies all constraints.")
            print("Consider simplifying the dataset or increasing the timeout.")
            return 1

        print("Solution found!")

        # Display the schedule
        print("\nFaculty Schedule:")
        scheduler.display_schedule()

        # Save the schedule to output file
        output_file = "output_schedule_csp.xlsx"
        if scheduler.save_schedule(output_file):
            print(f"Schedule successfully saved to {output_file}")
        else:
            print("Error: Failed to save the schedule.")
            return 1

        return 0

    except Exception as e:
        print(f"Error: {e}")
        return 1


if __name__ == "__main__":
    run_with_csp_only()