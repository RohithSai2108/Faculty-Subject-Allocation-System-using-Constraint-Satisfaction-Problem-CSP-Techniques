"""
Constraint Satisfaction Problem (CSP) Solver for Faculty Scheduling.
This module implements a generic CSP solver that can be used for faculty scheduling problems.
"""
from typing import Dict, List, Tuple, Set, Callable, Any, Optional


class CSPSolver:
    """
    A generic Constraint Satisfaction Problem solver using backtracking with forward checking
    and minimum remaining values (MRV) heuristic for variable selection.
    """
    
    def __init__(self):
        self.variables = []  # List of variables to assign
        self.domains = {}    # Dictionary mapping variables to their domains
        self.constraints = []  # List of constraint functions
        self.assignment = {}  # Current assignment of variables to values
        
    def add_variable(self, variable: Any, domain: List[Any]) -> None:
        """
        Add a variable to the CSP.
        
        Args:
            variable: The variable to add
            domain: List of possible values for the variable
        """
        self.variables.append(variable)
        self.domains[variable] = domain.copy()
        
    def add_constraint(self, constraint: Callable) -> None:
        """
        Add a constraint to the CSP.
        
        Args:
            constraint: A function that takes the current assignment and returns True if satisfied
        """
        self.constraints.append(constraint)
        
    def is_consistent(self, variable: Any, value: Any, assignment: Dict[Any, Any]) -> bool:
        """
        Check if assigning value to variable is consistent with the current assignment.
        
        Args:
            variable: The variable to assign
            value: The value to assign
            assignment: The current assignment
            
        Returns:
            bool: True if the assignment is consistent with all constraints
        """
        # Create a temporary assignment with the new variable-value pair
        test_assignment = assignment.copy()
        test_assignment[variable] = value
        
        # Check all constraints
        for constraint in self.constraints:
            if not constraint(test_assignment):
                return False
                
        return True
    
    def select_unassigned_variable(self) -> Optional[Any]:
        """
        Select an unassigned variable using the Minimum Remaining Values (MRV) heuristic
        and degree heuristic as a tie-breaker.
        
        Returns:
            The next variable to assign or None if all variables are assigned
        """
        unassigned = [v for v in self.variables if v not in self.assignment]
        if not unassigned:
            return None
        
        # Function to count constraints that involve a variable
        def count_constraints(var):
            count = 0
            for other_var in unassigned:
                if other_var != var:
                    # Check if these variables share any constraints
                    # For faculty scheduling, variables are now (subject_id, timeslot_id) tuples
                    # Variables interact if they share:
                    # - Same subject ID (would need to check faculty qualifications)
                    # - Same timeslot (for both faculty and classroom constraints)
                    if var[0] == other_var[0] or var[1] == other_var[1]:
                        count += 1
            return count
        
        # First use MRV heuristic - choose variable with fewest valid values in its domain
        # Then, break ties using degree heuristic - choose variable involved in most constraints
        return min(unassigned, 
                  key=lambda var: (len(self.domains[var]), -count_constraints(var)))
    
    def order_domain_values(self, var: Any) -> List[Any]:
        """
        Order domain values to try promising values first.
        For faculty scheduling, prefer:
        1. Faculty with more available hours
        2. Faculty with stronger preferences for the timeslot
        
        Args:
            var: The variable whose domain we're ordering
            
        Returns:
            Ordered list of values from the domain
        """
        # For faculty scheduling, the variable is (subject_id, timeslot_id)
        # and the values are (faculty_id, classroom_id) tuples
        
        # Just return the domain for now - we're already sorting faculty in faculty_scheduler.py
        # when creating the domains
        return self.domains[var]
    
    def forward_check(self, var: Any, value: Any) -> bool:
        """
        Simple forward checking to identify dead-ends early.
        Check if any unassigned variable has its domain emptied by this assignment.
        
        Args:
            var: The variable being assigned
            value: The value being assigned to var
            
        Returns:
            bool: False if making this assignment leads to a domain wipeout, True otherwise
        """
        # Create a temporary assignment
        temp_assignment = self.assignment.copy()
        temp_assignment[var] = value
        
        # Print debug info for this assignment
        if len(temp_assignment) == 11:  # Use the variable count to identify where we're getting stuck
            subject_id, timeslot_id = var
            faculty_id, classroom_id = value
            print(f"\nTrying to assign Subject {subject_id} at Timeslot {timeslot_id} to Faculty {faculty_id} in Classroom {classroom_id}")
            
        # Check domains of unassigned variables
        wipeouts = 0
        for other_var in self.variables:
            if other_var not in temp_assignment:
                # Check if there's at least one consistent value in the domain
                consistent_values = False
                for other_value in self.domains[other_var]:
                    if self.is_consistent(other_var, other_value, temp_assignment):
                        consistent_values = True
                        break
                
                # If no consistent values, this assignment leads to a dead-end
                if not consistent_values:
                    # Print debug info if we've reached our "stuck" point
                    if len(temp_assignment) == 11:
                        other_subject_id, other_timeslot_id = other_var
                        print(f"  DOMAIN WIPEOUT: Variable (Subject {other_subject_id}, Timeslot {other_timeslot_id}) has no valid values left")
                        wipeouts += 1
                    
                    # After reporting several wipeouts, exit to avoid excessive output
                    if wipeouts > 5:
                        return False
        
        # If we detected wipeouts but are in the debug section, finish the report
        if wipeouts > 0:
            print(f"  Total domain wipeouts: {wipeouts}")
            return False
            
        return True
    
    def backtrack(self) -> Optional[Dict[Any, Any]]:
        """
        Backtracking search for a solution to the CSP with forward checking.
        
        Returns:
            A complete assignment that satisfies all constraints, or None if no solution exists
        """
        # Check if assignment is complete
        if len(self.assignment) == len(self.variables):
            return self.assignment
            
        # Select unassigned variable
        var = self.select_unassigned_variable()
        
        if var is None:
            print("Warning: No unassigned variable but assignment is not complete")
            return None
            
        # Debug info - show current progress
        if len(self.assignment) % 10 == 0:
            # Use carriage return to overwrite the line instead of printing a new one
            print(f"CSP solving progress: {len(self.assignment)}/{len(self.variables)} variables assigned", end="\r")
        
        # Get ordered domain values 
        ordered_values = self.order_domain_values(var)
        
        # Try each value in the domain
        for value in ordered_values:
            # Check consistency with current assignment
            if self.is_consistent(var, value, self.assignment):
                # Perform forward checking
                if self.forward_check(var, value):
                    # Assign value to variable
                    self.assignment[var] = value
                    
                    # Recursive call
                    result = self.backtrack()
                    if result is not None:
                        return result
                        
                    # If we reach here, we need to backtrack
                    del self.assignment[var]
        
        # If we tried all values in the domain and none worked, report it for debugging
        # Only report backtracking occasionally to avoid excessive output
        if len(self.assignment) < 5 and len(self.assignment) % 10 == 0:  
            subject_id, timeslot_id = var
            print(f"Backtracking at variable: Subject {subject_id}, Timeslot {timeslot_id}...", end="\r")
                
        return None
    
    def solve(self) -> Optional[Dict[Any, Any]]:
        """
        Solve the CSP and return a solution if one exists.
        
        Returns:
            A complete assignment that satisfies all constraints, or None if no solution exists
        """
        self.assignment = {}
        return self.backtrack()
