"""
Faculty Scheduler module.
This module implements the faculty scheduling problem as a constraint satisfaction problem.

The scheduler handles these key constraints:
1. Faculty requirements (max teaching hours) are consistently addressed
2. No conflicts with faculty teaching different subjects at the same time
3. No collisions in classrooms (same classroom not used for different subjects at the same time)
4. Special handling for subjects with lab requirements
5. Faculty day preferences (which days they prefer to teach)
6. Faculty time period preferences (morning/afternoon/evening preferences)
7. Consecutive classes preferences (some faculty prefer/avoid back-to-back teaching)
8. Weighted preference scoring system (different faculty may have stronger preferences)

Two solution approaches are implemented:
1. Direct construction method (faster, greedy approach)
2. CSP solver with backtracking (more comprehensive but slower)

The system now includes a preference scoring system that optimizes faculty satisfaction
while still meeting all hard constraints.
"""
from typing import Dict, List, Tuple, Set, Any, Optional
import os
from csp_solver import CSPSolver
from excel_handler import ExcelHandler


class FacultyScheduler:
    """
    Faculty Scheduler that uses constraint satisfaction to allocate subjects to faculty members.
    """
    
    def __init__(self, verbose: bool = False):
        """
        Initialize the Faculty Scheduler.
        
        Args:
            verbose: Whether to print detailed debug information
        """
        self.faculty = []   # List of faculty members
        self.subjects = []  # List of subjects
        self.classrooms = []  # List of classrooms
        self.timeslots = []  # List of timeslots
        self.solver = CSPSolver()  # The CSP solver
        self.schedule = []  # The final schedule
        self.verbose = verbose  # Whether to print debug info
        
    def load_data(self, input_file: str) -> bool:
        """
        Load data from Excel file.
        
        Args:
            input_file: Path to the input Excel file
            
        Returns:
            bool: True if data was loaded successfully, False otherwise
        """
        excel_handler = ExcelHandler()
        
        # Load faculty, subjects, classrooms, and timeslots
        self.faculty = excel_handler.read_faculty_data(input_file)
        self.subjects = excel_handler.read_subject_data(input_file)
        self.classrooms = excel_handler.read_classroom_data(input_file)
        self.timeslots = excel_handler.read_timeslot_data(input_file)
        
        return (len(self.faculty) > 0 and len(self.subjects) > 0 and 
                len(self.classrooms) > 0 and len(self.timeslots) > 0)
                
    def setup_csp(self) -> None:
        """
        Set up the constraint satisfaction problem for faculty scheduling.
        
        Using a more efficient approach:
        1. Create variables for subject-timeslot pairs (instead of subject-timeslot-classroom)
        2. Values are tuples of (faculty_id, classroom_id)
        3. This reduces the number of variables and makes the CSP more tractable
        """
        # Clear existing solver data
        self.solver = CSPSolver()
        
        # Print information about setup
        print("\nSetting up CSP problem:")
        print(f"Subjects: {len(self.subjects)}")
        print(f"Faculty: {len(self.faculty)}")
        print(f"Classrooms: {len(self.classrooms)}")
        print(f"Timeslots: {len(self.timeslots)}")
        
        # Get subjects with labs
        subjects_with_labs = {s['id']: s.get('lab_hours', 0) for s in self.subjects if s.get('lab_hours', 0) > 0}
        classrooms_with_labs = {c['id']: c['name'] for c in self.classrooms if c.get('has_lab', False)}
        
        print(f"Subjects with labs: {len(subjects_with_labs)}")
        print(f"Classrooms with labs: {len(classrooms_with_labs)}")
        
        # Create variables: tuples of (subject_id, timeslot_id)
        # We need to be smarter about the order of assignments - first handle subjects with labs
        variables_count = 0
        
        # First add the lab subjects (which are more constrained)
        lab_subjects = [s for s in self.subjects if s.get('lab_hours', 0) > 0]
        non_lab_subjects = [s for s in self.subjects if s.get('lab_hours', 0) == 0]
        
        if self.verbose:
            print(f"Lab subjects first: {[s['name'] for s in lab_subjects]}")
            print(f"Non-lab subjects second: {[s['name'] for s in non_lab_subjects]}")
        
        # Process subjects in order (lab subjects first)
        ordered_subjects = lab_subjects + non_lab_subjects
        
        for subject in ordered_subjects:
            subject_id = subject['id']
            requires_lab = subject.get('lab_hours', 0) > 0
            
            for timeslot in self.timeslots:
                timeslot_id = timeslot['id']
                
                variables_count += 1
                variable = (subject_id, timeslot_id)
                
                # Create domain values: (faculty_id, classroom_id) pairs
                domain_values = []
                
                # Find faculty members who can teach this subject
                suitable_faculty = []
                for f in self.faculty:
                    qualified_subjects = f.get('qualified_subjects', '')
                    # Handle both list and comma-separated string formats
                    if isinstance(qualified_subjects, str):
                        qualified_subjects = [int(s.strip()) for s in qualified_subjects.split(',') if s.strip()]
                    
                    if subject_id in qualified_subjects:
                        suitable_faculty.append(f['id'])
                
                # Determine eligible classrooms for this subject
                eligible_classrooms = []
                if requires_lab:
                    eligible_classrooms = [c for c in self.classrooms if c.get('has_lab', False)]
                else:
                    eligible_classrooms = self.classrooms
                
                # Create all possible faculty-classroom combinations
                for faculty_id in suitable_faculty:
                    faculty = next((f for f in self.faculty if f['id'] == faculty_id), None)
                    if faculty:
                        # Calculate preference score for this faculty-timeslot combination
                        preference_score = self.calculate_faculty_preference_score(
                            faculty, timeslot, {})
                        
                        for classroom in eligible_classrooms:
                            classroom_id = classroom['id']
                            # Add (faculty_id, classroom_id) tuple to domain with preference score
                            domain_values.append(((faculty_id, classroom_id), preference_score))
                
                # Sort domain values by preference score (higher scores first)
                sorted_values = [v[0] for v in sorted(domain_values, key=lambda x: x[1], reverse=True)]
                
                # Add variable only if it has a non-empty domain
                if sorted_values:
                    self.solver.add_variable(variable, sorted_values)
        
        print(f"Created {variables_count} variables for the CSP problem")
        print(f"Added {len(self.solver.variables)} variables with valid domains to the solver")
        
        # Add constraints
        self.add_constraints()
        
    def add_constraints(self) -> None:
        """
        Add constraints to the CSP solver.
        """
        # Hard constraints (must be satisfied)
        
        # 1. Faculty requirements should be consistently addressed
        self.solver.add_constraint(self.faculty_requirements_constraint)
        
        # 2. No conflicts in classes for faculty members
        self.solver.add_constraint(self.no_faculty_conflicts_constraint)
        
        # 3. No collision in classrooms
        self.solver.add_constraint(self.no_classroom_collision_constraint)
        
        # 4. Process subjects with labs appropriately
        self.solver.add_constraint(self.handle_labs_constraint)
        
        # Soft constraints (preferences - used in scoring)
        # These are implemented in the direct_solve method and faculty_preference_score function
        # They include day preferences, time period preferences, and consecutive classes preferences
        
    def faculty_requirements_constraint(self, assignment: Dict[Tuple, Any]) -> bool:
        """
        Constraint 1: The faculty's requirements should be consistently addressed.
        Each faculty member should not exceed their maximum teaching hours.
        
        Args:
            assignment: Current assignment of variables to values
            
        Returns:
            bool: True if the constraint is satisfied
        """
        # Only add debug output if we're at the "stuck" point
        debug_mode = (len(assignment) == 21)  # Just after our progress gets stuck
        
        if debug_mode:
            print("\nChecking faculty requirements constraint...")
            
        # Get all subjects that have been assigned to each faculty member
        faculty_subjects = {}
        
        for variable, value in assignment.items():
            subject_id, _ = variable
            faculty_id, _ = value  # Now value is (faculty_id, classroom_id)
            
            # Add unique subject assignments only (avoid counting the same subject multiple times)
            if faculty_id not in faculty_subjects:
                faculty_subjects[faculty_id] = set()
            
            faculty_subjects[faculty_id].add(subject_id)
        
        # Calculate faculty hours
        faculty_hours = {}
        for faculty_id, subject_ids in faculty_subjects.items():
            hours = 0
            for subject_id in subject_ids:
                subject = next((s for s in self.subjects if s['id'] == subject_id), None)
                if subject:
                    hours += subject.get('hours', 0)
            
            faculty_hours[faculty_id] = hours
            
            if debug_mode:
                faculty = next((f for f in self.faculty if f['id'] == faculty_id), None)
                if faculty:
                    faculty_name = faculty.get('name', '')
                    max_hours = faculty.get('max_hours', 0)
                    print(f"  Faculty {faculty_id} ({faculty_name}): {hours}/{max_hours} hours assigned")
        
        # Check if any faculty exceeds their maximum hours
        for faculty_id, hours in faculty_hours.items():
            faculty = next((f for f in self.faculty if f['id'] == faculty_id), None)
            if not faculty:
                continue
                
            max_hours = faculty.get('max_hours', 0)
            faculty_name = faculty.get('name', '')
            
            if hours > max_hours:
                if debug_mode or len(assignment) > 15:  # Show more info for later assignments
                    print(f"CONSTRAINT FAILED: Faculty requirements - Faculty {faculty_id} ({faculty_name}) " +
                          f"has {hours} hours assigned but max is {max_hours}")
                return False
                
        if debug_mode:
            print("  Faculty requirements constraint: PASSED")
            
        return True
    
    def no_faculty_conflicts_constraint(self, assignment: Dict[Tuple, Any]) -> bool:
        """
        Constraint 2: Different faculty members may teach the same subject, 
        but there should not be any conflicts with their classes.
        A faculty member can't teach two different subjects at the same time.
        
        Args:
            assignment: Current assignment of variables to values
            
        Returns:
            bool: True if the constraint is satisfied
        """
        # Only add debug output if we're at the "stuck" point
        debug_mode = (len(assignment) == 21)  # Just after our progress gets stuck
        
        if debug_mode:
            print("\nChecking faculty conflicts constraint...")
            
        faculty_schedule = {}
        
        for variable, value in assignment.items():
            subject_id, timeslot_id = variable
            faculty_id, _ = value  # Now value is (faculty_id, classroom_id)
            
            # Add to faculty schedule
            if faculty_id not in faculty_schedule:
                faculty_schedule[faculty_id] = set()
                
            # Check if faculty is already teaching at this timeslot
            if timeslot_id in faculty_schedule[faculty_id]:
                if debug_mode or len(assignment) > 15:
                    faculty = next((f for f in self.faculty if f['id'] == faculty_id), None)
                    faculty_name = faculty.get('name', '') if faculty else f"Faculty {faculty_id}"
                    timeslot = next((t for t in self.timeslots if t['id'] == timeslot_id), None)
                    timeslot_info = f"{timeslot.get('day', '')} {timeslot.get('time', '')}" if timeslot else f"Timeslot {timeslot_id}"
                    
                    subject = next((s for s in self.subjects if s['id'] == subject_id), None)
                    subject_name = subject.get('name', '') if subject else f"Subject {subject_id}"
                    
                    print(f"CONSTRAINT FAILED: Faculty conflict - {faculty_name} is already assigned to another subject at {timeslot_info}")
                    print(f"  Attempted to assign: {subject_name}")
                return False
                
            faculty_schedule[faculty_id].add(timeslot_id)
        
        if debug_mode:
            print("  Faculty conflicts constraint: PASSED")
                
        return True
    
    def no_classroom_collision_constraint(self, assignment: Dict[Tuple, Any]) -> bool:
        """
        Constraint 3: There must not be a collision in the classroom 
        designated for each teaching member.
        
        Args:
            assignment: Current assignment of variables to values
            
        Returns:
            bool: True if the constraint is satisfied
        """
        # Only add debug output if we're at the "stuck" point
        debug_mode = (len(assignment) == 21)  # Just after our progress gets stuck
        
        if debug_mode:
            print("\nChecking classroom collision constraint...")
            
        classroom_schedule = {}
        
        for variable, value in assignment.items():
            subject_id, timeslot_id = variable
            _, classroom_id = value  # Now value is (faculty_id, classroom_id)
            
            # Create key for classroom-timeslot combination
            key = (classroom_id, timeslot_id)
            
            # Check if classroom is already in use at this timeslot
            if key in classroom_schedule:
                if debug_mode or len(assignment) > 15:
                    classroom = next((c for c in self.classrooms if c['id'] == classroom_id), None)
                    classroom_name = classroom.get('name', '') if classroom else f"Classroom {classroom_id}"
                    
                    timeslot = next((t for t in self.timeslots if t['id'] == timeslot_id), None)
                    timeslot_info = f"{timeslot.get('day', '')} {timeslot.get('time', '')}" if timeslot else f"Timeslot {timeslot_id}"
                    
                    existing_subject_id = classroom_schedule[key]
                    existing_subject = next((s for s in self.subjects if s['id'] == existing_subject_id), None)
                    existing_subject_name = existing_subject.get('name', '') if existing_subject else f"Subject {existing_subject_id}"
                    
                    new_subject = next((s for s in self.subjects if s['id'] == subject_id), None)
                    new_subject_name = new_subject.get('name', '') if new_subject else f"Subject {subject_id}"
                    
                    print(f"CONSTRAINT FAILED: Classroom collision - {classroom_name} is already assigned at {timeslot_info}")
                    print(f"  Existing subject: {existing_subject_name}")
                    print(f"  Attempted to add: {new_subject_name}")
                return False
                
            classroom_schedule[key] = subject_id  # Store the subject ID for better debugging
        
        if debug_mode:
            print("  Classroom collision constraint: PASSED")
                
        return True
    
    def handle_labs_constraint(self, assignment: Dict[Tuple, Any]) -> bool:
        """
        Constraint 4: The algorithm must process a subject if it has any lab, 
        subject to constraint 3 (no classroom collision).
        
        Args:
            assignment: Current assignment of variables to values
            
        Returns:
            bool: True if the constraint is satisfied
        """
        # Only add debug output if we're at the "stuck" point
        debug_mode = (len(assignment) == 21)  # Just after our progress gets stuck
        
        if debug_mode:
            print("\nChecking labs constraint...")
            
        # Get subjects with labs - this should return subjects that have lab_hours > 0
        subjects_with_labs = {s['id']: s.get('lab_hours', 0) for s in self.subjects if s.get('lab_hours', 0) > 0}
        
        if debug_mode:
            print(f"  Found {len(subjects_with_labs)} subjects with lab requirements:")
            for subject_id, lab_hours in subjects_with_labs.items():
                subject = next((s for s in self.subjects if s['id'] == subject_id), None)
                subject_name = subject.get('name', '') if subject else f"Subject {subject_id}"
                print(f"    - {subject_name} (Lab hours: {lab_hours})")
        
        # Print debug info if verbose mode is enabled but limit output
        if self.verbose and len(assignment) % 20 == 0 and not debug_mode:
            # Use a counter instead of printing every check
            print(f"Checking constraints, progress: {len(assignment)} variables", end="\r")
        
        # Check that lab subjects are assigned to classrooms with lab facilities
        for variable, value in assignment.items():
            subject_id, timeslot_id = variable
            _, classroom_id = value  # Now value is (faculty_id, classroom_id)
            
            # If subject has a lab
            if subject_id in subjects_with_labs:
                # Check if classroom has lab facilities
                classroom = next((c for c in self.classrooms if c['id'] == classroom_id), None)
                if classroom is None:
                    print(f"Warning: Classroom with ID {classroom_id} not found")
                    continue
                    
                if not classroom.get('has_lab', False):
                    if debug_mode or len(assignment) > 15:
                        subject = next((s for s in self.subjects if s['id'] == subject_id), None)
                        subject_name = subject.get('name', '') if subject else f"Subject {subject_id}"
                        classroom_name = classroom.get('name', '') if classroom else f"Classroom {classroom_id}"
                        
                        print(f"CONSTRAINT FAILED: Lab requirement - {subject_name} requires lab facilities")
                        print(f"  Attempted to assign to: {classroom_name} which has no lab")
                    return False
        
        if debug_mode:
            print("  Labs constraint: PASSED")
                    
        return True
    
    def calculate_faculty_preference_score(self, faculty: Dict[str, Any], timeslot: Dict[str, Any], 
                                         current_schedule: Dict[int, Set[int]]) -> float:
        """
        Calculate a preference score for a faculty-timeslot combination.
        Higher scores indicate better matches with faculty preferences.

        Args:
            faculty: Faculty member data
            timeslot: Timeslot data
            current_schedule: Current faculty schedule (faculty_id -> set of timeslot_ids)

        Returns:
            float: Preference score (higher is better)
        """
        score = 0.0
        faculty_id = faculty['id']
        
        # Get faculty preferences
        day_preferences = faculty.get('day_preferences', [])
        time_preferences = faculty.get('time_preferences', [])
        consecutive_preference = faculty.get('consecutive_classes_preference', 0)
        
        # Handle string preferences (convert to list if needed)
        if isinstance(day_preferences, str):
            day_preferences = [d.strip() for d in day_preferences.split(',')]
        if isinstance(time_preferences, str):
            time_preferences = [t.strip() for t in time_preferences.split(',')]
        
        # 1. Day preference (e.g., preferred teaching days)
        timeslot_day = timeslot.get('day', '')
        if timeslot_day in day_preferences:
            score += 3.0  # Strong bonus for preferred day
        
        # 2. Time period preference (morning/afternoon/evening)
        timeslot_period = timeslot.get('period', '')
        if timeslot_period in time_preferences:
            score += 2.0  # Bonus for preferred time period
        
        # 3. Consecutive classes preference
        # Check if this timeslot would create consecutive classes
        timeslot_id = timeslot.get('id', 0)
        
        # Find adjacent timeslots in same day
        adjacent_timeslots = set()
        for ts in self.timeslots:
            if ts['day'] == timeslot_day and ts['id'] != timeslot_id:
                # Check if timeslot is adjacent by comparing IDs or times
                # This is a simple approach - in real systems, more complex logic would be needed
                if abs(ts['id'] - timeslot_id) == 1:
                    adjacent_timeslots.add(ts['id'])
        
        has_adjacent_class = False
        if faculty_id in current_schedule:
            # Check if any adjacent timeslots are already in faculty's schedule
            if any(ts_id in current_schedule[faculty_id] for ts_id in adjacent_timeslots):
                has_adjacent_class = True
        
        # Adjust score based on consecutive preference (-5 to +5 scale)
        if has_adjacent_class:
            # If faculty prefers consecutive classes (positive value), add bonus
            # If faculty avoids consecutive classes (negative value), add penalty
            normalized_preference = consecutive_preference / 5.0  # Convert to [-1, 1] range
            score += normalized_preference * 2.0  # Scale impact
        
        # Weight the score by faculty's preference_weight (importance factor)
        preference_weight = faculty.get('preference_weight', 1.0)
        normalized_weight = min(5.0, max(0.1, preference_weight)) / 5.0  # Normalize to [0.02, 1] range
        
        # Apply weight as multiplier to the final score
        weighted_score = score * normalized_weight
        
        return weighted_score
    
    def direct_solve(self) -> bool:
        """
        Directly construct a valid solution without using CSP.
        This is a greedy approach that may not find all valid solutions,
        but is faster and simpler for small to medium problems.
        
        The algorithm works by using a priority-based assignment:
        1. Lab subjects are prioritized as they have more constraints
        2. Faculty with specialized skills (fewer qualified subjects) are favored
        3. Even distribution of hours across faculty is attempted
        
        Returns:
            bool: True if a solution was found, False otherwise
        """
        print("Using direct construction method...")
        
        # Clear any existing schedule
        self.schedule = []
        
        # Sort subjects by constraint level (lab requirements and hours)
        # Lab subjects first, then by total hours (descending)
        lab_subjects = sorted([s for s in self.subjects if s.get('lab_hours', 0) > 0], 
                              key=lambda s: s.get('hours', 0) + s.get('lab_hours', 0), 
                              reverse=True)
        
        non_lab_subjects = sorted([s for s in self.subjects if s.get('lab_hours', 0) == 0], 
                                  key=lambda s: s.get('hours', 0), 
                                  reverse=True)
        
        ordered_subjects = lab_subjects + non_lab_subjects
        
        if self.verbose:
            print(f"Subject assignment order (most constrained first):")
            for idx, s in enumerate(ordered_subjects):
                lab_info = " (lab required)" if s.get('lab_hours', 0) > 0 else ""
                print(f"{idx+1}. {s['name']}: {s.get('hours', 0)} hours{lab_info}")
        
        # Get available lab classrooms
        lab_classrooms = [c for c in self.classrooms if c.get('has_lab', True)]
        regular_classrooms = [c for c in self.classrooms if not c.get('has_lab', False)]
        
        # Track faculty teaching hours
        faculty_hours = {f['id']: 0 for f in self.faculty}
        
        # Track used timeslots per faculty
        faculty_timeslots = {f['id']: set() for f in self.faculty}
        
        # Track used classroom-timeslot combinations
        used_classroom_timeslots = set()  # (classroom_id, timeslot_id) pairs
        
        # First, assign lab subjects (they are more constrained)
        for subject in ordered_subjects:
            subject_id = subject['id']
            hours = subject.get('hours', 0)
            needs_lab = subject.get('lab_hours', 0) > 0
            
            # Determine possible classrooms for this subject
            possible_classrooms = lab_classrooms if needs_lab else self.classrooms
            
            # Find faculty qualified to teach this subject with weighted selection
            qualified_faculty = []
            for f in self.faculty:
                qualified_subjects = f.get('qualified_subjects', '')
                if isinstance(qualified_subjects, str):
                    qualified_subjects = [int(s.strip()) for s in qualified_subjects.split(',') if s.strip()]
                
                if subject_id in qualified_subjects:
                    # Create a tuple with faculty and their specialization score
                    # Specialization score - faculty with fewer qualified subjects are more specialized
                    specialization_score = 1.0 / max(1, len(qualified_subjects))
                    qualified_faculty.append((f, specialization_score))
            
            # Sort faculty using a composite score that balances:
            # 1. Current teaching hours (less is better)
            # 2. Specialization (higher is better)
            # This way we prefer specialized faculty but still maintain workload balance
            def faculty_score(faculty_tuple):
                f, specialization = faculty_tuple
                faculty_id = f['id']
                
                # Normalize hours to [0,1] range using max hours (lower is better)
                max_possible_hours = f.get('max_hours', 20)
                hour_ratio = faculty_hours[faculty_id] / max_possible_hours if max_possible_hours > 0 else 0
                
                # We'll modify the score based on preferences in the timeslot selection stage
                # This base score prioritizes workload balance and specialization
                
                # Final score combines workload balance (hour_ratio) and specialization
                # - Lower hour_ratio is better (faculty with fewer hours)
                # - Higher specialization is better (faculty with specialized skills)
                # - Lower final score is better overall
                return hour_ratio - (0.5 * specialization)
                
            # Sort by composite score (lower is better)
            qualified_faculty.sort(key=faculty_score)
            
            # Extract just the faculty objects
            qualified_faculty = [f[0] for f in qualified_faculty]
            
            if self.verbose and qualified_faculty:
                print(f"\nFaculty candidates for {subject['name']}:")
                for idx, f in enumerate(qualified_faculty):
                    current_hours = faculty_hours[f['id']]
                    max_hours = f.get('max_hours', 0)
                    print(f"  {idx+1}. {f['name']}: currently {current_hours}/{max_hours} hours")
            
            # Try to assign subject to a faculty member
            assigned = False
            for faculty in qualified_faculty:
                faculty_id = faculty['id']
                max_hours = faculty.get('max_hours', 0)
                
                # Skip if faculty would exceed max hours
                if faculty_hours[faculty_id] + hours > max_hours:
                    continue
                
                # Find an available timeslot and classroom, using a smarter selection approach
                
                # Get available timeslots for this faculty
                available_timeslots = []
                for timeslot in self.timeslots:
                    timeslot_id = timeslot['id']
                    
                    # Skip if faculty already has a class in this timeslot
                    if timeslot_id in faculty_timeslots[faculty_id]:
                        continue
                    
                    # Count how many classrooms are still available in this timeslot
                    available_classrooms_count = sum(
                        1 for classroom in possible_classrooms 
                        if (classroom['id'], timeslot_id) not in used_classroom_timeslots
                    )
                    
                    # Add to available timeslots with preference score
                    # We prefer timeslots with more available classrooms to maximize future options
                    available_timeslots.append((timeslot, available_classrooms_count))
                
                # Calculate preference scores for each available timeslot
                scored_timeslots = []
                for timeslot_tuple in available_timeslots:
                    timeslot, classroom_count = timeslot_tuple
                    
                    # Calculate faculty's preference score for this timeslot
                    preference_score = self.calculate_faculty_preference_score(
                        faculty, timeslot, faculty_timeslots)
                    
                    # Use a composite score that balances:
                    # 1. Resource constraints (fewer classrooms = more constrained)
                    # 2. Faculty preferences (higher score = better match)
                    # Lower classroom count and higher preference score are better
                    composite_score = (-preference_score, classroom_count)
                    
                    scored_timeslots.append((timeslot, classroom_count, composite_score))
                
                # Sort by composite score (lower is better)
                # This prioritizes faculty preferences while still considering constraints
                scored_timeslots.sort(key=lambda t: t[2])
                
                # Extract just the timeslot and classroom count
                available_timeslots = [(t[0], t[1]) for t in scored_timeslots]
                
                # Try each timeslot in order
                for timeslot_tuple in available_timeslots:
                    timeslot, _ = timeslot_tuple
                    timeslot_id = timeslot['id']
                    
                    # Get sorted list of available classrooms (preferring smaller rooms for non-lab subjects
                    # or specialized labs for lab subjects to avoid wasting resources)
                    available_classrooms = [
                        classroom for classroom in possible_classrooms
                        if (classroom['id'], timeslot_id) not in used_classroom_timeslots
                    ]
                    
                    # Sort classrooms by suitability
                    if needs_lab:
                        # For lab subjects, prioritize dedicated lab rooms
                        available_classrooms.sort(key=lambda c: 0 if c.get('has_lab', False) else 1)
                    else:
                        # For regular subjects, prioritize non-lab rooms to save lab rooms for lab subjects
                        available_classrooms.sort(key=lambda c: 1 if c.get('has_lab', False) else 0)
                    
                    # Try each classroom in order
                    for classroom in available_classrooms:
                        classroom_id = classroom['id']
                        
                        # If we got here, we found a valid assignment!
                        self.schedule.append({
                            'faculty_id': faculty_id,
                            'faculty_name': faculty.get('name', ''),
                            'subject_id': subject_id,
                            'subject_name': subject.get('name', ''),
                            'has_lab': needs_lab,
                            'timeslot_id': timeslot_id,
                            'timeslot_day': timeslot.get('day', ''),
                            'timeslot_time': timeslot.get('time', ''),
                            'classroom_id': classroom_id,
                            'classroom_name': classroom.get('name', ''),
                        })
                        
                        # Update tracking variables
                        faculty_hours[faculty_id] += hours
                        faculty_timeslots[faculty_id].add(timeslot_id)
                        used_classroom_timeslots.add((classroom_id, timeslot_id))
                        
                        assigned = True
                        break
                    
                    if assigned:
                        break
                
                if assigned:
                    break
            
            # If we couldn't assign this subject, direct construction fails
            if not assigned:
                print(f"Could not assign subject {subject['name']} - no valid faculty, timeslot, or classroom available")
                return False
        
        # If we got here, we've successfully assigned all subjects
        print("Successfully constructed a valid schedule!")
        print(f"Total schedule entries: {len(self.schedule)}")
        
        # Debug: show faculty hours
        for faculty_id, hours in faculty_hours.items():
            faculty = next((f for f in self.faculty if f['id'] == faculty_id), None)
            if faculty:
                print(f"Faculty {faculty.get('name', '')}: {hours}/{faculty.get('max_hours', 0)} hours assigned")
        
        return True

    def solve(self, pre_assignments=None) -> bool:
        """
        Solve the faculty scheduling problem using CSP.
        
        This enhanced version uses a hybrid approach:
        1. First, pre-assign lab subjects (which have more constraints)
        2. Then let the CSP solver handle the remaining variables
        
        Args:
            pre_assignments: Optional dict of pre-assigned variables {(subject_id, timeslot_id): (faculty_id, classroom_id)}
            
        Returns:
            bool: True if a solution was found, False otherwise
        """
        # Set up the CSP
        self.setup_csp()
        
        # If pre-assignments were provided, use them directly
        if pre_assignments:
            if self.verbose:
                print(f"\nUsing {len(pre_assignments)} provided pre-assignments")
                for var, value in pre_assignments.items():
                    subject_id, timeslot_id = var
                    faculty_id, classroom_id = value
                    subject = next((s for s in self.subjects if s['id'] == subject_id), None)
                    faculty = next((f for f in self.faculty if f['id'] == faculty_id), None)
                    print(f"  Pre-assigned: {subject['name'] if subject else subject_id} to " +
                          f"{faculty['name'] if faculty else faculty_id}")
            
            # Set the partial assignment in the solver
            self.solver.assignment = pre_assignments.copy()
            
            # Solve with these pre-assignments
            solution = self.solver.solve()
            
            if solution is not None:
                self._convert_solution_to_schedule(solution)
                return True
            else:
                if self.verbose:
                    print("Provided pre-assignments did not lead to a solution. Trying the standard approach.")
                # Reset and try standard approach
                self.solver.assignment = {}
        
        # Apply a hybrid approach - pre-assign lab subjects (more constrained)
        # First, identify lab subjects
        lab_subjects = [s['id'] for s in self.subjects if s.get('lab_hours', 0) > 0]
        lab_classrooms = [c['id'] for c in self.classrooms if c.get('has_lab', False)]
        
        if self.verbose:
            print("\nHybrid approach: Pre-assigning lab subjects to reduce combinatorial explosion")
            print(f"Lab subjects: {lab_subjects}")
            print(f"Lab classrooms: {lab_classrooms}")
        
        # Find variables for lab subjects
        lab_variables = []
        for var in self.solver.variables:
            subject_id, _ = var
            if subject_id in lab_subjects:
                lab_variables.append(var)
        
        # Sort lab variables to prioritize high-constraint variables
        lab_variables.sort(key=lambda var: (var[0], var[1]))  # Sort by subject_id, then timeslot_id
        
        if self.verbose:
            print(f"Pre-assigning {len(lab_variables)} lab subject variables")
        
        # Create a partial assignment for lab subjects
        partial_assignment = {}
        
        # Try to pre-assign lab subjects
        for var in lab_variables:
            subject_id, timeslot_id = var
            found_valid_value = False
            
            # Try each valid combination of faculty and lab classroom
            for value in self.solver.domains[var]:
                faculty_id, classroom_id = value
                
                # Skip if classroom doesn't have lab facilities
                if classroom_id not in lab_classrooms:
                    continue
                
                # Check if assignment is consistent with existing partial assignment
                if self.solver.is_consistent(var, value, partial_assignment):
                    partial_assignment[var] = value
                    found_valid_value = True
                    
                    if self.verbose:
                        subject = next((s for s in self.subjects if s['id'] == subject_id), None)
                        faculty = next((f for f in self.faculty if f['id'] == faculty_id), None)
                        timeslot = next((t for t in self.timeslots if t['id'] == timeslot_id), None)
                        classroom = next((c for c in self.classrooms if c['id'] == classroom_id), None)
                        
                        subject_name = subject.get('name', '') if subject else f"Subject {subject_id}"
                        faculty_name = faculty.get('name', '') if faculty else f"Faculty {faculty_id}"
                        timeslot_info = f"{timeslot.get('day', '')} {timeslot.get('time', '')}" if timeslot else f"Timeslot {timeslot_id}"
                        classroom_name = classroom.get('name', '') if classroom else f"Classroom {classroom_id}"
                        
                        print(f"Pre-assigned: {subject_name} to {faculty_name} at {timeslot_info} in {classroom_name}")
                    
                    break
            
            if not found_valid_value and self.verbose:
                subject = next((s for s in self.subjects if s['id'] == subject_id), None)
                subject_name = subject.get('name', '') if subject else f"Subject {subject_id}"
                print(f"Warning: Could not pre-assign lab subject {subject_name} at timeslot {timeslot_id}")
        
        if self.verbose:
            print(f"Successfully pre-assigned {len(partial_assignment)} of {len(lab_variables)} lab variables")
        
        # Set the partial assignment in the solver
        self.solver.assignment = partial_assignment.copy()
        
        # Now solve the remaining variables
        solution = self.solver.solve()
        
        if solution is None:
            print("No solution found with hybrid approach!")
            return False
            
        # Convert solution to schedule format
        self._convert_solution_to_schedule(solution)
            
        if self.verbose:
            print(f"Found a solution with {len(self.schedule)} assignments!")
            
        return True
    
    def calculate_preference_satisfaction(self) -> Dict[int, Dict[str, float]]:
        """
        Calculate how well the current schedule satisfies faculty preferences.
        
        Returns:
            Dictionary mapping faculty IDs to their preference satisfaction metrics
        """
        # Initialize satisfaction stats for each faculty
        satisfaction = {}
        
        # Group schedule by faculty
        faculty_schedules = {}
        for entry in self.schedule:
            faculty_id = entry['faculty_id']
            if faculty_id not in faculty_schedules:
                faculty_schedules[faculty_id] = []
            faculty_schedules[faculty_id].append(entry)
        
        # Analyze each faculty's schedule
        for faculty_id, schedule_entries in faculty_schedules.items():
            faculty = next((f for f in self.faculty if f['id'] == faculty_id), None)
            if not faculty:
                continue
                
            # Initialize faculty satisfaction metrics
            satisfaction[faculty_id] = {
                'day_satisfaction': 0.0,
                'time_satisfaction': 0.0,
                'consecutive_satisfied': True,
                'overall_satisfaction': 0.0
            }
            
            # Get faculty preferences
            day_prefs = faculty.get('day_preferences', [])
            if isinstance(day_prefs, str):
                day_prefs = [d.strip() for d in day_prefs.split(',')]
                
            time_prefs = faculty.get('time_preferences', [])
            if isinstance(time_prefs, str):
                time_prefs = [t.strip() for t in time_prefs.split(',')]
                
            consecutive_pref = faculty.get('consecutive_classes_preference', 0)
            
            # Calculate day preference satisfaction
            days_in_schedule = set(entry['timeslot_day'] for entry in schedule_entries)
            preferred_days_met = sum(1 for day in days_in_schedule if day in day_prefs)
            if len(days_in_schedule) > 0:
                day_satisfaction = (preferred_days_met / len(days_in_schedule)) * 100
            else:
                day_satisfaction = 0.0
            
            # Calculate time preference satisfaction
            # Get time periods from timeslots
            periods_in_schedule = []
            for entry in schedule_entries:
                timeslot_id = entry['timeslot_id']
                timeslot = next((t for t in self.timeslots if t['id'] == timeslot_id), {})
                period = timeslot.get('period', '')
                if period:
                    periods_in_schedule.append(period)
            
            unique_periods = set(periods_in_schedule)
            preferred_periods_met = sum(1 for period in unique_periods if period in time_prefs)
            if len(unique_periods) > 0:
                time_satisfaction = (preferred_periods_met / len(unique_periods)) * 100
            else:
                time_satisfaction = 0.0
            
            # Check consecutive classes satisfaction
            # Group schedule entries by day
            classes_by_day = {}
            for entry in schedule_entries:
                day = entry['timeslot_day']
                if day not in classes_by_day:
                    classes_by_day[day] = []
                classes_by_day[day].append(entry['timeslot_id'])
            
            # Check if there are consecutive classes
            has_consecutive = False
            for day, timeslot_ids in classes_by_day.items():
                if len(timeslot_ids) < 2:
                    continue
                    
                # Sort timeslot IDs
                sorted_ids = sorted(timeslot_ids)
                
                # Check for adjacent IDs
                for i in range(len(sorted_ids) - 1):
                    if sorted_ids[i + 1] - sorted_ids[i] == 1:
                        has_consecutive = True
                        break
            
            # Determine if consecutive preference is satisfied
            consecutive_satisfied = True
            if consecutive_pref > 0 and not has_consecutive:
                # Faculty prefers consecutive classes but doesn't have any
                consecutive_satisfied = False
            elif consecutive_pref < 0 and has_consecutive:
                # Faculty avoids consecutive classes but has some
                consecutive_satisfied = False
            
            # Calculate overall satisfaction
            # Weight day and time satisfaction more heavily than consecutive
            overall_satisfaction = (day_satisfaction * 0.4) + (time_satisfaction * 0.4)
            if consecutive_satisfied:
                overall_satisfaction += 20  # Add 20% for satisfied consecutive preference
            
            # Store results
            satisfaction[faculty_id]['day_satisfaction'] = day_satisfaction
            satisfaction[faculty_id]['time_satisfaction'] = time_satisfaction
            satisfaction[faculty_id]['consecutive_satisfied'] = consecutive_satisfied
            satisfaction[faculty_id]['overall_satisfaction'] = overall_satisfaction
        
        return satisfaction
    
    def _convert_solution_to_schedule(self, solution):
        """
        Convert a CSP solution to a schedule format.
        
        Args:
            solution: Dictionary of variable-value assignments from the CSP solver
        """
        self.schedule = []
        for variable, value in solution.items():
            subject_id, timeslot_id = variable
            faculty_id, classroom_id = value  # Now value is (faculty_id, classroom_id)
            
            # Get subject, faculty, timeslot, and classroom details
            subject = next((s for s in self.subjects if s['id'] == subject_id), {})
            faculty = next((f for f in self.faculty if f['id'] == faculty_id), {})
            timeslot = next((t for t in self.timeslots if t['id'] == timeslot_id), {})
            classroom = next((c for c in self.classrooms if c['id'] == classroom_id), {})
            
            # Add to schedule
            self.schedule.append({
                'faculty_id': faculty_id,
                'faculty_name': faculty.get('name', ''),
                'subject_id': subject_id,
                'subject_name': subject.get('name', ''),
                'has_lab': subject.get('lab_hours', 0) > 0,
                'timeslot_id': timeslot_id,
                'day': timeslot.get('day', ''),
                'time': timeslot.get('time', ''),
                'classroom_id': classroom_id,
                'classroom_name': classroom.get('name', ''),
            })
    
    def save_schedule(self, output_file: str) -> bool:
        """
        Save the schedule to an Excel file.
        
        Args:
            output_file: Path to the output Excel file
            
        Returns:
            bool: True if the schedule was saved successfully, False otherwise
        """
        if not self.schedule:
            print("No schedule to save!")
            return False
            
        return ExcelHandler.write_schedule(output_file, self.schedule)
    
    def _create_formatted_table(self, headers, rows, title=None, hr_after_header=True):
        """
        Create a formatted table string with custom styling
        
        Args:
            headers: List of column headers
            rows: List of rows (each row is a list of values)
            title: Optional title to display above the table
            hr_after_header: Whether to add a horizontal rule after the header
            
        Returns:
            A string containing the formatted table
        """
        # Define minimum and maximum column widths
        min_width = 12
        max_width = 25
        
        # Calculate column widths based on content
        col_widths = [max(min_width, min(max_width, len(h))) for h in headers]
        
        # Adjust column widths based on content
        for row in rows:
            for i, cell in enumerate(row):
                if i < len(col_widths):  # Ensure we don't exceed list bounds
                    cell_width = len(str(cell))
                    col_widths[i] = max(col_widths[i], min(max_width, cell_width))
        
        # Give more space to certain columns based on content type
        col_types = {
            "Faculty": 0,
            "Subject": 1,
            "Day": 2,
            "Time": 3,
            "Classroom": 4,
            "Has Lab": 5,
            "Preferred Days": 1,
            "Preferred Times": 2,
            "Consecutive Classes": 3,
            "Preference Weight": 4,
            "Day Preferences Met": 1,
            "Time Preferences Met": 2,
            "Overall Satisfaction": 4
        }
        
        # Adjust widths for specific columns
        for i, header in enumerate(headers):
            if header in col_types:
                col_type = col_types[header]
                if header == "Subject" or header == "Preferred Days" or header == "Preferred Times":
                    col_widths[i] = max(col_widths[i], 20)  # Give more space to these columns
                elif header == "Faculty" or header == "Classroom":
                    col_widths[i] = max(col_widths[i], 15)
                elif header == "Has Lab":
                    col_widths[i] = 8  # Fixed width for Yes/No
        
        # Create horizontal line
        hline = "+" + "+".join("-" * (w + 2) for w in col_widths) + "+"
        
        # Format a row with word wrapping
        def format_row(row):
            # First, determine if any cell needs wrapping
            needs_wrapping = False
            wrapped_cells = []
            
            for i, cell in enumerate(row):
                cell_text = str(cell)
                wrapped_text = []
                
                # If cell is longer than column width, wrap it
                if len(cell_text) > col_widths[i]:
                    # Simple word wrapping
                    words = cell_text.split()
                    current_line = ""
                    
                    for word in words:
                        if len(current_line) + len(word) + 1 <= col_widths[i]:
                            if current_line:
                                current_line += " " + word
                            else:
                                current_line = word
                        else:
                            wrapped_text.append(current_line)
                            current_line = word
                    
                    if current_line:
                        wrapped_text.append(current_line)
                    
                    needs_wrapping = True
                else:
                    wrapped_text = [cell_text]
                
                wrapped_cells.append(wrapped_text)
            
            # If no cell needs wrapping, return a simple formatted row
            if not needs_wrapping:
                cells = []
                for i, cell in enumerate(row):
                    cell_text = str(cell)
                    # Center or left align based on column type
                    align_center = i in [1, 2, 3, 4] and "Preferences Met" in headers[i]
                    if align_center or headers[i] == "Has Lab" or headers[i] == "Overall Satisfaction":
                        cells.append(f" {cell_text.center(col_widths[i])} ")
                    else:
                        cells.append(f" {cell_text.ljust(col_widths[i])} ")
                return "|" + "|".join(cells) + "|"
            
            # If wrapping is needed, handle multi-line output
            max_lines = max(len(cell) for cell in wrapped_cells)
            result_lines = []
            
            for line_idx in range(max_lines):
                cells = []
                for i, wrapped_text in enumerate(wrapped_cells):
                    if line_idx < len(wrapped_text):
                        cell_text = wrapped_text[line_idx]
                    else:
                        cell_text = ""
                    
                    # Center or left align based on column type
                    align_center = i in [1, 2, 3, 4] and "Preferences Met" in headers[i]
                    if align_center or headers[i] == "Has Lab" or headers[i] == "Overall Satisfaction":
                        cells.append(f" {cell_text.center(col_widths[i])} ")
                    else:
                        cells.append(f" {cell_text.ljust(col_widths[i])} ")
                
                result_lines.append("|" + "|".join(cells) + "|")
            
            return "\n".join(result_lines)
        
        # Build table
        table = []
        if title:
            table.append(f"\n{title}:")
        table.append(hline)
        table.append(format_row(headers))
        table.append(hline)
        
        for i, row in enumerate(rows):
            table.append(format_row(row))
            table.append(hline)
        
        return "\n".join(table)

    def display_schedule(self) -> None:
        """
        Display the schedule in a formatted table.
        """
        if not self.schedule:
            print("No schedule to display!")
            return
        
        # First show faculty preferences for context
        if self.verbose:
            # Prepare data for faculty preferences table
            headers = ["Faculty", "Preferred Days", "Preferred Times", "Consecutive Classes", "Preference Weight"]
            rows = []
            
            for faculty in self.faculty:
                day_prefs = faculty.get('day_preferences', [])
                if isinstance(day_prefs, str):
                    day_prefs = day_prefs.split(',')
                
                time_prefs = faculty.get('time_preferences', [])
                if isinstance(time_prefs, str):
                    time_prefs = time_prefs.split(',')
                
                consecutive_pref = faculty.get('consecutive_classes_preference', 0)
                consecutive_str = "Neutral"
                if consecutive_pref > 0:
                    consecutive_str = f"Prefers (+{consecutive_pref})" 
                elif consecutive_pref < 0:
                    consecutive_str = f"Avoids ({consecutive_pref})"
                
                rows.append([
                    faculty.get('name', ''),
                    ", ".join(day_prefs),
                    ", ".join(time_prefs),
                    consecutive_str,
                    faculty.get('preference_weight', 1.0)
                ])
            
            print(self._create_formatted_table(headers, rows, "Faculty Preferences"))
        
        # Display the main schedule
        headers = ["Faculty", "Subject", "Day", "Time", "Classroom", "Has Lab"]
        rows = []
        
        for item in self.schedule:
            rows.append([
                item['faculty_name'], 
                item['subject_name'], 
                item['timeslot_day'], 
                item['timeslot_time'], 
                item['classroom_name'],
                "Yes" if item['has_lab'] else "No"
            ])
        
        print(self._create_formatted_table(headers, rows, "Faculty Schedule"))
        
        # Analyze how well preferences were satisfied
        if self.verbose:
            preference_satisfaction = self.calculate_preference_satisfaction()
            
            headers = [
                "Faculty", "Day Preferences Met", "Time Preferences Met", 
                "Consecutive Classes", "Overall Satisfaction"
            ]
            rows = []
            
            for faculty_id, stats in preference_satisfaction.items():
                faculty = next((f for f in self.faculty if f['id'] == faculty_id), None)
                if faculty:
                    rows.append([
                        faculty.get('name', ''),
                        f"{stats['day_satisfaction']:.1f}%",
                        f"{stats['time_satisfaction']:.1f}%",
                        "Satisfied" if stats['consecutive_satisfied'] else "Not Satisfied",
                        f"{stats['overall_satisfaction']:.1f}%"
                    ])
            
            print(self._create_formatted_table(headers, rows, "Faculty Preference Satisfaction Analysis"))
