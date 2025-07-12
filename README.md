# Faculty Subject Allocation System (CSP-based)

A Constraint Satisfaction Problem (CSP) solution in Python for assigning subjects to faculty members based on availability, workload, and timing constraints.

---

## 📂 Repository Structure

├── FacultyAllocation.py # Main CSP solver (hill‑climbing local search)
├── subjects.py # Defines Subject class and instances
├── faculty.py # Defines Faculty class and instances
├── constraints.py # Availability, workload, and conflict constraints
├── run_allocation.py # Script to run full allocation end-to-end
└── data/ # (Optional) Place to store CSV/JSON input files


---

## 🛠️ Requirements

- Python 3.7+
- No external packages needed (standard library only)

---

## 🚀 How to Run

### 1. Clone this repository

```bash
git clone https://github.com/RohithSai2108/Faculty-Subject-Allocation-System-using-Constraint-Satisfaction-Problem-CSP-Techniques.git
cd Faculty-Subject-Allocation-System-using-Constraint-Satisfaction-Problem-CSP-Techniques 
```

### 2. Define your data
subjects.py: Edit the subject list and assign available faculty names.

faculty.py: Edit the faculty list with names and max_classes.

constraints.py: Tweak constraint logic if needed (e.g., time-slot overlap).

### 3. Test allocation
Run the CSP solver directly:
```
python FacultyAllocation.py
```
You’ll see the resulting subject-to-faculty mapping printed in your console.

### 4. Full workflow script
Alternatively, use:
```
python run_allocation.py
```
### This script:

Loads data from subjects.py, faculty.py, and constraints.py

Executes the CSP solver

Prints summary output:

Assigned faculty per subject

Number of constraint violations (should ideally be zero)

Total assignments

### ⚙️ Understanding Key Files
FacultyAllocation.py
- Implements hill-climbing with probabilistic acceptance to optimize assignments based on cumulative constraint satisfaction.

- subjects.py & faculty.py
- Define simple classes:

@dataclass
class Subject:
    name: str
    available_faculty: List[str]
    timeslot: Tuple[int, int]  # start-end slots

@dataclass
class Faculty:
    name: str
    max_classes: int

Add or update data directly here.

constraints.py
Each function checks:

Availability: faculty can teach the subject

Workload: within max_classes

Conflict: no timeslot overlap

Modify or extend as needed for custom rules.

### 🧪 Sample Run
```
python run_allocation.py
```
Example output:

Assignment:
  Math101 ➜ Dr. Smith
  Physics102 ➜ Dr. Patel
...

Total constraints satisfied: 30 / 30

### 📈 Tips & Customization
New constraints: write a function new_constraint(subj, fac, solution) and add it in constraints_list.

Data input: swap Python data with CSV/JSON by reading from data/ in run_allocation.py.

Improve search: experiment with simulated annealing or genetic algorithms.

### 📝 License
MIT License — feel free to reuse and adapt!

🤝 Feedback & Collaboration
Contributions are welcome! Please raise issues or pull requests if you'd like to add features — e.g., GUI, database support, scheduling-based constraints, etc.
