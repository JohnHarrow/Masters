import pandas as pd
from collections import deque
from pulp import LpProblem, LpVariable, LpMaximize, lpSum, LpBinary, PULP_CBC_CMD

# --- Load data ---
students_df = pd.read_csv("students.csv")
projects_df = pd.read_csv("projects.csv")
preallocated_df = pd.read_csv("preallocated.csv")
supervisors_df = pd.read_csv("supervisors.csv")  # supervisor_id, supervisor_name, capacity

# --- Validate student data structure and content ---
def validate_student_data(students_df, projects_df):
    errors = []
    required_fields = ['student_id', 'student_name', 'choice_1', 'choice_2', 'choice_3']
    missing_fields = [field for field in required_fields if field not in students_df.columns]
    if missing_fields:
        errors.append(f"Missing required columns: {missing_fields}")

    if students_df[required_fields].isnull().any().any():
        null_rows = students_df[required_fields].isnull().any(axis=1)
        null_ids = students_df[null_rows]['student_id'].tolist()
        errors.append(f"Missing data in required fields for student_ids: {null_ids}")

    if students_df['student_id'].duplicated().any():
        dup_ids = students_df[students_df['student_id'].duplicated()]['student_id'].tolist()
        errors.append(f"Duplicate student IDs found: {dup_ids}")

    valid_project_ids = set(projects_df['project_id'])
    for _, row in students_df.iterrows():
        sid = row['student_id']
        for col in ['choice_1', 'choice_2', 'choice_3']:
            if row[col] not in valid_project_ids:
                errors.append(f"Student {sid}: Invalid project ID '{row[col]}' in {col}")

    for _, row in students_df.iterrows():
        sid = row['student_id']
        choices = [row['choice_1'], row['choice_2'], row['choice_3']]
        if len(set(choices)) < 3:
            errors.append(f"Student {sid}: Duplicate project choices detected.")

    if errors:
        print("❌ STUDENT DATA VALIDATION ERRORS:")
        for err in errors:
            print(" -", err)
        raise ValueError("Student data validation failed.")
    else:
        print("✅ Student data validation passed.")

# --- Optional supervisor diversity check ---
def validate_supervisor_diversity(students_df, projects_df):
    warnings = []
    for _, row in students_df.iterrows():
        choices = [row['choice_1'], row['choice_2'], row['choice_3']]
        sup_ids = []
        for c in choices:
            pr = projects_df[projects_df['project_id'] == c]
            if not pr.empty:
                sup_ids.append(pr['supervisor_id'].values[0])
        if len(set(sup_ids)) < 3:
            warnings.append(row['student_id'])

    if warnings:
        print("⚠️ SUPERVISOR DIVERSITY WARNINGS (not errors):")
        print(" Students with multiple choices from same supervisor:", warnings)
    else:
        print("✅ Supervisor diversity check passed.")

# --- Capacities ---
supervisor_capacity = {
    row['supervisor_id']: int(row['capacity']) if pd.notna(row['capacity']) else 3
    for _, row in supervisors_df.iterrows()
}

project_capacity = {
    row['project_id']: int(row['max_students']) if pd.notna(row['max_students']) else None
    for _, row in projects_df.iterrows()
}

# --- Validations ---
validate_student_data(students_df, projects_df)
validate_supervisor_diversity(students_df, projects_df)

# --- Greedy matching function ---
def greedy_matching(students_df, projects_df, supervisor_capacity, project_capacity, preallocated_df):
    allocation = {}
    supervisor_load = {sup_id: 0 for sup_id in supervisor_capacity}
    project_load = {pid: 0 for pid in projects_df['project_id']}

    for _, row in preallocated_df.iterrows():
        sid = row['student_id']
        pid = row['project_id']
        sup = row['supervisor_id']
        allocation[sid] = pid
        supervisor_load[sup] += 1
        project_load[pid] += 1

    for _, row in students_df.iterrows():
        sid = row['student_id']
        if sid in allocation:
            continue

        for choice in ['choice_1', 'choice_2', 'choice_3']:
            pid = row[choice]
            project_row = projects_df[projects_df['project_id'] == pid]
            if project_row.empty:
                continue

            sup = project_row['supervisor_id'].values[0]
            max_proj_cap = project_capacity[pid]

            if supervisor_load[sup] < supervisor_capacity[sup]:
                if max_proj_cap is None or project_load[pid] < max_proj_cap:
                    allocation[sid] = pid
                    supervisor_load[sup] += 1
                    project_load[pid] += 1
                    break

    return allocation

# --- Stable marriage matching function ---
def stable_marriage_matching(students_df, projects_df, supervisor_capacity, project_capacity, preallocated_df):
    allocation = {}
    supervisor_load = {sup_id: 0 for sup_id in supervisor_capacity}
    project_load = {pid: 0 for pid in projects_df['project_id']}

    for _, row in preallocated_df.iterrows():
        sid = row['student_id']
        pid = row['project_id']
        sup = row['supervisor_id']
        allocation[sid] = pid
        supervisor_load[sup] += 1
        project_load[pid] += 1

    student_prefs = {
        row['student_id']: deque([row['choice_1'], row['choice_2'], row['choice_3']])
        for _, row in students_df.iterrows() if row['student_id'] not in allocation
    }

    free_students = deque(student_prefs.keys())

    while free_students:
        sid = free_students.popleft()
        if not student_prefs[sid]:
            continue

        pid = student_prefs[sid].popleft()
        project_row = projects_df[projects_df['project_id'] == pid]
        if project_row.empty:
            continue

        sup = project_row['supervisor_id'].values[0]
        max_proj_cap = project_capacity.get(pid, None)

        if supervisor_load[sup] < supervisor_capacity[sup] and \
           (max_proj_cap is None or project_load[pid] < max_proj_cap):
            allocation[sid] = pid
            supervisor_load[sup] += 1
            project_load[pid] += 1
        else:
            free_students.append(sid)

    return allocation

# --- Linear programming matching function ---
def linear_programming_matching(students_df, projects_df, supervisor_capacity, project_capacity, preallocated_df):
    prob = LpProblem("Student_Project_Matching", LpMaximize)

    students = list(students_df['student_id'])
    projects = list(projects_df['project_id'])

    # Build dictionaries for quick lookup
    project_supervisors = projects_df.set_index('project_id')['supervisor_id'].to_dict()
    student_choices = {}
    for _, row in students_df.iterrows():
        sid = row['student_id']
        # Map project to weight based on preference order (choice_1=3, choice_2=2, choice_3=1)
        student_choices[sid] = {
            row['choice_1']: 3,
            row['choice_2']: 2,
            row['choice_3']: 1
        }

    # Decision variables: x[(student, project)] = 1 if assigned, else 0
    x = LpVariable.dicts("assign",
                         [(s, p) for s in students for p in projects],
                         cat=LpBinary)

    # Objective: maximize total satisfaction (sum of assigned weights)
    prob += lpSum(x[(s, p)] * student_choices.get(s, {}).get(p, 0) for s in students for p in projects)

    # Constraint 1: Each student assigned to at most one project
    for s in students:
        prob += lpSum(x[(s, p)] for p in projects) <= 1

    # Constraint 2: Respect preallocation (force assigned)
    for _, row in preallocated_df.iterrows():
        sid = row['student_id']
        pid = row['project_id']
        for p in projects:
            if p == pid:
                prob += x[(sid, p)] == 1
            else:
                prob += x[(sid, p)] == 0

    # Constraint 3: Project capacity
    for p in projects:
        max_cap = project_capacity.get(p)
        if max_cap is not None:
            prob += lpSum(x[(s, p)] for s in students) <= max_cap

    # Constraint 4: Supervisor capacity
    for sup, sup_cap in supervisor_capacity.items():
        sup_projects = [p for p, sup_id in project_supervisors.items() if sup_id == sup]
        prob += lpSum(x[(s, p)] for s in students for p in sup_projects) <= sup_cap

    # Solve
    solver = PULP_CBC_CMD(msg=False)
    prob.solve(solver)

    # Extract allocation
    allocation = {}
    for s in students:
        for p in projects:
            var = x[(s, p)]
            if var.varValue is not None and var.varValue > 0.5:
                allocation[s] = p
                break

    return allocation

# --- Save results ---
def save_allocation_result(allocation, students_df, filename):
    output = []
    for _, row in students_df.iterrows():
        sid = row['student_id']
        assigned = allocation.get(sid, "UNASSIGNED")
        output.append({
            "student_id": sid,
            "student_name": row['student_name'],
            "assigned_project": assigned
        })
    pd.DataFrame(output).to_csv(filename, index=False)
    print(f"📄 Results saved to {filename}")

# --- Run all matching algorithms ---
allocation_greedy = greedy_matching(
    students_df, projects_df, supervisor_capacity, project_capacity, preallocated_df
)
allocation_stable = stable_marriage_matching(
    students_df, projects_df, supervisor_capacity, project_capacity, preallocated_df
)
allocation_lp = linear_programming_matching(
    students_df, projects_df, supervisor_capacity, project_capacity, preallocated_df
)

# --- Save output ---
save_allocation_result(allocation_greedy, students_df, "allocation_result_greedy.csv")
save_allocation_result(allocation_stable, students_df, "allocation_result_stable.csv")
save_allocation_result(allocation_lp, students_df, "allocation_result_linear_programming.csv")

# --- Summary ---
unmatched_greedy = [sid for sid in students_df['student_id'] if sid not in allocation_greedy]
unmatched_stable = [sid for sid in students_df['student_id'] if sid not in allocation_stable]
unmatched_lp = [sid for sid in students_df['student_id'] if sid not in allocation_lp]

print(f"\n🔁 Greedy unmatched students: {len(unmatched_greedy)}")
print(f"💍 Stable unmatched students: {len(unmatched_stable)}")
print(f"🧮 LP unmatched students: {len(unmatched_lp)}")