import pandas as pd

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

# --- Optional diversity check (warning only) ---
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

# --- Prepare supervisor capacities ---
supervisor_capacity = {
    row['supervisor_id']: int(row['capacity']) if pd.notna(row['capacity']) else 3
    for _, row in supervisors_df.iterrows()
}

# --- Prepare project capacities (NaN = unlimited within supervisor capacity) ---
project_capacity = {}
for _, row in projects_df.iterrows():
    pid = row['project_id']
    cap = row.get('max_students', None)
    project_capacity[pid] = int(cap) if pd.notna(cap) else None  # None means no limit

# --- Run validations ---
validate_student_data(students_df, projects_df)
validate_supervisor_diversity(students_df, projects_df)

# --- Initialize tracking ---
allocation = {}  # student_id -> project_id
supervisor_load = {sup_id: 0 for sup_id in supervisor_capacity}
project_load = {pid: 0 for pid in projects_df['project_id']}

# --- Preallocate ---
for _, row in preallocated_df.iterrows():
    sid = row['student_id']
    pid = row['project_id']
    sup = row['supervisor_id']
    allocation[sid] = pid
    supervisor_load[sup] += 1
    project_load[pid] += 1

# --- Greedy allocation ---
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
                break  # Assigned

# --- Report unassigned students ---
unmatched = [sid for sid in students_df['student_id'] if sid not in allocation]
print(f"\n🔍 Unmatched students ({len(unmatched)}): {unmatched}")

# --- Save results ---
output = []
for _, row in students_df.iterrows():
    sid = row['student_id']
    assigned = allocation.get(sid, "UNASSIGNED")
    output.append({
        "student_id": sid,
        "student_name": row['student_name'],
        "assigned_project": assigned
    })

pd.DataFrame(output).to_csv("allocation_result.csv", index=False)
print("📄 Allocation results saved to allocation_result.csv")