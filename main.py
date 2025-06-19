import pandas as pd

# --- Load data ---
students_df = pd.read_csv("students.csv")
projects_df = pd.read_csv("projects.csv")
preallocated_df = pd.read_csv("preallocated.csv")

# --- Define supervisor capacity (assume each can supervise up to 5 students) ---
supervisor_capacity = {sup_id: 5 for sup_id in projects_df['supervisor_id'].unique()}

# --- Validate student preferences ---
def validate_students(df, projects_df):
    valid_project_ids = set(projects_df['project_id'])
    errors = []

    for _, row in df.iterrows():
        choices = [row['choice_1'], row['choice_2'], row['choice_3']]

        # 1. Check for duplicate choices
        if len(set(choices)) < 3:
            errors.append((row['student_id'], "Duplicate choices"))

        # 2. Check for invalid project IDs
        for choice in choices:
            if choice not in valid_project_ids:
                errors.append((row['student_id'], f"Invalid project ID: {choice}"))

        # 3. Check if all choices are with different supervisors
        sup_ids = []
        for c in choices:
            project_row = projects_df[projects_df['project_id'] == c]
            if not project_row.empty:
                sup_id = project_row['supervisor_id'].values[0]
                sup_ids.append(sup_id)
        if len(set(sup_ids)) < 3:
            errors.append((row['student_id'], "Same supervisor in multiple choices"))

    return errors

validation_errors = validate_students(students_df, projects_df)
print("Validation errors:", validation_errors)

# --- Initialize allocation and supervisor load trackers ---
allocation = {}  # student_id -> project_id
supervisor_load = {sup_id: 0 for sup_id in supervisor_capacity}

# --- Step 1: Preallocate fixed students ---
for _, row in preallocated_df.iterrows():
    student_id = row['student_id']
    project_id = row['project_id']
    sup_id = row['supervisor_id']
    allocation[student_id] = project_id
    supervisor_load[sup_id] += 1

# --- Step 2: Greedy allocation based on student choices ---
for _, row in students_df.iterrows():
    student_id = row['student_id']
    if student_id in allocation:
        continue  # Skip already preallocated students

    for choice in ['choice_1', 'choice_2', 'choice_3']:
        project_id = row[choice]

        # Skip if project does not exist (defensive check)
        project_row = projects_df[projects_df['project_id'] == project_id]
        if project_row.empty:
            continue

        sup_id = project_row['supervisor_id'].values[0]

        # Check supervisor capacity
        if supervisor_load[sup_id] < supervisor_capacity[sup_id]:
            allocation[student_id] = project_id
            supervisor_load[sup_id] += 1
            break  # Stop at the first successful assignment

# --- Identify unassigned students ---
unmatched = [sid for sid in students_df['student_id'] if sid not in allocation]
print("Unmatched students:", unmatched)

# --- Output results to CSV ---
output = []
for _, row in students_df.iterrows():
    sid = row['student_id']
    assigned_project = allocation.get(sid, "UNASSIGNED")
    output.append({
        "student_id": sid,
        "student_name": row['student_name'],
        "assigned_project": assigned_project
    })

output_df = pd.DataFrame(output)
output_df.to_csv("allocation_result.csv", index=False)
print("Allocation results saved to allocation_result.csv")