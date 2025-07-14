# python -m streamlit run app.py

import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO
from collections import deque
from pulp import LpProblem, LpVariable, LpMaximize, lpSum, LpBinary, PULP_CBC_CMD
from openpyxl.styles import PatternFill
from openpyxl.formatting.rule import CellIsRule
from openpyxl import load_workbook
import matplotlib.pyplot as plt

st.set_page_config(page_title="Project Matching App", layout="wide")

st.title("Student-Project Allocation Tool")

# --- Generate Intitial Files ---
def generate_template_excel():
    from openpyxl import Workbook
    from io import BytesIO

    wb = Workbook()

    # --- Students Sheet ---
    ws_students = wb.active
    ws_students.title = "students"
    ws_students.append(["student_id", "student_name", "choice_1", "choice_2", "choice_3"])

    # --- Projects Sheet ---
    ws_projects = wb.create_sheet("projects")
    ws_projects.append(["project_id", "project_title", "supervisor_id", "supervisor_name", "max_students"])

    # --- Supervisors Sheet ---
    ws_supervisors = wb.create_sheet("supervisors")
    ws_supervisors.append(["supervisor_id", "supervisor_name", "capacity"])

    # --- Preallocated Sheet ---
    ws_preallocated = wb.create_sheet("preallocated")
    ws_preallocated.append(["student_id", "project_id", "supervisor_id"])

    # Save to BytesIO
    output = BytesIO()
    wb.save(output)
    output.seek(0)
    return output

# --- File Upload ---
st.sidebar.header("Upload Input Data")

# --- Template Download Button ---
st.sidebar.markdown("---")
st.sidebar.markdown("Need sample input files?")
if st.sidebar.button("Generate Input Template"):
    template = generate_template_excel()
    st.sidebar.download_button(
        label="Download Excel Template",
        data=template,
        file_name="input_template.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

# Option 1: Single Excel file
excel_file = st.sidebar.file_uploader("Upload Excel File with All Sheets", type="xlsx")

# Option 2: Separate CSVs
st.sidebar.markdown("Or upload individual CSV files:")
students_file = st.sidebar.file_uploader("students.csv", type="csv")
projects_file = st.sidebar.file_uploader("projects.csv", type="csv")
supervisors_file = st.sidebar.file_uploader("supervisors.csv", type="csv")
preallocated_file = st.sidebar.file_uploader("preallocated.csv", type="csv")


data_loaded = False

if excel_file:
    try:
        excel_data = pd.ExcelFile(excel_file)
        students_df = pd.read_excel(excel_data, sheet_name="students")
        projects_df = pd.read_excel(excel_data, sheet_name="projects")
        supervisors_df = pd.read_excel(excel_data, sheet_name="supervisors")
        preallocated_df = pd.read_excel(excel_data, sheet_name="preallocated")
        st.success("Excel file loaded successfully.")
        data_loaded = True
    except Exception as e:
        st.error(f"Failed to read Excel file: {e}")

elif all([students_file, projects_file, supervisors_file, preallocated_file]):
    try:
        students_df = pd.read_csv(students_file)
        projects_df = pd.read_csv(projects_file)
        supervisors_df = pd.read_csv(supervisors_file)
        preallocated_df = pd.read_csv(preallocated_file)
        st.success("CSV files loaded successfully.")
        data_loaded = True
    except Exception as e:
        st.error(f"Failed to read one or more CSV files: {e}")

else:
    st.info("⬅️ Please upload either a single Excel file or all four required CSV files to begin.")

# --- Only run logic if data loaded ---
if data_loaded:

    # --- Coerce numeric columns before validation ---
    projects_df['max_students'] = pd.to_numeric(projects_df['max_students'], errors='coerce')
    supervisors_df['capacity'] = pd.to_numeric(supervisors_df['capacity'], errors='coerce')

    # --- Data Validation ---
    def validate_student_data(students_df, projects_df):
        errors = []
        required_fields = ['student_id', 'student_name', 'choice_1', 'choice_2', 'choice_3']
        missing_fields = [field for field in required_fields if field not in students_df.columns]
        if missing_fields:
            errors.append(f"Missing required columns: {missing_fields}")

        if students_df[required_fields].isnull().any().any():
            null_ids = students_df[students_df[required_fields].isnull().any(axis=1)]['student_id'].tolist()
            errors.append(f"Missing required fields for student_ids: {null_ids}")

        if students_df['student_id'].duplicated().any():
            dup_ids = students_df[students_df['student_id'].duplicated()]['student_id'].tolist()
            errors.append(f"Duplicate student IDs: {dup_ids}")

        valid_project_ids = set(projects_df['project_id'])
        for _, row in students_df.iterrows():
            sid = row['student_id']
            for col in ['choice_1', 'choice_2', 'choice_3']:
                if row[col] not in valid_project_ids:
                    errors.append(f"Student {sid}: Invalid project ID '{row[col]}' in {col}")

            if len(set([row['choice_1'], row['choice_2'], row['choice_3']])) < 3:
                errors.append(f"Student {sid}: Duplicate project choices.")

        return errors

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
        return warnings
    
    def validate_project_data(projects_df):
        errors = []
        required_columns = {'project_id', 'project_title', 'supervisor_id', 'supervisor_name', 'max_students'}
        missing_columns = required_columns - set(projects_df.columns)
        if missing_columns:
            errors.append(f"Missing required columns in projects sheet: {list(missing_columns)}")
            return errors

        if projects_df['project_id'].duplicated().any():
            dup_ids = projects_df[projects_df['project_id'].duplicated()]['project_id'].tolist()
            errors.append(f"Duplicate project IDs found: {dup_ids}")

        # Allow blanks but no negative values
        invalid_rows = projects_df[
            projects_df['max_students'].notnull() & (projects_df['max_students'] < 0)
        ]
        if not invalid_rows.empty:
            invalid_ids = invalid_rows['project_id'].tolist()
            errors.append(f"Negative max_students values for project IDs: {invalid_ids}")

        return errors
    
    def validate_supervisor_data(supervisors_df):
        errors = []
        required_columns = {'supervisor_id', 'supervisor_name', 'capacity'}
        missing_columns = required_columns - set(supervisors_df.columns)
        if missing_columns:
            errors.append(f"Missing required columns in supervisors sheet: {list(missing_columns)}")
            return errors

        if supervisors_df['supervisor_id'].duplicated().any():
            dup_ids = supervisors_df[supervisors_df['supervisor_id'].duplicated()]['supervisor_id'].tolist()
            errors.append(f"Duplicate supervisor IDs found: {dup_ids}")

        # Allow blanks but no negative values
        invalid_rows = supervisors_df[
            supervisors_df['capacity'].notnull() & (supervisors_df['capacity'] < 0)
        ]
        if not invalid_rows.empty:
            invalid_ids = invalid_rows['supervisor_id'].tolist()
            errors.append(f"Negative capacity values for supervisor IDs: {invalid_ids}")

        return errors



    # --- Capacities ---
    supervisor_capacity = {
        row['supervisor_id']: int(row['capacity']) if pd.notna(row['capacity']) else 3
        for _, row in supervisors_df.iterrows()
    }

    project_capacity = {
        row['project_id']: int(row['max_students']) if pd.notna(row['max_students']) else None
        for _, row in projects_df.iterrows()
    }

    # Run all validation checks
    project_errors = validate_project_data(projects_df)
    supervisor_errors = validate_supervisor_data(supervisors_df)
    student_errors = validate_student_data(students_df, projects_df)
    diversity_warnings = validate_supervisor_diversity(students_df, projects_df)

    if student_errors or project_errors or supervisor_errors:
        st.error("Data validation failed:")
        if student_errors:
            st.subheader("Student Data Issues")
            for err in student_errors:
                st.write("-", err)
        if project_errors:
            st.subheader("Project Data Issues")
            for err in project_errors:
                st.write("-", err)
        if supervisor_errors:
            st.subheader("Supervisor Data Issues")
            for err in supervisor_errors:
                st.write("-", err)
        st.stop()

    if diversity_warnings:
        st.warning("Students with limited supervisor diversity in their choices:")
        st.write(diversity_warnings)

    st.success("Data validation passed!")

    # --- Matching Algorithms ---
    def greedy_matching(students_df, projects_df, supervisor_capacity, project_capacity, preallocated_df):
        allocation = {}
        supervisor_load = {sup_id: 0 for sup_id in supervisor_capacity}
        project_load = {pid: 0 for pid in projects_df['project_id']}
        for _, row in preallocated_df.iterrows():
            sid, pid, sup = row['student_id'], row['project_id'], row['supervisor_id']
            allocation[sid] = pid
            supervisor_load[sup] += 1
            project_load[pid] += 1

        students_df_shuffled = students_df.sample(frac=1).reset_index(drop=True)
        for _, row in students_df_shuffled.iterrows():
            sid = row['student_id']
            if sid in allocation:
                continue
            for choice in ['choice_1', 'choice_2', 'choice_3']:
                pid = row[choice]
                project_row = projects_df[projects_df['project_id'] == pid]
                if project_row.empty: continue
                sup = project_row['supervisor_id'].values[0]
                max_cap = project_capacity[pid]
                if supervisor_load[sup] < supervisor_capacity[sup] and (max_cap is None or project_load[pid] < max_cap):
                    allocation[sid] = pid
                    supervisor_load[sup] += 1
                    project_load[pid] += 1
                    break
        return allocation

    def stable_marriage_matching(students_df, projects_df, supervisor_capacity, project_capacity, preallocated_df):
        allocation = {}
        supervisor_load = {sup_id: 0 for sup_id in supervisor_capacity}
        project_load = {pid: 0 for pid in projects_df['project_id']}
        for _, row in preallocated_df.iterrows():
            sid, pid, sup = row['student_id'], row['project_id'], row['supervisor_id']
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
            if project_row.empty: continue
            sup = project_row['supervisor_id'].values[0]
            max_cap = project_capacity.get(pid)
            if supervisor_load[sup] < supervisor_capacity[sup] and (max_cap is None or project_load[pid] < max_cap):
                allocation[sid] = pid
                supervisor_load[sup] += 1
                project_load[pid] += 1
            else:
                free_students.append(sid)
        return allocation

    def linear_programming_matching(students_df, projects_df, supervisor_capacity, project_capacity, preallocated_df):
        prob = LpProblem("Student_Project_Matching", LpMaximize)
        students = list(students_df['student_id'])
        projects = list(projects_df['project_id'])
        project_supervisors = projects_df.set_index('project_id')['supervisor_id'].to_dict()
        student_choices = {
            row['student_id']: {
                row['choice_1']: 3,
                row['choice_2']: 2,
                row['choice_3']: 1
            } for _, row in students_df.iterrows()
        }
        x = LpVariable.dicts("assign", [(s, p) for s in students for p in projects], cat=LpBinary)
        prob += lpSum(x[(s, p)] * student_choices.get(s, {}).get(p, 0) for s in students for p in projects)
        for s in students:
            prob += lpSum(x[(s, p)] for p in projects) <= 1
        for _, row in preallocated_df.iterrows():
            sid, pid = row['student_id'], row['project_id']
            for p in projects:
                prob += x[(sid, p)] == int(p == pid)
        for p in projects:
            max_cap = project_capacity.get(p)
            if max_cap is not None:
                prob += lpSum(x[(s, p)] for s in students) <= max_cap
        for sup, sup_cap in supervisor_capacity.items():
            sup_projects = [p for p, s in project_supervisors.items() if s == sup]
            prob += lpSum(x[(s, p)] for s in students for p in sup_projects) <= sup_cap
        prob.solve(PULP_CBC_CMD(msg=False))
        allocation = {
            s: p for s in students for p in projects
            if x[(s, p)].varValue is not None and x[(s, p)].varValue > 0.5
        }
        return allocation

    # --- Run all matchings ---
    with st.spinner("Running matching algorithms..."):
        greedy = greedy_matching(students_df, projects_df, supervisor_capacity, project_capacity, preallocated_df)
        stable = stable_marriage_matching(students_df, projects_df, supervisor_capacity, project_capacity, preallocated_df)
        lp = linear_programming_matching(students_df, projects_df, supervisor_capacity, project_capacity, preallocated_df)

    st.success("Matching complete!")
    st.subheader("Match Summary")

    # --- Results Overview ---
    def summarize(allocation, method):
        total = len(students_df)
        matched = len(allocation)
        scores = []
        choice_counts = {1: 0, 2: 0, 3: 0}

        for _, row in students_df.iterrows():
            sid = row['student_id']
            pid = allocation.get(sid)

            if pid == row['choice_1']:
                scores.append(3)
                choice_counts[1] += 1
            elif pid == row['choice_2']:
                scores.append(2)
                choice_counts[2] += 1
            elif pid == row['choice_3']:
                scores.append(1)
                choice_counts[3] += 1
            elif pid is not None:
                scores.append(0)
            else:
                scores.append(0)

        return {
            "Method": method,
            "Matched": matched,
            "Unmatched": total - matched,
            "Avg Score": round(np.mean(scores), 2),
            "1st Choice": choice_counts[1],
            "2nd Choice": choice_counts[2],
            "3rd Choice": choice_counts[3],
        }

    # --- Summary Table ---
    summary_df = pd.DataFrame([
        summarize(greedy, "Greedy"),
        summarize(stable, "Stable Marriage"),
        summarize(lp, "Linear Programming")
    ])

    st.dataframe(summary_df, use_container_width=True)

    # --- Match Quality Analysis ---
    def analyze_match_quality(allocation, students_df, projects_df, supervisors_df, method_name="Method"):
        st.subheader(f"Match Quality Analysis – {method_name}")

        # 1. Choice preference distribution
        distribution = {'choice_1': 0, 'choice_2': 0, 'choice_3': 0, 'unmatched': 0, 'other': 0}
        for _, row in students_df.iterrows():
            sid = row['student_id']
            assigned = allocation.get(sid, None)
            if assigned is None:
                distribution['unmatched'] += 1
            elif assigned == row['choice_1']:
                distribution['choice_1'] += 1
            elif assigned == row['choice_2']:
                distribution['choice_2'] += 1
            elif assigned == row['choice_3']:
                distribution['choice_3'] += 1
            else:
                distribution['other'] += 1

        st.markdown("**Choice Preference Distribution:**")
        st.dataframe(pd.DataFrame(distribution.items(), columns=["Choice", "Count"]))

        # 2. Supervisor load distribution
        proj_to_sup = projects_df.set_index('project_id')['supervisor_id'].to_dict()
        supervisor_load = {sup: 0 for sup in supervisors_df['supervisor_id']}
        for sid, pid in allocation.items():
            sup = proj_to_sup.get(pid)
            if sup is not None:
                supervisor_load[sup] += 1

        loads = np.array(list(supervisor_load.values()))
        st.markdown("**Supervisor Load Distribution:**")
        st.write(f"Min load: {loads.min()} | Max load: {loads.max()} | Mean: {loads.mean():.2f} | Std Dev: {loads.std():.2f}")
        sup_load_df = pd.DataFrame.from_dict(supervisor_load, orient='index', columns=['Students Assigned'])
        st.dataframe(sup_load_df)

        # 3. Students assigned outside choices
        outside = []
        for _, row in students_df.iterrows():
            sid = row['student_id']
            assigned = allocation.get(sid, None)
            if assigned is not None and assigned not in [row['choice_1'], row['choice_2'], row['choice_3']]:
                outside.append(sid)

        st.markdown("**Assigned Outside Top 3 Choices:**")
        st.write(f"{len(outside)} student(s) assigned outside their top 3.")
        if outside:
            st.dataframe(pd.DataFrame(outside, columns=["Student ID"]))

        # 4. Project utilization
        usage = {pid: 0 for pid in projects_df['project_id']}
        for pid in allocation.values():
            usage[pid] = usage.get(pid, 0) + 1

        st.markdown("**Project Utilization:**")
        usage_df = pd.DataFrame.from_dict(usage, orient='index', columns=["Assigned Count"])
        usage_df.index.name = "Project ID"
        st.dataframe(usage_df)

    st.subheader("Analysis")
    with st.expander("View Greedy Matching Analysis"):
        analyze_match_quality(greedy, students_df, projects_df, supervisors_df, "Greedy Matching")
    with st.expander("View Stable Marriage Analysis"):
        analyze_match_quality(stable, students_df, projects_df, supervisors_df, "Stable Marriage")
    with st.expander("View Linear Programming Analysis"):
        analyze_match_quality(lp, students_df, projects_df, supervisors_df, "Linear Programming")

    # --- Satisfaction Analysis ---
    def compute_satisfaction_scores(allocation, students_df, method_name="Method"):
        st.subheader(f"Student Satisfaction Score – {method_name}")

        score_weights = {'choice_1': 3, 'choice_2': 2, 'choice_3': 1}
        scores = []
        unmatched = 0

        for _, row in students_df.iterrows():
            sid = row['student_id']
            assigned = allocation.get(sid)

            if assigned == row['choice_1']:
                scores.append(score_weights['choice_1'])
            elif assigned == row['choice_2']:
                scores.append(score_weights['choice_2'])
            elif assigned == row['choice_3']:
                scores.append(score_weights['choice_3'])
            elif assigned is None:
                unmatched += 1
                scores.append(0)
            else:
                scores.append(0)

        avg_score = np.mean(scores)

        st.markdown(f"**Average Satisfaction Score:** {avg_score:.2f}")
        st.markdown(f"**Unmatched Students:** {unmatched}")

        # Histogram plot using Streamlit
        fig, ax = plt.subplots(figsize=(6, 4))
        ax.hist(scores, bins=[-0.5, 0.5, 1.5, 2.5, 3.5], edgecolor='black', align='mid', rwidth=0.8)
        ax.set_xticks([0, 1, 2, 3])
        ax.set_xlabel("Satisfaction Score (0–3)")
        ax.set_ylabel("Number of Students")
        ax.set_title(f"Satisfaction Score Distribution – {method_name}")
        ax.grid(axis='y', linestyle='--', alpha=0.7)
        st.pyplot(fig)

    st.subheader("Satisfaction")
    with st.expander("Satisfaction – Greedy Matching"):
        compute_satisfaction_scores(greedy, students_df, "Greedy Matching")
    with st.expander("Satisfaction – Stable Marriage"):
        compute_satisfaction_scores(stable, students_df, "Stable Marriage")
    with st.expander("Satisfaction – Linear Programming"):
        compute_satisfaction_scores(lp, students_df, "Linear Programming")

    # --- Supervisor Load Analysis ---
    def analyze_supervisor_load(allocation, projects_df, supervisors_df, method_name="Method"):
        st.subheader(f"Supervisor Load Analysis – {method_name}")

        proj_to_sup = projects_df.set_index('project_id')['supervisor_id'].to_dict()
        sup_capacities = supervisors_df.set_index('supervisor_id')['capacity'].to_dict()

        supervisor_load = {sup: 0 for sup in sup_capacities}

        # Count assigned students per supervisor
        for pid in allocation.values():
            sup = proj_to_sup.get(pid)
            if sup in supervisor_load:
                supervisor_load[sup] += 1

        # Format results
        results = []
        for sup_id, load in supervisor_load.items():
            cap = sup_capacities.get(sup_id, 0)
            name = supervisors_df.loc[supervisors_df['supervisor_id'] == sup_id, 'supervisor_name'].values[0]

            if load > cap:
                status = "OVERLOADED"
            elif load < cap:
                status = "UNDERUSED"
            else:
                status = "OPTIMAL"

            results.append({
                "Supervisor Name": name,
                "Supervisor ID": sup_id,
                "Assigned": load,
                "Capacity": cap,
                "Status": status
            })

        df = pd.DataFrame(results)
        st.dataframe(df)

    st.subheader("Supervisor Load Analysis")
    with st.expander("Greedy Matching – Supervisor Load"):
        analyze_supervisor_load(greedy, projects_df, supervisors_df, "Greedy")
    with st.expander("Stable Marriage – Supervisor Load"):
        analyze_supervisor_load(stable, projects_df, supervisors_df, "Stable Marriage")
    with st.expander("Linear Programming – Supervisor Load"):
        analyze_supervisor_load(lp, projects_df, supervisors_df, "Linear Programming")

    # --- Project Popularity Analysis ---
    def analyze_project_popularity_and_utilization(students_df, projects_df, allocation, method_name="Method"):
        st.subheader(f"Project Popularity & Utilization – {method_name}")

        # Count how many students requested each project
        popularity = {pid: 0 for pid in projects_df['project_id']}
        for _, row in students_df.iterrows():
            for col in ['choice_1', 'choice_2', 'choice_3']:
                pid = row[col]
                if pid in popularity:
                    popularity[pid] += 1

        # Count how many were assigned to each project
        utilization = {pid: 0 for pid in projects_df['project_id']}
        for pid in allocation.values():
            if pid in utilization:
                utilization[pid] += 1

        # Build final table
        data = []
        for pid in projects_df['project_id']:
            requested = popularity.get(pid, 0)
            assigned = utilization.get(pid, 0)

            if requested == 0:
                status = "NEVER PICKED"
            elif assigned == 0:
                status = "NEVER ASSIGNED"
            elif assigned > requested:
                status = "OVER-ASSIGNED"
            elif assigned < requested:
                status = "UNDER-ASSIGNED"
            else:
                status = "MATCHED"

            data.append({
                "Project ID": pid,
                "Requested": requested,
                "Assigned": assigned,
                "Status": status
            })

        df = pd.DataFrame(data)
        st.dataframe(df, use_container_width=True)

    st.subheader("Project Popularity & Utilization")
    with st.expander("Greedy Matching – Project Popularity"):
        analyze_project_popularity_and_utilization(students_df, projects_df, greedy, "Greedy")
    with st.expander("Stable Marriage – Project Popularity"):
        analyze_project_popularity_and_utilization(students_df, projects_df, stable, "Stable Marriage")
    with st.expander("Linear Programming – Project Popularity"):
        analyze_project_popularity_and_utilization(students_df, projects_df, lp, "Linear Programming")


    # --- Downloads ---
    from tempfile import NamedTemporaryFile
    from openpyxl import Workbook

    def export_excel(allocations):
        from openpyxl import Workbook
        from openpyxl.utils.dataframe import dataframe_to_rows
        from io import BytesIO

        output = BytesIO()
        wb = Workbook()

        # Mapping for quick lookup
        project_lookup = projects_df.set_index('project_id')[['project_title', 'supervisor_id']].to_dict('index')
        supervisor_lookup = supervisors_df.set_index('supervisor_id')['supervisor_name'].to_dict()

        for name, alloc in allocations.items():
            ws = wb.create_sheet(title=name)
            data = []

            for _, row in students_df.iterrows():
                sid = row['student_id']
                sname = row['student_name']
                assigned_pid = alloc.get(sid, None)

                # Determine choice rank
                if assigned_pid == row['choice_1']:
                    choice_rank = "1st"
                elif assigned_pid == row['choice_2']:
                    choice_rank = "2nd"
                elif assigned_pid == row['choice_3']:
                    choice_rank = "3rd"
                elif assigned_pid is not None:
                    choice_rank = "Outside Top 3"
                else:
                    choice_rank = "Unassigned"

                if assigned_pid:
                    project_info = project_lookup.get(assigned_pid, {})
                    project_name = project_info.get('project_title', 'Unknown')
                    supervisor_id = project_info.get('supervisor_id', 'Unknown')
                    supervisor_name = supervisor_lookup.get(supervisor_id, 'Unknown')
                else:
                    assigned_pid = 'UNASSIGNED'
                    project_name = ''
                    supervisor_name = ''

                data.append({
                    "student_id": sid,
                    "student_name": sname,
                    "assigned_project_id": assigned_pid,
                    "assigned_project_name": project_name,
                    "supervisor_name": supervisor_name,
                    "assigned_choice": choice_rank
                })

            df = pd.DataFrame(data)
            for r in dataframe_to_rows(df, index=False, header=True):
                ws.append(r)

        # Remove default sheet if exists
        if "Sheet" in wb.sheetnames:
            del wb["Sheet"]

        wb.save(output)
        output.seek(0)
        return output

    excel_data = export_excel({
        "Greedy": greedy,
        "Stable Marriage": stable,
        "Linear Programming": lp
    })

    st.download_button("Download Excel Results", data=excel_data, file_name="matchings.xlsx")

# else:
    # st.info("Please upload all required CSV files to begin.")

