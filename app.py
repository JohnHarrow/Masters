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

st.title("🎓 Student-Project Allocation Tool")

# --- File Upload ---
st.sidebar.header("📂 Upload Required CSV Files")
students_file = st.sidebar.file_uploader("Upload students.csv", type="csv")
projects_file = st.sidebar.file_uploader("Upload projects.csv", type="csv")
supervisors_file = st.sidebar.file_uploader("Upload supervisors.csv", type="csv")
preallocated_file = st.sidebar.file_uploader("Upload preallocated.csv", type="csv")

if all([students_file, projects_file, supervisors_file, preallocated_file]):
    students_df = pd.read_csv(students_file)
    projects_df = pd.read_csv(projects_file)
    supervisors_df = pd.read_csv(supervisors_file)
    preallocated_df = pd.read_csv(preallocated_file)

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

    # --- Capacities ---
    supervisor_capacity = {
        row['supervisor_id']: int(row['capacity']) if pd.notna(row['capacity']) else 3
        for _, row in supervisors_df.iterrows()
    }

    project_capacity = {
        row['project_id']: int(row['max_students']) if pd.notna(row['max_students']) else None
        for _, row in projects_df.iterrows()
    }

    errors = validate_student_data(students_df, projects_df)
    if errors:
        st.error("🚫 Data validation failed:")
        for err in errors:
            st.write("-", err)
        st.stop()

    warnings = validate_supervisor_diversity(students_df, projects_df)
    if warnings:
        st.warning("⚠️ Students with limited supervisor diversity in choices:")
        st.write(warnings)

    st.success("✅ Data validation passed!")

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

    st.success("✅ Matching complete!")
    st.subheader("📊 Match Summary")

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
        st.subheader(f"📊 Match Quality Analysis – {method_name}")

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

        st.markdown("**🎯 Choice Preference Distribution:**")
        st.dataframe(pd.DataFrame(distribution.items(), columns=["Choice", "Count"]))

        # 2. Supervisor load distribution
        proj_to_sup = projects_df.set_index('project_id')['supervisor_id'].to_dict()
        supervisor_load = {sup: 0 for sup in supervisors_df['supervisor_id']}
        for sid, pid in allocation.items():
            sup = proj_to_sup.get(pid)
            if sup is not None:
                supervisor_load[sup] += 1

        loads = np.array(list(supervisor_load.values()))
        st.markdown("**🧑‍🏫 Supervisor Load Distribution:**")
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

        st.markdown("**📍 Assigned Outside Top 3 Choices:**")
        st.write(f"{len(outside)} student(s) assigned outside their top 3.")
        if outside:
            st.dataframe(pd.DataFrame(outside, columns=["Student ID"]))

        # 4. Project utilization
        usage = {pid: 0 for pid in projects_df['project_id']}
        for pid in allocation.values():
            usage[pid] = usage.get(pid, 0) + 1

        st.markdown("**📦 Project Utilization:**")
        usage_df = pd.DataFrame.from_dict(usage, orient='index', columns=["Assigned Count"])
        usage_df.index.name = "Project ID"
        st.dataframe(usage_df)

    with st.expander("🔍 View Greedy Matching Analysis"):
        analyze_match_quality(greedy, students_df, projects_df, supervisors_df, "Greedy Matching")
    with st.expander("🔍 View Stable Marriage Analysis"):
        analyze_match_quality(stable, students_df, projects_df, supervisors_df, "Stable Marriage")
    with st.expander("🔍 View Linear Programming Analysis"):
        analyze_match_quality(lp, students_df, projects_df, supervisors_df, "Linear Programming")

    # --- Satisfaction Analysis ---
    def compute_satisfaction_scores(allocation, students_df, method_name="Method"):
        st.subheader(f"🎯 Student Satisfaction Score – {method_name}")

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


    with st.expander("📈 Satisfaction – Greedy Matching"):
        compute_satisfaction_scores(greedy, students_df, "Greedy Matching")
    with st.expander("📈 Satisfaction – Stable Marriage"):
        compute_satisfaction_scores(stable, students_df, "Stable Marriage")
    with st.expander("📈 Satisfaction – Linear Programming"):
        compute_satisfaction_scores(lp, students_df, "Linear Programming")

    # --- Supervisor Load Analysis ---
    def analyze_supervisor_load(allocation, projects_df, supervisors_df, method_name="Method"):
        st.subheader(f"📋 Supervisor Load Analysis – {method_name}")

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
                status = "❌ OVERLOADED"
            elif load < cap:
                status = "⚠️ UNDERUSED"
            else:
                status = "✅ OPTIMAL"

            results.append({
                "Supervisor Name": name,
                "Supervisor ID": sup_id,
                "Assigned": load,
                "Capacity": cap,
                "Status": status
            })

        df = pd.DataFrame(results)
        st.dataframe(df)

    st.subheader("🧑‍🏫 Supervisor Load Analysis")
    with st.expander("🟦 Greedy Matching – Supervisor Load"):
        analyze_supervisor_load(greedy, projects_df, supervisors_df, "Greedy")
    with st.expander("🟪 Stable Marriage – Supervisor Load"):
        analyze_supervisor_load(stable, projects_df, supervisors_df, "Stable Marriage")
    with st.expander("🟧 Linear Programming – Supervisor Load"):
        analyze_supervisor_load(lp, projects_df, supervisors_df, "Linear Programming")

    # --- Project Popularity Analysis ---
    def analyze_project_popularity_and_utilization(students_df, projects_df, allocation, method_name="Method"):
        st.subheader(f"📈 Project Popularity & Utilization – {method_name}")

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
                status = "🧊 NEVER PICKED"
            elif assigned == 0:
                status = "⚠️ NEVER ASSIGNED"
            elif assigned > requested:
                status = "❌ OVER-ASSIGNED"
            elif assigned < requested:
                status = "🔼 UNDER-ASSIGNED"
            else:
                status = "✅ MATCHED"

            data.append({
                "Project ID": pid,
                "Requested": requested,
                "Assigned": assigned,
                "Status": status
            })

        df = pd.DataFrame(data)
        st.dataframe(df, use_container_width=True)

    st.subheader("📦 Project Popularity & Utilization")
    with st.expander("🟦 Greedy Matching – Project Popularity"):
        analyze_project_popularity_and_utilization(students_df, projects_df, greedy, "Greedy")
    with st.expander("🟪 Stable Marriage – Project Popularity"):
        analyze_project_popularity_and_utilization(students_df, projects_df, stable, "Stable Marriage")
    with st.expander("🟧 Linear Programming – Project Popularity"):
        analyze_project_popularity_and_utilization(students_df, projects_df, lp, "Linear Programming")


    # --- Downloads ---
    from tempfile import NamedTemporaryFile
    from openpyxl import Workbook

    def export_excel(allocations):
        from openpyxl import Workbook
        from openpyxl.utils.dataframe import dataframe_to_rows
        from openpyxl.styles import Font
        from io import BytesIO

        output = BytesIO()
        wb = Workbook()
        for name, alloc in allocations.items():
            ws = wb.create_sheet(title=name)
            data = []
            for _, row in students_df.iterrows():
                sid = row['student_id']
                assigned = alloc.get(sid, 'UNASSIGNED')
                data.append({
                    "student_id": sid,
                    "student_name": row['student_name'],
                    "assigned_project": assigned
                })
            df = pd.DataFrame(data)
            for r in dataframe_to_rows(df, index=False, header=True):
                ws.append(r)
        del wb["Sheet"]
        wb.save(output)
        return output

    excel_data = export_excel({
        "Greedy": greedy,
        "Stable Marriage": stable,
        "Linear Programming": lp
    })

    st.download_button("📥 Download Excel Results", data=excel_data, file_name="matchings.xlsx")

else:
    st.info("⬅️ Please upload all required CSV files to begin.")

