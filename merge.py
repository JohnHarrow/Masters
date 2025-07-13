import pandas as pd

# Load the CSVs
students = pd.read_csv("students.csv")
projects = pd.read_csv("projects.csv")
supervisors = pd.read_csv("supervisors.csv")
preallocated = pd.read_csv("preallocated.csv")

# Write to Excel
with pd.ExcelWriter("combined_input.xlsx", engine="openpyxl") as writer:
    students.to_excel(writer, sheet_name="students", index=False)
    projects.to_excel(writer, sheet_name="projects", index=False)
    supervisors.to_excel(writer, sheet_name="supervisors", index=False)
    preallocated.to_excel(writer, sheet_name="preallocated", index=False)