import pandas as pd

# -----------------------------
# Load Excel file
# -----------------------------
file_path = "interviews.xlsx"
df = pd.read_excel(file_path)
print("Current Interview Data:")
print(df)

# -----------------------------
# Add new candidate
# -----------------------------
new_candidate = {
    "CandidateID": 103,
    "CandidateName": "Ankit Sharma",
    "Email": "ankit@example.com",
    "Phone": "7777777777",
    "ResumeLink": "link_resume.pdf",
    "InterviewDate": "2026-02-27",
    "InterviewerName": "Alice Johnson",
    "Status": "Pending",
    "Remarks": ""
}

df = df.append(new_candidate, ignore_index=True)
print("\nNew candidate added successfully!")

# -----------------------------
# Update interview status
# -----------------------------
candidate_id_to_update = 101  # Example: Dipak Kumar
df.loc[df['CandidateID'] == candidate_id_to_update, 'Status'] = "Completed"
df.loc[df['CandidateID'] == candidate_id_to_update, 'Remarks'] = "Good technical skills"
print(f"\nCandidate ID {candidate_id_to_update} status updated!")

# -----------------------------
# Save updated Excel
# -----------------------------
df.to_excel(file_path, index=False)
print("\nExcel file updated successfully!")

# -----------------------------
# Generate summary report
# -----------------------------
total = len(df)
completed = len(df[df['Status'] == "Completed"])
pending = len(df[df['Status'] == "Pending"])

print("\n--- Summary Report ---")
print(f"Total Interviews: {total}")
print(f"Completed Interviews: {completed}")
print(f"Pending Interviews: {pending}")

# Interviewer-wise summary
interviewer_summary = df.groupby('InterviewerName')['Status'].value_counts()
print("\nInterviewer-wise Status Summary:")
print(interviewer_summary)