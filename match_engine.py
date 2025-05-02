
import pandas as pd
from collections import defaultdict

# Load mentor expertise data
mentors_df = pd.read_excel("data/mentors_2024.xlsx")
startups_df = pd.read_excel("data/startups_2024.xlsx")

# Convert to dictionaries for easier processing
MentorData = {}
StartupData = {}

# Preparing mentor data structure
for _, row in mentors_df.iterrows():
    mentor_id = row["Mentor"]
    expertise_1 = str(row["Q1"])
    expertise_2 = str(row["Q2"])
    expertise_3 = str(row["Q3"])
    max_sessions = int(row["Max_Sessions"])
    MentorData[mentor_id] = [expertise_1, expertise_2, expertise_3, max_sessions]

# Preparing startup preference structure
for _, row in startups_df.iterrows():
    startup_id = str(row["Startup"])
    pref_1 = str(row["Q1"])
    pref_2 = str(row["Q2"])
    pref_3 = str(row["Q3"])
    StartupData[startup_id] = [pref_1, pref_2, pref_3, 0]  # last entry for session count

# Matching logic based on overlap of Q1–Q3 preferences
pairing_scores = {}

for mentor_id, mentor_prefs in MentorData.items():
    mentor_expertise = mentor_prefs[:3]
    temp_scores = []
    
    for startup_id, startup_prefs in StartupData.items():
        score = 0
        s_q1, s_q2, s_q3 = startup_prefs[:3]
        m_q1, m_q2, m_q3 = mentor_expertise

        # Custom weighted scoring
        if s_q1 == m_q1: score += 100
        if s_q2 == m_q2: score += 100
        if s_q3 == m_q3: score += 200  # Higher weight to deep alignment

        if score >= 100:
            temp_scores.append((score, startup_id))
    
    temp_scores.sort(reverse=True)
    pairing_scores[mentor_id] = temp_scores

# Final allocation
final_matches = defaultdict(list)
nosessions = defaultdict(int)
total_sessions = 0

all_pairings = []
for mentor_id, matches in pairing_scores.items():
    for score, startup_id in matches:
        all_pairings.append((score, (mentor_id, startup_id)))

# Prioritize higher scores
all_pairings.sort(reverse=True)

# Allocate sessions
for score, (mentor_id, startup_id) in all_pairings:
    if MentorData[mentor_id][3] > 0 and StartupData[startup_id][3] < 4:
        MentorData[mentor_id][3] -= 1
        StartupData[startup_id][3] += 1
        final_matches[mentor_id].append((startup_id, score))
        nosessions[startup_id] += 1
        total_sessions += 1

# Save output to Excel
from pandas import ExcelWriter

with ExcelWriter("output/final_schedule.xlsx") as writer:
    mentor_view = []
    startup_view = defaultdict(list)

    for mentor_id, startups in final_matches.items():
        row = [mentor_id] + [f"S{sid} ({score})" for sid, score in startups]
        mentor_view.append(row)
        for sid, _ in startups:
            startup_view[sid].append(mentor_id)

    pd.DataFrame(mentor_view).to_excel(writer, sheet_name="Mentor to Startups", index=False)

    startup_rows = []
    for startup_id, mentors in startup_view.items():
        row = [f"S{startup_id}"] + mentors
        startup_rows.append(row)
    
    pd.DataFrame(startup_rows).to_excel(writer, sheet_name="Startups to Mentors", index=False)

print("Matching complete. Output saved to 'output/final_schedule.xlsx'")
