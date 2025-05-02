
import pandas as pd

def load_preferences():
    mentors_df = pd.read_excel("data/mentors_2024.xlsx")
    startups_df = pd.read_excel("data/startups_2024.xlsx")
    return mentors_df, startups_df

def save_schedule(matches):
    df = pd.DataFrame(matches, columns=["Mentor", "Startup", "Score"])
    df.to_excel("output/final_schedule.xlsx", index=False)
