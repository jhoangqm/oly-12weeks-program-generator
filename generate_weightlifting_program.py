import pandas as pd

# Define the structure of the Olympic weightlifting program
week_structure = {
    "Day 1": [
        {"Exercise": "Snatch", "Sets": 5, "Reps": 3, "Weight (lbs/kg)": "70%", "Notes": "Focus on form and speed"},
        {"Exercise": "Clean & Jerk", "Sets": 4, "Reps": 2, "Weight (lbs/kg)": "75%", "Notes": "Powerful extension"},
        {"Exercise": "Push Press", "Sets": 4, "Reps": 5, "Weight (lbs/kg)": "", "Notes": ""},
        {"Exercise": "Snatch-Grip Deadlift", "Sets": 4, "Reps": 3, "Weight (lbs/kg)": "", "Notes": ""},
    ],
    "Day 2": [
        {"Exercise": "Hang Power Clean", "Sets": 4, "Reps": 3, "Weight (lbs/kg)": "", "Notes": "Focus on speed under the bar"},
        {"Exercise": "Front Squat", "Sets": 5, "Reps": 3, "Weight (lbs/kg)": "70%", "Notes": ""},
        {"Exercise": "Bench Press", "Sets": 4, "Reps": 5, "Weight (lbs/kg)": "", "Notes": ""},
        {"Exercise": "Bent-over Rows", "Sets": 4, "Reps": 8, "Weight (lbs/kg)": "", "Notes": ""},
        {"Exercise": "Face Pulls", "Sets": 3, "Reps": 12, "Weight (lbs/kg)": "", "Notes": ""},
    ],
    "Day 4": [
        {"Exercise": "Hang Snatch", "Sets": 4, "Reps": 3, "Weight (lbs/kg)": "", "Notes": "Precision over weight"},
        {"Exercise": "Back Squat", "Sets": 5, "Reps": 5, "Weight (lbs/kg)": "65%", "Notes": ""},
        {"Exercise": "Bulgarian Split Squats", "Sets": 3, "Reps": 8, "Weight (lbs/kg)": "", "Notes": "Each leg"},
        {"Exercise": "GHD Hamstring Curls", "Sets": 3, "Reps": 10, "Weight (lbs/kg)": "", "Notes": ""},
    ],
    "Day 5": [
        {"Exercise": "Power Snatch", "Sets": 5, "Reps": 3, "Weight (lbs/kg)": "", "Notes": "Explosive movement"},
        {"Exercise": "Bench Press", "Sets": 4, "Reps": 5, "Weight (lbs/kg)": "", "Notes": ""},
        {"Exercise": "Pendlay Row", "Sets": 4, "Reps": 6, "Weight (lbs/kg)": "", "Notes": ""},
        {"Exercise": "Pull-Ups", "Sets": 3, "Reps": 10, "Weight (lbs/kg)": "Bodyweight", "Notes": ""},
    ],
}

# Define a function to create periodized data for weeks
def generate_week_data(phase, week_num):
    # Adjustments based on phase
    if phase == "Volume":
        intensity_factor = 0.7  # Moderate intensity
        notes = "Focus on building base strength and technique."
    elif phase == "Technique":
        intensity_factor = 0.75  # Slightly higher intensity
        notes = "Prioritize precision and speed under the bar."
    elif phase == "Peaking":
        intensity_factor = 0.85  # Higher intensity, lower volume
        notes = "Peak strength and power."

    # Adjust reps and weights for each phase
    week_data = {}
    for day, exercises in week_structure.items():
        week_exercises = []
        for exercise in exercises:
            adjusted_weight = (
                f"{int(float(exercise['Weight (lbs/kg)'].strip('%')) * intensity_factor)}%"
                if exercise['Weight (lbs/kg)']
                else ""
            )
            week_exercises.append({
                "Exercise": exercise["Exercise"],
                "Sets": exercise["Sets"],
                "Reps": exercise["Reps"],
                "Weight (lbs/kg)": adjusted_weight,
                "Notes": notes if not exercise["Notes"] else exercise["Notes"],
            })
        week_data[day] = pd.DataFrame(week_exercises)
    return week_data

# Define the phases and their respective weeks
phases = {
    "Volume": range(1, 5),  # Weeks 1–4
    "Technique": range(5, 9),  # Weeks 5–8
    "Peaking": range(9, 13),  # Weeks 9–12
}

# Generate data for all weeks
all_weeks_data = {}
for phase, weeks in phases.items():
    for week_num in weeks:
        all_weeks_data[f"Week {week_num}"] = generate_week_data(phase, week_num)

# Export all weeks into a single Excel file with a tab for each week
file_path_weeks = "Olympic_Weightlifting_Program_Weeks_1_to_12.xlsx"
with pd.ExcelWriter(file_path_weeks, engine="xlsxwriter") as writer:
    for week, days_data in all_weeks_data.items():
        for day, df in days_data.items():
            sheet_name = f"{week} {day}"
            # Limit sheet name to 31 characters (Excel's max)
            sheet_name = sheet_name[:31]
            df.to_excel(writer, sheet_name=sheet_name, index=False)

print(f"Program saved as {file_path_weeks}")
