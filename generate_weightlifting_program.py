import pandas as pd

# Define the structure of the Olympic weightlifting program
week_structure = {
    "Day 1": [
        {"Exercise": "Snatch", "Sets": 5, "Reps": 3, "Weight (lbs/kg)": "", "Intensity (%)": "", "Notes": ""},
        {"Exercise": "Clean & Jerk", "Sets": 4, "Reps": 2, "Weight (lbs/kg)": "", "Intensity (%)": "", "Notes": ""},
        {"Exercise": "Push Press", "Sets": 4, "Reps": 5, "Weight (lbs/kg)": "", "Intensity (%)": "", "Notes": ""},
        {"Exercise": "Snatch-Grip Deadlift", "Sets": 4, "Reps": 3, "Weight (lbs/kg)": "", "Intensity (%)": "", "Notes": ""},
    ],
    "Day 2": [
        {"Exercise": "Hang Power Clean", "Sets": 4, "Reps": 3, "Weight (lbs/kg)": "", "Intensity (%)": "", "Notes": ""},
        {"Exercise": "Front Squat", "Sets": 5, "Reps": 3, "Weight (lbs/kg)": "", "Intensity (%)": "", "Notes": ""},
        {"Exercise": "Bench Press", "Sets": 4, "Reps": 5, "Weight (lbs/kg)": "", "Intensity (%)": "", "Notes": ""},
        {"Exercise": "Bent-over Rows", "Sets": 4, "Reps": 8, "Weight (lbs/kg)": "", "Intensity (%)": "", "Notes": ""},
        {"Exercise": "Face Pulls", "Sets": 3, "Reps": 12, "Weight (lbs/kg)": "", "Intensity (%)": "", "Notes": ""},
    ],
    "Day 4": [
        {"Exercise": "Hang Snatch", "Sets": 4, "Reps": 3, "Weight (lbs/kg)": "", "Intensity (%)": "", "Notes": ""},
        {"Exercise": "Back Squat", "Sets": 5, "Reps": 5, "Weight (lbs/kg)": "", "Intensity (%)": "", "Notes": ""},
        {"Exercise": "Bulgarian Split Squats", "Sets": 3, "Reps": 8, "Weight (lbs/kg)": "", "Intensity (%)": "", "Notes": ""},
        {"Exercise": "GHD Hamstring Curls", "Sets": 3, "Reps": 10, "Weight (lbs/kg)": "", "Intensity (%)": "", "Notes": ""},
    ],
    "Day 5": [
        {"Exercise": "Power Snatch", "Sets": 5, "Reps": 3, "Weight (lbs/kg)": "", "Intensity (%)": "", "Notes": ""},
        {"Exercise": "Bench Press", "Sets": 4, "Reps": 5, "Weight (lbs/kg)": "", "Intensity (%)": "", "Notes": ""},
        {"Exercise": "Pendlay Row", "Sets": 4, "Reps": 6, "Weight (lbs/kg)": "", "Intensity (%)": "", "Notes": ""},
        {"Exercise": "Pull-Ups", "Sets": 3, "Reps": 10, "Weight (lbs/kg)": "Bodyweight", "Intensity (%)": "", "Notes": ""},
    ],
}

# Function to ask the user for 1RM values
def get_one_rm_values():
    print("Please enter the 1RM values for the following lifts (in pounds):")
    clean_jerk = float(input("Clean & Jerk: "))
    snatch = float(input("Snatch: "))
    front_squat = float(input("Front Squat: "))
    back_squat = float(input("Back Squat: "))
    
    return {
        "Clean & Jerk": clean_jerk,
        "Snatch": snatch,
        "Front Squat": front_squat,
        "Back Squat": back_squat
    }

# Function to adjust weights based on intensity
def adjust_weights_based_on_intensity(exercise, intensity_percentage, one_rm_values):
    lift_name = exercise["Exercise"]
    if lift_name in one_rm_values:
        one_rm = one_rm_values[lift_name]
        adjusted_weight = one_rm * (intensity_percentage / 100)
        return round(adjusted_weight, 2)  # Round to 2 decimal places
    return 0  # If it's not a 1RM lift, return 0

# Function to add the Week and Day columns (with no repetition)
def add_week_day_column(week_num, week_data):
    for day, exercises in week_data.items():
        # Add the week number to the first exercise of the day
        first_exercise = True
        for exercise in exercises:
            if first_exercise:
                exercise["Week"] = week_num
                exercise["Day"] = day
                first_exercise = False
            else:
                exercise["Week"] = ""
                exercise["Day"] = ""
    return week_data

# Function to create periodized data for weeks
def generate_week_data(phase, week_num, one_rm_values):
    # Intensity range for each phase
    if phase == "Volume":
        intensity_range = {
            "Snatch": 65,
            "Clean & Jerk": 70,
            "Front Squat": 75,
            "Back Squat": 75,
        }
    elif phase == "Technique":
        intensity_range = {
            "Snatch": 75,
            "Clean & Jerk": 80,
            "Front Squat": 80,
            "Back Squat": 80,
        }
    elif phase == "Peaking":
        intensity_range = {
            "Snatch": 85,
            "Clean & Jerk": 90,
            "Front Squat": 85,
            "Back Squat": 90,
        }

    # Adjust reps and weights for each phase
    week_data = {}
    for day, exercises in week_structure.items():
        week_exercises = []
        for exercise in exercises:
            # Set intensity only for 1RM lifts, else leave it blank
            if exercise["Exercise"] in intensity_range:
                intensity_percentage = intensity_range[exercise["Exercise"]]
                adjusted_weight = adjust_weights_based_on_intensity(exercise, intensity_percentage, one_rm_values)
                week_exercises.append({
                    "Exercise": exercise["Exercise"],
                    "Sets": exercise["Sets"],
                    "Reps": exercise["Reps"],
                    "Weight (lbs/kg)": adjusted_weight,  # Adjusted weight based on 1RM and intensity
                    "Intensity (%)": f"{intensity_percentage}%",
                    "Notes": exercise["Notes"] if exercise["Notes"] else "",
                })
            else:
                week_exercises.append({
                    "Exercise": exercise["Exercise"],
                    "Sets": exercise["Sets"],
                    "Reps": exercise["Reps"],
                    "Weight (lbs/kg)": "",  # Leave weight empty for non-1RM lifts
                    "Intensity (%)": "",  # Leave intensity empty for non-1RM lifts
                    "Notes": exercise["Notes"] if exercise["Notes"] else "",
                })
        week_data[day] = week_exercises
    
    # Add the 'Week' and 'Day' columns to the week_data
    week_data = add_week_day_column(week_num, week_data)
    
    return week_data

# Define the phases and their respective weeks
phases = {
    "Volume": range(1, 5),  # Weeks 1–4
    "Technique": range(5, 9),  # Weeks 5–8
    "Peaking": range(9, 13),  # Weeks 9–12
}

# Get the 1RM values from the user
one_rm_values = get_one_rm_values()

# Generate data for all weeks
all_weeks_data = {}
for phase, weeks in phases.items():
    for week_num in weeks:
        all_weeks_data[f"Week {week_num}"] = generate_week_data(phase, week_num, one_rm_values)

# Export all weeks into a single Excel file with a tab for each week block
file_path_weeks = "C:/Users/kyflu/Downloads/Olympic_Weightlifting_Program_Weeks_1_to_12_with_1RM_adjusted.xlsx"
with pd.ExcelWriter(file_path_weeks, engine="xlsxwriter") as writer:
    # Export the week data for each block (Weeks 1–4, 5–8, 9–12)
    for block, weeks in zip(["Week 1-4", "Week 5-8", "Week 9-12"], [range(1, 5), range(5, 9), range(9, 13)]):
        block_data = []
        for week_num in weeks:
            week_data = all_weeks_data[f"Week {week_num}"]
            # Add the adjusted weight and intensity to the block data
            for day, exercises in week_data.items():
                for exercise in exercises:
                    block_data.append({
                        "Week": exercise["Week"],
                        "Day": exercise["Day"],
                        "Exercise": exercise["Exercise"],
                        "Sets": exercise["Sets"],
                        "Reps": exercise["Reps"],
                        "Weight (lbs/kg)": exercise["Weight (lbs/kg)"],
                        "Intensity (%)": exercise["Intensity (%)"],
                        "Notes": exercise["Notes"],
                    })
        # Write to Excel
        block_df = pd.DataFrame(block_data)
        block_df.to_excel(writer, sheet_name=block, index=False)
    
    # Add 1RM data to a separate tab
    one_rm_data = {
        "Lift": ["Clean & Jerk", "Snatch", "Front Squat", "Back Squat"],
        "1RM (lbs)": [
            one_rm_values["Clean & Jerk"], 
            one_rm_values["Snatch"], 
            one_rm_values["Front Squat"], 
            one_rm_values["Back Squat"]
        ]
    }
    one_rm_df = pd.DataFrame(one_rm_data)
    one_rm_df.to_excel(writer, sheet_name="1RM Data", index=False)

print(f"Program saved as {file_path_weeks}")
