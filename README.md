# Olympic Weightlifting Program Script

This script generates a 12-week Olympic weightlifting program tailored for strength and technique development. It creates an Excel file with individual tabs for each week and structured workout plans for Days 1, 2, 4, and 5. The program progressively adjusts intensity across three phases: Volume, Technique, and Peaking.

## Features

- **12-Week Plan**: Automatically generates periodized workouts.
- **Phase-Based Programming**: Divided into three 4-week phases:
  - **Volume Phase** (Weeks 1–4): Build base strength with moderate intensity.
  - **Technique Phase** (Weeks 5–8): Focus on precision and form with slightly higher intensity.
  - **Peaking Phase** (Weeks 9–12): Prepare for peak performance with high intensity and reduced volume.
- **Exercise Breakdown**: Each day includes structured exercises, sets, reps, and weight recommendations.
- **Adjustable Intensity Based on 1RM**: User is prompted to input their 1RM values for the four key lifts (Clean & Jerk, Snatch, Front Squat, and Back Squat). The intensity for each exercise is adjusted accordingly, and the program dynamically calculates the weight based on 1RM and intensity percentage.
- **Non-1RM Lifts Exclusion**: For exercises that do not have a 1RM value, the weight and intensity fields are left blank.
- **Customizable Output**: Easy to modify for personalized training needs.

## Prerequisites

Make sure you have the following installed on your system:

- Python 3.7+
- Required Python libraries:
  ```bash
  pip install pandas xlsxwriter
  ```

## How to Use

1. **Download the Script**
   Copy the provided Python script into a file named `generate_weightlifting_program.py`.

2. **Run the Script**
   Execute the script in your Python environment:

   ```bash
   python generate_weightlifting_program.py
   ```

3. **Input 1RM values** When prompted enter your 1RM values for the following exercises:

   - **Clean & Jerk**
   - **Snatch**
   - **Front Squat**
   - **Back Squat**

4. **Output File**
   The script will generate an Excel file named `Olympic_Weightlifting_Program_Weeks_1_to_12.xlsx` in the same directory as the script.

## Customization

Feel free to modify the script:

- **Adjust Phases**: Change intensity factors or week ranges in the `phases` dictionary.
- **Modify Exercises**: Edit the `week_structure` dictionary to include or remove exercises.
- **Change File Name**: Update the `file_path_weeks` variable to save the file with a custom name.

## Example Output for Week 1, Day 1

### Day 1

| Exercise             | Sets | Reps | Weight (lbs/kg) | Intensity (%) | Notes                   |
| -------------------- | ---- | ---- | --------------- | ------------- | ----------------------- |
| Snatch               | 5    | 3    | 135 lbs         | 70%           | Focus on form and speed |
| Clean & Jerk         | 4    | 2    | 140 lbs         | 75%           | Powerful extension      |
| Push Press           | 4    | 5    |                 |               |                         |
| Snatch-Grip Deadlift | 4    | 3    |                 |               |                         |

## Support

If you encounter any issues, feel free to reach out or adjust the script to match your specific requirements.

Happy lifting!
