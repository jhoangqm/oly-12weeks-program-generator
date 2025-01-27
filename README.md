# Olympic Weightlifting Program Script

This script generates a 12-week Olympic weightlifting program tailored for strength and technique development. It creates an Excel file with individual tabs for each week and structured workout plans for Days 1, 2, 4, and 5. The program progressively adjusts intensity across three phases: Volume, Technique, and Peaking.

## Features

- **12-Week Plan**: Automatically generates periodized workouts.
- **Phase-Based Programming**: Divided into three 4-week phases:
  - **Volume Phase** (Weeks 1–4): Build base strength with moderate intensity.
  - **Technique Phase** (Weeks 5–8): Focus on precision and form with slightly higher intensity.
  - **Peaking Phase** (Weeks 9–12): Prepare for peak performance with high intensity and reduced volume.
- **Exercise Breakdown**: Each day includes structured exercises, sets, reps, and weight recommendations.
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
   Copy the provided Python script into a file named `generate_program.py`.

2. **Run the Script**
   Execute the script in your Python environment:

   ```bash
   python generate_program.py
   ```

3. **Output File**
   The script will generate an Excel file named `Olympic_Weightlifting_Program_Weeks_1_to_12.xlsx` in the same directory as the script.

4. **Navigate the Excel File**
   - Each week has its own set of tabs (e.g., `Week 1 Day 1`, `Week 1 Day 2`, etc.).
   - Columns include:
     - **Exercise**: Name of the movement.
     - **Sets/Reps**: Prescribed sets and repetitions.
     - **Weight**: Percentage of your 1-rep max or a weight to be used.
     - **Notes**: Helpful tips or focus points for the exercise.

## Customization

Feel free to modify the script:

- **Adjust Phases**: Change intensity factors or week ranges in the `phases` dictionary.
- **Modify Exercises**: Edit the `week_structure` dictionary to include or remove exercises.
- **Change File Name**: Update the `file_path_weeks` variable to save the file with a custom name.

## Example Output

**Week 1 Day 1**:
| Exercise | Sets | Reps | Weight (lbs/kg) | Notes |
|---------------------|------|------|-----------------|-----------------------------|
| Snatch | 5 | 3 | 70% | Focus on form and speed |
| Clean & Jerk | 4 | 2 | 75% | Powerful extension |
| Push Press | 4 | 5 | | |
| Snatch-Grip Deadlift| 4 | 3 | | |

## Support

If you encounter any issues, feel free to reach out or adjust the script to match your specific requirements.

Happy lifting!
