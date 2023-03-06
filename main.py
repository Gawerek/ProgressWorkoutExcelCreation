import pandas as pd
import subprocess
from datetime import date, timedelta

# Install xlsxwriter module
subprocess.check_call(["python", "-m", "pip", "install", "xlsxwriter"])

# Define the exercises and their sets and reps for each variant
variant_a = {
    'Squats': '5x5',
    'Bench press': '5x5',
    'Rowing': '5x5',
    'Side raises': '3x8',
    'Biceps + triceps': '3x8',
    'Allahs': '3x8',
    'Calves': '3x8'
}

variant_b = {
    'Classic deadlift': '5x5',
    'Soldier press': '5x5',
    'Narrow pull-ups': '5x5',
    'Biceps': '3x8',
    'Plank': 'x3',
    'Calves': '3x8'
}

# Create a Pandas DataFrame with the exercise data for each variant
variant_a_df = pd.DataFrame({'Exercise': list(variant_a.keys()), 'Sets and Reps': list(variant_a.values())})
variant_b_df = pd.DataFrame({'Exercise': list(variant_b.keys()), 'Sets and Reps': list(variant_b.values())})

# Create a new Excel file and write the DataFrame to it
writer = pd.ExcelWriter('training_schedule.xlsx', engine='xlsxwriter')
variant_a_df.to_excel(writer, sheet_name='Variant A', index=False)
variant_b_df.to_excel(writer, sheet_name='Variant B', index=False)

# Add a sheet for tracking progress for each variant
start_date = date(2023, 3, 1)
num_weeks = 4
variant_a_tracking = pd.DataFrame(columns=['Date', 'Squats (kg)', 'Bench press (kg)', 'Rowing (kg)', 'Side raises (kg)',
                                           'Biceps + triceps (kg)', 'Allahs (kg)', 'Calves (kg)'])
variant_b_tracking = pd.DataFrame(
    columns=['Date', 'Classic deadlift (kg)', 'Soldier press (kg)', 'Narrow pull-ups (kg)',
             'Biceps (kg)', 'Plank', 'Calves (kg)'])

for i in range(num_weeks * 2):
    curr_date = start_date + timedelta(days=7 * i)
    variant_a_tracking.loc[i] = [curr_date] + [0] * 7
    variant_b_tracking.loc[i] = [curr_date] + [0] * 6

variant_a_tracking.to_excel(writer, sheet_name='Variant A Progress', index=False)
variant_b_tracking.to_excel(writer, sheet_name='Variant B Progress', index=False)

# Save and close the Excel file
writer.save()

# Read the progress tracker sheets into Pandas DataFrames
variant_a_progress = pd.read_excel('training_schedule.xlsx', sheet_name='Variant A Progress')
variant_b_progress = pd.read_excel('training_schedule.xlsx', sheet_name='Variant B Progress')

# Create a pivot table that summarizes the progress by exercise for each variant
variant_a_pivot = pd.pivot_table(variant_a_progress, index='Date', aggfunc='sum')
variant_b_pivot = pd.pivot_table(variant_b_progress, index='Date', aggfunc='sum')

print(variant_a_pivot)
print(variant_b_pivot)
