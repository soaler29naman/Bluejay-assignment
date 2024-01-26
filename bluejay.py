import pandas as pd
from datetime import datetime, timedelta
import sys

# Function to load the Excel file and perform analysis
def analyze_excel_file(file_path, consecutive_days_threshold=7):
    try:
        # Read the Excel sheet into a DataFrame
        df = pd.read_excel(file_path)

        # Trim column names to remove leading/trailing spaces
        df.columns = df.columns.str.strip()

        # Initialize sets to keep track of printed employees
        consecutive_printed = set()
        short_break_printed = set()
        long_shift_printed = set()

        # Redirect output to a file
        sys.stdout = open('output.txt', 'w')

        print("Name of Employee and their Position Id, who has worked for 7 consecutive days ")

        for index, row in df.iterrows():
            employee_name = row['Employee Name']
            position_id = row['Position ID']

            if employee_name in consecutive_printed:
                continue

            # Check for consecutive days worked
            if index > 0 and employee_name == df.at[index - 1, 'Employee Name']:
                consecutive_days = 1
                for i in range(index - 1, -1, -1):
                    if df.at[i, 'Employee Name'] == employee_name:
                        consecutive_days += 1
                    else:
                        break
                if consecutive_days >= consecutive_days_threshold:
                    print(f"Employee: {employee_name}, Position: {position_id}")
                    consecutive_printed.add(employee_name)
        print("""
        
        """)
        print(
            "Name of Employee and their Position Id, who have less than 10 hours of time between shifts but greater than 1 hour")

        employee_breaks = {}  # Dictionary to track breaks for each employee
        short_break_printed = set()

        for index, row in df.iterrows():
            employee_name = row['Employee Name']
            position_id = row['Position ID']

            if employee_name in short_break_printed:
                continue

            if employee_name in employee_breaks:
                last_time_out = employee_breaks[employee_name]


                time_in = row['Time']
                if pd.notna(last_time_out) and pd.notna(time_in):
                    last_time_out = last_time_out.strftime('%m/%d/%Y %I:%M %p')
                    time_in = time_in.strftime('%m/%d/%Y %I:%M %p')
                else:
                    continue

                if isinstance(time_in, str) and isinstance(last_time_out, str):
                    time_in = datetime.strptime(time_in, '%m/%d/%Y %I:%M %p')
                    last_time_out = datetime.strptime(last_time_out, '%m/%d/%Y %I:%M %p')

                    time_diff = (time_in - last_time_out).total_seconds() / 3600
                    if 1 < time_diff < 10:
                        print(f"Employee: {employee_name}, Position: {position_id}")
                        short_break_printed.add(employee_name)
                else:
                    time_in = None

            employee_breaks[employee_name] = row['Time Out']

        print("""

                """)

        print("Name and Position Id of Employee who has worked for more than 14 hour in a single shift")


        long_break_printed = set()

        for index, row in df.iterrows():
            employee_name = row['Employee Name']
            position_id = row['Position ID']

            if employee_name in long_break_printed:
                continue

            if employee_name in employee_breaks:
                last_time_out = row['Time Out']

                time_in = row['Time']
                if pd.notna(last_time_out) and pd.notna(time_in):
                    last_time_out = last_time_out.strftime('%m/%d/%Y %I:%M %p')
                    time_in = time_in.strftime('%m/%d/%Y %I:%M %p')
                else:
                    continue

                if isinstance(time_in, str) and isinstance(last_time_out, str):
                    time_in = datetime.strptime(time_in, '%m/%d/%Y %I:%M %p')
                    last_time_out = datetime.strptime(last_time_out, '%m/%d/%Y %I:%M %p')

                    time_diff = (last_time_out - time_in).total_seconds() / 3600
                    if time_diff>14:
                        print(f"Employee: {employee_name}, Position: {position_id}")
                        long_break_printed.add(employee_name)


    except FileNotFoundError:
        print(f"File not found: {file_path}")
    except Exception as e:
        print(f"An error occurred: {str(e)}")
    finally:
        # Close the file and restore stdout
        sys.stdout.close()
        sys.stdout = sys.__stdout__

if __name__ == "__main__":
    file_path = 'Assignment_Timecard.xlsx'
    analyze_excel_file(file_path, consecutive_days_threshold=7)
