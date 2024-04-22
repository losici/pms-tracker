from datetime import datetime, timedelta

import numpy as np
import pandas as pd
import xlsxwriter

from utilities import input_in_range

MIN_VALUE_SCALE = 0
MAX_VALUE_SCALE = 5

class PMSData:
    """Class to handle data for PMS tracking with user input."""
    def __init__(self):
        self.entries = []

    def add_entry(self):
        """Method for user to input data interactively."""
        while True:
            date_str = input("Enter the date for the entry (YYYY-MM-DD): ")
            try:
                date = datetime.strptime(date_str, '%Y-%m-%d')
                break  # Exit loop if date is valid
            except ValueError:
                print("Invalid date format! Please enter the date in YYYY-MM-DD format.")
        day_of_cycle = int(input("Enter the day of the cycle: "))
        mood_swings = input_in_range("Rate your mood swings today (0-5): ", MIN_VALUE_SCALE, MAX_VALUE_SCALE)
        cramps = input_in_range("Rate your cramps today (0-5): ", MIN_VALUE_SCALE, MAX_VALUE_SCALE)
        bloating = input_in_range("Rate your bloating today (0-5): ", MIN_VALUE_SCALE, MAX_VALUE_SCALE)
        headaches = input_in_range("Rate your headaches today (0-5):" , MIN_VALUE_SCALE, MAX_VALUE_SCALE)
        fatigue = input_in_range("Rate your fatigue today (0-5): ", MIN_VALUE_SCALE, MAX_VALUE_SCALE)
        other_symptoms = input("List any other symptoms you experienced: ")
        stress_level = input_in_range("Rate your stress level today (0-5): ", MIN_VALUE_SCALE, MAX_VALUE_SCALE)
        notes = input("Any additional notes? ")

        if not other_symptoms:
            other_symptoms = 0
        
        if not notes:
            notes = 0

        entry = {
            "Date": date,
            "Day of Cycle": day_of_cycle,
            "Mood Swings": mood_swings,
            "Cramps": cramps,
            "Bloating": bloating,
            "Headaches": headaches,
            "Fatigue": fatigue,
            "Stress Level": stress_level,
            "Other Symptoms": other_symptoms,
            "Notes": notes
        }
        self.entries.append(entry)
        print("Entry added successfully!")

    def get_data_frame(self):
        """Convert entries to a pandas DataFrame."""
        return pd.DataFrame(self.entries)

class ExcelReporter:
    """Class to handle Excel reporting for PMS data."""
    def __init__(self, data_frame):
        self.data_frame = data_frame

    def save_to_excel(self, file_path='PMS_Tracking.xlsx'):
        """Save data to an Excel file with analysis and formatting."""
        with pd.ExcelWriter(file_path, engine='xlsxwriter') as writer:
            self.data_frame.to_excel(writer, sheet_name='Daily Log', index=False)
            workbook = writer.book
            worksheet = writer.sheets['Daily Log']

            # Set the format for the date column
            date_format = workbook.add_format({'num_format': 'yyyy-mm-dd'})
            worksheet.set_column('A:A', 15, date_format)

            # Optionally, add charts
            self._add_chart(workbook, worksheet)

    def _add_chart(self, workbook, worksheet):
        """Add a chart in the Excel file."""
        chart = workbook.add_chart({'type': 'line'})
        chart.add_series({
            'name': 'Mood Swings',
            'categories': '=Daily Log!$A$2:$A$91',
            'values': '=Daily Log!$C$2:$C$91',
        })
        chart.set_title({'name': 'Mood Swings Over Time'})
        chart.set_x_axis({'name': 'Date'})
        chart.set_y_axis({'name': 'Severity'})
        worksheet.insert_chart('K2', chart)

def main():
    pms_data = PMSData()
    # Example adding entries
    num_entries = int(input("How many days would you like to log? "))
    for _ in range(num_entries):
        pms_data.add_entry()

    data_frame = pms_data.get_data_frame()
    excel_reporter = ExcelReporter(data_frame)
    excel_reporter.save_to_excel()

if __name__ == "__main__":
    main()
