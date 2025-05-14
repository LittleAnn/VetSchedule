import pandas as pd
import random
import os
import subprocess
import calendar
import openpyxl  # type: ignore
from openpyxl import load_workbook  # type: ignore
from openpyxl.styles import PatternFill  # type: ignore
import time

# --- Define scheduling function ---
def generate_schedule(file_path, save_path, year, month):
    random.seed(time.time())  # Ensure different output every time

    xls = pd.ExcelFile(file_path)
    data = pd.read_excel(xls, sheet_name='Dyspozycje')
    limits = pd.read_excel(xls, sheet_name='Limit zmian')
    preferences = pd.read_excel(xls, sheet_name='Preferencje zmian')
    fixed_df = pd.read_excel(xls, sheet_name='Ustalone zmiany')
    try:
        vacations = pd.read_excel(xls, sheet_name='Urlopy')
        vacation_days = {(row['Pracownik'], int(row['Dzień'])) for idx, row in vacations.iterrows()}
    except:
        vacation_days = set()

    availability_marker = 1

    year = int(year)
    month = int(month)
    cal = calendar.Calendar()
    weekends = [day for day, weekday in cal.itermonthdays2(year, month) if day != 0 and weekday >= 5]

    shift_limits = {}
    for _, row in limits.iterrows():
        employee = row['Pracownik']
        shift_limits[employee] = {
            'day': row['Dzień'],
            'night': row['Noc'],
            'weekend': row['Weekend']
        }

    shift_preferences = {}
    for _, row in preferences.iterrows():
        employee = row['Pracownik']
        shift_preferences[employee] = {
            'day': row['Dzień'],
            'night': row['Noc'],
            'weekend': row['Weekend']
        }

    assigned_shifts = {
        emp: {'day': 0, 'night': 0, 'weekend': 0} for emp in shift_limits
    }
    last_shift = {emp: None for emp in shift_limits}

    # --- Weekly counters per shift type ---
    weekly_shift_counts = {
        emp: {
            week: {'day': 0, 'night': 0, 'weekend': 0} for week in range(1, 6)
        } for emp in shift_limits
    }

    assignments = []
    max_day = data.shape[0]
    all_employees = list(shift_limits.keys())
    shift_matrix = pd.DataFrame(index=range(1, max_day + 1), columns=all_employees).fillna('free')

    fixed_assignments = {}
    for _, row in fixed_df.iterrows():
        day = int(row['Dzień'])
        shift_type = row['Typ zmiany']
        emp = row['Pracownik']
        if day not in fixed_assignments:
            fixed_assignments[day] = {}
        fixed_assignments[day][shift_type] = emp
        if emp in assigned_shifts:
            assigned_shifts[emp][shift_type] += 1
            last_shift[emp] = shift_type
            week = (day - 1) // 7 + 1
            weekly_shift_counts[emp][week][shift_type] += 1
            shift_matrix.at[day, emp] = shift_type

    for idx, row in data.iterrows():
        day = idx + 1
        week = (day - 1) // 7 + 1
        available_employees = [col for col in data.columns[1:] if row[col] == availability_marker]

        shift_type_day = 'weekend' if day in weekends else 'day'
        shift_type_night = 'night'

        for shift_type, max_assign in [(shift_type_day, 2), (shift_type_night, 1)]:
            if day in fixed_assignments and shift_type in fixed_assignments[day]:
                continue

            eligible_emps = [
                emp for emp in available_employees
                if emp in assigned_shifts
                and assigned_shifts[emp][shift_type] < shift_limits[emp][shift_type]
                and shift_preferences[emp][shift_type] == 1
                and not (last_shift[emp] == 'night' and shift_type in ['day', 'weekend'])
                and (emp, day) not in vacation_days
                and weekly_shift_counts[emp][week][shift_type] < (2 if shift_type == 'night' else 4)
            ]

            def fairness_score(emp):
                return (
                    assigned_shifts[emp][shift_type],
                    sum(assigned_shifts[emp].values()),
                    random.random()
                )

            eligible_sorted = sorted(eligible_emps, key=fairness_score)
            assigned_today = eligible_sorted[:max_assign]

            for emp in assigned_today:
                assigned_shifts[emp][shift_type] += 1
                last_shift[emp] = shift_type
                weekly_shift_counts[emp][week][shift_type] += 1
                shift_matrix.at[day, emp] = shift_type

    label_translation = {'day': 'Dzień', 'night': 'Noc', 'weekend': 'Dzień', 'free': 'Wolne'}
    shift_matrix = shift_matrix.applymap(lambda x: label_translation.get(x, x))

    summary_data = []
    for emp, shifts in assigned_shifts.items():
        summary_data.append({
            'Pracownik': emp,
            'Dniowe zmiany': shifts['day'],
            'Nocne zmiany': shifts['night'],
            'Weekendowe zmiany': shifts['weekend'],
            'Suma zmian': sum(shifts.values())
        })
    summary_df = pd.DataFrame(summary_data)

    with pd.ExcelWriter(save_path, engine='openpyxl') as writer:
        shift_matrix.to_excel(writer, sheet_name='Grafik', index_label='Dzień')
        summary_df.to_excel(writer, sheet_name='Podsumowanie', index=False)

    wb = load_workbook(save_path)
    ws = wb['Grafik']

    fill_day = PatternFill(start_color="66B902", end_color="66B902", fill_type="solid")
    fill_night = PatternFill(start_color="2875D7", end_color="2875D7", fill_type="solid")
    fill_free = PatternFill(start_color="E8E8A5", end_color="E8E8A5", fill_type="solid")
    fill_weekend = PatternFill(start_color="C112CD", end_color="C112CD", fill_type="solid")

    for row in ws.iter_rows(min_row=2, min_col=2):
        day_cell = row[0].row
        is_weekend = (day_cell - 1) in weekends
        for cell in row:
            if cell.value == 'Dzień':
                cell.fill = fill_day
            elif cell.value == 'Noc':
                cell.fill = fill_night
            elif cell.value == 'Wolne':
                cell.fill = fill_weekend if is_weekend else fill_free

    for column_cells in ws.columns:
        max_length = 0
        column = column_cells[0].column_letter
        for cell in column_cells:
            if cell.value:
                max_length = max(max_length, len(str(cell.value)))
        ws.column_dimensions[column].width = max_length + 2

    wb.save(save_path)
