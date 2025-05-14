import pandas as pd
import random
import os
import subprocess
import calendar
import openpyxl  # type: ignore
from openpyxl import load_workbook  # type: ignore
from openpyxl.styles import PatternFill  # type: ignore

# --- Define scheduling function ---
def generate_schedule(file_path, save_path, year, month):
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

    weekly_shifts = {
        emp: {week: 0 for week in range(1, 6)} for emp in shift_limits
    }

    assignments = []

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
            weekly_shifts[emp][week] += 1

    for idx, row in data.iterrows():
        day = idx + 1
        week = (day - 1) // 7 + 1
        available_employees = [col for col in data.columns[1:] if row[col] == availability_marker]

        shift_type_day = 'weekend' if day in weekends else 'day'
        if day in fixed_assignments and shift_type_day in fixed_assignments[day]:
            emp = fixed_assignments[day][shift_type_day]
            assignments.append({'Dzień': day, 'Typ zmiany': shift_type_day, 'Pracownik': emp})
        else:
            eligible_day = [
                emp for emp in available_employees
                if emp in assigned_shifts
                and assigned_shifts[emp][shift_type_day] < shift_limits[emp][shift_type_day]
                and shift_preferences[emp][shift_type_day] == 1
                and (last_shift[emp] != 'night')
                and (emp, day) not in vacation_days
                and weekly_shifts[emp][week] < 4
            ]
            random.shuffle(eligible_day)
            eligible_day_sorted = sorted(eligible_day, key=lambda emp: sum(assigned_shifts[emp].values()))
            num_to_assign_day = min(2, len(eligible_day_sorted))
            assigned_day = eligible_day_sorted[:num_to_assign_day]
            for emp in assigned_day:
                assigned_shifts[emp][shift_type_day] += 1
                last_shift[emp] = shift_type_day
                weekly_shifts[emp][week] += 1
                assignments.append({'Dzień': day, 'Typ zmiany': shift_type_day, 'Pracownik': emp})

        shift_type_night = 'night'
        if day in fixed_assignments and shift_type_night in fixed_assignments[day]:
            emp = fixed_assignments[day][shift_type_night]
            assignments.append({'Dzień': day, 'Typ zmiany': shift_type_night, 'Pracownik': emp})
        else:
            eligible_night = [
                emp for emp in available_employees
                if emp in assigned_shifts
                and assigned_shifts[emp][shift_type_night] < shift_limits[emp][shift_type_night]
                and shift_preferences[emp][shift_type_night] == 1
                and (emp, day) not in vacation_days
                and weekly_shifts[emp][week] < 4
            ]
            random.shuffle(eligible_night)
            eligible_night_sorted = sorted(eligible_night, key=lambda emp: sum(assigned_shifts[emp].values()))
            num_to_assign_night = min(1, len(eligible_night_sorted))
            assigned_night = eligible_night_sorted[:num_to_assign_night]
            for emp in assigned_night:
                assigned_shifts[emp][shift_type_night] += 1
                last_shift[emp] = shift_type_night
                weekly_shifts[emp][week] += 1
                assignments.append({'Dzień': day, 'Typ zmiany': shift_type_night, 'Pracownik': emp})

    all_employees = list(shift_limits.keys())
    max_day = data.shape[0]
    shift_matrix = pd.DataFrame(index=range(1, max_day + 1), columns=all_employees).fillna('free')

    for assign in assignments:
        day = assign['Dzień']
        shift_type = assign['Typ zmiany']
        emp = assign['Pracownik']
        if emp and emp != 'Brak dostępnych pracowników':
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
