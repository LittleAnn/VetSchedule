import pandas as pd
import random
import os
import subprocess
import calendar
import openpyxl # type: ignore
from openpyxl import load_workbook # type: ignore
from openpyxl.styles import PatternFill # type: ignore

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

    # Availability marker
    availability_marker = 1

    # Calculate weekends automatically
    cal = calendar.Calendar()
    weekends = [day for day, weekday in cal.itermonthdays2(year, month) if day != 0 and weekday >= 5]

    shift_limits = {}
    for idx, row in limits.iterrows():
        employee = row['Pracownik']
        shift_limits[employee] = {
            'day': row['Dzień'],
            'night': row['Noc'],
            'weekend': row['Weekend']
        }

    shift_preferences = {}
    for idx, row in preferences.iterrows():
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

    assignments = []

    fixed_shifts = []
    for idx, row in fixed_df.iterrows():
        fixed_shifts.append({
            'Dzień': int(row['Dzień']),
            'Typ zmiany': row['Typ zmiany'],
            'Pracownik': row['Pracownik']
        })

    fixed_assignments = {}
    for entry in fixed_shifts:
        day = entry['Dzień']
        shift_type = entry['Typ zmiany']
        emp = entry['Pracownik']
        if day not in fixed_assignments:
            fixed_assignments[day] = {}
        fixed_assignments[day][shift_type] = emp
        if emp in assigned_shifts:
            assigned_shifts[emp][shift_type] += 1
            last_shift[emp] = shift_type

    for idx, row in data.iterrows():
        day = idx + 1
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
            ]

            if eligible_day:
                eligible_day_sorted = sorted(
                    eligible_day,
                    key=lambda emp: sum(assigned_shifts[emp].values())
                )
                num_to_assign_day = min(2, len(eligible_day_sorted))
                assigned_day = eligible_day_sorted[:num_to_assign_day]
                for emp in assigned_day:
                    assigned_shifts[emp][shift_type_day] += 1
                    last_shift[emp] = 'day'
                for emp in assigned_day:
                    assignments.append({'Dzień': day, 'Typ zmiany': shift_type_day, 'Pracownik': emp})
            else:
                assignments.append({'Dzień': day, 'Typ zmiany': shift_type_day, 'Pracownik': 'Brak dostępnych pracowników'})

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
            ]

            if eligible_night:
                eligible_night_sorted = sorted(
                    eligible_night,
                    key=lambda emp: sum(assigned_shifts[emp].values())
                )
                num_to_assign_night = min(1, len(eligible_night_sorted))
                assigned_night = eligible_night_sorted[:num_to_assign_night]
                for emp in assigned_night:
                    assigned_shifts[emp][shift_type_night] += 1
                    last_shift[emp] = 'night'
                for emp in assigned_night:
                    assignments.append({'Dzień': day, 'Typ zmiany': shift_type_night, 'Pracownik': emp})
            else:
                assignments.append({'Dzień': day, 'Typ zmiany': shift_type_night, 'Pracownik': 'Brak dostępnych pracowników'})

    # --- Create Shift Assignment Matrix ---
    all_employees = list(shift_limits.keys())
    max_day = data.shape[0]
    shift_matrix = pd.DataFrame(index=range(1, max_day + 1), columns=all_employees)

    # Initialize with 'free'
    shift_matrix = shift_matrix.fillna('free')

    for assign in assignments:
        day = assign['Dzień']
        shift_type = assign['Typ zmiany']
        emp = assign['Pracownik']

        if emp and emp != 'Brak dostępnych pracowników':
            if ',' in emp:
                for e in emp.split(', '):
                    shift_matrix.at[day, e] = 'day'
            else:
                if shift_type == 'night':
                    shift_matrix.at[day, emp] = 'night'
                else:
                    shift_matrix.at[day, emp] = 'day'

    # --- Translate labels to Polish ---
    label_translation = {
        'day': 'Dzień',
        'night': 'Noc',
        'free': 'Wolne'
    }
    shift_matrix = shift_matrix.applymap(lambda x: label_translation.get(x, x))

    # --- Create Shift Assignment Summary ---
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

    # --- Save to Excel ---
    with pd.ExcelWriter(save_path, engine='openpyxl') as writer:
        shift_matrix.to_excel(writer, sheet_name='Grafik', index_label='Dzień')
        summary_df.to_excel(writer, sheet_name='Podsumowanie', index=False)

    # --- Apply colors ---
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
                cell.fill = fill_free
            if is_weekend:
                cell.fill = fill_weekend

    # --- Auto-adjust column widths ---
    for column_cells in ws.columns:
        max_length = 0
        column = column_cells[0].column_letter
        for cell in column_cells:
            try:
                if cell.value:
                    max_length = max(max_length, len(str(cell.value)))
            except:
                pass
        adjusted_width = max_length + 2
        ws.column_dimensions[column].width = adjusted_width

    wb.save(save_path)