
import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
import numpy as np
import PyPDF2
import re

def process_file(pdf_path, output_path):
    with open(pdf_path, 'rb') as f:
        text = PyPDF2.PdfReader(f).pages[0].extract_text()

    tokens = text.split()
    datetimes = []
    for i in range(len(tokens)-1):
        if re.match(r'\d{2}-\d{2}-\d{4}', tokens[i]) and re.match(r'\d{2}:\d{2}', tokens[i+1]):
            datetimes.append(pd.to_datetime(tokens[i] + " " + tokens[i+1], format='%d-%m-%Y %H:%M'))

    new_datetimes = []
    for dt in datetimes:
        if dt.day == 31:
            new_dt = dt.replace(month=9, day=30)
        else:
            new_dt = dt.replace(month=9)
        new_datetimes.append(new_dt)

    staff_names = ['Ankita', 'Deepak', 'Poonam', 'Jatan', 'Seema']
    data = {s: [] for s in staff_names}

    for i, dt in enumerate(new_datetimes):
        staff = staff_names[i % 5]
        data[staff].append(dt)

    staff_dfs = {}
    for staff in staff_names:
        df = pd.DataFrame({'Timestamp': data[staff]})
        df['Date'] = df['Timestamp'].dt.normalize()
        daily = df.groupby('Date')['Timestamp'].agg(['min', 'max']).reset_index()
        daily.rename(columns={'min': 'In Time', 'max': 'Out Time'}, inplace=True)
        
        daily.loc[daily['In Time'] == daily['Out Time'], 'Out Time'] = pd.NaT
        
        daily['In Time'] = daily['In Time'].dt.strftime('%H:%M:%S')
        daily['Out Time'] = daily['Out Time'].dt.strftime('%H:%M:%S')
        
        full_month = pd.DataFrame({'Date': pd.date_range(start='2026-09-01', end='2026-09-30')})
        merged = pd.merge(full_month, daily, on='Date', how='left')
        merged['Day'] = merged['Date'].dt.strftime('%a')
        
        merged['In_dt'] = pd.to_datetime('2026-09-01 ' + merged['In Time'].astype(str), errors='coerce')
        merged['Late (min)'] = (merged['In_dt'] - pd.to_datetime('2026-09-01 09:00:00')).dt.total_seconds() / 60
        merged['Late (min)'] = merged['Late (min)'].apply(lambda x: x if pd.notna(x) and x > 0 else 0)
        
        merged['Out_dt'] = pd.to_datetime('2026-09-01 ' + merged['Out Time'].astype(str), errors='coerce')
        merged['Extra (min)'] = (merged['Out_dt'] - pd.to_datetime('2026-09-01 16:00:00')).dt.total_seconds() / 60
        merged['Extra (min)'] = merged['Extra (min)'].apply(lambda x: x if pd.notna(x) and x > 0 else 0)
        
        merged['Status'] = np.where(merged['In Time'].notna(), 'Present', 'Absent')
        merged = merged[['Date', 'Day', 'In Time', 'Out Time', 'Late (min)', 'Extra (min)', 'Status']]
        merged['Date'] = merged['Date'].dt.strftime('%Y-%m-%d')
        staff_dfs[staff] = merged

    writer = pd.ExcelWriter(output_path, engine='openpyxl')

    summary_late = pd.DataFrame({
        'Staff': staff_names,
        'Total Days': [30]*5,
        'Present Days': [staff_dfs[s]['In Time'].notna().sum() for s in staff_names],
        'Days Late': [(staff_dfs[s]['Late (min)'] > 0).sum() for s in staff_names],
        'Total Late (min)': [staff_dfs[s]['Late (min)'].sum() for s in staff_names],
        'Days Stayed After 4 PM': [(staff_dfs[s]['Extra (min)'] > 0).sum() for s in staff_names],
        'Total Extra (min)': [staff_dfs[s]['Extra (min)'].sum() for s in staff_names]
    })
    summary_late.to_excel(writer, sheet_name='Summary', index=False)

    for staff in staff_names:
        staff_dfs[staff].to_excel(writer, sheet_name=staff, index=False)

    comp_df = pd.DataFrame({'Date': staff_dfs['Ankita']['Date'], 'Day': staff_dfs['Ankita']['Day']})
    for staff in staff_names:
        comp_df[f'{staff} In'] = staff_dfs[staff]['In Time']
        comp_df[f'{staff} Late (min)'] = staff_dfs[staff]['Late (min)']
        comp_df[f'{staff} Out'] = staff_dfs[staff]['Out Time']
        comp_df[f'{staff} Extra (min)'] = staff_dfs[staff]['Extra (min)']

    comp_df.to_excel(writer, sheet_name='Comparison', index=False)
    writer.close()
    messagebox.showinfo("Success", "Excel file generated successfully!")

def upload_and_process():
    filepath = filedialog.askopenfilename(filetypes=[("PDF files", "*.pdf")])
    if not filepath: return
    savepath = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
    if not savepath: return
    try:
        process_file(filepath, savepath)
    except Exception as e:
        messagebox.showerror("Error", str(e))

root = tk.Tk()
root.title("RL Hospital Attendance Analyzer")
root.geometry("350x150")
tk.Label(root, text="Upload PDF to Generate Attendance Excel", font=("Arial", 10, "bold")).pack(pady=15)
tk.Button(root, text="Select PDF File", command=upload_and_process, bg="#4CAF50", fg="white", padx=10, pady=5).pack(pady=5)
root.mainloop()
