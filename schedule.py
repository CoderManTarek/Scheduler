import tkinter as tk
from tkinter import ttk, filedialog
import pandas as pd
import pulp
from datetime import timedelta, datetime
from tkcalendar import DateEntry


def daterange(date1, date2):
    for n in range(int((date2 - date1).days) + 1):
        yield date1 + timedelta(n)

def schedule(backlog, time_mapping, start_date, end_date):

    # Create a list of all weekdays between the start and end date
    date_list = [day for day in daterange(start_date, end_date) if day.weekday() < 5]

    # Merge the two dataframes
    merged = backlog.merge(time_mapping, left_on='Procedure Type', right_on='ProcedureType')

    # Constants
    num_seats = 10
    hours_per_day = 9

    # Create a Linear Programming Problem
    prob = pulp.LpProblem("SchedulingProblem", pulp.LpMaximize)

    # Create binary decision variables
    x = pulp.LpVariable.dicts("Choice", ((i, j, d, h) for i in range(len(merged))
                                         for j in range(num_seats)
                                         for d in range(len(date_list))
                                         for h in range(hours_per_day)), 0, 1, pulp.LpBinary)

    # Objective: Maximize the number of scheduled procedures
    prob += pulp.lpSum(x[i, j, d, h] for i in range(len(merged))
                       for j in range(num_seats)
                       for d in range(len(date_list))
                       for h in range(hours_per_day))

    # Constraints

    # Ensure no overlapping times for the same procedure
    for i in range(len(merged)):
        for j in range(num_seats):
            for d in range(len(date_list)):
                for h in range(hours_per_day - int(merged.iloc[i]['TurnAroundTime']) + 1):
                    prob += pulp.lpSum(x[i, j, d, k] for k in range(h, h + int(merged.iloc[i]['TurnAroundTime']))) <= 1

    # Ensure each procedure is scheduled only once
    for i in range(len(merged)):
        prob += pulp.lpSum(x[i, j, d, h] for j in range(num_seats)
                           for d in range(len(date_list))
                           for h in range(hours_per_day)) <= 1

    # Ensure a seat has only one procedure at a time
    for j in range(num_seats):
        for d in range(len(date_list)):
            for h in range(hours_per_day):
                prob += pulp.lpSum(x[i, j, d, h] for i in range(len(merged))) <= 1

    # Ensure maximum patients are in seats at all times
    for d in range(len(date_list)):
        for h in range(hours_per_day):
            prob += pulp.lpSum(x[i, j, d, h] for i in range(len(merged)) for j in range(num_seats)) >= min(10, len(merged))

    # Solve the problem
    prob.solve()

    # Extract results
    output = []
    for i in range(len(merged)):
        for j in range(num_seats):
            for d in range(len(date_list)):
                for h in range(hours_per_day):
                    if x[i, j, d, h].varValue == 1:
                        start_time = datetime.combine(date_list[d], datetime.min.time()) + timedelta(hours=8 + h)
                        end_time = start_time + timedelta(hours=int(merged.iloc[i]['TurnAroundTime']))
                        output.append([j,
                                       merged.iloc[i]['Patient Name'],
                                       merged.iloc[i]['Procedure Type'],
                                       merged.iloc[i]['TurnAroundTime'],
                                       start_time,
                                       f"{start_time.strftime('%Y-%m-%d %H:%M')} - {end_time.strftime('%H:%M')}"])

    # Create a DataFrame from the results
    output_df = pd.DataFrame(output, columns=["Seat Number", "Patient Name", "Procedure Type", "TurnAroundTime", "DateTime", "Scheduled Time"])

    # Sort the DataFrame by 'DateTime' and 'Seat Number'
    output_df = output_df.sort_values(by=['DateTime', 'Seat Number'])

    return output_df.drop(columns=['DateTime'])

# GUI Functions
def choose_file_procedures():
    file_path = filedialog.askopenfilename()
    global procedures_file
    procedures_file = file_path
    lbl_procedures_file["text"] = file_path.split("/")[-1]

def choose_file_turnaround():
    file_path = filedialog.askopenfilename()
    global turnaround_file
    turnaround_file = file_path
    lbl_turnaround_file["text"] = file_path.split("/")[-1]

def schedule_procedures():
    # Read the CSV files
    backlog = pd.read_csv(procedures_file)
    time_mapping = pd.read_csv(turnaround_file)

    # Extract start and end date from date pickers
    start_date = start_date_entry.get_date()
    end_date = end_date_entry.get_date()

    # Call the schedule function
    result = schedule(backlog, time_mapping, start_date, end_date)

    # Clear existing items in the treeview
    for row in tree.get_children():
        tree.delete(row)

    # Insert the resulting dataframe into the treeview
    for index, row in result.iterrows():
        tree.insert("", "end", values=(row["Seat Number"], row["Patient Name"], row["Procedure Type"], row["TurnAroundTime"], row["Scheduled Time"]))

def save_csv():
    file_path = filedialog.asksaveasfilename(defaultextension=".csv", filetypes=[("CSV files", "*.csv")])
    if file_path:
        # Convert treeview content to a dataframe
        items = tree.get_children()
        rows = [tree.item(item)["values"] for item in items]
        df = pd.DataFrame(rows, columns=["Seat Number", "Patient Name", "Procedure Type", "TurnAroundTime", "Scheduled Time"])
        df.to_csv(file_path, index=False)

# Initialize the main window
app = tk.Tk()
app.title("Procedure Scheduler")
app.geometry("800x600")
app.configure(background="#2E2E2E")

# Styling
style = ttk.Style()
style.theme_use("default")
style.configure("TButton", background="#383838", foreground="white", bordercolor="#383838", highlightthickness=0, bd=0)
style.map('TButton', background=[('pressed', '!disabled', 'black'), ('active', 'blue')])
style.configure("TLabel", background="#2E2E2E", foreground="white")
style.configure("TDateEntry", background="black", foreground="white")

# Add buttons and labels for file selection
btn_procedures = ttk.Button(app, text="Choose Patient Procedures File", command=choose_file_procedures)
btn_procedures.pack(pady=20)
lbl_procedures_file = ttk.Label(app, text="No file chosen for Patient Procedures")
lbl_procedures_file.pack(pady=10)

btn_turnaround = ttk.Button(app, text="Choose Procedure Turnaround Times File", command=choose_file_turnaround)
btn_turnaround.pack(pady=20)
lbl_turnaround_file = ttk.Label(app, text="No file chosen for Procedure Turnaround Times")
lbl_turnaround_file.pack(pady=10)

# Add Start Date Picker
lbl_start_date = ttk.Label(app, text="Start Date:")
lbl_start_date.pack(pady=10)
start_date_entry = DateEntry(app, width=12, background='black', foreground='white', borderwidth=2, selectbackground='white', selectforeground='black')
start_date_entry.pack(pady=10)

# Add End Date Picker
lbl_end_date = ttk.Label(app, text="End Date:")
lbl_end_date.pack(pady=10)
end_date_entry = DateEntry(app, width=12, background='black', foreground='white', borderwidth=2, selectbackground='white', selectforeground='black')
end_date_entry.pack(pady=10)

# Add a Schedule button
btn_schedule = ttk.Button(app, text="Schedule Procedures", command=schedule_procedures)
btn_schedule.pack(pady=20)

# Create a treeview for displaying the schedule
tree = ttk.Treeview(app, columns=("Seat Number", "Patient Name", "Procedure Type", "TurnAroundTime", "Scheduled Time"), show='headings')  # Hide default column by setting show to 'headings'
tree.heading("Seat Number", text="Seat Number")
tree.heading("Patient Name", text="Patient Name")
tree.heading("Procedure Type", text="Procedure Type")
tree.heading("TurnAroundTime", text="TurnAroundTime")
tree.heading("Scheduled Time", text="Scheduled Time")
tree.pack(pady=20, fill=tk.BOTH, expand=True)

# Add a Save to CSV button
btn_save = ttk.Button(app, text="Save Schedule to CSV", command=save_csv)
btn_save.pack(pady=20)

app.mainloop()
