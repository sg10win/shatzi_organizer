import pandas as pd
import random

def create_test_input():
    # Generate names for Mashaks and Katzins
    mashaks_names = [f"מש״ק {i+1}" for i in range(40)]
    katzins_names = [f"קצין {i+1}" for i in range(15)]

    # Define the date range for two months
    date_range = pd.date_range("2025-03-01", "2025-04-30").strftime("%d/%m/%Y").tolist()

    # Generate Mashaks constraints (date-based)
    mashaks_constraints_data = {"Date": date_range}
    cant_work = {date: [] for date in date_range}
    for name in mashaks_names:
        # Assign 5-10 days they can't work for some, 0-3 for others
        num_days = random.choice(range(5, 11) if random.random() < 0.3 else range(0, 4))
        cant_work_days = random.sample(date_range, num_days)
        for day in cant_work_days:
            cant_work[day].append(name)
    mashaks_constraints_data["Can't Work"] = [
        ", ".join(cant_work[date]) for date in date_range
    ]

    # Generate Katzins constraints (date-based)
    katzins_constraints_data = {"Date": date_range}
    cant_work_katzins = {date: [] for date in date_range}
    for name in katzins_names:
        # Assign 5-10 days they can't work for some, 0-3 for others
        num_days = random.choice(range(5, 11) if random.random() < 0.3 else range(0, 4))
        cant_work_days = random.sample(date_range, num_days)
        for day in cant_work_days:
            cant_work_katzins[day].append(name)
    katzins_constraints_data["Can't Work"] = [
        ", ".join(cant_work_katzins[date]) for date in date_range
    ]

    # Create DataFrames
    potential_mashaks_df = pd.DataFrame(mashaks_names, columns=["Full Name"])
    potential_katzins_df = pd.DataFrame(katzins_names, columns=["Full Name"])
    mashaks_constraints_df = pd.DataFrame(mashaks_constraints_data)
    katzins_constraints_df = pd.DataFrame(katzins_constraints_data)

    # Save to Excel
    with pd.ExcelWriter('inputs.xlsx', engine='openpyxl') as writer:
        potential_mashaks_df.to_excel(writer, sheet_name='potential mashaks', index=False)
        potential_katzins_df.to_excel(writer, sheet_name='potential katzins', index=False)
        mashaks_constraints_df.to_excel(writer, sheet_name='mashaks constraints', index=False)
        katzins_constraints_df.to_excel(writer, sheet_name='katzins constraints', index=False)

# Call the function to create the input file
create_test_input()