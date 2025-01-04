import pandas as pd


def run_schedule_simulation(potential_mashaks, potential_katzins, mashaks_constraints, katzins_constraints):
    mashaks = potential_mashaks['Full Name'].tolist()
    katzins = potential_katzins['Full Name'].tolist()

    mashaks_unavailable = {
        pd.to_datetime(row['Date'], dayfirst=True).date(): str(row["Can't Work"]).split(', ') for _, row in
        mashaks_constraints.iterrows()
    }
    katzins_unavailable = {
        pd.to_datetime(row['Date'], dayfirst=True).date(): str(row["Can't Work"]).split(', ') for _, row in
        katzins_constraints.iterrows()
    }

    schedule = []
    mashak_assignments = {mashak: 0 for mashak in mashaks}
    katzin_assignments = {katzin: 0 for katzin in katzins}

    for date in mashaks_unavailable.keys():
        current_date = pd.to_datetime(date)
        weekday = current_date.weekday()
        unavailable_mashaks = mashaks_unavailable.get(date, [])
        unavailable_katzins = katzins_unavailable.get(date, [])
        available_mashaks = [m for m in mashaks if m not in unavailable_mashaks]
        available_katzins = [k for k in katzins if k not in unavailable_katzins]

        if weekday == 4:
            next_day = (current_date + pd.Timedelta(days=1)).date()
            if next_day in mashaks_unavailable.keys():
                unavailable_mashaks_next = mashaks_unavailable.get(next_day, [])
                unavailable_katzins_next = katzins_unavailable.get(next_day, [])
                available_mashaks = [m for m in available_mashaks if m not in unavailable_mashaks_next]
                available_katzins = [k for k in available_katzins if k not in unavailable_katzins_next]

            selected_mashak = None
            selected_katzin = None
            if available_mashaks:
                selected_mashak = min(available_mashaks, key=lambda m: mashak_assignments[m])
                mashak_assignments[selected_mashak] += 2
            if available_katzins:
                selected_katzin = min(available_katzins, key=lambda k: katzin_assignments[k])
                katzin_assignments[selected_katzin] += 2

            schedule.append({
                "Date": date,
                "Mashak": selected_mashak,
                "Katzin": selected_katzin,
                "Mashaks Unavailable": f"{', '.join(unavailable_mashaks)}",
                "Katzins Unavailable": f"{', '.join(unavailable_katzins)}"
            })
            schedule.append({
                "Date": next_day,
                "Mashak": selected_mashak,
                "Katzin": selected_katzin,
                "Mashaks Unavailable": f"{', '.join(mashaks_unavailable.get(next_day, []))}",
                "Katzins Unavailable": f"{', '.join(katzins_unavailable.get(next_day, []))}"
            })

        elif weekday != 5:
            selected_mashak = None
            selected_katzin = None
            if available_mashaks:
                selected_mashak = min(available_mashaks, key=lambda m: mashak_assignments[m])
                mashak_assignments[selected_mashak] += 1
            if available_katzins:
                selected_katzin = min(available_katzins, key=lambda k: katzin_assignments[k])
                katzin_assignments[selected_katzin] += 1

            schedule.append({
                "Date": date,
                "Mashak": selected_mashak,
                "Katzin": selected_katzin,
                "Mashaks Unavailable": f"{', '.join(unavailable_mashaks)}",
                "Katzins Unavailable": f"{', '.join(unavailable_katzins)}"
            })

    schedule_df = pd.DataFrame(schedule)
    mashak_assignments_df = pd.DataFrame(
        [{'Name': name, 'Working days': mashak_assignments[name]} for name in mashak_assignments.keys()]
    )
    katzin_assignments_df = pd.DataFrame(
        [{'Name': name, 'Working days': katzin_assignments[name]} for name in katzin_assignments.keys()]
    )

    return schedule_df, mashak_assignments_df, katzin_assignments_df


# Parameters
num_simulations = 100
results = []

# Load input data
potential_mashaks = pd.read_excel('inputs.xlsx', sheet_name='potential mashaks', dtype=str)
potential_katzins = pd.read_excel('inputs.xlsx', sheet_name='potential katzins', dtype=str)
mashaks_constraints = pd.read_excel('inputs.xlsx', sheet_name='mashaks constraints', dtype=str)
katzins_constraints = pd.read_excel('inputs.xlsx', sheet_name='katzins constraints', dtype=str)

# Run multiple simulations
for _ in range(num_simulations):
    # Randomize the order of names
    potential_mashaks = potential_mashaks.sample(frac=1).reset_index(drop=True)
    potential_katzins = potential_katzins.sample(frac=1).reset_index(drop=True)
    # Run the scheduling simulation
    schedule_df, mashak_assignments_df, katzin_assignments_df = run_schedule_simulation(
        potential_mashaks, potential_katzins, mashaks_constraints, katzins_constraints)
    # Calculate balance score (difference between max and min assignments)
    mashak_balance = mashak_assignments_df['Working days'].max() - mashak_assignments_df['Working days'].min()
    katzin_balance = katzin_assignments_df['Working days'].max() - katzin_assignments_df['Working days'].min()
    total_balance = katzin_balance + mashak_balance
    results.append((total_balance, schedule_df, mashak_assignments_df, katzin_assignments_df))

# Sort results by balance score (lower is better)
results.sort(key=lambda x: x[0])

# Save top 3 results to Excel files
for i in range(3):
    _, schedule_df, mashak_assignments_df, katzin_assignments_df = results[i]
    with pd.ExcelWriter(f'output_best_{i + 1}.xlsx', engine='openpyxl') as writer:
        schedule_df.to_excel(writer, sheet_name='schedule', index=False)
        mashak_assignments_df.to_excel(writer, sheet_name='mashak_assignments', index=False)
        katzin_assignments_df.to_excel(writer, sheet_name='katzin_assignments', index=False)
