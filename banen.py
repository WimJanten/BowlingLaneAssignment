import pandas as pd
from openpyxl import Workbook
from datetime import datetime, timedelta

# ---- CONFIGURATION ----
input_file = 'Geboekte producten.xlsx'  # Your reservation Excel file
output_file_full = 'assigned_lanes_full.xlsx'  # Full schedule
output_file_compact = 'assigned_lanes_compact.xlsx'  # Compact schedule

# ---- STEP 1: Load Data ----
df = pd.read_excel(input_file)

# ---- STEP 2: Process Data ----

# Decide number of lanes needed
def calculate_lanes_needed(row):
    if row['Aantal'] >= 4:
        return (row['Aantal'] + 5) // 6  # 6 persons per lane, round up
    else:
        return row['Aantal']  # Aantal is lanes if < 4

df['lanes_needed'] = df.apply(calculate_lanes_needed, axis=1)

# Initialize lane booking memory
lanes = {i: [] for i in range(1, 9)}  # lanes 1 to 8
assignments = []

# Sort by group and start time
df = df.sort_values(by=['Groep', 'Begindatum'])

# Helper function to check if a lane is free
def is_lane_free(lane, new_start, new_end):
    for booked_start, booked_end in lanes[lane]:
        if new_start < booked_end and new_end > booked_start:
            return False
    return True

# Facing pairs
facing_pairs = [(1,2), (3,4), (5,6), (7,8)]

previous_assignment = {}

# Process each reservation
for idx, row in df.iterrows():
    start_time = pd.to_datetime(row['Begindatum'])
    groep = row['Groep']
    lanes_needed = row['lanes_needed']
    end_time = start_time + timedelta(minutes=55)

    if start_time.minute == 0:
        possible_lanes = range(1, 5)  # 1-4
        possible_pairs = [(1,2), (3,4)]
    elif start_time.minute == 30:
        possible_lanes = range(5, 9)  # 5-8
        possible_pairs = [(5,6), (7,8)]
    else:
        print(f"Skipping {groep} at {start_time} — invalid start time (must be :00 or :30).")
        continue

    assigned_lanes = []

    # Try to continue previous lane assignment if group matches and times are consecutive
    if groep in previous_assignment:
        last_end_time, last_lanes = previous_assignment[groep]
        if abs((start_time - last_end_time).total_seconds()) <= 5 * 60:
            if all(is_lane_free(lane, start_time, end_time) for lane in last_lanes) and len(last_lanes) == lanes_needed:
                assigned_lanes = last_lanes

    # Otherwise, assign new lanes
    if not assigned_lanes:
        # First, try full facing pairs
        for pair in possible_pairs:
            if all(is_lane_free(lane, start_time, end_time) for lane in pair):
                if lanes_needed == 2:
                    assigned_lanes = list(pair)
                    break
                elif lanes_needed > 2:
                    assigned_lanes.extend(list(pair))
        
        # If still need more lanes (e.g. lanes_needed > 2)
        if len(assigned_lanes) < lanes_needed:
            for lane in possible_lanes:
                if lane not in assigned_lanes and is_lane_free(lane, start_time, end_time):
                    assigned_lanes.append(lane)
                if len(assigned_lanes) == lanes_needed:
                    break

    if len(assigned_lanes) < lanes_needed:
        print(f"⚠️ Not enough lanes available for {groep} at {start_time}. Only {len(assigned_lanes)} assigned.")
        continue

    # Book the lanes
    for lane in assigned_lanes:
        lanes[lane].append((start_time, end_time))

    assignments.append({
        'Groep': groep,
        'Starttijd': start_time,
        'Eindtijd': end_time,
        'Lanes': assigned_lanes
    })

    previous_assignment[groep] = (end_time, assigned_lanes)

# ---- STEP 3: Build the schedule dictionary ----

# Build time slots: from 10:00 to 23:30 in half-hour increments
start_time_of_day = datetime.strptime("10:00", "%H:%M")
timeslots = [start_time_of_day + timedelta(minutes=30*i) for i in range(28)]  # 10:00 to 23:30

# Create base schedule
schedule = {}
for time in timeslots:
    schedule[time.strftime("%H:%M")] = {lane: "" for lane in range(1, 9)}

# Fill reservations
for entry in assignments:
    groep = entry['Groep']
    start_time = entry['Starttijd']
    lanes = entry['Lanes']
    time_str = start_time.strftime("%H:%M")
    
    if time_str in schedule:
        for lane in lanes:
            schedule[time_str][lane] = groep
    else:
        print(f"⚠️ Warning: Time {time_str} not in basic schedule range.")

# ---- STEP 4: Save Full Version (even empty slots) ----

wb_full = Workbook()
ws_full = wb_full.active
ws_full.title = "Bowling Schedule Full"

# Write headers
headers = ["Time"] + [f"{i}" for i in range(1, 9)]
ws_full.append(headers)

for time_str in schedule:
    row = [time_str] + [schedule[time_str][lane] for lane in range(1, 9)]
    ws_full.append(row)

wb_full.save(output_file_full)
print(f"✅ Full schedule saved to {output_file_full}")

# ---- STEP 5: Save Compact Version (only booked times) ----

wb_compact = Workbook()
ws_compact = wb_compact.active
ws_compact.title = "Bowling Schedule Compact"

# Write headers
ws_compact.append(headers)

for time_str, lanes_dict in schedule.items():
    if any(lanes_dict[lane] != "" for lane in range(1, 9)):  # Only times with at least 1 reservation
        row = [time_str] + [lanes_dict[lane] for lane in range(1, 9)]
        ws_compact.append(row)

wb_compact.save(output_file_compact)
print(f"✅ Compact schedule saved to {output_file_compact}")
