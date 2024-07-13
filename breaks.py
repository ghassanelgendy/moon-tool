import pandas as pd
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
from datetime import datetime, timedelta


def generate_break_schedule(agent_names, start_time, break_schema):
    # Define the columns for the DataFrame
    if break_schema == '1':
        columns = ['Agent Name', 'Start Time', '15 Min Break', '30 Min Break', '15 Min Break', 'End Time']
    elif break_schema == '2':
        columns = ['Agent Name', 'Start Time', '30 Min Break', '30 Min Break', 'End Time']
    else:
        raise ValueError("Invalid break schema. Choose '1' for schema: 15-30-15 or '2' for schema : 30-30.")
    
    data = []
    
    start = datetime.strptime(start_time, '%I:%M %p')
    end = start + timedelta(hours=9)
    
    # Initial break times for the first agent
    break_time = start + timedelta(hours=2)

    for i, agent_name in enumerate(agent_names):
        if break_schema == '1':
            first_break = break_time
            second_break = first_break + timedelta(minutes=15) + timedelta(hours=2)
            third_break = second_break + timedelta(minutes=30) + timedelta(hours=2)
            break_time += timedelta(minutes=15)  # Next agent's first break is 15 mins later
            row = [agent_name, start.strftime('%I:%M %p'), first_break.strftime('%I:%M %p'), 
                   second_break.strftime('%I:%M %p'), third_break.strftime('%I:%M %p'), end.strftime('%I:%M %p')]
        elif break_schema == '2':
            first_break = break_time
            second_break = first_break + timedelta(minutes=30) + timedelta(hours=3)
            break_time += timedelta(minutes=30)  # Next agent's first break is 30 mins later
            row = [agent_name, start.strftime('%I:%M %p'), first_break.strftime('%I:%M %p'), 
                   second_break.strftime('%I:%M %p'), end.strftime('%I:%M %p')]
        
        data.append(row)
    
    df = pd.DataFrame(data, columns=columns)
    return df



def main():
    # Get user input for shift start time
    shift_start_times = {
        9: '09:00 AM',
        7: '07:00 AM',
        11: '11:00 AM',
        10: '10:00 PM',
        4: '04:00 PM',
        1: '01:00 PM'
    }

    today = datetime.today()
    shift_choice = int(input("Enter the shift start time: "))
    start_time = shift_start_times.get(shift_choice)
    filename = 'breaks shift ' + str(shift_choice) +' '+ str(today.day) + '-' +str(today.month)+'.xlsx' 
    if not start_time:
        print("Invalid shift start time.")
        return

    # Get user input for agent names in one line, space-separated
    agent_names = input("Enter the names of agents, space-separated: ").split()

    # Get user input for break schema
    break_schema = input("Enter the break schema \n\t1 = (15-30-15) \n\t2 = (30-30)\n=> ")

    # Generate schedule and save to Excel
    schedule_df = generate_break_schedule(agent_names, start_time, break_schema)
    save_to_excel(schedule_df, filename)
    print(f"Break schedule saved to {filename}")
if __name__ == "__main__":
    main()
