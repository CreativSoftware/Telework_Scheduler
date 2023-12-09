import random
import pandas
import os

users = ["Edward", "Kabba", "Joe", "Gary", "Carol"]
days_of_week = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday"]


def get_time_off():
    user_vacation_days = {} 
    selecting_name = True
    vacation_answer = input("Anybody taking vacation this week? (Y/N): ").lower()
    while selecting_name:
        if vacation_answer == "y":
            chosen_name = input("Enter the name of the person taking vacation: ").title()
            if chosen_name not in users:
                print("Invalid Name...")
                continue
            selecting_days = True
            while selecting_days:
                vacation_days = input("Enter the days off (comma-separated, e.g., Monday, Wednesday): ").title().replace(' ','').split(',')
                for day in vacation_days:
                    if day not in days_of_week:
                        print("Invalid Day...")
                        continue
                    else:
                        break
                break       
        elif vacation_answer == "n":
            break
        else:
            print("Invalid input. Please enter 'yes' or 'no'.")
        user_vacation_days[chosen_name] = vacation_days
        vacation_answer_two = input("Anybody ELSE taking vacation this week? (Y/N): ").lower()
        if vacation_answer_two ==  "y":
            continue
        else:
            break
    return user_vacation_days
        

def scheduler():
    time_off = get_time_off()
    days_assigned = {}
    removed_days = set()
    for user in users:
        for name, vacation_days in time_off.items():
            if name == user:
                for vacation_day in vacation_days:
                    if vacation_day in days_of_week:
                        days_of_week.remove(vacation_day)
                        removed_days.add(tuple(vacation_days))

        work_home = random.sample(days_of_week, 2)
        work_office = []
        for day in days_of_week:
            if day not in work_home:
                work_office.append(day)

        work_home = set(work_home)
        work_office = set(work_office)
        days_assigned[user] = {
            "Work from Home": work_home,
            "Work in Office": work_office
        }
        days_of_week.extend(list(removed_days))

    data = pandas.DataFrame.from_dict(days_assigned, orient='index')
    file_path = os.path.join(os.getcwd(), ".\Telework Schedule.xlsx")
    data.to_excel(file_path, engine='openpyxl', index_label="User")
    return days_assigned
           

scheduler()
