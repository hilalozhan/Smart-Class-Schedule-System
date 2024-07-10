import pandas as pd
from tabulate import tabulate

# Load Excel file
file_path = 'path your dataset'
classrooms_df = pd.read_excel(file_path, sheet_name='Classrooms')
fall_semester_df = pd.read_excel(file_path, sheet_name='Fall Semester')
spring_semester_df = pd.read_excel(file_path, sheet_name='Spring Semester')
students_df = pd.read_excel(file_path, sheet_name='Students')

# Remove spaces from column names
fall_semester_df.columns = fall_semester_df.columns.str.strip()
spring_semester_df.columns = spring_semester_df.columns.str.strip()

# Process the data
classrooms = classrooms_df[['Classroom', 'Capacity']].to_dict(orient='records')
courses_fall = fall_semester_df[['Course', 'Semester', 'ECTS']].to_dict(orient='records')
courses_spring = spring_semester_df[['Course', 'Semester', 'ECTS']].to_dict(orient='records')
students = students_df[['Student', 'Class']].to_dict(orient='records')

# Define elective courses
elective_courses = [
    # Define your elective courses here
]

professional_elective_courses = [
    # Define your professional elective courses here
]

general_culture_courses = [
    # Define your general culture elective courses here
]

# Define time slots and weekdays
time_slots = [(f'{hour:02d}:{minute:02d}', f'{hour + 2:02d}:{minute:02d}') for hour, minute in [(8, 30), (11, 0), (13, 30), (16, 0)]]
weekdays = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday']

def is_slot_available(schedule, classroom, day, start_time, end_time, student_count, capacity):
    start_minutes = time_to_minutes(start_time)
    end_minutes = time_to_minutes(end_time)
    
    if student_count > capacity:
        return False
    
    for entry in schedule:
        if entry['Classroom'] == classroom and entry['Day'] == day:
            entry_start = time_to_minutes(entry['Start Time'])
            entry_end = time_to_minutes(entry['End Time'])
            if not (end_minutes + 30 <= entry_start or start_minutes >= entry_end + 30):
                return False
    return True

def time_to_minutes(time_str):
    hours, minutes = map(int, time_str.split(':'))
    return hours * 60 + minutes

def backtrack(schedule, courses, classrooms, weekdays, time_slots, index=0):
    if index == len(courses):
        return schedule
    
    course = courses[index]
    for classroom in classrooms:
        for day in weekdays:
            for start_time, end_time in time_slots:
                student_count = sum(1 for student in students if student['Class'] == course['Semester'])
                if is_slot_available(schedule, classroom['Classroom'], day, start_time, end_time, student_count, classroom['Capacity']):
                    schedule.append({
                        'Course': course['Course'],
                        'Class': course.get('Semester', 'Elective'),
                        'Classroom': classroom['Classroom'],
                        'Day': day,
                        'Start Time': start_time,
                        'End Time': end_time,
                        'Capacity': classroom['Capacity']
                    })
                    result = backtrack(schedule, courses, classrooms, weekdays, time_slots, index + 1)
                    if result is not None:
                        return result
                    schedule.pop()
    return None

# Combine all courses
all_courses = courses_fall + courses_spring + elective_courses + professional_elective_courses + general_culture_courses

initial_schedule = []
final_schedule = backtrack(initial_schedule, all_courses, classrooms, weekdays, time_slots)

# Print the results as a table
if final_schedule:
    print(tabulate(final_schedule, headers="keys", tablefmt="grid"))
else:
    print("No suitable schedule found.")
