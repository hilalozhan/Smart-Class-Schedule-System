#!/usr/bin/env python
# coding: utf-8

# In[ ]:


import pandas as pd
import random
from tabulate import tabulate
from datetime import datetime


classrooms_df = pd.read_excel(file_path, sheet_name='Classrooms')
fall_semester_df = pd.read_excel(file_path, sheet_name='Fall Semester')
spring_semester_df = pd.read_excel(file_path, sheet_name='Spring Semester')
students_df = pd.read_excel(file_path, sheet_name='Students')

# Remove extra spaces in column names
fall_semester_df.columns = fall_semester_df.columns.str.strip()
spring_semester_df.columns = spring_semester_df.columns.str.strip()

# Process the data
classrooms = classrooms_df[['Classroom', 'Capacity']].to_dict(orient='records')
courses_fall = fall_semester_df[['Course', 'Semester', 'ECTS']].to_dict(orient='records')
courses_spring = spring_semester_df[['Course', 'Semester', 'ECTS']].to_dict(orient='records')
students = students_df[['Student', 'Class']].to_dict(orient='records')

# Define elective courses
elective_courses = [
....
]

professional_courses = [
.....
]

general_culture_courses = [
.....
]
# Update the courses list by removing the placeholder elective course
courses_fall = [course for course in courses_fall if "Elective Courses for Field Education" not in course['Course'] and "Elective Vocational Knowledge" not in course['Course'] and "Elective Courses for General Culture" not in course['Course']]
courses_fall.extend(elective_courses)
courses_fall.extend(professional_courses)
courses_fall.extend(general_culture_courses)

# Update the courses list by removing the placeholder elective course
courses_spring = [course for course in courses_spring if "Elective Courses for Field Education" not in course['Course'] and "Elective Vocational Knowledge" not in course['Course'] and "Elective Courses for General Culture" not in course['Course']]
courses_spring.extend(elective_courses)
courses_spring.extend(professional_courses)
courses_spring.extend(general_culture_courses)


# Define time intervals by excluding lunch break
all_time_slots = [
    ('08:30', '10:30'),
    ('11:00', '13:00'),
    ('13:30', '15:30'),
    ('16:00', '18:00')
]

# Exclude lunch break by filtering time intervals
time_slots = [(start, end) for start, end in all_time_slots if not ('11:00' <= start < '13:30')]
# Define weekdays
weekdays = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday']

# Function to convert time strings to minutes
def time_to_minutes(time_str):
    hours, minutes = map(int, time_str.split(':'))
    return hours * 60 + minutes

# Function to check if a time slot is available
def is_slot_available(occupied_slots, classroom, day, start_time, end_time, term_slots):
    start_minutes = time_to_minutes(start_time)
    end_minutes = time_to_minutes(end_time)
    
    # Check classroom availability
    for (start, end) in occupied_slots[day][classroom]:
        if not (end_minutes + 30 <= time_to_minutes(start) or start_minutes >= time_to_minutes(end) + 30):
            return False
    
    # Check term (semester) availability if term_slots is provided
    if term_slots:
        for (start, end) in term_slots[day]:
            if not (end_minutes + 30 <= time_to_minutes(start) or start_minutes >= time_to_minutes(end) + 30):
                return False
    
    return True

# Function to check if a time slot is available
def is_slot_available(occupied_slots, classroom, day, start_time, end_time, term_slots):
    start_minutes = time_to_minutes(start_time)
    end_minutes = time_to_minutes(end_time)
    
    # Check classroom availability
    for (start, end) in occupied_slots[day][classroom]:
        if not (end_minutes + 30 <= time_to_minutes(start) or start_minutes >= time_to_minutes(end) + 30):
            return False
    
    # Check term (semester) availability if term_slots is provided
    if term_slots:
        for (start, end) in term_slots[day]:
            if not (end_minutes + 30 <= time_to_minutes(start) or start_minutes >= time_to_minutes(end) + 30):
                return False
    
    return True

# Function to create a random schedule
def create_random_schedule(classrooms, courses, weekdays, time_slots):
    schedule = []
    occupied_slots = {day: {classroom['Classroom']: [] for classroom in classrooms} for day in weekdays}
    occupied_slots_by_term = {term: {day: [] for day in weekdays} for term in set(course['Semester'] for course in courses if 'Semester' in course)}
    
    for course in courses:
        assigned = False
        while not assigned:
            classroom = random.choice(classrooms)
            day = random.choice(weekdays)
            start_time, end_time = random.choice(time_slots)
            
            # Check if it's an elective course
            if 'Semester' not in course:
                term = None
                capacity_check = False  # No capacity check for elective courses
            else:
                term = course['Semester']
                capacity_check = True  # Capacity check for non-elective courses
            
            # Check if capacity check is required
            if capacity_check:
                # Calculate student count for the course
                if 'ECTS' in course:
                    student_count = sum(1 for student in students if student['Class'] == course['Semester'])
                else:
                    student_count = 0
                
                # Check if classroom capacity is sufficient
                if student_count > classroom['Capacity']:
                    continue  # Try again with a different slot if capacity is exceeded
            
            # Check if the time slot is available
            if is_slot_available(occupied_slots, classroom['Classroom'], day, start_time, end_time, occupied_slots_by_term.get(term)):
                occupied_slots[day][classroom['Classroom']].append((start_time, end_time))
                if term:
                    occupied_slots_by_term[term][day].append((start_time, end_time))
                schedule.append({
                    'Course': course['Course'],
                    'Class': course.get('Semester', 'Elective'),
                    'Classroom': classroom['Classroom'],
                    'Day': day,
                    'Start Time': start_time,
                    'End Time': end_time,
                    'Capacity': classroom['Capacity']
                })
                assigned = True
    
    return schedule

# Function to create an initial population
def create_initial_population(population_size, classrooms, courses, weekdays, time_slots):
    return [create_random_schedule(classrooms, courses, weekdays, time_slots) for _ in range(population_size)]

# Example usage
population_size = 3000  # Define the size of the population
initial_population_fall = create_initial_population(population_size, classrooms, courses_fall, weekdays, time_slots)
initial_population_spring = create_initial_population(population_size, classrooms, courses_spring, weekdays, time_slots)

# Genetic Algorithm for Course Scheduling

def fitness(schedule):
    # Calculate fitness based on constraints and student preferences
    total_conflicts = calculate_conflicts(schedule)
    total_travel_time = calculate_travel_time(schedule)
    return  (1 + total_conflicts + total_travel_time)

def calculate_conflicts(schedule):
    # Calculate total conflicts in the schedule
    conflicts = 0
    for i in range(len(schedule)):
        for j in range(i + 1, len(schedule)):
            if courses_overlap(schedule[i], schedule[j]):
                conflicts += 1
    return conflicts

def calculate_travel_time(schedule):
    # Calculate total travel time for students
    travel_time = 0
    for student in students:
        student_courses = [course for course in schedule if course['Class'] == student['Class']]
        if len(student_courses) > 1:
            travel_time += calculate_student_travel_time(student_courses)
    return travel_time

def calculate_student_travel_time(student_courses):
    # Calculate travel time for a student between consecutive courses
    total_time = 0
    for i in range(len(student_courses) - 1):
        end_time_current = datetime.strptime(student_courses[i]['End Time'], '%H:%M')
        start_time_next = datetime.strptime(student_courses[i+1]['Start Time'], '%H:%M')
        if start_time_next > end_time_current:
            total_time += (start_time_next - end_time_current).seconds / 60
    return total_time

def courses_overlap(course1, course2):
    # Check if two courses overlap in time and classroom
    if (course1['Day'] == course2['Day'] and
        course1['Classroom'] == course2['Classroom']):
        start_time1 = datetime.strptime(course1['Start Time'], '%H:%M')
        end_time1 = datetime.strptime(course1['End Time'], '%H:%M')
        start_time2 = datetime.strptime(course2['Start Time'], '%H:%M')
        end_time2 = datetime.strptime(course2['End Time'], '%H:%M')

        if (start_time1 < end_time2 and start_time2 < end_time1):
            return True
    return False

# Calculate fitness values for the schedules
fitness_values_fall = [(schedule, fitness(schedule)) for schedule in initial_population_fall]
fitness_values_spring = [(schedule, fitness(schedule)) for schedule in initial_population_spring]

# Sort schedules by fitness values
sorted_fitness_values_fall = sorted(fitness_values_fall, key=lambda x: x[1], reverse=False)
sorted_fitness_values_spring = sorted(fitness_values_spring, key=lambda x: x[1], reverse=False)

# Print the best schedule for Fall Semester
print("Best Schedule for Fall Semester:")
print(tabulate(sorted_fitness_values_fall[0][0], headers="keys"))
print("Fitness Value:", sorted_fitness_values_fall[0][1])

# Print the best schedule for Spring Semester
print("\nBest Schedule for Spring Semester:")
print(tabulate(sorted_fitness_values_spring[0][0], headers="keys"))
print("Fitness Value:", sorted_fitness_values_spring[0][1])

# Print all fitness values for Fall Semester
print("\nFitness Values for Fall Semester Schedules:")
for i, (schedule, fitness_value) in enumerate(sorted_fitness_values_fall):
    print(f"Schedule {i+1}: Fitness Value = {fitness_value}")

# Print all fitness values for Spring Semester
print("\nFitness Values for Spring Semester Schedules:")
for i, (schedule, fitness_value) in enumerate(sorted_fitness_values_spring):
    print(f"Schedule {i+1}: Fitness Value = {fitness_value}")

