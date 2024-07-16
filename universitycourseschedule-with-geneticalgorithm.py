#!/usr/bin/env python
# coding: utf-8

# In[ ]:


import pandas as pd
import random
from tabulate import tabulate
from datetime import datetime

# Load the Excel file
file_path = 'path your dataset'
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

def courses_overlap(course1, course2):
    if course1['Day'] == course2['Day']:
        start_time1 = datetime.strptime(course1['Start Time'], '%H:%M')
        end_time1 = datetime.strptime(course1['End Time'], '%H:%M')
        start_time2 = datetime.strptime(course2['Start Time'], '%H:%M')
        end_time2 = datetime.strptime(course2['End Time'], '%H:%M')

        if (start_time1 < end_time2 and start_time2 < end_time1):
            return True
    return False

def calculate_conflicts(schedule):
    conflicts = 0
    for i in range(len(schedule)):
        for j in range(i + 1, len(schedule)):
            if courses_overlap(schedule[i], schedule[j]):
                conflicts += 1
    return conflicts

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

def mutate(schedule, classrooms, time_slots, mutation_rate=0.2):
    occupied_slots = {day: {classroom['Classroom']: [] for classroom in classrooms} for day in weekdays}
    occupied_slots_by_term = {term: {day: [] for day in weekdays} for term in set(course['Class'] for course in schedule if 'Class' in course)}

    for course in schedule:
        day = course['Day']
        classroom = course['Classroom']
        start_time = course['Start Time']
        end_time = course['End Time']
        term = course.get('Class', None)

        occupied_slots[day][classroom].append((start_time, end_time))
        if term:
            occupied_slots_by_term[term][day].append((start_time, end_time))

    for i in range(len(schedule)):
        if random.random() < mutation_rate:
            course = schedule[i]
            day = course['Day']
            classroom = random.choice(classrooms)
            start_time, end_time = random.choice(time_slots)

            term = course.get('Class', None)
            term_slots = occupied_slots_by_term.get(term, None)

            if is_slot_available(occupied_slots, classroom['Classroom'], day, start_time, end_time, term_slots):
                occupied_slots[course['Day']][course['Classroom']].remove((course['Start Time'], course['End Time']))
                if term:
                    occupied_slots_by_term[term][course['Day']].remove((course['Start Time'], course['End Time']))

                occupied_slots[day][classroom['Classroom']].append((start_time, end_time))
                if term:
                    occupied_slots_by_term[term][day].append((start_time, end_time))

                schedule[i] = {
                    'Course': course['Course'],
                    'Class': term,
                    'Classroom': classroom['Classroom'],
                    'Day': day,
                    'Start Time': start_time,
                    'End Time': end_time,
                    'Capacity': classroom['Capacity']
                }
    return schedule

def fitness(schedule):
    # Calculate fitness based on constraints and student preferences
    total_conflicts = calculate_conflicts(schedule)
    total_travel_time = calculate_travel_time(schedule)
    return  total_conflicts + total_travel_time

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
        student_courses = [course for course in schedule if course['Sınıf'] == student['Sınıf']]
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

def crossover(schedule1, schedule2):
    crossover_point = random.randint(1, len(schedule1) - 1)
    new_schedule = schedule1[:crossover_point] + schedule2[crossover_point:]
    return new_schedule

def select_best_population(population, retain_ratio=0.2, random_select=0.3):
    sorted_population = sorted(population, key=lambda x: calculate_conflicts(x), reverse=False)
    retain_length = int(len(sorted_population) * retain_ratio)
    parents = sorted_population[:retain_length]

    for individual in sorted_population[retain_length:]:
        if random_select > random.random():
            parents.append(individual)
    return parents

def tournament_selection(population, k):
    tournament = random.sample(population, k)
    best_individual = min(tournament, key=lambda x: calculate_conflicts(x))
    return best_individual

def evolve_with_tournament(population, classrooms, time_slots, k=2):
    parents = select_best_population(population)
    desired_length = len(population) - len(parents)
    children = []

    while len(children) < desired_length:
        parent1 = tournament_selection(parents, k)
        parent2 = tournament_selection(parents, k)
        if parent1 != parent2:
            child = crossover(parent1, parent2)
            child = mutate(child, classrooms, time_slots)
            
            conflicts = calculate_conflicts(child)
            if conflicts == 0:
                children.append(child)
            elif len(children) < desired_length:
                children.append(child)

    parents.extend(children)
    return parents
    
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
            
            if 'Semester' not in course:
                term = None
                capacity_check = False
            else:
                term = course['Semester']
                capacity_check = True
            
            if capacity_check:
                student_count = sum(1 for student in students if student['Class'] == course['Semester'])
                
                if student_count > classroom['Capacity']:
                    continue
            
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

# Initialize populations
population_size = 50
generations = 15
initial_population_fall = create_initial_population(population_size, classrooms, courses_fall, weekdays, time_slots)
initial_population_spring = create_initial_population(population_size, classrooms, courses_spring, weekdays, time_slots)
population_fall = initial_population_fall
population_spring = initial_population_spring

best_schedule_fall = None
best_fitness_fall = float('inf')
best_schedule_spring = None
best_fitness_spring = float('inf')

# Evolve populations
for generation in range(generations):
    population_fall = evolve_with_tournament(population_fall, classrooms, time_slots)
    population_spring = evolve_with_tournament(population_spring, classrooms, time_slots)
    
    fitness_values_fall = [(schedule, fitness(schedule)) for schedule in population_fall]
    fitness_values_spring = [(schedule, fitness(schedule)) for schedule in population_spring]

    sorted_fitness_values_fall = sorted(fitness_values_fall, key=lambda x: x[1], reverse=False)
    sorted_fitness_values_spring = sorted(fitness_values_spring, key=lambda x: x[1], reverse=False)

    if sorted_fitness_values_fall[0][1] < best_fitness_fall:
        best_fitness_fall = sorted_fitness_values_fall[0][1]
        best_schedule_fall = sorted_fitness_values_fall[0][0]

    if sorted_fitness_values_spring[0][1] < best_fitness_spring:
        best_fitness_spring = sorted_fitness_values_spring[0][1]
        best_schedule_spring = sorted_fitness_values_spring[0][0]

    print(f"Generation {generation + 1}: Best Fitness (Fall) = {best_fitness_fall}, Best Fitness (Spring) = {best_fitness_spring}")

def print_schedule(schedule):
    headers = schedule[0].keys()
    table = [course.values() for course in schedule]
    print(tabulate(table, headers=headers, tablefmt='grid'))

print("\nBest Schedule (Fall):")
print(f"Best Fitness (Fall) = {best_fitness_fall}")
print_schedule(best_schedule_fall)

print("\nBest Schedule (Spring):")
print(f"Best Fitness (Spring) = {best_fitness_spring}")
print_schedule(best_schedule_spring)

  


