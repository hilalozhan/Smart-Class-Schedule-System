import pandas as pd
import numpy as np
from datetime import datetime
from tabulate import tabulate
import math

# Load sample data set
file_path = 'path your dataset'
classrooms_df = pd.read_excel(file_path, sheet_name='Classrooms')
fall_semester_df = pd.read_excel(file_path, sheet_name='Fall Semester')
spring_semester_df = pd.read_excel(file_path, sheet_name='Spring Semester')
students_df = pd.read_excel(file_path, sheet_name='Students')

# Data preprocessing
fall_semester_df.columns = fall_semester_df.columns.str.strip()
spring_semester_df.columns = spring_semester_df.columns.str.strip()

# Prepare the list of classrooms and courses
classrooms = classrooms_df[['Classroom', 'Capacity']].to_dict(orient='records')
courses_fall = fall_semester_df[['Course', 'Term', 'ECTS']].to_dict(orient='records')
courses_spring = spring_semester_df[['Course', 'Term', 'ECTS']].to_dict(orient='records')
students = students_df[['Student', 'Grade']].to_dict(orient='records')

# Define elective courses
major_elective_courses = [
..
]

professional_elective_courses = [
..
]

general_culture_courses = [
 ..   
]

# Function to convert time strings to minutes
def time_to_minutes(time_str):
    time_parts = time_str.split(':')
    return int(time_parts[0]) * 60 + int(time_parts[1])

# Define time slots (2-hour intervals with 30-minute breaks from 08:30 to 17:30)
time_slots = [(f'{hour:02d}:{minute:02d}', f'{hour + 2:02d}:{minute:02d}') for hour, minute in [(8, 30), (11, 0), (13, 30), (16, 0)]]

weekdays = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday']

all_courses_fall = courses_fall + major_elective_courses + professional_elective_courses + general_culture_courses
all_courses_spring = courses_spring + major_elective_courses + professional_elective_courses + general_culture_courses

# Particle Swarm Optimization (PSO) parameters
POPULATION = 50
MAX_ITER = 300
V_MAX = 4
B_LO = 0
B_HI = len(classrooms) - 1
DIMENSIONS = len(all_courses_fall)
PERSONAL_C = 2
SOCIAL_C = 2
GLOBAL_BEST = 0  # Assumed value for convergence criteria
CONVERGENCE = 1e-6

def is_slot_available(occupied_slots, classroom, day, start_time, end_time, schedule):
    for slot in occupied_slots[day][classroom]:
        if (slot[0] < end_time and start_time < slot[1]):
            return False
    return True

def check_student_schedule_conflict(student_schedule, student_class, day, start_time, end_time):
    for slot in student_schedule:
        if (slot[1] == day and slot[2] < end_time and start_time < slot[3]):
            return False
    return True

def generate_final_schedule(best_schedule, all_courses):
    final_schedule = []
    for i, course_index in enumerate(best_schedule):
        course = all_courses[i]
        classroom = classrooms[int(course_index)]
        day = weekdays[i % len(weekdays)]
        start_time, end_time = time_slots[i % len(time_slots)]
        final_schedule.append({
            'Course': course['Course'],
            'Grade': course.get('Term', 'Elective'),
            'Classroom': classroom['Classroom'],
            'Day': day,
            'Start Time': start_time,
            'End Time': end_time
        })
    return final_schedule

def calculate_student_travel_time(student_courses):
    total_time = 0
    for i in range(len(student_courses) - 1):
        end_time_current = datetime.strptime(student_courses[i]['End Time'], '%H:%M')
        start_time_next = datetime.strptime(student_courses[i+1]['Start Time'], '%H:%M')
        if start_time_next > end_time_current:
            total_time += (start_time_next - end_time_current).seconds / 60
    return total_time

def calculate_travel_time(schedule):
    travel_time = 0
    for student in students:
        student_courses = [course for course in schedule if course['Grade'] == student['Grade']]
        if len(student_courses) > 1:
            travel_time += calculate_student_travel_time(student_courses)
    return travel_time

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

def fitness_function(schedule, all_courses):
    conflicts = 0
    occupied_slots = {day: {classroom['Classroom']: [] for classroom in classrooms} for day in weekdays}
    student_schedule = {student['Student']: [] for student in students}
    final_schedule = []
    
    for i, course_index in enumerate(schedule):
        course = all_courses[i]
        classroom = classrooms[int(course_index)]
        day = weekdays[i % len(weekdays)]
        start_time, end_time = time_slots[i % len(time_slots)]
        
        # Check if it's an elective course
        if 'Term' not in course:
            term = None
            capacity_check = False  # No capacity check for elective courses
        else:
            term = course['Term']
            capacity_check = True  # Capacity check for non-elective courses
        
        # Check if capacity check is required
        if capacity_check:
            # Calculate student count for the course
            if 'ECTS' in course:
                student_count = sum(1 for student in students if student['Grade'] == course['Term'])
            else:
                student_count = 0
            
            # Check if classroom capacity is sufficient
            if student_count > classroom['Capacity']:
                continue  # Try again with a different slot if capacity is exceeded
        
        # Check classroom overlaps
        if not is_slot_available(occupied_slots, classroom['Classroom'], day, start_time, end_time, None):
            conflicts += 1
        
        # Check student overlaps
        for student in students:
            if student['Grade'] == course.get('Term', 'Elective'):
                if not check_student_schedule_conflict(student_schedule[student['Student']], student['Grade'], day, start_time, end_time):
                    conflicts += 1

        # Mark the time slot if no conflicts
        occupied_slots[day][classroom['Classroom']].append((start_time, end_time))
        final_schedule.append({
            'Course': course['Course'],
            'Grade': course.get('Term', 'Elective'),
            'Classroom': classroom['Classroom'],
            'Day': day,
            'Start Time': start_time,
            'End Time': end_time
        })
    
    travel_time = calculate_travel_time(final_schedule)
    fitness_value = conflicts + travel_time
    return fitness_value

class Particle:
    def __init__(self, dimensions, b_lo, b_hi, v_max):
        self.position = np.random.uniform(b_lo, b_hi, dimensions)
        self.velocity = np.random.uniform(-v_max, v_max, dimensions)
        self.best_position = self.position.copy()
        self.best_fitness = float('inf')
        self.fitness = float('inf')
    
    def update_velocity(self, global_best_position, w, personal_c, social_c):
        r1 = np.random.random(self.position.shape)
        r2 = np.random.random(self.position.shape)
        
        personal_velocity = personal_c * r1 * (self.best_position - self.position)
        social_velocity = social_c * r2 * (global_best_position - self.position)
        self.velocity = w * self.velocity + personal_velocity + social_velocity
    
    def update_position(self, b_lo, b_hi):
        self.position += self.velocity
        self.position = np.clip(self.position, b_lo, b_hi)

def particle_swarm_optimization(dimensions, b_lo, b_hi, v_max, personal_c, social_c, w_max, w_min, population_size, max_iter, all_courses):
    particles = [Particle(dimensions, b_lo, b_hi, v_max) for _ in range(population_size)]
    global_best_position = np.zeros(dimensions)
    global_best_fitness = float('inf')
    w = w_max
    
    for iteration in range(max_iter):
        for particle in particles:
            particle.fitness = fitness_function(particle.position, all_courses)
            
            if particle.fitness < particle.best_fitness:
                particle.best_fitness = particle.fitness
                particle.best_position = particle.position.copy()
            
            if particle.fitness < global_best_fitness:
                global_best_fitness = particle.fitness
                global_best_position = particle.position.copy()
        
        for particle in particles:
            particle.update_velocity(global_best_position, w, personal_c, social_c)
            particle.update_position(b_lo, b_hi)
        
        w = w_max - (w_max - w_min) * iteration / max_iter
    
    return global_best_position

best_schedule_fall = particle_swarm_optimization(
    dimensions=DIMENSIONS,
    b_lo=B_LO,
    b_hi=B_HI,
    v_max=V_MAX,
    personal_c=PERSONAL_C,
    social_c=SOCIAL_C,
    w_max=0.9,
    w_min=0.2,
    population_size=POPULATION,
    max_iter=MAX_ITER,
    all_courses=all_courses_fall
)

best_schedule_spring = particle_swarm_optimization(
    dimensions=DIMENSIONS,
    b_lo=B_LO,
    b_hi=B_HI,
    v_max=V_MAX,
    personal_c=PERSONAL_C,
    social_c=SOCIAL_C,
    w_max=0.9,
    w_min=0.2,
    population_size=POPULATION,
    max_iter=MAX_ITER,
    all_courses=all_courses_spring
)

final_schedule_fall = generate_final_schedule(best_schedule_fall, all_courses_fall)
final_schedule_spring = generate_final_schedule(best_schedule_spring, all_courses_spring)

# Print final schedules
print("Fall Semester Schedule:")
print(tabulate(final_schedule_fall, headers='keys', tablefmt='pretty'))
print("\nSpring Semester Schedule:")
print(tabulate(final_schedule_spring, headers='keys', tablefmt='pretty'))


# Particle Swarm Optimization (PSO) parameters
POPULATION = 50
MAX_ITER = 300
V_MAX = 4
B_LO = 0
B_HI = len(classrooms) - 1
DIMENSIONS = len(all_courses_fall)
PERSONAL_C = 2
SOCIAL_C = 2
GLOBAL_BEST = 0  # Assumed value for convergence criteria
CONVERGENCE = 1e-6

def is_slot_available(occupied_slots, classroom, day, start_time, end_time, schedule):
    for slot in occupied_slots[day][classroom]:
        if (slot[0] < end_time and start_time < slot[1]):
            return False
    return True

def check_student_schedule_conflict(student_schedule, student_class, day, start_time, end_time):
    for slot in student_schedule:
        if (slot[1] == day and slot[2] < end_time and start_time < slot[3]):
            return False
    return True

def generate_final_schedule(best_schedule, all_courses):
    final_schedule = []
    for i, course_index in enumerate(best_schedule):
        course = all_courses[i]
        classroom = classrooms[int(course_index)]
        day = weekdays[i % len(weekdays)]
        start_time, end_time = time_slots[i % len(time_slots)]
        final_schedule.append({
            'Course': course['Course'],
            'Grade': course.get('Semester', 'Elective'),
            'Classroom': classroom['Classroom'],
            'Day': day,
            'Start Time': start_time,
            'End Time': end_time
        })
    return final_schedule

def calculate_student_travel_time(student_courses):
    total_time = 0
    for i in range(len(student_courses) - 1):
        end_time_current = datetime.strptime(student_courses[i]['End Time'], '%H:%M')
        start_time_next = datetime.strptime(student_courses[i+1]['Start Time'], '%H:%M')
        if start_time_next > end_time_current:
            total_time += (start_time_next - end_time_current).seconds / 60
    return total_time

def calculate_travel_time(schedule):
    travel_time = 0
    for student in students:
        student_courses = [course for course in schedule if course['Grade'] == student['Grade']]
        if len(student_courses) > 1:
            travel_time += calculate_student_travel_time(student_courses)
    return travel_time

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

def fitness_function(schedule, all_courses):
    conflicts = 0
    occupied_slots = {day: {classroom['Classroom']: [] for classroom in classrooms} for day in weekdays}
    student_schedule = {student['Student']: [] for student in students}
    final_schedule = []
    
    for i, course_index in enumerate(schedule):
        course = all_courses[i]
        classroom = classrooms[int(course_index)]
        day = weekdays[i % len(weekdays)]
        start_time, end_time = time_slots[i % len(time_slots)]
        
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
                student_count = sum(1 for student in students if student['Grade'] == course['Semester'])
            else:
                student_count = 0
            
            # Check if classroom capacity is sufficient
            if student_count > classroom['Capacity']:
                continue  # Try again with a different slot if capacity is exceeded
        
        # Check for classroom overlaps
        if not is_slot_available(occupied_slots, classroom['Classroom'], day, start_time, end_time, None):
            conflicts += 1
        
        # Check for student overlaps
        for student in students:
            if student['Grade'] == course.get('Semester', 'Elective'):
                if not check_student_schedule_conflict(student_schedule[student['Student']], student['Grade'], day, start_time, end_time):
                    conflicts += 1

        # If no conflicts, mark the time slot as occupied
        occupied_slots[day][classroom['Classroom']].append((start_time, end_time))
        final_schedule.append({
            'Course': course['Course'],
            'Grade': course.get('Semester', 'Elective'),
            'Classroom': classroom['Classroom'],
            'Day': day,
            'Start Time': start_time,
            'End Time': end_time
        })
    
    travel_time = calculate_travel_time(final_schedule)
    fitness_value = conflicts + travel_time
    return fitness_value

class Particle:
    def __init__(self, dimensions, b_lo, b_hi, v_max):
        self.position = np.random.uniform(b_lo, b_hi, dimensions)
        self.velocity = np.random.uniform(-v_max, v_max, dimensions)
        self.best_position = self.position.copy()
        self.best_fitness = float('inf')
        self.fitness = float('inf')
    
    def update_velocity(self, global_best_position, w, personal_c, social_c):
        r1 = np.random.random(self.position.shape)
        r2 = np.random.random(self.position.shape)
        
        personal_velocity = personal_c * r1 * (self.best_position - self.position)
        social_velocity = social_c * r2 * (global_best_position - self.position)
        self.velocity = w * self.velocity + personal_velocity + social_velocity
    
    def update_position(self, b_lo, b_hi):
        self.position += self.velocity
        self.position = np.clip(self.position, b_lo, b_hi)

def particle_swarm_optimization(dimensions, b_lo, b_hi, v_max, personal_c, social_c, w_max, w_min, population_size, max_iter, all_courses):
    particles = [Particle(dimensions, b_lo, b_hi, v_max) for _ in range(population_size)]
    global_best_position = np.random.uniform(b_lo, b_hi, dimensions)
    global_best_fitness = float('inf')
    
    for iteration in range(max_iter):
        w = w_max - (w_max - w_min) * iteration / max_iter
        
        for particle in particles:
            particle.fitness = fitness_function(particle.position, all_courses)
            if particle.fitness < particle.best_fitness:
                particle.best_position = particle.position.copy()
                particle.best_fitness = particle.fitness
            
            if particle.fitness < global_best_fitness:
                global_best_position = particle.position.copy()
                global_best_fitness = particle.fitness
        
        for particle in particles:
            particle.update_velocity(global_best_position, w, personal_c, social_c)
            particle.update_position(b_lo, b_hi)
        
        if math.isclose(global_best_fitness, GLOBAL_BEST, rel_tol=CONVERGENCE):
            break
    
    return global_best_position

# Run PSO for fall semester courses
best_schedule_fall = particle_swarm_optimization(DIMENSIONS, B_LO, B_HI, V_MAX, PERSONAL_C, SOCIAL_C, 1.4, 0.1, POPULATION, MAX_ITER, all_courses_fall)
final_schedule_fall = generate_final_schedule(best_schedule_fall, all_courses_fall)

# Run PSO for spring semester courses
best_schedule_spring = particle_swarm_optimization(DIMENSIONS, B_LO, B_HI, V_MAX, PERSONAL_C, SOCIAL_C, 1.4, 0.1, POPULATION, MAX_ITER, all_courses_spring)
final_schedule_spring = generate_final_schedule(best_schedule_spring, all_courses_spring)

# Display final schedules
print("Final Schedule for Fall Semester:")
print(tabulate(final_schedule_fall, headers='keys', tablefmt='grid'))
print("\nFinal Schedule for Spring Semester:")
print(tabulate(final_schedule_spring, headers='keys', tablefmt='grid'))
