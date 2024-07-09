import pandas as pd
import random
from tabulate import tabulate
from datetime import datetime


# Load the Excel file
file_path = 'path to your file'
derslikler_df = pd.read_excel(file_path, sheet_name='Derslikler')
guz_yariyili_df = pd.read_excel(file_path, sheet_name='Güzyarıyılı')
bahar_yariyili_df = pd.read_excel(file_path, sheet_name='Baharyarıyılı ')
ogrenciler_df = pd.read_excel(file_path, sheet_name='Öğrenciler')

# Remove extra spaces in column names
guz_yariyili_df.columns = guz_yariyili_df.columns.str.strip()
bahar_yariyili_df.columns = bahar_yariyili_df.columns.str.strip()

# Process the data
classrooms = derslikler_df[['Derslik', 'Kapasite']].to_dict(orient='records')
courses_guz = guz_yariyili_df[['Ders', 'Yarıyıl', 'AKTS']].to_dict(orient='records')
courses_bahar = bahar_yariyili_df[['Ders', 'Yarıyıl', 'AKTS']].to_dict(orient='records')
students = ogrenciler_df[['Öğrenci', 'Sınıf']].to_dict(orient='records')

# Define elective courses
alansecmeli_courses = [
....
]

mesleksecmeli_courses = [
......
]

# Define general culture elective courses
genelkultur_courses = [
.....
]


# Update the courses list by removing the placeholder elective course
courses_guz = [course for course in courses_guz if "Alan Eğitimi Seçmeli Dersleri" not in course['Ders'] and "Meslek Bilgisi Seçmeli" not in course['Ders']and "Genel Kültür Seçmeli Dersleri" not in course['Ders']]
courses_guz.extend(alansecmeli_courses)
courses_guz.extend(mesleksecmeli_courses)
courses_guz.extend(genelkultur_courses)

# Update the courses list by removing the placeholder elective course
courses_bahar = [course for course in courses_bahar if "Alan Eğitimi Seçmeli Dersleri" not in course['Ders'] and "Meslek Bilgisi Seçmeli" not in course['Ders']and "Genel Kültür Seçmeli Dersleri" not in course['Ders']]
courses_bahar.extend(alansecmeli_courses)
courses_bahar.extend(mesleksecmeli_courses)
courses_bahar.extend(genelkultur_courses)


# Define time slots (2-hour intervals with 30-minute breaks from 08:30 to 17:30)
time_slots = [(f'{hour:02d}:{minute:02d}', f'{hour + 2:02d}:{minute:02d}') for hour, minute in [(8, 30), (11, 0), (13, 30), (16, 0)]]
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
    
    # Check term (yarıyıl) availability if term_slots is provided
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
    
    # Check term (yarıyıl) availability if term_slots is provided
    if term_slots:
        for (start, end) in term_slots[day]:
            if not (end_minutes + 30 <= time_to_minutes(start) or start_minutes >= time_to_minutes(end) + 30):
                return False
    
    return True

# Function to create a random schedule
def create_random_schedule(classrooms, courses, weekdays, time_slots):
    schedule = []
    occupied_slots = {day: {classroom['Derslik']: [] for classroom in classrooms} for day in weekdays}
    occupied_slots_by_term = {term: {day: [] for day in weekdays} for term in set(course['Yarıyıl'] for course in courses if 'Yarıyıl' in course)}
    
    for course in courses:
        assigned = False
        while not assigned:
            classroom = random.choice(classrooms)
            day = random.choice(weekdays)
            start_time, end_time = random.choice(time_slots)
            
            # Check if it's an elective course
            if 'Yarıyıl' not in course:
                term = None
                capacity_check = False  # No capacity check for elective courses
            else:
                term = course['Yarıyıl']
                capacity_check = True  # Capacity check for non-elective courses
            
            # Check if capacity check is required
            if capacity_check:
                # Calculate student count for the course
                if 'AKTS' in course:
                    student_count = sum(1 for student in students if student['Sınıf'] == course['Yarıyıl'])
                else:
                    student_count = 0
                
                # Check if classroom capacity is sufficient
                if student_count > classroom['Kapasite']:
                    continue  # Try again with a different slot if capacity is exceeded
            
            # Check if the time slot is available
            if is_slot_available(occupied_slots, classroom['Derslik'], day, start_time, end_time, occupied_slots_by_term.get(term)):
                occupied_slots[day][classroom['Derslik']].append((start_time, end_time))
                if term:
                    occupied_slots_by_term[term][day].append((start_time, end_time))
                schedule.append({
                    'Ders': course['Ders'],
                    'Sınıf': course.get('Yarıyıl', 'Seçmeli'),
                    'Derslik': classroom['Derslik'],
                    'Gün': day,
                    'Başlama Zamanı': start_time,
                    'Bitiş Zamanı': end_time,
                    'Kontenjan': classroom['Kapasite']
                })
                assigned = True
    
    return schedule

# Function to create an initial population
def create_initial_population(population_size, classrooms, courses, weekdays, time_slots):
    return [create_random_schedule(classrooms, courses, weekdays, time_slots) for _ in range(population_size)]

# Example usage
population_size = 20  # Define the size of the population
initial_population_guz = create_initial_population(population_size, classrooms, courses_guz, weekdays, time_slots)
initial_population_bahar = create_initial_population(population_size, classrooms, courses_bahar, weekdays, time_slots)

# Genetic Algorithm for Course Scheduling

def fitness(schedule):
    # Calculate fitness based on constraints and student preferences
    total_conflicts = calculate_conflicts(schedule)
    total_travel_time = calculate_travel_time(schedule)
    return 1 / (1 + total_conflicts + total_travel_time)

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
        end_time_current = datetime.strptime(student_courses[i]['Bitiş Zamanı'], '%H:%M')
        start_time_next = datetime.strptime(student_courses[i+1]['Başlama Zamanı'], '%H:%M')
        if start_time_next > end_time_current:
            total_time += (start_time_next - end_time_current).seconds / 60
    return total_time

def courses_overlap(course1, course2):
    # Check if two courses overlap in time and classroom
    if (course1['Gün'] == course2['Gün'] and
        course1['Derslik'] == course2['Derslik']):
        start_time1 = datetime.strptime(course1['Başlama Zamanı'], '%H:%M')
        end_time1 = datetime.strptime(course1['Bitiş Zamanı'], '%H:%M')
        start_time2 = datetime.strptime(course2['Başlama Zamanı'], '%H:%M')
        end_time2 = datetime.strptime(course2['Bitiş Zamanı'], '%H:%M')

        if (start_time1 < end_time2 and start_time2 < end_time1):
            return True
    return False

def crossover(schedule1, schedule2):
    # Perform crossover between two schedules
    crossover_point = random.randint(1, len(schedule1) - 1)
    new_schedule = schedule1[:crossover_point] + schedule2[crossover_point:]
    return new_schedule


def mutate(schedule, mutation_rate=0.2):
    for i in range(len(schedule)):
        if random.random() < mutation_rate:
            course = schedule[i]
            classroom = random.choice(classrooms)
            time_slot = random.choice(time_slots)
            # check tuple for secure access
            if len(time_slot) >= 3:
                schedule[i] = {
                    'Ders': course['Ders'],
                    'AKTS': course['AKTS'],
                    'Derslik': classroom['Derslik'],
                    'Kapasite': classroom['Kapasite'],
                    'Gün': time_slot[0],         # day info access
                    'Başlama Zamanı': time_slot[1],  # start time  info access
                    'Bitiş Zamanı': time_slot[2]   # end time info
                }
          #  else:
          #      print(f"Warning: Invalid time_slot tuple found: {time_slot}")
    return schedule



def select_best_population(population, retain_ratio=0.3, random_select=0.5):
    # Select the best individuals from the population to retain for the next generation
    sorted_population = sorted(population, key=lambda x: fitness(x), reverse=True)
    retain_length = int(len(sorted_population) * retain_ratio)
    parents = sorted_population[:retain_length]

    # Add some random individuals to promote genetic diversity
    for individual in sorted_population[retain_length:]:
        if random_select > random.random():
            parents.append(individual)
    return parents

def evolve(population):
    # Evolve the population using crossover and mutation
    parents = select_best_population(population)
    desired_length = len(population) - len(parents)
    children = []

    while len(children) < desired_length:
        male = random.randint(0, len(parents) - 1)
        female = random.randint(0, len(parents) - 1)
        if male != female:
            male = parents[male]
            female = parents[female]
            child = crossover(male, female)
            child = mutate(child)
            children.append(child)

    parents.extend(children)
    return parents

generations = 300
population_guz = initial_population_guz
population_bahar = initial_population_bahar

# Find the best schedule
best_schedule_guz = None
best_fitness_guz = 0.0  # Initialize with a low value or based on your fitness function

best_schedule_bahar = None
best_fitness_bahar = 0.0  # Initialize with a low value or based on your fitness function

for schedule in population_guz:
    current_fitness = fitness(schedule)
    if current_fitness > best_fitness_guz:
        best_fitness_guz = current_fitness
        best_schedule_guz = schedule

for schedule in population_bahar:
    current_fitness = fitness(schedule)
    if current_fitness > best_fitness_bahar:
        best_fitness_bahar = current_fitness
        best_schedule_bahar = schedule

for generation in range(generations):
    population_guz = evolve(population_guz)
    population_bahar = evolve(population_bahar)

    best_schedule_guz = max(population_guz, key=lambda x: fitness(x))
    best_schedule_bahar = max(population_bahar, key=lambda x: fitness(x))

    print(f"Generation {generation + 1} - Fitness (Güzyarıyılı): {fitness(best_schedule_guz)}")
    print(tabulate(best_schedule_guz, headers="keys", tablefmt="grid"))
    print()

    print(f"Generation {generation + 1} - Fitness (Baharyarıyılı): {fitness(best_schedule_bahar)}")
    print(tabulate(best_schedule_bahar, headers="keys", tablefmt="grid"))
    print()

# Print the best schedule's fitness value and details
print("En iyi programın Fitness Değeri (Güzyarıyılı):", best_fitness_guz)
print("En iyi programın Detayları (Güzyarıyılı):")
print(tabulate(best_schedule_guz, headers='keys'))

print("En iyi programın Fitness Değeri (Baharyarıyılı):", best_fitness_bahar)
print("En iyi programın Detayları (Baharyarıyılı):")
print(tabulate(best_schedule_bahar, headers='keys'))