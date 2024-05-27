import pandas as pd
from datetime import datetime, timedelta
import openpyxl

def load_data(file_name):
    data = pd.read_csv(file_name)
    data = data[data['Parent task'].isna()]  # Exclude subtasks
    data = data[data['Projects'] != 'Статистика']  # Exclude 'Статистика' project
    return data

def calculate_task_days(data):
    data['Start Date'] = pd.to_datetime(data['Start Date'])
    data['Due Date'] = pd.to_datetime(data['Due Date'])
    data['task_days'] = (data['Due Date'] - data['Start Date']).dt.days
    return data

def calculate_project_completion(data):
    data = calculate_task_days(data)
    total_task_days = data.groupby('Projects')['task_days'].sum()
    incomplete_tasks = data[data['Completed At'].isna()]
    incomplete_task_days = incomplete_tasks.groupby('Projects')['task_days'].sum()
    project_completion = 1 - (incomplete_task_days / total_task_days)
    return project_completion

data = load_data('project.csv')
project_completion = calculate_project_completion(data)
project_completion = project_completion.copy().mul(100).round(0)
project_completion.to_excel('Данные по теспу ведения проекта.xlsx')