import pandas as pd
from datetime import datetime, timedelta
import openpyxl

def load_data(file_name):
    data = pd.read_csv(file_name)
    data = data[data['Parent task'].isna()]  # Exclude subtasks
    return data

def get_tasks_per_month(data):
    data['Created At'] = pd.to_datetime(data['Created At'])
    data['month_year'] = data['Created At'].dt.to_period('M')
    tasks_per_month = data.groupby('month_year').size()
    return tasks_per_month

def get_completed_tasks(data):
    data['Completed At'] = pd.to_datetime(data['Completed At'])
    completed_tasks = data[data['Completed At'].notna()]
    return completed_tasks

def get_completion_rate(data):
    tasks_per_month = get_tasks_per_month(data)
    completed_tasks_per_month = get_tasks_per_month(get_completed_tasks(data))
    completion_rate = completed_tasks_per_month / tasks_per_month
    return completion_rate

def get_tasks_per_week(data):
    data['week_year'] = data['Created At'].dt.to_period('W')
    tasks_per_week = data.groupby('week_year').size()
    return tasks_per_week

def get_overdue_tasks(data):
    data['Due Date'] = pd.to_datetime(data['Due Date'])
    overdue_tasks = data[data['Completed At'] > data['Due Date']]
    return overdue_tasks

def get_overdue_tasks_by_project(data):
    overdue_tasks = get_overdue_tasks(data)
    overdue_tasks_by_project = overdue_tasks.groupby('Projects').size()
    return overdue_tasks_by_project

def get_overdue_tasks_by_assignee(data):
    overdue_tasks = get_overdue_tasks(data)
    overdue_tasks_by_assignee = overdue_tasks.groupby('Assignee').size()
    return overdue_tasks_by_assignee

def get_avg_deviation(data):
    completed_tasks = get_completed_tasks(data)
    completed_tasks['deviation'] = (completed_tasks['Completed At'] - completed_tasks['Due Date']).dt.days
    avg_deviation = completed_tasks['deviation'].mean()
    return avg_deviation

def write_to_excel(data, file_name):
    with pd.ExcelWriter(file_name) as writer:
        for key, value in data.items():
            if isinstance(value, pd.DataFrame) or isinstance(value, pd.Series):
                value.to_excel(writer, sheet_name=key)
            else:
                pd.DataFrame([value], columns=[key]).to_excel(writer, sheet_name=key)

data = load_data('project.csv')
results = {
    'Рейтинг выполнения задач': get_completion_rate(data),
    'Задачи за неделю': get_tasks_per_week(data),
    'Выполненные задачи по проекту': get_overdue_tasks_by_project(data),
    'Выполненные задачи по исполнителю': get_overdue_tasks_by_assignee(data),
    'Среднее отклонение от выполнения задач': get_avg_deviation(data)
}
write_to_excel(results, 'Результат.xlsx')