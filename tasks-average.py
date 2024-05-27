

import pandas as pd
from datetime import datetime, timedelta
import openpyxl

def load_data(file_name):
    data = pd.read_csv(file_name)
    data = data[data['Parent task'].isna()]  # Exclude subtasks
    print(data)
    return data

def get_completion_rate(data):
    data['Created At'] = pd.to_datetime(data['Created At'])
    data['Completed At'] = pd.to_datetime(data['Completed At'])
    data['month_year'] = data['Created At'].dt.to_period('M')
    tasks_per_month = data.groupby(['Assignee', 'month_year']).size()
    completed_tasks = data[data['Completed At'].notna()]
    completed_tasks_per_month = completed_tasks.groupby(['Assignee', 'month_year']).size()
    completion_rate = completed_tasks_per_month / tasks_per_month
    return completion_rate.unstack()

data = load_data('project.csv')
completion_rate = get_completion_rate(data)
completion_rate.to_excel('рейтинг выполнения.xlsx')

