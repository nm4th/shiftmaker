#!/usr/bin/env python
# coding: utf-8

# In[202]:


get_ipython().system('pip install pulp')
import pulp
import pandas as pd

# 月の日数とシフトタイプを設定
days_in_month = 30  # 月の日数を適切に設定してください
shift_types = ['早番', '遅番', '夜勤', '休日']
work_types = ['早番', '遅番', '夜勤']
# モデルの作成
model = pulp.LpProblem("Shift_Scheduling", pulp.LpMinimize)

# 変数の定義
employees = range(1, 18)  # 従業員は17人
days = range(1, days_in_month + 1)
shift_vars = pulp.LpVariable.dicts("Shift", (employees, days, shift_types), cat='Binary')

# 目的関数の定義
model += pulp.lpSum(shift_vars[e][d][shift] for e in employees for d in days for shift in shift_types)

# 制約条件の追加
# 連続勤務時は同じシフト
for e in employees:
    for d in range(1, days_in_month):
        model += pulp.lpSum(shift_vars[e][d]['夜勤'] + shift_vars[e][d + 1]['早番']) <=1
        model += pulp.lpSum(shift_vars[e][d]['夜勤'] + shift_vars[e][d + 1]['遅番']) <=1
        model += pulp.lpSum(shift_vars[e][d]['遅番'] + shift_vars[e][d + 1]['早番']) <=1
        model += pulp.lpSum(shift_vars[e][d]['遅番'] + shift_vars[e][d + 1]['夜勤']) <=1
        model += pulp.lpSum(shift_vars[e][d]['早番'] + shift_vars[e][d + 1]['夜勤']) <=1
        model += pulp.lpSum(shift_vars[e][d]['早番'] + shift_vars[e][d + 1]['遅番']) <=2
        

#夜勤明けの制約        
for e in employees:
    for d in range(1, days_in_month-1):
        model +=shift_vars[e][d]['夜勤'] + shift_vars[e][d + 1]['休日'] + shift_vars[e][d + 2]['早番'] <=2
        model +=shift_vars[e][d]['夜勤'] + shift_vars[e][d + 1]['休日'] + shift_vars[e][d + 2]['遅番'] <=2

#休みは３連休まで            
for e in employees:
    for d in range(1, days_in_month-1):
        model += pulp.lpSum([shift_vars[e][d]['休日'] + shift_vars[e][d + 1]['休日'] + shift_vars[e][d + 2]['休日']]) <= 3

#４連勤まで     
    for d in range(1, days_in_month-3):
        model += pulp.lpSum(shift_vars[e][d]['休日'] + shift_vars[e][d + 1]['休日'] + shift_vars[e][d + 2]['休日']+ shift_vars[e][d + 3]['休日']+ shift_vars[e][d + 4]['休日']) >= 1
# 各従業員は飛び石連休なし  
for e in employees:
    for d in range(1,days_in_month-1):
        model +=shift_vars[e][d]['休日']+shift_vars[e][d+2]['休日'] <=1
        
# 一日に担当できるシフトタイプは1つのみ
for e in employees:
    for d in days:
        model += pulp.lpSum(shift_vars[e][d][shift] for shift in shift_types) == 1
        
# 各従業員は8日以下夜勤でなければならない       
for e in employees:
    model += pulp.lpSum(shift_vars[e][d]['夜勤'] for d in days) >= 4
    model += pulp.lpSum(shift_vars[e][d]['夜勤'] for d in days) <= 8
    
# 各従業員の休日の合計は10日でなければならない
for e in employees:
    model += pulp.lpSum(shift_vars[e][d]['休日'] for d in days) == 10
    

# 各日において、早番は3人以上、遅番は1人以上、夜勤は4人以上にならなければならない
for d in days:
    model += pulp.lpSum(shift_vars[e][d]['早番'] for e in employees) >= 3
    #model += pulp.lpSum(shift_vars[e][d]['早番'] for e in employees) <= 4
    model += pulp.lpSum(shift_vars[e][d]['遅番'] for e in employees) >= 1
    #model += pulp.lpSum(shift_vars[e][d]['遅番'] for e in employees) <= 2
    model += pulp.lpSum(shift_vars[e][d]['夜勤'] for e in employees) >= 4
# 最適化
model.solve()

# シフトをExcelファイルとして出力
shift_schedule = pd.DataFrame(index=employees, columns=days)

for e in employees:
    for d in days:
        for shift in shift_types:
            if pulp.value(shift_vars[e][d][shift]) == 1:
                shift_schedule.at[e, d] = shift

status = model.solve()
print("Status:",pulp.LpStatus[status])
shift_schedule.to_excel("shift_schedule.xlsx")

