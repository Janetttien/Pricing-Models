#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Tue Feb 20 23:11:13 2024

@author: yangjiatian
"""

from pyomo.environ import *
import pandas as pd
import numpy 

file_name='WCG_DataSetV1.xlsx'
df_demand = pd.read_excel(file_name,'Demand', index_col=0)
df_price = pd.read_excel(file_name,'Price', index_col=0)
df_returns = pd.read_excel(file_name,'Return', index_col=0)

df_demand.fillna(0, inplace=True) 
df_price.fillna(0, inplace=True)
df_returns.fillna(0, inplace=True)

demand = df_demand.to_numpy()
price = df_price.to_numpy()
returns = df_returns.to_numpy()

#for creating index to variables
weeks = range(len(demand[:,0])) #number of rows
rates = range(len(price[0,:])) #number of columns


#for calculating in the constraints
lengths = [1,4,8,16]
initial_capacity = 77

#create model
model = ConcreteModel()
model.x = Var(weeks,rates, domain=NonNegativeReals)
model.c = Var(weeks, domain=NonNegativeReals)

#define obj
def obj_rule(model):
    return sum(7*model.x[i,j]*price[i,j]*lengths[j] for i in weeks for j in rates)
model.rev=Objective(rule=obj_rule, sense=maximize)

#Demand constraint
def const_rule(model,i,j): #weeks,rates=>i,j means there are 200 constraints in this rule
    return model.x[i,j]<=demand[i,j]
#there is no "for" because each product will have its own demand individually
model.constr=Constraint(weeks,rates, rule=const_rule)

#Capacity constraints
def cap_inventory(model,m):
    if m == 0:
        return model.c[m] == initial_capacity
    elif m == 1:
        return model.c[m] == model.c[m-1] + sum(returns[m,j] for j in rates)
    elif 2<=m<=4:
        return model.c[m] == model.c[m-1] - sum(model.x[m-1,j] for j in rates) + model.x[m-1,0] + sum(returns[m,j] for j in rates if j >=1)  
    elif 5<=m<=8: 
        return model.c[m] == model.c[m-1] - sum(model.x[m-1,j] for j in rates) + sum(model.x[m-1,j] for j in rates if j ==0) + sum(model.x[m-4,j] for j in rates if j ==1) + sum(returns[m,j] for j in rates if j >=2)
    elif 9<=m<=16:  
        return model.c[m] == model.c[m-1] - sum(model.x[m-1,j] for j in rates) + sum(model.x[m-1,j] for j in rates if j ==0) + sum(model.x[m-4,j] for j in rates if j ==1) + sum(model.x[m-8,j] for j in rates if j ==2) + sum(returns[m,j] for j in rates if j ==3)
    elif m>=17:
        return model.c[m] == model.c[m-1] - sum(model.x[m-1,j] for j in rates) + sum(model.x[m-1,j] for j in rates if j==0) + sum(model.x[m-4,j] for j in rates if j==1) + sum(model.x[m-8,j] for j in rates if j==2) + sum(model.x[m-16,j] for j in rates if j==3)
model.inv=Constraint(weeks,rule=cap_inventory)

def cap_constr(model,i):
        return sum(model.x[i,j] for j in rates) <= model.c[i]
model.cap=Constraint(weeks, rule=cap_constr)

model.pprint()


solver = SolverFactory('glpk')
results = solver.solve(model)
if (results.solver.status == SolverStatus.ok) and (results.solver.termination_condition == TerminationCondition.optimal):
    print('total revenue is: ', model.rev())
for i in weeks:
    for j in rates:
       print('on week',i+1,':',model.x[i,j](),'container(s) are accept for renting',lengths[j], 'week(s)')


#model.display()
