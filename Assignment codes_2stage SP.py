#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Fri Jul 21 10:31:47 2023

@author: yushi
"""

import pandas as pd
import numpy as np
import gurobipy as gp
from gurobipy import GRB



# import industrial case data
pro_Capa = '/Users/yushi/Documents/WBS Doctoral Materials/Year 1/Term 3/ODM/Assignment/production capacity.xlsx'
trans_Capa = '/Users/yushi/Documents/WBS Doctoral Materials/Year 1/Term 3/ODM/Assignment/transportation capacity.xlsx'
processing_time = '/Users/yushi/Documents/WBS Doctoral Materials/Year 1/Term 3/ODM/Assignment/processing time.xlsx'
inventory_cost = '/Users/yushi/Documents/WBS Doctoral Materials/Year 1/Term 3/ODM/Assignment/unit inventory cost.xlsx'
production_cost = '/Users/yushi/Documents/WBS Doctoral Materials/Year 1/Term 3/ODM/Assignment/unit production cost.xlsx'
transportation_cost = '/Users/yushi/Documents/WBS Doctoral Materials/Year 1/Term 3/ODM/Assignment/unit transportation cost.xlsx'
demand_scenario = '/Users/yushi/Documents/WBS Doctoral Materials/Year 1/Term 3/ODM/Assignment/uncertain demand.xlsx'

# read data & transform data
procapacity = pd.read_excel(pro_Capa, index_col=0)
transcapacity = pd.read_excel(trans_Capa, index_col=0)
procetim = pd.read_excel(processing_time, index_col=0)
invcost = pd.read_excel(inventory_cost, index_col=0)
procost = pd.read_excel(production_cost, index_col=0)
transcost = pd.read_excel(transportation_cost, index_col=0)
dmd_prob = pd.read_excel(demand_scenario, sheet_name = 'Sheet1', index_col=0)

# create a gurobipy stochastic model
model = gp.Model('stochastic version')

# generate decision variables (DV): model.addVar()
plant = list(range(1,9))
period = list(range(1,9))
product = list(range(1,3))
scenario = list(range(1,65))

## production amounts of product
production = {}
for k in product:
    for i in plant:
        for t in period:
            production[(k, i, t)] = model.addVar(vtype=GRB.CONTINUOUS, lb=0, name=f"production_{k}_{i}_{t}")

## amount of end of period inventory of product
finivt = {}
for k in product:
    for i in plant:
        for t in period:
            for s in scenario:
                finivt[(k, i, t, s)] = model.addVar(vtype=GRB.CONTINUOUS, lb=0, name=f"finivt_{k}_{i}_{t}_{s}")

## amount of end of period inventory of semi-finished product
semiivt = {}
for k in product:
    for i in plant:
        for t in period:
            for s in scenario:
                semiivt[(k, i, t, s)] = model.addVar(vtype=GRB.CONTINUOUS, lb=0, name=f"semiivt_{k}_{i}_{t}_{s}")

## backorder amounts of finished product
backorder = {}
for k in product:
        for t in period:
            for s in scenario:
                backorder[(k, t, s)] = model.addVar(vtype=GRB.CONTINUOUS, lb=0, name=f"backorder_{k}_{t}_{s}")
                
## amounts of product transported between plants
middletrans = {}
for k in product:
        for t in period:
            for i in plant:
                for j in plant:
                    middletrans[(k, t, i, j)] = model.addVar(vtype=GRB.CONTINUOUS, lb=0, name=f"middletrans_{k}_{t}_{i}_{j}")

## amounts of product transported from the last plant to customer
endtrans = {}
for k in product:
    for t in period:
        for s in scenario:
            endtrans[(k, t, s)] = model.addVar(vtype=GRB.CONTINUOUS, lb=0, name=f"endtrans_{k}_{t}_{s}")
            
## amounts of product received by each plant
rcvamount = {}
for k in product:
    for i in plant:
        for t in period:
                rcvamount[(k, i, t)] = model.addVar(vtype=GRB.CONTINUOUS, lb=0, name=f"rcvamount_{k}_{i}_{t}")

# formulate the objective function: model.setObjective()
## 1st stage costs
Cost1_pro = sum(procost.iat[k-1, i-1] * production[(k, i , t)] for k in product for i in plant for t in period)
Cost1_trans = sum(transcost.iat[0, 1] * middletrans[(k, t, 1, 2)] + transcost.iat[1, 2] * middletrans[(k, t, 2, 3)] + transcost.iat[2, 3] * middletrans[(k, t, 3, 4)] + transcost.iat[2, 4] * middletrans[(k, t, 3, 5)] + transcost.iat[2, 5] * middletrans[(k, t, 3, 6)] + transcost.iat[2, 6] * middletrans[(k, t, 3, 7)] + transcost.iat[4, 7] * middletrans[(k, t, 5, 8)] + transcost.iat[5, 7] * middletrans[(k, t, 6, 8)]+ transcost.iat[6, 7] * middletrans[(k, t, 7, 8)] for k in product for t in period )
## 2nd stage costs
Cost2_inv = sum(dmd_prob.iat[s-1, -1] * (invcost.iat[0, i-1] * (finivt[(k, i, t, s)] + semiivt[(k, i, t, s)])) for k in product for i in plant for t in period for s in scenario)
Cost2_endtrans = sum(dmd_prob.iat[s-1, -1] * (1/2 * endtrans[(k, t, s)]) for k in product for t in period for s in scenario)
## backorder cost: due to the lack of industrial case data provided by the research paper, I set simulated data of the unit backorder cost according to the averaage demand level: product 1 - 0.3, product 2 - 0.1
Cost2_backorder = sum(dmd_prob.iat[s-1, -1] * (0.3 * backorder[(1, t, s)] + 0.1 * backorder[(2, t, s)]) for t in period for s in scenario)
## total expected revenue: unit sales price of product 1 is 2.21, product 2 is 2.82
## transported amount to customer = selling amount <= demand
Revenue = sum(dmd_prob.iat[s-1, -1] * (2.21 * endtrans[(1, t, s)] + 2.82 * endtrans[(2, t, s)]) for t in period for s in scenario)
## Set the objective function
model.setObjective(Revenue - (Cost2_endtrans + Cost2_inv + Cost2_backorder) - (Cost1_pro + Cost1_trans), GRB.MAXIMIZE)


# set the model constraints
## inventroy balance of production stage (excluding the last stage)
for k in product:
        for s in scenario:
            for i in range(1, 3):
                model.addConstr(finivt[(k, i, 1, s)] <= production[(k, i, 1)] - middletrans[(k, 1, i, i + 1)]) 
                model.addConstr(finivt[(k, i, 1, s)] >= production[(k, i, 1)] - middletrans[(k, 1, i, i + 1)])
                
for k in product:
    for t in range (2, 9):
        for s in scenario:
            for i in range(1, 3):
                model.addConstr(finivt[(k, i, t, s)] <= finivt[(k, i, t-1, s)] + production[(k, i, t)] - middletrans[(k, t, i, i + 1)]) 
                model.addConstr(finivt[(k, i, t, s)] >= finivt[(k, i, t-1, s)] + production[(k, i, t)] - middletrans[(k, t, i, i + 1)]) 


for k in product:
        for s in scenario:
                model.addConstr(finivt[(k, 3, 1, s)] <= production[(k, 3, 1)] - middletrans[(k, 1, 3, 4)] - middletrans[(k, 1, 3, 5)] - middletrans[(k, 1, 3, 6)] - middletrans[(k, 1, 3, 7)]) 
                model.addConstr(finivt[(k, 3, 1, s)] >= production[(k, 3, 1)] - middletrans[(k, 1, 3, 4)] - middletrans[(k, 1, 3, 5)] - middletrans[(k, 1, 3, 6)] - middletrans[(k, 1, 3, 7)]) 


for k in product:
    for t in range (2, 9):
        for s in scenario:
                model.addConstr(finivt[(k, 3, t, s)] <= finivt[(k, 3, t-1, s)] + production[(k, 3, t)] - middletrans[(k, t, 3, 4)] - middletrans[(k, t, 3, 5)] - middletrans[(k, t, 3, 6)] - middletrans[(k, t, 3, 7)]) 
                model.addConstr(finivt[(k, 3, t, s)] >= finivt[(k, 3, t-1, s)] + production[(k, 3, t)] - middletrans[(k, t, 3, 4)] - middletrans[(k, t, 3, 5)] - middletrans[(k, t, 3, 6)] - middletrans[(k, t, 3, 7)]) 


for k in product:
        for s in scenario:
            for i in range(4, 8):
                model.addConstr(finivt[(k, i, 1, s)] <= production[(k, i, 1)] - middletrans[(k, 1, i, 8)]) 
                model.addConstr(finivt[(k, i, 1, s)] >= production[(k, i, 1)] - middletrans[(k, 1, i, 8)]) 

                
for k in product:
    for t in range (2, 9):
        for s in scenario:
            for i in range(4, 8):
                model.addConstr(finivt[(k, i, t, s)] <= finivt[(k, i, t-1, s)] + production[(k, i, t)] - middletrans[(k, t, i, 8)]) 
                model.addConstr(finivt[(k, i, t, s)] >= finivt[(k, i, t-1, s)] + production[(k, i, t)] - middletrans[(k, t, i, 8)]) 


## inventory balance in the last production stage
for k in product:
    for s in scenario:
        model.addConstr(finivt[(k, 8, 1, s)] <= production[(k, 8, t)] - endtrans[(k, 1, s)])
        model.addConstr(finivt[(k, 8, 1, s)] >= production[(k, 8, t)] - endtrans[(k, 1, s)])

for k in product:
    for s in scenario:
        for t in range(2, 9):
            model.addConstr(finivt[(k, 8, t, s)] <= finivt[(k, 8, t-1, s)] + production[(k, 8, t)] - endtrans[(k, t, s)])
            model.addConstr(finivt[(k, 8, t, s)] >= finivt[(k, 8, t-1, s)] + production[(k, 8, t)] - endtrans[(k, t, s)])
            
            
##inventory balance for the semi-finished products
for k in product:
    for s in scenario:
        model.addConstr(semiivt[(k, i, 1, s)] <= rcvamount[(k, i, 1)] - production[(k, i, 1)])
        model.addConstr(semiivt[(k, i, 1, s)] >= rcvamount[(k, i, 1)] - production[(k, i, 1)])

for k in product:
    for s in scenario:
        for t in range(2, 9):
            model.addConstr(semiivt[(k, i, t, s)] <= semiivt[(k, i, t-1 , s)] + rcvamount[(k, i, t)] - production[(k, i, t)])
            model.addConstr(semiivt[(k, i, t, s)] >= semiivt[(k, i, t-1 , s)] + rcvamount[(k, i, t)] - production[(k, i, t)])

## balance equation for shortage: backorder and uncertain demand
for s in scenario:
    for t in range(2,9):
        model.addConstr(backorder[(1, t, s)] <= backorder[(1, t-1, s)] + dmd_prob.iat[s-1, t-1] - endtrans[(1, t, s)])
        model.addConstr(backorder[(1, t, s)] >= backorder[(1, t-1, s)] + dmd_prob.iat[s-1, t-1] - endtrans[(1, t, s)])

for s in scenario:
    for t in range(2,9):
        model.addConstr(backorder[(2, t, s)] <= backorder[(2, t-1, s)] + dmd_prob.iat[s-1, t+7] - endtrans[(2, t, s)])
        model.addConstr(backorder[(2, t, s)] >= backorder[(2, t-1, s)] + dmd_prob.iat[s-1, t+7] - endtrans[(2, t, s)])

## balance for transporation between different production plants ------ omitted constraint 1

# inequality constraint
## production capacity constraint
for i in plant:
    for t in period:
        practicalcapacity = sum((procetim.iat[k-1, i-1] * production[(k, i, t)]) for k in product)
        model.addConstr(practicalcapacity <= procapacity.iat[i-1, t-1])

## storage capacity ------ omitted constraint 2
## transportation capacity (between different plants)
for i in plant:
    for j in plant:
        for t in period:
            if j > i:
                practicaltransamount = sum(middletrans[(k, t, i, j)] for k in product)
                model.addConstr(practicaltransamount <= transcapacity.iat[i-1, j-1])
                
## transportation capacity (between the last plant and customer) - new-added constraint
for t in period:
        for s in scenario:
            endtransamount = sum(endtrans[(k, t, s)] for k in product)
            model.addConstr(endtransamount <= transcapacity.iat[7, 8])

## non-negativity constraints are met in the setting of DVs

# solve the model
model.optimize()

## print the solution and optimized value
if model.status == GRB.OPTIMAL:
    print("Optimization successful.")

    # Print optimal objective value
    print("Optimal value:", model.objVal)
else:
    print("Optimization failed.")
