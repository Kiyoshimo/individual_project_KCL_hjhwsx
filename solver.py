import pyomo.environ as pyo
import pyomo.kernel as pmo
import pandas as pd
from pyomo.environ import *

#-----------------------
# declare Param 设置参数
#-----------------------
provider_num = 468 #供应商数量
power_station_num = 14 #电厂数量
fuel_num = 6 #燃料种类
distance = pd.read_excel('data2.xlsx', sheet_name='距离').iloc[0:, 2:] 
capacity = pd.read_excel('data2.xlsx', sheet_name='Sheet3').iloc[:,:]   
demand = pd.read_excel('data2.xlsx', sheet_name='Sheet2').iloc[1:, 4:]  
Distance, Capacity, Demand = {}, {}, {}
fuel_price = [0, 49, 49, 49, 70, 55, 53] # 生物质价格 biomass-price[lue1,]
#mwe = [11, 10, 10, 10, 10.5, 6.3, 10, 4.5, 20, 8.4, 9, 4.5, 4.8, 15] #发电厂瓦数


def generate_d():
    for i in range(distance.shape[0]):
        for j in range(distance.shape[1]):
            Distance[i+1, j+1] = distance.iloc[i, j]


def generate_c():
    for i in range(capacity.shape[0]):
        for j in range(capacity.shape[1]):
            Capacity[i+1, j+1] = capacity.iloc[i, j]


def generate_D():
    for i in range(demand.shape[0]):
        for j in range(demand.shape[1]):
            Demand[i+1, j+1] = demand.iloc[i, j]

generate_d()
generate_c()
generate_D()

model = pyo.ConcreteModel()

model.X = pyo.RangeSet(1, provider_num) #燃料供应商
model.Y = pyo.RangeSet(1, power_station_num) #发电站
model.F = pyo.RangeSet(1, fuel_num) #燃料种类
model.CX = pyo.Var(model.X, model.F, domain=pmo.Binary)
model.CY = pyo.Var(model.Y, model.F, domain=pmo.Binary)
model.N = pyo.Var(model.X, model.Y, model.F, domain=pyo.NonNegativeReals) #送的货
model.D = pyo.Param(model.X, model.Y, initialize=Distance) #距离
model.Demand = pyo.Param(model.Y, model.F, initialize=Demand) #发电厂各个燃料种类的需求量
model.Capacity = pyo.Param(model.X, model.F, initialize=Capacity) #供应商6种燃料的产量

#-----------------------
#setting model 模型设置
#-----------------------

def DemandConstraint(model):
    return [sum([model.N[x, y, f] for x in range(1, provider_num+1)]) >= model.Demand[y, f] * model.CY[y, f]
            for f in range(1, fuel_num+1) for y in range(1, power_station_num+1)]
def CapacityConstraint(model):
    return [sum([model.N[x, y, f] for y in range(1, power_station_num+1)]) <= model.Capacity[x, f] * model.CX[x, f]
            for f in range(1, fuel_num+1) for x in range(1, provider_num+1)]
def DemandOne(model):
    return [sum([model.CY[y, f] for f in range(1, fuel_num + 1)]) == 1
            for y in range(1, power_station_num + 1)]
def CapacityOne(model):
    return [sum([model.CX[x, f] for f in range(1, fuel_num + 1)]) == 1
            for x in range(1, provider_num + 1)]
def N(model):
    return [model.N[x, y, f] <= model.CX[x,f]*model.CY[y,f]*100000 for y in range(1, power_station_num+1)
          for f in range(1, fuel_num+1) for x in range(1, provider_num+1)]


model.NC = pyo.ConstraintList(rule=N)
model.DemandConstraint = pyo.ConstraintList(rule=DemandConstraint)
model.CapacityConstraint = pyo.ConstraintList(rule=CapacityConstraint)
model.DemandOne = pyo.ConstraintList(rule=DemandOne)
model.CapacityOne = pyo.ConstraintList(rule=CapacityOne)


def objective(model):
    CostT = sum(model.N[x, y, f]/19.3*model.D[x, y]*(32/100*1.3+1.2) for f in range(1, fuel_num+1) \
                for y in range(1, power_station_num+1) \
                for x in range(1, provider_num+1))

    return CostT

model.OBJ = pyo.Objective(rule=objective)
#-----------------------
# optimizate 优化求解
#-----------------------
print("开始优化等等等")
SolverFactory('ipopt', executable='C:/Users/hjhws/Desktop/ipopt').solve(model).write()#求解器
#SolverFactory('glpk', executable='C:/Windows/System32/glpk-4.65/w64/glpsol').solve(model).write()
counter = 0
res = pd.DataFrame(columns=['land-grip', 'power station', 'biomass', 'N']) #【供应土地id，发电厂id，燃料种类，运输量】
for i in range(1, provider_num+1):
  for j in range(1, power_station_num+1):
    for k in range(1, fuel_num+1):
      val = model.N[i,j,k]()
      if val < 1:
        val = 0
      res.loc[counter] = [i, j, k, val]
      counter += 1

res.to_csv('res.csv')