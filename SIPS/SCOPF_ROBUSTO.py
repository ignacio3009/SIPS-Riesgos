# -*- coding: utf-8 -*-

import numpy as np
import xpress as xp
from openpyxl import Workbook
import os
import xlwings as xw

M = 10000
NGen = 3
NLin = 3
NBar = 3
NEle = NGen + NLin
NFal = NGen + NLin
NSce = NFal + 1 
numLin = range(NLin)
numBar = range(NBar)
numGen = range(NGen)
numFal = range(NFal)
numEle = range(NEle)
numSce = range(NSce)


PMIN =  np.array([0,30,10])
PMAX =  np.array([100,150,150])
FMAX =  np.array([50,100,100])
CV   =  np.array([5,30,120])
CR   =  np.array([10,10,10])
D =     np.array([30,60,120])
RAMPMAX = np.array([100,100,100])

A = np.ones((NGen,NSce))
B = np.ones((NLin,NSce))

for i in numGen:
    A[i,i+1] = 0
    
for i in range(NLin):
    B[i,NGen+i+1] = 0


#Define problem
model = xp.problem()

#Variables
p =     np.array([[xp.var(lb=0, vartype=xp.continuous) for s in numSce] for g in numGen])
x =     np.array([xp.var(vartype = xp.binary)for g in numGen])
f =     np.array([[xp.var(lb= -xp.infinity,vartype=xp.continuous) for s in numSce] for l in numLin])
theta = np.array([[xp.var(lb= -xp.infinity, vartype=xp.continuous) for s in numSce] for n in numBar])
rup =   np.array([xp.var(lb = 0, vartype=xp.continuous)for g in numGen])
rdown = np.array([xp.var(lb = 0, vartype=xp.continuous)for g in numGen])


#Objective
totalcost = xp.Sum (p[g,0]*CV[g] + CR[g]*(rup[g] + rdown[g]) for g in numGen) 

#Init Variables
for s in numSce:
    for g in numGen:
        model.addVariable(p[g,s])
    for l in numLin:
        model.addVariable(f[l,s])
    for n in numBar:
        model.addVariable(theta[n,s])
for g in numGen:
    model.addVariable(x[g],rup[g],rdown[g])

#Limits of generators
for g in numGen:
    model.addConstraint(p[g,0] + rup[g]<=PMAX[g]*x[g])
    model.addConstraint(p[g,0] - rdown[g]>=PMIN[g]*x[g])
    model.addConstraint(rup[g] <= RAMPMAX[g])
    model.addConstraint(rdown[g] <= RAMPMAX[g])
    

#Nodal Balancing
for s in numSce:
    model.addConstraint(p[0,s] == D[0] + f[0,s] + f[2,s])
    model.addConstraint(p[1,s] + f[0,s] == D[1] + f[1,s])
    model.addConstraint(p[2,s] + f[2,s] + f[1,s] == D[2])


for s in numSce:
    for g in numGen:
        model.addConstraint(p[g,s] <= A[g,s]*(p[g,0] + rup[g]))
        model.addConstraint(p[g,s] >= A[g,s]*(p[g,0] - rdown[g]))
    for l in numLin:
        model.addConstraint(f[l,s] <= FMAX[l]*B[l,s])
        model.addConstraint(f[l,s] >= -FMAX[l]*B[l,s])
    
    model.addConstraint(f[0,s] <= theta[0,s] - theta[1,s] + M*(1-B[0,s]))
    model.addConstraint(f[0,s] >= theta[0,s] - theta[1,s] - M*(1-B[0,s]))
    model.addConstraint(f[1,s] <= theta[1,s] - theta[2,s] + M*(1-B[1,s]))
    model.addConstraint(f[1,s] >= theta[1,s] - theta[2,s] - M*(1-B[1,s]))
    model.addConstraint(f[2,s] <= theta[0,s] - theta[2,s] + M*(1-B[2,s]))
    model.addConstraint(f[2,s] >= theta[0,s] - theta[2,s] - M*(1-B[2,s]))
    


# =============================================================================
# SOLVE AND SAVE SOLUTIONS    
# =============================================================================
model.setObjective(totalcost)
model.solve()

value = model.getObjVal()
print("---Total Cost----")
print(value)

psol =      model.getSolution(p)
fsol =      model.getSolution(f)
xsol =      model.getSolution(x)
thetasol =  model.getSolution(theta)
rupsol =    model.getSolution(rup)
rdownsol =  model.getSolution(rdown)


startcol='C'
filename = 'data.xlsx'
curFol = os.path.realpath('')
path_filename = curFol + '\\' + filename
batcol = chr(ord(startcol)+1)
rupcol = chr(ord(startcol)+2)
rdowncol = chr(ord(startcol)+3)
wb = xw.Book(path_filename)
sht = wb.sheets['ROBUSTO']
sht.range(startcol+str(3)).value = psol
sht.range(startcol+str(6)).value = fsol

sht.range(rupcol+str(10)).value = rupsol[0]
sht.range(rupcol+str(11)).value = rupsol[1]
sht.range(rupcol+str(12)).value = rupsol[2]

sht.range(rdowncol+str(10)).value = rdownsol[0]
sht.range(rdowncol+str(11)).value = rdownsol[1]
sht.range(rdowncol+str(12)).value = rdownsol[2]

sht.range(rdowncol+str(10)).value = rdownsol[0]
sht.range(rdowncol+str(11)).value = rdownsol[1]
sht.range(rdowncol+str(12)).value = rdownsol[2]

wb.save()