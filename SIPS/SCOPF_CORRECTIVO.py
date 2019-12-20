# -*- coding: utf-8 -*-

import numpy as np
import xpress as xp
import saveresults as sv
from openpyxl import Workbook
import os
import xlwings as xw
from openpyxl import Workbook
import os
import xlwings as xw
import pandapower as pp
import pandapower.plotting.simple_plot
from SCOPF_EVALUACION_TECNICA import*


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
RAMPMAX = np.array([50,50,50])
PROBSCEN = np.array([0,0.01,0.01,0.01,0.01,0.01, 0.01])
CRUP = np.array([100,20,10])
CRDOWN = np.array([80,50,10])

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
rup =   np.array([[xp.var(lb = 0, vartype=xp.continuous)for s in numSce] for g in numGen])
rdown = np.array([[xp.var(lb = 0, vartype=xp.continuous)for s in numSce]for g in numGen])

#Objective
totalcost = xp.Sum (p[g,0]*CV[g] for g in numGen) + xp.Sum(PROBSCEN[s]*xp.Sum(p[g,0]*CV[g] + rup[g,s]*CRUP[g] + rdown[g,s]*CRDOWN[g] for g in numGen) for s in numSce) 


#Init Variables
for s in numSce:
    for g in numGen:
        model.addVariable(p[g,s])
        model.addVariable(rup[g,s])
        model.addVariable(rdown[g,s])
    for l in numLin:
        model.addVariable(f[l,s])
    for n in numBar:
        model.addVariable(theta[n,s])
for g in numGen:
    model.addVariable(x[g])


#Limits of generators
for s in numSce:
    for g in numGen:
        model.addConstraint(p[g,0] + rup[g,s]<=PMAX[g]*x[g])
        model.addConstraint(p[g,0] - rdown[g,s]>=PMIN[g]*x[g])
        model.addConstraint(rup[g,s] <= RAMPMAX[g])
        model.addConstraint(rdown[g,s] <= RAMPMAX[g])
    

#Nodal Balancing
for s in numSce:
    model.addConstraint(p[0,s] == D[0] + f[0,s] + f[2,s])
    model.addConstraint(p[1,s] + f[0,s] == D[1] + f[1,s])
    model.addConstraint(p[2,s] + f[2,s] + f[1,s] == D[2])


for s in numSce:
    for g in numGen:
        model.addConstraint(p[g,s] <= A[g,s]*(p[g,0] + rup[g,s]))
        model.addConstraint(p[g,s] >= A[g,s]*(p[g,0] - rdown[g,s]))
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

costsol = model.getObjVal()
print("---Total Cost----")
print(costsol)

psol =      model.getSolution(p)
fsol =      model.getSolution(f)
xsol =      model.getSolution(x)
thetasol =  model.getSolution(theta)
rupsol =    model.getSolution(rup)
rdownsol =  model.getSolution(rdown)



filename = 'RES_CORRECTIVO.xlsx'
wb = xw.Book(filename)
sht = wb.sheets['Economicos']


sht.range('B2').value = psol
sht.range('B5').value = fsol
sht.range('B8').value = rupsol
sht.range('B11').value = rdownsol
sht.range('K2').value = costsol

for i in range(1,4):
    sht.range('I'+str(i+1)).value = xsol[i-1]
    # sht.range('L'+str(i+1)).value = rupsol[i-1]
    # sht.range('M'+str(i+1)).value = rdownsol[i-1]
wb.save()


# =============================================================================
# Tecnical Simulations
# =============================================================================
Qgen = np.array([0,0,0])
Qcon= np.array([0,0,0])
QMAX = np.array([100,100,100])
QMIN = np.array([0,0,0])
clearContents(filename)
for s in numSce:
    evaluacionTecEco(psol[:,0],Qgen,D,Qcon,PMAX,QMAX,PMIN,QMIN,filename,s,print_results=False)
wb.save()                    