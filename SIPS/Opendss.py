# -*- coding: utf-8 -*-
"""
Created on Wed Dec 18 14:04:47 2019

@author: ignac
"""

# =============================================================================
# import xpress as xp
# import numpy as np
# 
# x = xp.var(vartype=xp.integer, name='x1', lb=-10, ub=10)
# y = xp.var(name='x2')
# p = xp.problem(name='myexample')
# 
# p.addVariable(x,y)
# p.setObjective(x**2 + 2*y)
# p.addConstraint(x + 3*y >= 4)
# p.solve()
# print ("solution: {0} = {1}; {2} = {3}".format (x.name, p.getSolution(x), y.name, p.getSolution(y)))
# 
# 
# #==============================================================================
# # DSS
# #==============================================================================
# import win32com.client
# 
# class DSS(object):
#      def __init__(self, dssFileName):
#             # Create a new instance of the DSS
#             self.dssObj = win32.com.client.Dispatch("OpenDSSEngine.DSS")
#             # Start the DSS
#             if self.dssObj.Start(0)=false:
#                 print("DSS Failed to Start")
#             else:
#                 #Assign a variable to each of the nterface for easier access
#                 self.dssText = self.dssObj.Text
#                 self.dssCircuit = self.dssObj.ActiveCircuit
#                 self.dssSolution = self.dssCircuit.Solution
#                 
#             self.dssObj.ClearAll()
#             
#             self.dssText.Command = "Compile " + dssFileName
#             
#         def mySolve(self):
#             self.dssSolution.Solve()
#     
# if __name__ = '__main__':
#     file = getCurrentFolder + 'Red.txt'
#     myObject = DSS('')
# =============================================================================
# import win32com.client
# import matplotlib.pyplot as plt
# import pandas as pd
# import os
# import sys

# curFol = os.path.dirname(os.path.realpath(sys.argv[0]))
# #Load DSS Engine
# DSSObj = win32com.client.Dispatch("OpenDSSEngine.DSS")
# DSSText = DSSObj.Text
# DSSCircuit = DSSObj.ActiveCircuit
# DSSSolution = DSSCircuit.Solution
# DSSBus=DSSCircuit.ActiveBus
# DSSCtrlQueue=DSSCircuit.CtrlQueue
# DSSObj.Start(0)
# #Read and Compile

# filename = curFol + '\\Grid.txt'
# DSSObj.Text.Command='compile '+filename

# # data = pd.read_csv('Test_Mon_e_1.csv') 
# # values = data.values
# # plt.plot(values[:,2])

import win32com.client
import os
import sys
import numpy as np

busesNames = ['A','B','C']

def compileCircuit(PLoad, QLoad, Pgen, Qgen):
    engine=win32com.client.Dispatch("OpenDSSEngine.DSS")
    engine.Start("0")
    engine.Text.Command='clear'
    circuit=engine.ActiveCircuit
    curFol = os.path.dirname(os.path.realpath(sys.argv[0]))
    filename = curFol + '\\Grid.txt'
    engine.Text.Command = 'compile '+filename
    #Add loads
    for i in range(len(PLoad)):
        com = "new load.Load{0} bus1={1} phases=3 kV=220 kW={2} kvar={3} model=1".format(str(i+1),busesNames[i],PLoad[i],QLoad[i])
        print(com)
        engine.Text.Command = com
        
    #Add Generators
    for i in range(len(Pgen)):
        com =  "new generator.Gen{0} bus1=G{1} phases=3 kV=220 kW={2} kvar={3} model=1".format(str(i+1),busesNames[i],Pgen[i],Qgen[i])
        print(com)
        engine.Text.Command = com
        
    
    
    engine.Text.Command = 'Set controlmode = static'
    engine.Text.Command = 'Set mode = snapshot'
    engine.Text.Command = 'Solve'
    engine.Text.Command = 'Show Voltages'
    
Pgen =   np.array([30,30,40])*1000
PLoad =  np.array([20,40,40])*1000
QLoad =  np.array([0,0,0])*1000
Qgen =   np.array([0,0,0])*1000

compileCircuit(PLoad,QLoad,Pgen,Qgen)



