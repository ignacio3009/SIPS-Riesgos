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
import win32com.client
import matplotlib.pyplot as plt
import pandas as pd
import os

#curFol = os.getcwd()
##Load DSS Engine
#DSSObj = win32com.client.Dispatch("OpenDSSEngine.DSS")
#DSSText = DSSObj.Text
#DSSCircuit = DSSObj.ActiveCircuit
#DSSSolution = DSSCircuit.Solution
#DSSBus=DSSCircuit.ActiveBus
#DSSCtrlQueue=DSSCircuit.CtrlQueue
#DSSObj.Start(0)
##Read and Compile

filename = curFol + '\\Red.txt'
DSSObj.Text.Command='compile '+filename

data = pd.read_csv('Test_Mon_a_1.csv') 
values = data.values
print(values.shape)

