# -*- coding: utf-8 -*-
"""
Created on Fri Dec 20 17:06:00 2019

@author: ignac
"""

import pandapower as pp
net = pp.create_empty_network() 

#create buses
b1 = pp.create_bus(net, vn_kv=20., name="Bus 1")
b2 = pp.create_bus(net, vn_kv=0.4, name="Bus 2")
b3 = pp.create_bus(net, vn_kv=0.4, name="Bus 3")

#create bus elements
pp.create_ext_grid(net, bus=b1, vm_pu=1.02, name="Grid Connection")
pp.create_load(net, bus=b3, p_mw=0.1, q_mvar=0.05, name="Load")

#create branch elements
tid = pp.create_transformer(net, hv_bus=b1, lv_bus=b2,  std_type="0.4 MVA 20/0.4 kV", name="Trafo")
pp.create_line(net, from_bus=b2, to_bus=b3, length_km=0.1, name="Line",std_type="NAYY 4x50 SE")  

pp.runpp(net,numba=False)
print(net.res_bus)

# =============================================================================
# EXAMPLE
# =============================================================================
# import SCOPF_EVALUACION_TECNICA as et
# import numpy as np
# PMIN =  np.array([0,30,10]).tolist()
# PMAX =  np.array([100,150,150]).tolist()
# FMAX =  np.array([50,100,100]).tolist()
# CV   =  np.array([5,30,120]).tolist()
# CR   =  np.array([10,10,10]).tolist()
# D =     np.array([30,60,120]).tolist()
# RAMPMAX = np.array([50,50,50]).tolist()
# PROBSCEN = np.array([0,0.01,0.01,0.01,0.01,0.01, 0.01]).tolist()
# CRUP = np.array([100,20,10]).tolist()
# CRDOWN = np.array([80,50,10]).tolist()

# filename = 'RES_CORRECTIVO.xlsx'
# # wb = xw.Book(filename)
# Qgen = np.array([0,0,0]).tolist()
# Qcon= np.array([0,0,0]).tolist()
# QMAX = np.array([100,100,100]).tolist()
# QMIN = np.array([0,0,0]).tolist()

# psol = np.array([90,60,90]).tolist()
# et.evaluacionTec(psol,Qgen,D,Qcon,PMAX,QMAX,PMIN,QMIN,filename,2,print_results=True)