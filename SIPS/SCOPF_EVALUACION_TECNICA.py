# import pandapower as pp
# import matplotlib.pyplot as plt
# from pandapower.plotting.simple_plot import *


# Pgen = [100,50,60]
# Qgen = [10,10,10]

# Pcon = [50,50,50]
# Qcon = [10,10,10]

# PMAX = [300,300,300]
# QMAX = [300,300,300]
# PMIN = [0,0,0]
# QMIN = [0,0,0]

# #create empty net
# net = pp.create_empty_network()

# #create buses
# busA = pp.create_bus(net, vn_kv=220, name="Bus A",geodata=(1,2))
# busB = pp.create_bus(net, vn_kv=220, name="Bus B",geodata=(0,4))
# busC = pp.create_bus(net, vn_kv=220, name="Bus C",geodata=(2,4))


# busGA = pp.create_bus(net, vn_kv=13.8, name="Bus GA",geodata=(1,1))
# busGB = pp.create_bus(net, vn_kv=13.8, name="Bus GB",geodata=(0,5))
# busGC = pp.create_bus(net, vn_kv=13.8, name="Bus GC",geodata=(2,5))

# #create bus elements
# # pp.create_ext_grid(net, bus=bus1, vm_pu=1.02, name="Grid Connection")
# #pp.create_load(net, bus=bus3, p_mw=0.100, q_mvar=0.05, name="Load")



# #create standard type
# # trafo_data = {"sn_mva": 220, "vn_hv_kv": 220, "vn_lv_kv": 13.8, "vk_percent": 12.2,"vkr_percent": 0.25, "pfe_kw": 65,"i0_percent": 0.06,"shift_degree": 0}
# # pp.create_std_type(net, trafo_data, "160 MVA 380/110 kV", element="trafo")
# #To create data go to
# #C:\users\ignac\Anaconda3\lib\site-packages\pandapower\std_types.py

# #create branch elements
# trafoA = pp.create_transformer(net, hv_bus=busA, lv_bus=busGA, std_type="250 MVA 220/13.8 kV", name="TrafoA")
# trafoB = pp.create_transformer(net, hv_bus=busB, lv_bus=busGB, std_type="250 MVA 220/13.8 kV", name="TrafoB")
# trafoC = pp.create_transformer(net, hv_bus=busC, lv_bus=busGC, std_type="250 MVA 220/13.8 kV", name="TrafoC")

# line1 = pp.create_line(net, from_bus=busA, to_bus=busB, length_km=100, std_type="490-AL1/64-ST1A 220.0", name="Line1")
# line2 = pp.create_line(net, from_bus=busB, to_bus=busC, length_km=100, std_type="490-AL1/64-ST1A 220.0", name="Line2")
# line3 = pp.create_line(net, from_bus=busA, to_bus=busC, length_km=100, std_type="490-AL1/64-ST1A 220.0", name="Line3")


# pp.create_load(net, bus=busA, p_mw = Pcon[0], q_mvar=Qcon[0], name= "LoadA")
# pp.create_load(net, bus=busB, p_mw = Pcon[1], q_mvar=Qcon[1], name= "LoadB")
# pp.create_load(net, bus=busC, p_mw = Pcon[2], q_mvar=Qcon[2], name= "LoadC")


# pp.create_gen(net,bus=busGA, p_mw = Pgen[0], sn_mva = 100, name="Gen1", in_service=True, 
#                 max_p_mw = PMAX[0], min_p_mw=PMIN[0], max_q_mvar=QMAX[0], min_q_mvar=QMIN[0], controllable=True, 
#                 slack=True)

# pp.create_gen(net,bus=busGB, p_mw = Pgen[1], sn_mva = 100, name="Gen2", in_service=True, 
#                 max_p_mw = PMAX[1], min_p_mw=PMIN[1], max_q_mvar=QMAX[1], min_q_mvar=QMIN[1], controllable=True)

# pp.create_gen(net,bus=busGC, p_mw = Pgen[2], sn_mva = 100, name="Gen3", in_service=True, 
#                 max_p_mw = PMAX[2], min_p_mw=PMIN[2], max_q_mvar=QMAX[2], min_q_mvar=QMIN[2], controllable=True)



# #Run power flow
# pp.runpp(net)
# print(net.res_bus)
# simple_plot(net)

# simple_plot(net,library='networkx')




# =============================================================================
# 
# =============================================================================
# import pandapower as pp
# #create empty net
# net = pp.create_empty_network() 

# #create buses
# b1 = pp.create_bus(net, vn_kv=20., name="Bus 1")
# b2 = pp.create_bus(net, vn_kv=0.4, name="Bus 2")
# b3 = pp.create_bus(net, vn_kv=0.4, name="Bus 3")

# #create bus elements
# pp.create_ext_grid(net, bus=b1, vm_pu=1.02, name="Grid Connection")
# pp.create_load(net, bus=b3, p_mw=0.1, q_mvar=0.05, name="Load")

# #create branch elements
# tid = pp.create_transformer(net, hv_bus=b1, lv_bus=b2, std_type="0.4 MVA 20/0.4 kV", name="Trafo")
# pp.create_line(net, from_bus=b2, to_bus=b3, length_km=0.1, name="Line",std_type="NAYY 4x50 SE")  

# pp.runpp(net)
# print(net.res_bus)

# =============================================================================
# =============================================================================
# # OTRA VERSION
# =============================================================================
# =============================================================================

from openpyxl import Workbook
import os
import xlwings as xw
import pandapower as pp
import pandapower.plotting.simple_plot

def evaluacionTecEco(Pgen,Qgen,Pcon,Qcon,PMAX,QMAX,PMIN,QMIN, filename, scenario, print_results=False):
    # from pandapower.plotting.simple_plot import *
    #create empty net
    net = pp.create_empty_network()
    
    #create buses    
    busGA = pp.create_bus(net, vn_kv=13.8, name="Bus GA",geodata=(1,1))
    busGB = pp.create_bus(net, vn_kv=13.8, name="Bus GB",geodata=(0,5))
    busGC = pp.create_bus(net, vn_kv=13.8, name="Bus GC",geodata=(2,5))
    
    busA = pp.create_bus(net, vn_kv=220, name="Bus A",geodata=(1,2))
    busB = pp.create_bus(net, vn_kv=220, name="Bus B",geodata=(0,4))
    busC = pp.create_bus(net, vn_kv=220, name="Bus C",geodata=(2,4))
    #create bus elements
    # pp.create_ext_grid(net, bus=bus1, vm_pu=1.02, name="Grid Connection")
    #pp.create_load(net, bus=bus3, p_mw=0.100, q_mvar=0.05, name="Load")
    
    
    
    #create standard type
    # trafo_data = {"sn_mva": 220, "vn_hv_kv": 220, "vn_lv_kv": 13.8, "vk_percent": 12.2,"vkr_percent": 0.25, "pfe_kw": 65,"i0_percent": 0.06,"shift_degree": 0}
    # pp.create_std_type(net, trafo_data, "160 MVA 380/110 kV", element="trafo")
    #To create data go to
    #C:\users\ignac\Anaconda3\lib\site-packages\pandapower\std_types.py
    
    #create branch elements
    pp.create_transformer(net, hv_bus=busA, lv_bus=busGA, std_type="250 MVA 220/13.8 kV", name="TrafoA")
    pp.create_transformer(net, hv_bus=busB, lv_bus=busGB, std_type="250 MVA 220/13.8 kV", name="TrafoB")
    pp.create_transformer(net, hv_bus=busC, lv_bus=busGC, std_type="250 MVA 220/13.8 kV", name="TrafoC")
    
    pp.create_line(net, from_bus=busA, to_bus=busB, length_km=100, std_type="490-AL1/64-ST1A 220.0", name="Line1")
    pp.create_line(net, from_bus=busB, to_bus=busC, length_km=100, std_type="490-AL1/64-ST1A 220.0", name="Line2")
    pp.create_line(net, from_bus=busA, to_bus=busC, length_km=100, std_type="490-AL1/64-ST1A 220.0", name="Line3")
    
    
    pp.create_load(net, bus=busA, p_mw = Pcon[0], q_mvar=Qcon[0], name= "LoadA")
    pp.create_load(net, bus=busB, p_mw = Pcon[1], q_mvar=Qcon[1], name= "LoadB")
    pp.create_load(net, bus=busC, p_mw = Pcon[2], q_mvar=Qcon[2], name= "LoadC")
    
    
    pp.create_gen(net,bus=busGA, p_mw = Pgen[0], sn_mva = 100, name="Gen1", in_service=True, 
                    max_p_mw = PMAX[0], min_p_mw=PMIN[0], max_q_mvar=QMAX[0], min_q_mvar=QMIN[0], controllable=True, 
                    slack=True)
    
    pp.create_gen(net,bus=busGB, p_mw = Pgen[1], sn_mva = 100, name="Gen2", in_service=True, 
                    max_p_mw = PMAX[1], min_p_mw=PMIN[1], max_q_mvar=QMAX[1], min_q_mvar=QMIN[1], controllable=True)
    
    pp.create_gen(net,bus=busGC, p_mw = Pgen[2], sn_mva = 100, name="Gen3", in_service=True, 
                    max_p_mw = PMAX[2], min_p_mw=PMIN[2], max_q_mvar=QMAX[2], min_q_mvar=QMIN[2], controllable=True)
    
    pp.runpp(net)
    if(print_results):
        print('----- Buses Results -----')
        print(net.res_bus)
        print('----- Lines Results -----')
        print(net.res_line)
        print('----- Transformers Results -----')
        print(net.res_trafo)
        print('----- Generators Results -----')
        print(net.res_gen)
    
    row = [3,13,23,33,43,53,63]
    wb = xw.Book(filename)
    sht = wb.sheets['Tecnicos']
    sht.range('B'+str(row[scenario])).value = net.res_bus
    sht.range('I'+str(row[scenario])).value = net.res_line
    sht.range('Z'+str(row[scenario])).value = net.res_trafo
    sht.range('AP'+str(row[scenario])).value = net.res_gen
    return 

def clearContents(filename):
    wb = xw.Book(filename)
    sht = wb.sheets['Tecnicos']
    rows = [3,13,23,33,43,53,63]
    rowe = [9,19,29,39,49,59,69]
    for i,j in zip(rows,rowe):
        sht.range('B'+str(i)+':F'+str(j)).clear_contents()
        sht.range('I'+str(i)+':W'+str(j)).clear_contents()
        sht.range('Z'+str(i)+':AM'+str(j)).clear_contents()
        sht.range('AP'+str(i)+':AT'+str(j)).clear_contents()
        
        
