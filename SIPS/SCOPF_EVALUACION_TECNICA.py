import xlwings as xw
import pandapower as pp
import pandapower.plotting.simple_plot

def evaluacionTec(Pgen,Qgen,Pcon,Qcon,PMAX,QMAX,PMIN,QMIN, filename, s, print_results=False):
    # from pandapower.plotting.simple_plot import *
    #create empty net
    
    net = pp.create_empty_network()
    
    #create buses    
    pp.create_bus(net, index = 3, vn_kv=13.8, name="Bus GA",geodata=(1,1))
    pp.create_bus(net, index = 4, vn_kv=13.8, name="Bus GB",geodata=(0,5))
    pp.create_bus(net, index = 5, vn_kv=13.8, name="Bus GC",geodata=(2,5))
    
    pp.create_bus(net, index = 0, vn_kv=220, name="Bus A",geodata=(1,2))
    pp.create_bus(net, index = 1, vn_kv=220, name="Bus B",geodata=(0,4))
    pp.create_bus(net, index = 2, vn_kv=220, name="Bus C",geodata=(2,4))

    #Create Generators
    pp.create_gen(net,bus=0, p_mw = Pgen[0], sn_mva = 100, name="Gen1", in_service=(s!=1), 
                    max_p_mw = PMAX[0], min_p_mw=PMIN[0], max_q_mvar=QMAX[0], min_q_mvar=QMIN[0], controllable=True, 
                    slack=True)
    
    if(s==1):
        pp.create_gen(net,bus=1, p_mw = Pgen[1], sn_mva = 100, name="Gen2", in_service=True, 
                    max_p_mw = PMAX[1], min_p_mw=PMIN[1], max_q_mvar=QMAX[1], min_q_mvar=QMIN[1], controllable=True,
                    slack=True)
    else:
        pp.create_gen(net,bus=1, p_mw = Pgen[1], sn_mva = 100, name="Gen2", in_service=True, 
                    max_p_mw = PMAX[1], min_p_mw=PMIN[1], max_q_mvar=QMAX[1], min_q_mvar=QMIN[1], controllable=True, 
                    slack=False)
    
    
    pp.create_gen(net,bus=2, p_mw = Pgen[2], sn_mva = 100, name="Gen3", in_service=(s!=3), 
                    max_p_mw = PMAX[2], min_p_mw=PMIN[2], max_q_mvar=QMAX[2], min_q_mvar=QMIN[2], controllable=True)
    
    #create branch elements
    pp.create_transformer(net, hv_bus=0, lv_bus=3, std_type="250 MVA 220/13.8 kV", name="TrafoA")
    pp.create_transformer(net, hv_bus=1, lv_bus=4, std_type="250 MVA 220/13.8 kV", name="TrafoB")
    pp.create_transformer(net, hv_bus=2, lv_bus=5, std_type="250 MVA 220/13.8 kV", name="TrafoC")
    
    pp.create_line(net, from_bus=0, to_bus=1, length_km=100, std_type="490-AL1/64-ST1A 220.0", name="Line1",in_service=(s!=4))
    pp.create_line(net, from_bus=1, to_bus=2, length_km=100, std_type="490-AL1/64-ST1A 220.0", name="Line2",in_service=(s!=5))
    pp.create_line(net, from_bus=0, to_bus=2, length_km=100, std_type="490-AL1/64-ST1A 220.0", name="Line3",in_service=(s!=6))
    
    #Create Loads
    pp.create_load(net, bus=0, p_mw = Pcon[0], q_mvar=Qcon[0], name= "LoadA")
    pp.create_load(net, bus=1, p_mw = Pcon[1], q_mvar=Qcon[1], name= "LoadB")
    pp.create_load(net, bus=2, p_mw = Pcon[2], q_mvar=Qcon[2], name= "LoadC")
    
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
    sht.range('B'+str(row[s])).value = net.res_bus
    sht.range('I'+str(row[s])).value = net.res_line
    sht.range('Z'+str(row[s])).value = net.res_trafo
    sht.range('AP'+str(row[s])).value = net.res_gen
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
        
        
        
        
# =============================================================================
# EXAMPLE
# =============================================================================
# import numpy as np
def solveAll():
    filename = 'RES_CORRECTIVO.xlsx'
    clearContents(filename)
    for s in range(0,7):
        print('Simulating Scenario '+str(s)+'. Please wait...')
        letter = ['B','C','D','E','F','G','H']
        PMIN =  [0,30,10]
        PMAX =  [100,150,150]
        # FMAX =  [50,100,100]
        # CV   =  [5,30,120]
        # CR   =  [10,10,10]
        D    =  [30,60,120]
        # RAMPMAX = [50,50,50]
        # PROBSCEN = [0,0.01,0.01,0.01,0.01,0.01, 0.01]
        # CRUP = [100,20,10]
        # CRDOWN = [80,50,10]
        
    
        Qgen = [0,0,0]
        Qcon=  [0,0,0]
        QMAX = [100,100,100]
        QMIN = [0,0,0]
        
        wb = xw.Book(filename)
        sht = wb.sheets['Economicos']
        psol = sht.range(letter[s]+'2:'+letter[s]+'4').value
        wb = xw.Book(filename)
        evaluacionTec(psol,Qgen,D,Qcon,PMAX,QMAX,PMIN,QMIN,filename,s,print_results=False)
        wb.save()
    
    print('End Simulations!')
    
solveAll()