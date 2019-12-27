import xlwings as xw
import pandapower as pp

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
        
        
        
def evalTec(NUMNODES,PGEN,GENDATA,DEMDATA,LINDATA,BATDATA,slack_gen1,slack_gen2, filename_res,s, print_results=False):
    
    numgen = len(GENDATA)
    numlin = len(LINDATA)
    numdem = len(DEMDATA)
    # =========================================================================
    # CREATE EMPTY NET
    # =========================================================================
    net = pp.create_empty_network(f_hz=50.0, sn_mva=100)

    # =========================================================================
    # CREATE BUSES
    # =========================================================================
    for i in range(numgen):
        pp.create_bus(net,vn_kv=220, index = i, max_vm_pu=1.1, min_vm_pu=0.9)
        
    cntidx = NUMNODES
    j=-1
    
    
    for i in range(numgen):
        barcon = GENDATA[i][1]
        if(barcon>j):
            j=j+1
            pp.create_bus(net, index = cntidx+j, vn_kv=13.8, max_vm_pu=1.03, min_vm_pu=0.97)
            
    # =========================================================================
    # CREATE GENERATORS
    # =========================================================================    
    j=-1
    for i in range(numgen):
        pp.create_gen(net, bus=GENDATA[i][1], p_mw = PGEN[i], sn_mva = GENDATA[i][11], in_service=(s!=(i+1)),
                      max_p_mw = GENDATA[i][7], min_p_mw = GENDATA[i][8],max_q_mvar=GENDATA[i][9], min_q_mvar=GENDATA[i][10], controllable=False, 
                      slack=False)
        #create trafos     
        barcon = GENDATA[i][1]
        if(barcon>j):
            j=j+1
        pp.create_transformer(net, hv_bus=GENDATA[i][1], lv_bus=cntidx+j, std_type="250 MVA 220/13.8 kV")
    
    
    #set slack gen
    if(s==1):
        net.gen['slack'][slack_gen2]=True
    else:
        net.gen['slack'][slack_gen1]=True
        
    # =========================================================================
    # CREATE LINES
    # =========================================================================
    for i in range(numlin):
        fmax = LINDATA[i][1]/(3**0.5*220)
        ltype = {'typ'+str(i):{"r_ohm_per_km": LINDATA[i][5], "x_ohm_per_km": LINDATA[i][4], "c_nf_per_km": 10, "max_i_ka": fmax, "type": "ol", "qmm2":490, "alpha":4.03e-3}}
        pp.create_std_types(net,data=ltype,element='line')
        
        pp.create_line(net, from_bus=LINDATA[i][2], to_bus=LINDATA[i][3], length_km=LINDATA[i][6], std_type="typ"+str(i), in_service=(s!=numgen+1+i))
        
    # =========================================================================
    # CREATE LOADS
    # =========================================================================
    #create loads
    for i in range(numdem):
        pp.create_load(net,  bus=DEMDATA[i][0], p_mw = DEMDATA[i][1], q_mvar=DEMDATA[i][2])

    # =========================================================================
    # PRINT
    # =========================================================================

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
    
    # =========================================================================
    # SAVE
    # =========================================================================
    row = [3,13,23,33,43,53,63]
    wb = xw.Book(filename_res)
    sht = wb.sheets['Tecnicos']
    sht.range('B'+str(row[s])).value = net.res_bus
    sht.range('I'+str(row[s])).value = net.res_line
    sht.range('Z'+str(row[s])).value = net.res_trafo
    sht.range('AP'+str(row[s])).value = net.res_gen
    return 

    

        
        
# =============================================================================
# TAKE RESULTS AND SIMULATE
# =============================================================================
def solveAll(filename_data,filename_res):
    clearContents(filename_res)
    for s in range(0,7):
        print('Simulating Scenario '+str(s)+'. Please wait...')
        letter = ['B','C','D','E','F','G','H']

        wbd = xw.Book(filename_data)
        sd = wbd.sheets['data']
        
        GENDATA = sd.range('A4:L6').value
        DEMDATA = sd.range('N4:P6').value
        LINDATA = sd.range('R4:X6').value
        BATDATA = sd.range('Z4:AK6').value
        NUMNODES = 3
        
        wb = xw.Book(filename_res)
        sht = wb.sheets['Economicos']
        PGEN = sht.range(letter[s]+'2:'+letter[s]+'4').value
        slack_gen1 = 0
        slack_gen2 = 1
        evalTec(NUMNODES,PGEN,GENDATA,DEMDATA,LINDATA,BATDATA,slack_gen1,slack_gen2, filename_res,s, print_results=False)

    print('End Simulations!')
    
filename_res = 'res_reservas.xlsx'
filename_data = 'data_3_bus.xlsx'
solveAll(filename_data,filename_res)