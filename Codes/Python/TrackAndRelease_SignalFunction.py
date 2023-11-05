# -*- coding: utf-8 -*-
"""
Produced on Tue Mar 15 21:28:44 2022
@author: KUY3IB
"""
import pandas as pd
import numpy as np
import re

def signal_check(config_path): # pass while debugging
    
    # this function gets the user x selections about signals in config file
    # for individual projects, gets individual selected signal data to use
    # after getting these datas, checks some areas to inspect and warn the
    # user if the user did not select the signal but it is used in an area
    
    # gives back meta corpse signal labels and mismatch signal labels
    # go to the end of the script for more information and parameters
    
    # %% Signal Sheet
    # hard coded - indicate the config excel file full path
    # excel_file = r"C:\EATB_Config_Generator\config.xlsx"
    
    if not 'config_path' in locals():
        config_path = r"C:\EATB_Config_Generator\config.xlsx"
        
    excel_file = config_path
    
    # read the excel file and construct the signals sheet dataframe
    signal_dataframe = pd.read_excel(excel_file, sheet_name="Signals", engine='openpyxl')
    
    # reads the excel headers as a list, thanks to the well organized sheet
    signal_headers = list(signal_dataframe.columns.values)
    
    # checks if the signal sheet is in default headers before project titles
    if ((signal_headers[0] == 'EATB signal') and (signal_headers[1] == 'A2L label')\
                                             and (signal_headers[2] == 'Raster')):
        pass
    else:
        print("Signal headers are not in default state. Please change the headers as:")
        print("Cell A1 : 'EATB signal'")
        print("Cell A2 : 'A2L label'")
        print("Cell A3 : 'Raster'")
        print("If the problem persists, please reach the author.")
        # return # in standalone testing, erase this line to avoid error 
        
    # discover project names from after default signal sheet headers
    project_list = signal_headers[3:]
    # del signal_headers
    
    # getting eatb signal labels in a list (reference if needed)
    eatb_list = signal_dataframe['EATB signal'].tolist()
    
    # getting a2l signal labels in a list (reference if needed)
    a2l_list = signal_dataframe['A2L label'].tolist()
    
    # getting project number
    project_number = len(project_list)
    
    # checking user signal choices for individual projects and saving the
    # index value for individual projects when the user entered x
    # may be completed faster using remove NaN related methods
    sig_cho_dict = {}
    for i in range (project_number):
        choices_ind = signal_dataframe[project_list[i]].tolist()
        cho_indices = []
        for i2 in range (len(choices_ind)):
            if choices_ind[i2] == "x":
                cho_indices.append(i2)
        sig_cho_dict[i] = cho_indices
    # del i,i2,cho_indices,choices_ind
    
    # constructing an eatb dict to represent only selected signals
    eatb_dict = {}
    for i in range (project_number):
        temp = list(sig_cho_dict.get(i))
        eatb_label= []
        for i2 in range (len(temp)):
            eatb_label.append(eatb_list[temp[i2]])
        eatb_dict[i] = eatb_label
    # del i, i2, eatb_label, temp, eatb_list
        
    # constructing an a2l dict to represent only selected signals
    a2l_dict = {}
    for i in range (project_number):
        temp = list(sig_cho_dict.get(i))
        a2l_label= []
        for i2 in range (len(temp)):
            a2l_label.append(a2l_list[temp[i2]])
        a2l_dict[i] = a2l_label    
    # del i, i2, temp, a2l_label, a2l_list
    
    # %% Calculation Sheet
    
    # loading the calculation sheet as a dataframe
    # note a function may be added to Calculation sheet dataframe preparation
    calc_dataframe = pd.read_excel(excel_file, sheet_name="Calculation", engine='openpyxl')
    
    # finding project name column index numbers
    header_number = len(calc_dataframe.columns)
    index_temp = []
    for i in range(header_number):
        if not pd.isnull(calc_dataframe.iat[0,i]): 
            index_temp.append(i)
    # del i, header_number
     
    # changing header names and deleting old header row
    old_col = list(calc_dataframe.columns)
    for i in range (len(index_temp)):
        calc_dataframe.rename(columns={old_col[index_temp[i]]: calc_dataframe.\
                                       iat[0, index_temp[i]]}, inplace=True)
    calc_dataframe = calc_dataframe.drop(0)
    # del i, index_temp, old_col
        
    # get non-NaN row index values at block column
    block_list = calc_dataframe['Block'].tolist()
    block_index= []
    for i in range(len(block_list)):
        if not pd.isna(block_list[i]):
            block_index.append(i)
    # del block_list, i
    
    # filling NaN values in block array with ffill method and removing empty rows
    calc_dataframe.loc[:,'Block'] = calc_dataframe.loc[:,'Block'].ffill()
    calc_dataframe.loc[:,'Condition'] = calc_dataframe.loc[:,'Condition'].ffill()
    for i in range(len(block_index)):
        calc_dataframe = calc_dataframe.drop(block_index[i]+1)
    # del block_index, i
    
    # reset index
    calc_dataframe.reset_index(drop = True, inplace = True)

    # datasheet is configured, getting user choices again in this sheet
    cal_cho_dict = {}
    for i in range (project_number):
        choices_ind = calc_dataframe[project_list[i]].tolist()
        cho_indices = []
        for i2 in range (len(choices_ind)):
            if choices_ind[i2] == "x":
                cho_indices.append(i2)
        cal_cho_dict[i] = cho_indices    
    # del i, i2, choices_ind, cho_indices
         
    # getting all 3 column contents and saving them based on project choices
    cal_sig_dict = {}
    cal_cal_dict = {}
    cal_con_dict = {}
    
    calc_sig_list = calc_dataframe['Signal'].tolist()
    calc_cal_list = calc_dataframe['Calculation'].tolist()
    calc_con_list = calc_dataframe['Condition'].tolist()
    
    for i in range (project_number):
        calc_sig_list_pro = []
        calc_cal_list_pro = []
        calc_con_list_pro = []
        
        cho = list(cal_cho_dict.get(i))
        for i2 in range(len(cho)):
            calc_sig_list_pro.append(calc_sig_list[cho[i2]])
            calc_cal_list_pro.append(calc_cal_list[cho[i2]])
            calc_con_list_pro.append(calc_con_list[cho[i2]])
            
        cal_sig_dict[i] = calc_sig_list_pro
        cal_cal_dict[i] = calc_cal_list_pro
        cal_con_dict[i] = calc_con_list_pro
    # del calc_sig_list, calc_cal_list, calc_con_list, i , i2, cho
    # del calc_sig_list_pro, calc_cal_list_pro, calc_con_list_pro
        
    # disecting the signal data for different projects
    raw_cal_sig = {}
    raw_cal_cal = {}
    raw_cal_con = {}
    
    for i in range (project_number):
        sig = list(cal_sig_dict.get(i))      
        sig2 = re.findall('\<<.*?\>>', str(sig))
        sig3 = []
        for i2 in range (len(sig2)):
            temp = sig2[i2].replace("<<","")
            temp = temp.replace(">>","")
            sig3.append(temp)
        sig3 = list(dict.fromkeys(sig3))
        raw_cal_sig[i] = sig3
    # del i, i2, sig, sig2, sig3, temp
        
    for i in range (project_number):
        cal = list(cal_cal_dict.get(i))      
        cal2 = re.findall('\<<.*?\>>', str(cal))
        cal3 = []
        for i2 in range (len(cal2)):
            temp = cal2[i2].replace("<<","")
            temp = temp.replace(">>","")
            cal3.append(temp)
        cal3 = list(dict.fromkeys(cal3))
        raw_cal_cal[i] = cal3
    # del i, i2, cal, cal2, cal3, temp
        
    for i in range (project_number):
        con = list(cal_con_dict.get(i))      
        con2 = re.findall('\<<.*?\>>', str(con))
        con3 = []
        for i2 in range (len(con2)):
            temp = con2[i2].replace("<<","")
            temp = temp.replace(">>","")
            con3.append(temp)
        con3 = list(dict.fromkeys(con3))
        raw_cal_con[i] = con3
    # del i, i2, con, con2, con3, temp
    # del cal_cal_dict, cal_con_dict, cal_sig_dict
    
    # %% Graph Sheet
         
    # getting data from calculation sheet is completed as well
    # for next, getting data from graphs sheet:
    graph_dataframe = pd.read_excel(excel_file, sheet_name="Graphs", engine='openpyxl')
    
    # removing extra columns after comment column
    i_ref = graph_dataframe.columns.get_loc("Comment")
    i_end = len(graph_dataframe.columns)
    
    k = 1 # temporary constant
    while i_ref < (i_end - 1):
        graph_dataframe.drop(graph_dataframe.columns[i_ref + k], axis = 1, inplace = True)
        i_end -= 1
    # del k, i_end, i_ref
    
    # non official ffill method on the column headers to change unnamed values
    old_col = old_col2 = graph_dataframe.columns.values.tolist()
    indexes = [old_col.index(l) for l in old_col if not l.startswith('Unnamed')]
    for i in range(len(indexes) - 1):
        if (indexes[i+1]-indexes[i]) > 1:
            k = indexes[i+1] - indexes[i]
            for i2 in range(k-1):
                old_col[indexes[i]+i2+1] = old_col[indexes[i]]
    # del old_col2, indexes
    
    # appointing the new column names
    graph_dataframe.columns = old_col
    
    # appointing new headers with sub_strings
    sub_header = graph_dataframe.loc[0,:].tolist()    
    new_list = []
    
    for i in range (len(sub_header)):
        if str(sub_header[i]) != "nan":
            new_list.append(old_col[i] + "_" + sub_header[i])
        elif str(sub_header[i]) == "nan":
            new_list.append(old_col[i])
    # del sub_header
            
    graph_dataframe.columns = new_list
    graph_dataframe = graph_dataframe.drop(0)
    
    # del new_list, old_col
    
    # forward fill on chapter names
    graph_dataframe["Chapter"] = graph_dataframe["Chapter"].replace('none', np.nan)
    chap_list = graph_dataframe['Chapter'].tolist()
    chap_index = []
    
    for i in range(len(chap_list)):
        if not pd.isna(chap_list[i]):
            chap_index.append(i)
    graph_dataframe.loc[:,'Chapter'].ffill(inplace = True)
    # del chap_list
    
    # reset index
    graph_dataframe.reset_index(drop = True, inplace = True)
    
    # deleting the predefined chapter rows
    for i in range(len(chap_index)):
        graph_dataframe = graph_dataframe.drop(chap_index[i])
    # del chap_index
        
    # reset index
    graph_dataframe.reset_index(drop = True, inplace = True)
    
    # repeating the proces on section names
    graph_dataframe["Section"] = graph_dataframe["Section"].replace('none', np.nan)
    sect_list = graph_dataframe['Section'].tolist()
    sect_index = []
    
    for i in range(len(sect_list)):
        if not pd.isna(sect_list[i]):
            sect_index.append(i)
    # del sect_list
    
    graph_dataframe.loc[:,'Section'].ffill(inplace = True)
    
    for i in range(len(sect_index)):
        graph_dataframe = graph_dataframe.drop(sect_index[i])
    # del sect_index
        
    # reset index
    graph_dataframe.reset_index(drop = True, inplace = True)
    
    # getting user choices based on projects
    gra_cho_dict = {}
    for i in range (project_number):
        choice_ind = graph_dataframe["Project_" + project_list[i]].tolist()
        index_cho  = []
        for i2 in range (len(choice_ind)):
            if choice_ind[i2] == "x":
                index_cho.append(i2)
        gra_cho_dict[i] = index_cho
    # del choice_ind, index_cho
        
    # getting all 3 column contents and saving them based on project choices 
    gra_sig_dict = {}
    gra_con_dict = {}
    gra_tri_dict = {}
    
    gra_sig_list = graph_dataframe['Signal(s)_Name'].tolist()
    gra_con_list = graph_dataframe['Condition_Signal'].tolist()
    gra_tri_list = graph_dataframe['Trigger_Signal'].tolist()
    
    for i in range (project_number):
        gra_sig_list_pro = []
        gra_con_list_pro = []
        gra_tri_list_pro = []
        
        cho = list(gra_cho_dict.get(i))
        for i2 in range(len(cho)):
            gra_sig_list_pro.append(gra_sig_list[cho[i2]])
            gra_con_list_pro.append(gra_con_list[cho[i2]])
            gra_tri_list_pro.append(gra_tri_list[cho[i2]])
    
        gra_sig_dict[i] = gra_sig_list_pro
        gra_con_dict[i] = gra_con_list_pro
        gra_tri_dict[i] = gra_tri_list_pro
        
        
    # del gra_cho_dict
    # del gra_sig_list_pro, gra_con_list_pro, gra_tri_list_pro
    # del cho, gra_sig_list, gra_con_list, gra_tri_list
    
    # disecting the signal data for different projects
    raw_gra_sig = {}
    raw_gra_con = {}
    raw_gra_tri = {}            
        
    for i in range(project_number):
        signal = list (gra_sig_dict.get(i))
        while "none" in signal: signal.remove("none")
        while np.nan in signal: signal.remove(np.nan)
        
        gra_tt1 = signal
        gra_tt1 = [map(lambda x: x.strip(), item.split(',')) for item in gra_tt1]
        signal = [item for sub_list in gra_tt1 for item in sub_list]
        # extend two lines on one index
        final_signal = list(dict.fromkeys(signal))
        raw_gra_sig[i] = final_signal
     
        condition = list (gra_con_dict.get(i))
        while "none" in condition: condition.remove("none")
        while np.nan in condition: condition.remove(np.nan)
        # extend two lines on one index
        gra_tt2 = condition
        gra_tt2 = [map(lambda x: x.strip(), item.split(',')) for item in gra_tt2]
        condition = [item for sub_list in gra_tt2 for item in sub_list]
        
        final_condition = list(dict.fromkeys(condition))
        raw_gra_con[i] = final_condition
        
        trigger = list (gra_tri_dict.get(i))
        while "none" in trigger: trigger.remove("none")
        while np.nan in trigger: trigger.remove(np.nan)
        
        gra_tt3 = trigger
        gra_tt3 = [map(lambda x: x.strip(), item.split(',')) for item in gra_tt3]
        trigger = [item for sub_list in gra_tt3 for item in sub_list]
        # extend two lines on one index
        final_trigger = list(dict.fromkeys(trigger))
        raw_gra_tri[i] = final_trigger
         # multiple items - repeat above process if similar results in other sheets
        
    # del signal, condition, trigger
    # del final_signal, final_condition, final_trigger
    # del gra_sig_dict, gra_con_dict, gra_tri_dict
    # del i, i2, k
        
    # project based dictionaries:
    # signals sheet:
        # eatb signal names: eatb_dict
        # a2l  signal names: a2l_dict
        # user choice index: sig_cho_dict
    # calculation sheet:
        # signal column signals: raw_cal_sig
        # calcul.column signals: raw_cal_cal
        # condit.column signals: raw_cal_con
    # graphs sheet:
        # signal(s) name column: raw_gra_sig
        # condition name column: raw_gra_con
        # trigger   name column: raw_gra_tri

    # %% Comparisons
    
    # 2 returning variables
    mismatch = {}
    corpse_meta = {}
    
    # run project times individually
    for i in range(project_number):
        
        # measured signals
        m1 = list(eatb_dict.get(i))
        mt = m1
        
        #calculated signals
        c1 = list(raw_cal_sig.get(i))
        ct = c1
        
        #measured and calculated signals with no duplicates
        mct= list(set(mt + ct))
        
        # used signals
        u1 = list(raw_gra_sig.get(i))
        u2 = list(raw_gra_con.get(i))
        u3 = list(raw_gra_tri.get(i))
        u4 = list(raw_cal_cal.get(i))
        u5 = list(raw_cal_con.get(i))
        
        # all used signals with no duplicates
        ut = list(set(u1+u2+u3+u4+u5))
        
        # temporary list for dictionary data collection
        mism_list = []
        core_mela = []
        
        # check if used signal is available
        for sig1 in ut:
            count = 0
            for sig2 in mct:
                if sig1 == sig2:
                    count += 1
            if count == 0:
                mism_list.append(sig1)
                
        mism_list = sorted(mism_list)  # remove sorting if warning
        mismatch[i] = mism_list
        # del count, sig1, sig2
        
        # check if available signals are used
        for sig1 in mct:
            count = 0
            for sig2 in ut:
                if sig1 == sig2:
                    count += 1
            if count == 0:
                core_mela.append(sig1)               
                
        core_mela = sorted(core_mela) # remove sorting if warning
        corpse_meta[i] = core_mela
        # del count, sig1, sig2
        
    return mismatch, corpse_meta
                
    # %% Developer Info for Customization
    
    # Hello reader, if you want to customize the code,
    # please take a look at the parameters below which
    # are generated in a successfull run. If needed, use
    # them as quick access and if needed contact me.
    
    # Signal Sheet Parameters:
        # a2l_dict: all a2l labels divided by project indices
        # eatb_dict: all eatb labels divided by project indices
        # project_list: singular project labels in a list
        # project_number: number of projects
        # sig_cho_dict: user choices divided by project indices
        # signal_dataframe: total dataframe of excel sheet
        # excel_file: optional config path for function debug
        
    # Calculation Sheet Parameters:
        # cal_cho_dict: user choices divided by project indices
        # calc_dataframe: total dataframe of excel sheet
        # raw_cal_cal: disected signal labels in calculation column
        # raw_cal_con: disected signal labels in condition column
        # raw_cal_sig: disected signal labels in signal column
        
    # Graph Sheet Parameters:
        # graph_dataframe: processed graph sheet dataframe
        # raw_gra_con: disected signal labels in condition column
        # raw_gra_sig: disected signal labels in signal column
        # raw_gra_tri: disected trigger labels in signal column
        
    # Finalized Parameters for Quick Access
        # 5 used signals
        # 1 measured signal
        # 1 calculated signal
    
    # Recommendations:        
        # produce dictionary items for individual projects
        # you may prepare strings like: (in project loop)
            # "in project {}, signal {}" has been used.
            # "total number of used signals are {}"
            # mismatch and metaphorical corpse variables:
        # also, you may use excel to dataframe blocks externally
        # if you have recommendations, please contant me: KUY3IB
        # thank you for reading and happy developing...        

# %%
# sample test in this file (Spyder Python)
# click the empty space above upper comment
# ctrl + shift + home to select all content above
# run the selection (f9 in spyder)
# make sure that config file is in directory
# run the 2 lines below:
    
# excel = "C:\EATB_Config_Generator\config.xlsx"
# (mismatch, corpse_meta) = signal_check (excel)

# tested for config excel file
# tested for FIE excel file
# tested for EPM excel file