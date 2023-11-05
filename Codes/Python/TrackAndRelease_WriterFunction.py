"""
Produced on Sat Mar 19 20:59:59 2022
@author: KUY3IB
"""

# %% class organization

"""
config class (parent)
--------------------
calc class (child)
conf class (child)
filt class (child)
func class (child)
grap class (child)
sign class (child)
"""

# %% initialization

# %reset -f
import pandas as pd
import numpy as np
import shutil as sh
import warnings as wr
import sys
import os
import re

# r"C:\EATB_Config_Generator\eatb-exc\EATB_Diesel_FIE_v2.02.xlsm"
# r"C:\EATB_Config_Generator\eatb-exc\EATB_EPM_v0.07.xlsm"

# %% config - main class

class config:
    output_folder_name = "Output" ;    
    excel_path = r"C:\Users\KUY3IB\Desktop\temp_ref\EATB_EPM_v0.07.xlsm"
    output_env = r"C:\Users\KUY3IB\Desktop\temp_result"
    
    def __init__(self):
        self.set_warning_management("ignore")  
        self.output_path = self.join_path (self.output_env, self.output_folder_name)
        self.check_and_make_folder (self.output_path)
        self.excel_dict = self.read_excel(self.excel_path)
        self.signals_df = self.get_dataframe_from_dict(self.excel_dict, "Signals")
        self.signals_hd = self.get_headers_from_dataframe(self.signals_df)
        self.raster_ind = self.find_index_in_list("Raster", self.signals_hd)
        self.base_projects = self.get_list_after_index(self.raster_ind, self.signals_hd)
        self.project_num = len(self.base_projects)
        self.make_project_folders(self.base_projects, self.output_path)   
        self.project_paths = self.project_folder_paths(self.base_projects, self.output_path,{})
        self.functions_df = self.get_dataframe_from_dict(self.excel_dict, "Functions")
        del self.signals_df, self.signals_hd, self.raster_ind
        
    def set_warning_management (self, method):
        """ sets the warning management method for the whole script """
        wr.simplefilter(action=method, category=Warning)

    def join_path(self, target_folder_path, file_name):
        """ combines a file name with a target path, to make a file path """
        return os.path.join(target_folder_path, file_name)
        
    def check_and_make_folder (self, path):
        """ checks if a folder exists at the given path and resets if exists or makes path only """    
        (sh.rmtree(path), os.mkdir(path)) if os.path.exists(path) else os.mkdir(path) 

    def read_excel (self, input_path):
        """ reads the input excel file and returns a dictionary item """
        return pd.read_excel(input_path,sheet_name=None,engine='openpyxl' )

    def get_dataframe_from_dict (self, excel_dict, worksheet):
        """ takes a dataframe from a dict in main class if necessary """
        return excel_dict [worksheet]

    def get_headers_from_dataframe (self, dataframe):
        """ collects non multi-index header list from a dataframe """
        return dataframe.columns.tolist()
    
    def find_index_in_list (self, item, e_list):
        """ finds the index value of an item in list """
        return e_list.index(item)
    
    def get_list_after_index (self, index, e_list):
        """ from a list, gets sub-list after an index """
        return e_list[(index+1):]
    
    def make_project_folders (self, project_list, output_path):
        """ generates project folders in output folder """
        for pr in project_list: os.mkdir(os.path.join(output_path, pr)); 
        
    def project_folder_paths(self, p_list, output_path, e_dict):
        """ collects project folder paths in a dict item, based on project indices """
        for i in range(len(p_list)): e_dict[i]=os.path.join(output_path, p_list[i])
        return e_dict
    
    # additional functions for general use
    def set_input_path (self, path):
        """ changes excel file path if necessary """
        self.excel_path = rf"{path}"
        
    def set_output_name (self, name):
        """ changes output folder name-Output to another name """
        self.output_folder = rf"{name}"
        
    def set_output_environment_path (self, path):
        """ changes the target output path if necessary """
        self.output_env = rf"{path}"
        
    def set_output_path (self, path):
        """ changes the final output folder path if necessary """
        self.output_path = rf"{path}"
     
    def get_sub_list_df(self, dataframe, indices, col_name):
        """ returns a sub list from a list based on another list of index locations """
        return [(dataframe[col_name].tolist())[i] for i in indices]
       
    def change_unnamed_to_nan (self, e_l):
        """ in a list: change Unnamed: 1, Unnamed: 2... values to np.nan """
        for i in range(len(e_l)): e_l[i]=np.nan if e_l[i].startswith("Unnamed") else e_l[i]
        
    def change_dataframe_headers(self, new_header, dataframe):
        """ change columns of a dataframe with a list of new columns """
        dataframe.columns = new_header
     
    def ffill_dataframe_header_nan(self, dataframe):
        """ fills dataframe header from left to right with respect to nan values """ 
        dataframe.columns = dataframe.columns.to_series().mask(lambda x: x==np.nan).ffill()

    def ffill_dataframe_header_unnamed(self, dataframe):
        """ fills dataframe header from left to right with respect to "Unnamed" values """
        dataframe.columns = dataframe.columns.to_series() \
                            .mask(lambda x: x.str.startswith('Unnamed')).ffill()
        
    def combine_non_nan_multiindex_headers (self, df):
        """ while row directly below header is not nan, add to header with space """
        for i in range(len(df.columns)): df.columns.values[i] = df.columns[i] + \
            " " + df.iat[0,i] if not pd.isnull(df.iat[0,i]) else df.columns[i]
        
    def dataframe_drop_row(self, dataframe, row_index):
        """ drops a single row from a dataframe """
        dataframe.drop(row_index, inplace=True); 
        
    def dataframe_reset_index(self, dataframe):
        """ resets index and removes the old index """
        dataframe.reset_index(drop = True , inplace = True)
        
    def dataframe_not_nan_row_indices(self, dataframe,column_index):
        """ saves index locations of non-nan locations of row headers """
        return np.where(dataframe.iloc[:,column_index].notnull())[0].tolist()
        
    def ffill_dataframe_row_header_nan(self, df, col_index):
        """ fills dataframe row header from up to down with respect to nan values """
        df.iloc[:,col_index].mask(lambda x: x==np.nan).ffill(inplace = True)
     
    def dataframe_drop_rows(self, dataframe, row_indices):
        """ drops multiple rows based on indices from a dataframe """
        for i in row_indices: dataframe.drop(i, inplace = True)
        
    def get_single_value(self, header_name, dataframe):
        """ this is designed for dataframes with one header and one value """
        return dataframe.at[0,header_name]
    
    def change_single_value(self, header_name, dataframe, new_value):
        """ this is designed for dataframes with one header and one value """
        dataframe.at[0,header_name] = new_value
        
    def find_column_index_df(self, dataframe,column_name):
        """ finds the column index of a header name in a dataframe """
        return dataframe.columns.get_loc(column_name)
        
    def dataframe_drop_columns_after_index (self, dataframe, index):
        """ drops all columns after a column to the end of the dataframe """
        dataframe.drop(dataframe.columns[(index+1):], axis=1, inplace = True)
        
    def dataframe_col_to_list(self, dataframe, column_name):
        return dataframe.loc[:,column_name].tolist()
    
    def non_duplicate_list (self, dup_list):
        """ from a list with duplicate values, get non-duplicate value list """
        return list(dict.fromkeys(dup_list))
    
    def join_paths(self, target_folder_path_dict, file_name, e_dict):
        """ combines a file name with multiple target paths, to make file paths """
        for i in range(len(target_folder_path_dict)): 
            e_dict[i] = os.path.join(target_folder_path_dict[i], file_name)
        return e_dict
    
# %% calculation sheet child class

class calc (config):
    ws_name = 'Calculation'
    def __init__(self, parent):
        self.dataframe = parent.get_dataframe_from_dict (parent.excel_dict, self.ws_name)
        parent.ffill_dataframe_header_unnamed(self.dataframe)
        parent.combine_non_nan_multiindex_headers (self.dataframe)
        self.header = parent.get_headers_from_dataframe (self.dataframe)
        parent.dataframe_drop_row(self.dataframe, 0)
        parent.dataframe_reset_index(self.dataframe)
        self.row_indices = parent.dataframe_not_nan_row_indices(self.dataframe,0)
        parent.ffill_dataframe_row_header_nan(self.dataframe, 0)
        parent.dataframe_drop_rows(self.dataframe, self.row_indices)
        parent.dataframe_reset_index(self.dataframe)
                                                      
# %% config sheet child class
        
class conf (config): 
    ws_name = "Config"
    def __init__(self, parent):
        self.dataframe = parent.get_dataframe_from_dict (parent.excel_dict, self.ws_name)
        self.header = self.get_config_header(self.dataframe)
        self.values = self.get_config_values(self.dataframe)
        self.change_config_dataframe(self.header, self.values)
        self.output_path = parent.get_single_value("Output path", self.dataframe)
        self.debug_config_diff=parent.get_single_value("Debug config_diff", self.dataframe)
        self.encrypter = parent.get_single_value("Encrypter", self.dataframe)
        
    def get_config_header(self, dataframe):
        """ gets worksheet headers from different locations to a list """
        return [dataframe.columns[0],dataframe.iat[0,0],dataframe.iat[1,0]]
        
    def get_config_values(self, dataframe):
        """ gets worksheet values from different locations to a list """
        return [dataframe.columns[1],dataframe.iat[0,1],dataframe.iat[1,1]]
    
    def change_config_dataframe (self, header, values):
        """ with new headers and values, makes a new dataframe as active """
        self.dataframe = pd.DataFrame([values], columns=header)

# %% filter sheet child class
    
class filt (config): 
    ws_name = "Filters"
    def __init__(self, parent):
        self.dataframe = parent.get_dataframe_from_dict (parent.excel_dict, self.ws_name)
        parent.dataframe_drop_row(self.dataframe, 0)
        parent.dataframe_reset_index(self.dataframe)
        self.header = parent.get_headers_from_dataframe (self.dataframe)
        
# %% function sheet child class
        
class func (config):
    ws_name = "Functions"
    outputs = "result, err"
    output_list = ["result","err"]
    def __init__(self, parent):
        self.dataframe = parent.get_dataframe_from_dict(parent.excel_dict, self.ws_name)
        parent.dataframe_drop_row(self.dataframe, 0)
        parent.dataframe_reset_index(self.dataframe)        
        self.header = parent.get_headers_from_dataframe (self.dataframe)
        self.row_indices = parent.dataframe_not_nan_row_indices(self.dataframe,0) 
        parent.ffill_dataframe_row_header_nan(self.dataframe, 0)       
        parent.dataframe_drop_rows(self.dataframe, self.row_indices)
        parent.dataframe_reset_index(self.dataframe)
        self.dataframe = self.dataframe.replace("return", "result")
        self.function_info = parent.non_duplicate_list (self.dataframe["Block"].tolist())
        self.function_num = len(self.function_info)
        self.function_name = self.get_function_labels(self.function_info,[])
        self.all_inputs = self.get_function_inputs(self.function_info, [])
        self.ind_inputs = self.individual_inputs(self.all_inputs, self.function_num, {})
        self.func_indices (self.function_num, self.dataframe, self.function_info, [], {})
        self.calc_info = self.prepare_calculation(self.dataframe, {})
        # self.ffill_else (self.function_indices, self.dataframe) don't enable, if working
        # below part may be edited later hopefully, for smaller code and organization
        self.m_text_start = self.start_text (self.function_num, self.outputs, self.function_name, \
                                             self.ind_inputs, self.output_list, self.all_inputs, [])
        self.m_text_middle, self.count = self.cont_text (self.function_num, self.function_indices, \
                                                         self.dataframe, self.calc_info, {}, {})
        self.m_text_end = self.final_text ()
        self.function_paths = parent.join_paths(parent.project_paths, "functions", {})
        
        self.make_the_text (self.m_text_start, self.m_text_middle, self.m_text_end, \
                            parent.project_num, self.function_name, self.function_paths, self.count)

#         self.make_m_from_the_text(self.label, self.m_text, parent.project_paths)
        
        
    def get_function_labels(self, func_list, e_list):    
        """ takes the label part of function infos before paranthesis, it is also .m file label """
        for item in func_list: e_list.append(item.rpartition('(')[0])
        return e_list
        
    def get_function_inputs(self, func_list, e_list):    
        """ gets the input part of function names between paranthesis """
        for item in func_list: e_list.append(item[item.find('(')+1 : item.find(')')])
        return e_list
        
    def individual_inputs (self, input_list, func_num, e_dict):
        """ gets the individual inputs as list in dict for different functions """
        for i in range(func_num): e_dict[i] = input_list[i].split(", ")
        return e_dict
        
    def func_indices (self, func_num, dataframe, func_info, e_list, e_dict):
        """ in block column ffilled dataframe, finds index values for functions TBE"""
        for i in range(func_num):
            e_list = []
            for i2 in range(len(dataframe)):
                if dataframe.at[i2,"Block"]==func_info[i]: e_list.append(i2)
            e_dict[i] = e_list
        self.function_indices = e_dict
    
    def prepare_calculation (self, dataframe, e_dict):
        """ disects the inputs at calculation sheet into a dict and eliminates comments TBE """
        for i in range(dataframe.shape[0]): e_dict[i]= dataframe.iat[i, 2].split("\n")
        for i in range(len(e_dict)): 
            for i2 in range(len(e_dict[i])): 
                e_dict[i][i2]=e_dict[i][i2].strip().split("%")[0]
        return e_dict
    
    def ffill_else (self, f_indices, dataframe):
        """ in same function cells, ffills the remaining cells after excel string """
        for i in range(len(f_indices)):
            dataframe[dataframe.iloc[f_indices[i]].ffill() == "else"] = "else"
            
    def start_text (self, function_num, outputs, label, inputs, output_list, all_inputs, e_list):
        for i in range(function_num):
            m_start = f"function[{outputs}]={label[i]}({all_inputs[i]})\n" +\
                      (f"{output_list[0]}={inputs[i][0]};\n" if len(inputs[i])!=0 else None) +\
                      f"{output_list[1]}= 0;\n" +\
                      "try\n"
            e_list.append(m_start)
        return e_list
                      
    def cont_text (self, f_num, f_indices, d_frame, c_info, storage_dict, e_dict):
        for i in range(f_num):
            storage = []
            count = 0
            for i2 in f_indices[i]:
                
                if pd.isnull(d_frame.at[i2,'Condition']):
                    if pd.isnull(d_frame.at[i2,'Signal']):
                        calc_joined = "\n".join(c_info[i2])
                        storage.append(f"{calc_joined}")
                        
                    if not pd.isnull(d_frame.at[i2,'Signal']):
                        storage.append(f"{d_frame.at[i2,'Signal']}={d_frame.at[i2,'Calculation']};")
                        
                if not pd.isnull(d_frame.at[i2,'Condition']): 
                    
                    if pd.isnull(d_frame.at[i2,'Signal']):
                        calc_joined = "\n".join(c_info[i2])
                        temp = "else" if d_frame.at[i2,'Condition'].startswith("else") else "if"
                        ttemp = f"{temp}" + (f" {d_frame.at[i2,'Condition']}\n" if temp == "if" \
                                              else "\n") + f"{calc_joined}\n" + "end"
                        storage.append(ttemp)
                        
                    if not pd.isnull(d_frame.at[i2,'Signal']):
                        temp = "else" if d_frame.at[i2,'Condition'].startswith("else") else "if"
                        ttemp= f"{temp}" + (f" {d_frame.at[i2,'Condition']}\n" if temp=="if" else "\n")\
                                         + f"{d_frame.at[i2,'Signal']}={d_frame.at[i2,'Calculation']};"
                        if temp == "if": count += 1
                        storage.append(ttemp)
            storage_dict[i] = storage
            e_dict[i] = count
        return storage_dict, e_dict
    
    def final_text (self):
        teex = "catch me \nerr=1; \nend \nclear temp*"
        return teex
                
    def make_the_text (self, start, middle, end, p_num, f_name, f_paths, count):
        for i in range(p_num):
            os.mkdir(f_paths[i])
            for i2 in range(len(f_name)):
                txt = open(os.path.join(f_paths[i],f_name[i2]+".m"), "w")
                
                txt.write(start[i2] + ("\n".join(middle[i2])) + "\n" + "\nend\n"*count[i2] + end)
                txt.close
        
# %% graph sheet child class
        
class grap (config):
    ws_name = "Graphs"
    def __init__(self, parent):
        self.dataframe = parent.get_dataframe_from_dict(parent.excel_dict, self.ws_name)
        self.comment_index = parent.find_column_index_df(self.dataframe,"Comment")
        parent.dataframe_drop_columns_after_index (self.dataframe, self.comment_index)
        parent.ffill_dataframe_header_unnamed(self.dataframe)
        parent.combine_non_nan_multiindex_headers (self.dataframe)
        self.header = parent.get_headers_from_dataframe (self.dataframe)
        parent.dataframe_drop_row(self.dataframe, 0)
        parent.dataframe_reset_index(self.dataframe)
        for i in range(4):
            if i < 2: self.row_indices = parent.dataframe_not_nan_row_indices(self.dataframe,i)
            parent.ffill_dataframe_row_header_nan(self.dataframe, i)
            if i < 2: parent.dataframe_drop_rows(self.dataframe, self.row_indices)
            if i < 2: parent.dataframe_reset_index(self.dataframe)
        
# %% signal sheet child class
    
class sign (config): 
    ws_name = "Signals"
    def __init__(self, parent):
        self.dataframe = parent.get_dataframe_from_dict(parent.excel_dict, self.ws_name)
        self.header = parent.get_headers_from_dataframe (self.dataframe)
        
#%% functions - testing

# con = config()
# cal = calc(con)
# can = conf(con)
# fil = filt(con)
# gra = grap(con)
# sig = sign(con)
# if len(con.functions_df) >= 4: fun = func(con)