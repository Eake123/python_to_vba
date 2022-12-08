import json
from datetime import datetime
import pandas as pd
import subprocess 
from config_script import config_python_to_vba
import string
def int_to_col(column_int:int):
    start_index = 1   #  it can start either at 0 or at 1
    letter = ''
    while column_int > 25 + start_index:   
        letter += chr(65 + int((column_int-start_index)/26) - 1)
        column_int = column_int - (int((column_int-start_index)/26))*26
    letter += chr(65 - start_index + (int(column_int)))
    return letter
        
def col_to_int(column_letter):
    s = 0
    for count,i in enumerate(column_letter):
        if count == len(column_letter) - 1:
            s += string.ascii_uppercase.index(i) + 1
        else:
            s += 26 * (string.ascii_uppercase.index(i) + 1)
    return s




class VBA:
    def __init__(self,data,**kwargs) -> None:
        '''
        PARAMS
        data: REQUIRED VBA class accepts three datatypes for data.
        1. dictionary:
        if passing a dictionary it must be configured like this
        {excel_sheetname: {column letter: {row number: value}}}
        2. pandas dataframe
        if passing dataframe you must also pass a parameter "sheet", where it is the sheet name that you want edited
        3. list of pandas dataframe

        sheet: REQUIRED IF USING DATAFRAME OR LIST OF DATAFRAME. If a single DataFrame send in a str with the desired sheet name for editing.
        If a list of DataFrames you need to send in a list that is equal in length to the list of DataFrames ex. ["Sheet1","Sheet2"]
        
        start_row: OPTIONAL, ONLY USED FOR DataFrame and list of DataFrames. It is the row that the DataFrame begins pasting in excel.
        If sending a list of DataFrames you can also use a list, as long as it is equal length to the DataFrames

        start_col: OPTIONAL, ONLY USED FOR DataFrame and list of DataFrames. ACCEPTS int or string. It is the column that the DataFrame begins pasting in excel.
        If sending a list of DataFrames you can also use a list, as long as it is equal length to the DataFrames

        index: OPTIONAL, ONLY USED FOR DataFrame and list of DataFrames. ACCEPTS bool or list. If index is True or not defined it will paste the index on the start_col.
        If sending a list of DataFrames you can also use a list, as long as it is equal length to the DataFrames

        index_name: OPTIONAL, ONLY USED FOR DataFrame and list of DataFrames. ACCEPTS any excel printable d_type. Makes the header name for the index column, if not define it defaults to empty string.
        If sending a list of DataFrames you can also use a list, as long as it is equal length to the DataFrames


        
        EXAMPLE with dictionary to add "value" in the first cell
        
        df = {
            'Sheet1':{
                'A':{
                    1:'value'
                }
            }
        }

        v = VBA(df)

        v.to_file('test.json')

        EXAMPLE with single DataFrame
        
        length = 10

        df = {
            'a':[random.randint(0,100) for x in range(length)],
            'b':[random.randint(0,100) for x in range(length)],
            'c':[random.randint(0,100) for x in range(length)],
            'd':[random.randint(0,100) for x in range(length)]
        }

        df = pd.DataFrame(df)

        v = VBA(df,sheet='Sheet1')

        v.to_file('test.json')

        EXAMPLE with two DataFrames
        
        length = 10000

        df = {
            'a':[random.randint(0,100) for x in range(length)],
            'b':[random.randint(0,100) for x in range(length)],
            'c':[random.randint(0,100) for x in range(length)],
            'd':[datetime(2022,random.randint(1,12),1) for x in range(length)]
        }

        df2 = {
            'a':[random.randint(0,100) for x in range(length)],
            'b':[random.randint(0,100) for x in range(length)],
            'c':[random.randint(0,100) for x in range(length)],
            'd':[random.randint(0,100) for x in range(length)]
        }

        sheet = ['Sheet1','Sheet2']

        df = pd.DataFrame(df)

        df2 = pd.DataFrame(df2)

        l = [df,df2]

        v = VBA(l,sheet=sheet)
        v.copy_to_clipboard()

        '''
        if isinstance(data,pd.DataFrame):
            data = self.df_to_dict(data,kwargs)
        elif isinstance(data,list):
            data = self.multiple_dfs_to_dict(data,kwargs)
        self.data = self.__constructor(data)

    


    def __constructor(self,data):
        '''makes sure that the dictionary is constructed as its supposed to. Also adds the datatype to the end'''
        df = {}
        for sheet in data.keys():
            if isinstance(sheet,str) == False:
                raise ValueError(f'Sheet (the first keys) must be a str not {type(sheet)}')
            df[sheet] = {}
            try:
                for col in data[sheet].keys():
                    df[sheet].update({col:{}})
                    try:
                        for row,value in data[sheet][col].items():
                            if row < 1:
                                raise ValueError('row is less then 1, it must be 1 or greater')
                            value,data_type = self.get_type(value) # adds the datatype so vba knows how to convert it.
                            value = {value:data_type}
                            df[sheet][col].update({row:value})


                    except AttributeError:
                        raise ValueError('Missing row dictionary')
            except AttributeError:
                raise ValueError('Missing column dictionary')
        return df
    
    def df_to_dict(self,data:pd.DataFrame,kwargs:dict):
        sheet = kwargs.get('sheet')
        if sheet is None:
            raise ValueError('sheet parameter required if passing in DataFrame as data')
        start_row = kwargs.get('start_row')
        if start_row is None:
            start_row = 1
        start_col = kwargs.get('start_col')
        if start_col is None:
            start_col = 1
        elif isinstance(start_col,str):
            start_col = col_to_int(start_col)
        
        index = kwargs.get('index')
        if index is None:
            index = True
        elif isinstance(index,bool) is False:
            raise ValueError(f'Index must be a bool not {type(index)}')

        index_name = kwargs.get('index_name')
        if index_name is None:
            index_name = ''
            
            


        sheet = kwargs.get('sheet')
        df = {
            sheet:{}
        }
        if index:
            df,start_col = self.create_column(df,sheet,start_col,start_row,index_name,data.index)

        for col in data.columns:
            df,start_col = self.create_column(df,sheet,start_col,start_row,col,data[col])
        return df

    def create_column(self,df,sheet,start_col,start_row,header_name,data_list):
        col_letter = int_to_col(start_col)   
        df_column,row = self.add_header(start_row,col_letter,header_name)
        df_row = df_column[col_letter]
        df_column[col_letter].update(self.add_rows(data_list,df_row,row))
        df[sheet].update(df_column)
        start_col += 1
        return df,start_col

    def add_header(self,start_row,col_letter,value):
            row = start_row
            df_column = {
                col_letter:{
                    row:value
                }
            }
            row += 1
            return df_column,row

    def add_rows(self,data:pd.Series,df_row:dict,row):
        for value in data:
            df_row.update({
                row:value
            })
            row += 1
        return df_row
            

    def multiple_dfs_to_dict(self,data:list,kwargs:dict):
        sheet = kwargs.get('sheet')
        if sheet == None:
            raise ValueError('sheet parameter required if passing in DataFrame as data')
        elif isinstance(sheet,list) == False:
            raise ValueError('sheet parameter must be a list if passing in multiple dataframes')
        elif len(sheet) != len(data):
            raise ValueError(f'Not matching lengths of sheets and dataframes. sheet length is {len(sheet)} and list of dataframes is {len(data)}')
        start_row = kwargs.get('start_row')
        if start_row == None:
            start_row = 1
        elif isinstance(start_row,list):
            if len(start_row) != len(data):
                raise ValueError(f'Not matching lengths of start row and dataframes. start row length is {len(start_row)} and list of dataframes is {len(data)}')
        start_col = kwargs.get('start_col')
        if start_col is None:
            start_col = 1
        elif isinstance(start_col,str):
            start_col = col_to_int(start_col)
        elif isinstance(start_col,list):
            if len(start_col) != len(data):
                raise ValueError(f'Not matching lengths of start col and dataframes. start col length is {len(start_col)} and list of dataframes is {len(data)}')
        index = kwargs.get('index')
        if index is None:
            index = True
        elif isinstance(index,list):
            if len(index) != len(data):
                raise ValueError(f'Not matching lengths of index  and dataframes. index  length is {len(index)} and list of dataframes is {len(data)}')
        elif isinstance(index,bool) is False:
            raise ValueError(f'Index must be a bool not {type(index)}')

        index_name = kwargs.get('index_name')
        if index_name is None:
            index_name = ''
        elif isinstance(index,list):
            if len(index) != len(data):
                raise ValueError(f'Not matching lengths of index_name and dataframes. index_name length is {len(index_name)} and list of dataframes is {len(data)}')

        complete_df = {}
        for count,df in enumerate(data):
            params = {
                'sheet':sheet[count],
                'start_row' : start_row[count] if isinstance(start_row,list) else start_row,
                'start_col':start_col[count] if isinstance(start_col,list) else start_col,
                'index':index[count] if isinstance(index,list) else index,
                'index_name':index_name[count] if isinstance(index_name,list) else index_name
            }
            complete_df.update(
                self.df_to_dict(df,params)
            )
        
        return complete_df

            
    

    @staticmethod
    def get_type(value):
        if isinstance(value,int):
            data_type = 'Integer'
        elif isinstance(value,float):
            data_type = 'Float'
        elif isinstance(value,datetime):
            data_type = 'Datetime'
            value = value.strftime('%m-%d-%Y')
        else:
            data_type = 'String'
        return value,data_type

    def to_file(self,file):
        with open(file,'w') as file:
            file.write(json.dumps(self.data))

    def __add__(self,__o):
        self.data.update(__o.data)
    
    def __str__(self) -> str:
        return self.data
    
    def copy_to_clipboard(self):
        cmd=f'echo {json.dumps(self.data)}|clip'
        subprocess.check_call(cmd, shell=True)
