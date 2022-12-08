import json
from datetime import datetime
import pandas as pd
import string
import subprocess 
from config_script import config_python_to_vba

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

        NOTE: as of now indexes are not reflected in the vba creation

        
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
        if sheet == None:
            raise ValueError('sheet parameter required if passing in DataFrame as data')
        start_row = kwargs.get('start_row')
        if start_row == None:
            start_row = 1
        sheet = kwargs.get('sheet')
        df = {
            sheet:{}
        }
        
        for count,col in enumerate(data.columns):
            row = start_row
            col_letter = string.ascii_uppercase[count]
            df_column = {
                col_letter:{
                    row:col
                }
            }
            row += 1
            df_row = df_column[col_letter]
            for value in data[col]:
                df_row.update({
                    row:value
                })
                row += 1
            df[sheet].update(df_column)
        return df
                

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
        complete_df = {}
        for count,df in enumerate(data):
            params = {
                'sheet':sheet[count],
                'start_row' : start_row[count] if isinstance(start_row,list) else start_row
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

import random
def make_dict():
    sheets = ['Sheet1','Sheet2']
    cols = ['A','B','C','D','E','F','G','H','I','J']
    rows = [x for x in range(1,1000)]
    df = {}
    for sheet in sheets:
        df[sheet] = {}
        for col in cols:
            df[sheet].update({col:{}})
            for row in rows:
                df[sheet][col].update({row:random.randint(0,1000)})
    return df
if __name__ == '__main__':
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
    v = VBA()
    VBA()
    config_python_to_vba()
    # first = {
    #     'sheet1':{