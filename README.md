This Script builds on the work of VBA-Json where you can find a link to it here https://github.com/VBA-tools/VBA-JSON 

What it does.

The purpose of this script is to output data from a python script into vba. It achieves this with the VBA class that reads either a nested dictionary
(see below for the required format), pandas DataFrame, or list of pandas DataFrame. It also creates a bas and bat file to run your custom script in VBA.

After downloading this script import the JsonConverter.bas and Pythonconverter.bas files into the excel file in the developer mode of excel. 

You will then need to add the reference Microsoft Scripting Runtime if you are reading from a file with the VBA().to_file(). https://www.automateexcel.com/vba/using-the-filesystemobject-in-excel-vba/

If reading from the clipboard using VBA().copy_to_clipboard() you will need to add the reference Microsoft Forms 2.0 Object Library


The VBA Class works as follows


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
        
        
   At the end of these examples you will need to either add v.to_file(file_name) or v.copy_to_clipboard()
   
   Once this is complete you will then need to configure the bat and bas file as follows
   
       '''
    PARAMS

    script: must be __file__

    file_or_clip: if you are using VBA().to_file() to run the script you must make this the file name that goes into the to_file(). 

    directory: OPTIONAL this makes the directory for the bat file, and the bas file 

    EXAMPLE 1: making a configuration for a python script that creates a file with VBA().to_file() in the current directory

    config_python_to_vba(__file__,'test.json')

    EXAMPLE 2: making a configuration for a python script that creates a file with VBA().to_clipboard() in C:\\Users

    config_python_to_vba(__file__,directory='C:\\Users')
    
    '''
    
    
    The last thing you will need to do is import the bas file created from this into the excel sheet. The bas file generated is named after the script that runs the config_python_to_vba.
    
    
    
If you want to send params to a batch file make your Sub vba script looks like this
```
Sub python_to_vba_file()
    Dim filename As String
    filename = "C:\Users\erikj\Desktop\scripts\python-vba\test.json"
    Dim bat_file As String
    Dim param1 As String
    param1 = ""
    Dim param2 As String
    param2 = ""
    bat_file = "c:\Users\python_to_vba.bat" & " " + param2 + " " + param1
    Dim get_from_file As Boolean
    get_from_file = True
    Call PythonConverter(filename, bat_file, get_from_file)
    End Sub
 ```
 Where param1 and param2 are sent to your python script.
 
 Make your batch file look like this to accept them
 ```
 @echo off
set arg1=%1
set arg2=%2
"C:\Users\python.exe" "c:\Users\python_to_vba.py" %1 %arg1% %2 %arg2%
```

   
   
   
   
   
