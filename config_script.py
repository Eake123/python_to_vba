import sys,os





def write_to_file(filename,data_str):
    with open(filename,'w') as file:
        file.write(data_str)
    return filename

def config_python_to_vba(
    script:str,
    file_or_clip=False,
    directory=None
    ):
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

    file_no_extension = os.path.splitext(os.path.abspath(script))[0]

    bat_file = construct_bat(script,directory,file_no_extension)
    construct_bas(file_or_clip,bat_file,file_no_extension,directory)


def construct_bas(file_or_clip,bat_file,file_no_extension,directory):

    if isinstance(file_or_clip,str):
        vba_str = construct_vba_file(file_or_clip,bat_file,file_no_extension)
    else:
        vba_str = construct_vba_clip(bat_file,file_no_extension)
    bas = f'{file_no_extension}.bas'
    if directory is not None:
        bas = f'{directory}\\{os.path.basename(bas)}'
    write_to_file(bas,vba_str)


def construct_vba_clip(bat_file,file_no_extension):
    vba_str = f'''Sub {os.path.basename(file_no_extension)}_clip()
    Dim bat_file As String
    bat_file = "{bat_file}"
    Dim get_from_file As Boolean
    get_from_file = False
    Call PythonConverter(filename, bat_file, get_from_file)
    End Sub'''
    return vba_str

def construct_vba_file(file,bat_file,file_no_extension):
    vba_str = f'''Sub {os.path.basename(file_no_extension)}_file()
    Dim filename As String
    filename = "{file}"
    Dim bat_file As String
    bat_file = "{bat_file}"
    Dim get_from_file As Boolean
    get_from_file = True
    Call PythonConverter(filename, bat_file, get_from_file)
    End Sub'''
    return vba_str
    


def construct_bat(script,directory,file_no_extension):
    python = sys.executable
    file = os.path.abspath(script)
    bat = f'{file_no_extension}.bat'
    if directory is not None:
        file = f'{directory}\\{os.path.basename(file)}'
        bat = f'{directory}\\{os.path.basename(bat)}'
        
    
    bat_str = f'@echo off\n"{python}" "{file}"'
    return write_to_file(bat,bat_str)
    


if __name__ == '__main__':
    config_python_to_vba()

