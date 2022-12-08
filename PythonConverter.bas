
Public Function PythonConverter(filename As String, bat_file As String, get_from_file As Boolean)
    'main function that runs the bat file, reads the data, and adds it into the worksheet
    Dim file_str As String
    Call run_code(bat_file)
    If get_from_file = True Then
        file_str = read_file(filename)
    Else
        file_str = read_clipboard
    End If
    Call loop_dictionary(file_str)
    
    
    
    


End Function


Private Function run_code(bat_file As String)
    'runs the bat file, waits until it is complete as well
    Dim wsh As Object
    Set wsh = VBA.CreateObject("WScript.Shell")
    Dim waitOnReturn As Boolean: waitOnReturn = True
    Dim windowStyle As Integer: windowStyle = 1
    wsh.Run bat_file, windowStyle, waitOnReturn
End Function

Private Function read_file(filename As String) As String
    'reads the file generated in the bat file
    Dim FSO As New FileSystemObject
    Set FSO = CreateObject("Scripting.FileSystemObject")
    Set FileToRead = FSO.OpenTextFile(filename) 'add here the path of your text file
    
    TextString = FileToRead.ReadAll
    
    FileToRead.Close
    read_file = TextString
End Function

Private Function read_clipboard() As String
    'reads the clipboard after the bat file is run

'Tools -> References -> Microsoft Forms 2.0 Object Library
'of you will get a "Compile error: user-defined type not defined"
  Dim DataObj As New MSForms.DataObject
  Dim S As String
  DataObj.GetFromClipboard
  S = DataObj.GetText
  read_clipboard = S 'print code in the Intermediate box in the Macro editor
End Function


Private Function loop_dictionary(file_str As String)
        'loops through the dictionary
        Dim Json As Dictionary
        Set Json = JsonConverter.ParseJson(file_str)
        Dim row_dict As Dictionary
        Dim col_dict As Dictionary
        Dim value_dict As Dictionary
        Dim data_type As String
        Dim value_typing As Variant
        For Each sheet In Json.Keys
            'gets the sheet name
            Set col_dict = Json(sheet) 'initializes the next nest of the dictionary with the column letters
            For Each col In col_dict.Keys
                'gets the column letter
                Set row_dict = col_dict(col) 'initializes the next nest of the dictionary with the row numbers
                For Each row In row_dict.Keys
                    'gets the row number
                    Set value_dict = row_dict(row) 'initializes the next nest of the dictionary with the value variant
                    For Each Value In value_dict.Keys
                        'gets the value variant
                        data_type = value_dict(Value) 'gets the data type value
                        value_typing = get_typing(Value, data_type) 'converts the value from a string to the designated value
                        Call add_to_sheet(CStr(sheet), CStr(col), CStr(row), value_typing) 'sets the col and row as a string so it can add them together to make the cell ex. A + 1 = A1
                        
                        Next Value
                        
                    
                    Next row
                    
                    
                    
                
                Next col
                
            
            
            
            
            
                
                    

        Next sheet
    
End Function

Private Function get_typing(Value As Variant, data_type As String) As Variant
    'changes the variant to the type designated in the json
    Dim value_type As Variant
    If data_type = "Integer" Then
        value_type = CInt(Value)
    ElseIf data_type = "Float" Then
        value_type = CLng(Value)
    ElseIf data_type = "datetime" Then
        value_type = CDate(Value)
    Else
        value_type = CStr(Value)
    End If
    
    get_typing = value_type
End Function


Private Function add_to_sheet(sheet As String, col As String, row As String, value_typing As Variant)
    'concatenates the column letter and the row number and adds the value to excel
    Dim cell_cord As String
    cell_cord = col + row
    
    Sheets(sheet).Range(cell_cord).Value = value_typing
    
End Function




