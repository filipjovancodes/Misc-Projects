Sub FormatOptionsData()

    Dim Index As Integer
    Dim i As Integer
    Dim row_num As Integer
    Dim col_num As Integer
    Dim table_row As Integer
    Dim table_col As Integer
    Dim Calls As Boolean
    Dim initial_row As Integer

    Application.ScreenUpdating = False
    
    ' Stop taking in values once a whitespace is reached
    NumRows = Range("A1", Range("A1").End(xlDown)).Rows.Count
    
    ' Initial values
    data_entry = 2
    table_row = 0
    
    ' Loop through all data entries
    For Index = 1 To NumRows

        ' If the string "Strike" is reached, then start the transition to listing puts
        If Cells(data_entry, 1).Value = "Strike" Then
            ' Set the table column to 11 to input the strikes
            table_col = 11
            
            ' Move the date to the header
            Cells(initial_row - 1, 11).Value = Cells(table_row, 2).Value
            Cells(table_row, 2).Value = Clear
            
            ' List the strikes
            For i = initial_row To table_row - 1
                Cells(i, table_col).Value = Cells(data_entry, 1).Value
                data_entry = data_entry + 1
                Next i
            
            data_entry = data_entry + 1 'Skip the "Puts" string
            table_col = 12 'Start of Puts section column
            'Reset the table row
            table_row = initial_row
            'Start aligning the table for pubs
            Calls = False
        End If
            
        ' Start a new table
        If Cells(data_entry, 1).Value = "Calls" Then
            ' Reset the table column
            table_col = 2
            ' Shift the new table 2 spaces down
            table_row = table_row + 2
            ' Store the current tables initial row
            initial_row = table_row
            ' Skip the "Calls" data entry
            data_entry = data_entry + 1
            ' Start aligning the table for calls
            Calls = True
        End If
        
        
        'Set the current table row and column value to the current data entry
        Cells(table_row, table_col).Value = Cells(data_entry, 1).Value
        data_entry = data_entry + 1 'Move to the next data entry
        table_col = table_col + 1 'Move to the next table column
        
        ' Calls table reallignment
        If Calls = True Then
            If table_col Mod 10 = 1 Then table_row = table_row + 1
            If table_col Mod 10 = 1 Then table_col = 2
        End If
        
        ' Puts table reallignment
        If Calls = False Then
            If table_col Mod 20 = 1 Then table_row = table_row + 1
            If table_col Mod 20 = 1 Then table_col = 12
        End If
        
        Next

    Application.ScreenUpdating = True
    
End Sub



