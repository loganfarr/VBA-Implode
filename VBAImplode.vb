Sub IMPLODE()
    Call IMPLODE_FXN(Sheet1.Range("A2"))
End Sub
Function IMPLODE_FXN(start As Range)

    'Defines and sets beginning values for the base variables
    Dim count As Integer 'Number of pictures one ID has
    Dim unique_ids As Integer 'The number of unique IDs
    Dim id As Integer 'ID of the picture
    Dim start_row As Integer 'The starting row for the comparison
    
    count = 0 'The amount of pictures one ID has
    id = start.value 'Gets the beginning ID
    'id = 7971
    unique_ids = 2 'Number used to calculate what row the new ID should go on
    start_row = start.row 'Gets the row of the starting cell
    
    Dim temp_id As Integer 'The temporary ID for comparing
    Dim temp_a_range As String 'Temporary Holder for A column cells
    Dim temp_b_range As String 'Temporary holder for B column cells
    
    'Gets the rows in one worksheet
    Dim a As Range, b As Range
    Set a = Sheet1.Rows
    
    'Defines the current row variables
    Dim current_row As Integer
    
    'Defines the string of picture URLs
    Dim pictures As String
    
    current_row = start_row 'Starts the current row on the starting cell's row
    
    'For each row in the sheet
    'For Each b In a.Rows
    For i = 0 To 5000 '5000 is just a random number to make sure it goes through all of it
        Dim temp_range As String
        temp_range = "A" & current_row
        'temp_id = Worksheets("pictures").Range(temp_range).value 'Temporary ID
        temp_id = Sheet1.Range(temp_range).value 'Temporary ID
       
        'If the temporary ID is the same as the selected ID
        If temp_id = id Then
            count = count + 1 'Adding to the number of times the ID appears
            current_row = current_row + 1 'Moving down to the next row
            
            temp_b_range = "B" & current_row
            Sheet1.Range(temp_b_range).value = ""
        Else 'If the ID is different
            unique_ids = unique_ids + 1
            pictures = ""
            For j = 1 To count 'For each picture
                pictures = pictures & "public://pics/" & id & "-" & j & ".jpg|" 'Add picture-file text
            Next j
            
            temp_a_range = "A" & unique_ids
            temp_b_range = "B" & unique_ids
            
            Sheet1.Range(temp_a_range).value = temp_id
            Sheet1.Range(temp_b_range).value = pictures 'Inserts picture text
            
            count = 0 'Reset count
            start_row = current_row 'Move on to the next row
            id = temp_id
        End If
    Next
    
End Function
