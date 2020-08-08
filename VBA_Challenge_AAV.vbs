Sub ticker_summary()
    
    'Define variables
    ' - - - - - - - - -- - - -  -- - -  -- - - - - - - -  --  - -- - - - -  - - --
    'Loop variables
    Dim i As Integer 'count the acive wbs
    Dim j As LongLong 'loop for normal homework
    Dim k As LongLong 'loop for challenge
    ' - - - - - - - - -- - - -  -- - -  -- - - - - - - -  --  - -- - - - -  - - --
    'Variable to count the leght of the df (number of rows)
    Dim nrow As Variant 'count rows in general homework
    Dim ncol As Variant 'count colums just to check dimension of df
    Dim nrow_c As Variant 'count the length of the rows of the tciker summary (for the challenge)
    ' - - - - - - - - -- - - -  -- - -  -- - - - - - - -  --  - -- - - - -  - - --
    'Counters
    Dim conta_1 As Variant 'Count how many times there is a change in the name of an action
    Dim conta_2 As Variant 'Count how many times there is no change in the name of an action
    Dim conta_3 As Variant 'It is not properly a counter but it will sum the volume of equal name cases
    ' - - - - - - - - -- - - -  -- - -  -- - - - - - - -  --  - -- - - - -  - - --
    'Variables to temporarly keep values
    Dim op As Variant 'To keep the opoening price of a given ticker
    Dim cp As Variant 'To keep the close price of a given ticker
    Dim tv As Variant 'To keep the total volume of a given ticker
    ' - - - - - - - - -- - - -  -- - -  -- - - - - - - -  --  - -- - - - -  - - --
    'Variables for the challenge
    Dim gpi As Variant 'To keep the maximun percent change
    Dim gpd As Variant 'To keep the minimum percent change
    Dim gtv As Variant 'To keep the maximum total volum
    Dim ngpi As Variant 'To keep the name of the ticker with the maximun percent change
    Dim ngpd As Variant 'To keep the name of the ticker with the minimum percent change
    Dim ngtv As Variant 'To keep the name of the ticker with the maximum total volume
    
    
    'Dim fila_1 As Variant 'Guardara el número de fila en donde escribir el resultado
    
    
    'Initiate loop to count throug wbs
    'Set WS_Count equal to the number of worksheets in the active workbook.
    WS_Count = ActiveWorkbook.Worksheets.Count

    ' Begin the loop for worksheets
    For i = 1 To WS_Count
        'i = 1
        Worksheets(i).Activate
        MsgBox ("Hoja" & Str(i))
        
        ' - - - - - - - - -- - - -  -- - -  -- - - - - - - -  --  - -- - - - -  - - --
        'Count number of rows with info
        nrow = Cells(Rows.Count, 1).End(xlUp).Row
        MsgBox ("There are " & Str(nrow) & " rows in this file")
        'Count number of columns with info
        ncol = Cells(1, Columns.Count).End(xlToLeft).Column
        MsgBox ("There are " & Str(ncol) & " cols in this file")
        
        
        ' - - - - - - - - -- - - -  -- - -  -- - - - - - - -  --  - -- - - - -  - - --
        'Write names of the columns
        
        Cells(1, 9) = "Ticker"
        Cells(1, 10) = "Yearly Change"
        Cells(1, 11) = "Percent Change"
        Cells(1, 12) = "Total Stock Volume"
        
        
        ' - - - - - - - - -- - - -  -- - -  -- - - - - - - -  --  - -- - - - -  - - --
        'Define initial values of counter and "cumulative" sums variables
        conta_1 = 2
        conta_2 = 0
        conta_3 = 0
        ' - - - - - - - - -- - - -  -- - -  -- - - - - - - -  --  - -- - - - -  - - --
        'Begin loop for regular homework
        ' - - - - - - - - -- - - -  -- - -  -- - - - - - - -  --  - -- - - - -  - - --
        For j = 2 To nrow
            ' - - - - - - - - -- - - -  -- - -  -- - - - - - - -  --  - -- - - - -  - - --
            'Ask if the name of the ticker in a cell is the same as the next cell
            If Cells(j, 1) <> Cells(j + 1, 1) Then
                'Define the row where to start writing the info
                            
                ' - - - - - - - - -- - - -  -- - -  -- - - - - - - -  --  - -- - - - -  - - --
                'Name of ticker
                Cells(conta_1, 9) = Cells(j - conta_2, 1) 'Extract the name of the ticker
                'Cells(conta_1, 10) = Cells(j, 1) 'Ultima
                
                
                ' - - - - - - - - -- - - -  -- - -  -- - - - - - - -  --  - -- - - - -  - - --
                'Extract the first opening price
                op = Cells(j - conta_2, 3) 'Precio apertura(Esta se debe quedar)
                ' - - - - - - - - -- - - -  -- - -  -- - - - - - - -  --  - -- - - - -  - - --
                'Extract the last close price
                cp = Cells(j, 6) 'Precio apertura(Esta se debe quedar)
                
                ' - - - - - - - - -- - - -  -- - -  -- - - - - - - -  --  - -- - - - -  - - --
                'Yearly Change
                Cells(conta_1, 10) = cp - op
                
                ' - - - - - - - - -- - - -  -- - -  -- - - - - - - -  --  - -- - - - -  - - --
                'Conditional formating if positive green if negative red
                If Cells(conta_1, 10) < 0 Then
                    Cells(conta_1, 10).Interior.ColorIndex = 3 'red
                ElseIf Cells(conta_1, 10) > 0 Then
                    Cells(conta_1, 10).Interior.ColorIndex = 4 'green
                Else
                    Cells(conta_1, 10).Interior.ColorIndex = 2 'no color
                End If
                
                ' - - - - - - - - -- - - -  -- - -  -- - - - - - - -  --  - -- - - - -  - - --
                
                'Percent Change
                'To avoid dividing by zero
                If op <> 0 Then
                    Cells(conta_1, 11) = Format((cp - op) / op, "Percent")
                Else
                    Cells(conta_1, 11) = Format(0, "Percent")
                End If
                
                ' - - - - - - - - -- - - -  -- - -  -- - - - - - - -  --  - -- - - - -  - - --
                'Sum the total volume
                tv = conta_3
                Cells(conta_1, 12) = tv + Cells(j, 7)
                
                ' - - - - - - - - -- - - -  -- - -  -- - - - - - - -  --  - -- - - - -  - - --
                'Reset or fix counters
                'Add one to counter of number of different names of ticker
                conta_1 = conta_1 + 1
                'Reset counter of negative cases (equal ticker name)
                conta_2 = 0
                'Reset sum of total volume
                conta_3 = 0
                ' - - - - - - - - -- - - -  -- - -  -- - - - - - - -  --  - -- - - - -  - - --
            Else
                ' - - - - - - - - -- - - -  -- - -  -- - - - - - - -  --  - -- - - - -  - - --
                'Cells(j, 22) = "equal"
                'Counts how many times enter to the else (how many observations -1 are of a given ticker
                conta_2 = conta_2 + 1
                'Add the volume of a given ticker before the name of the ticker changes
                conta_3 = conta_3 + Cells(j, 7)
                
            End If
            
        Next j
        
            
        ' - - - - - - - - -- - - -  -- - -  -- - - - - - - -  --  - -- - - - -  - - --
        'Lets do the challenge
        'Look for the max of % the min of % and the max of volume
        'Lets define the length of the row with the ticker summary
        nrow_c = Cells(Rows.Count, 10).End(xlUp).Row
        ' - - - - - - - - -- - - -  -- - -  -- - - - - - - -  --  - -- - - - -  - - --
        'Define initial values for variables
        gpi = Cells(2, 11)
        gpd = Cells(2, 11)
        gtv = Cells(2, 12)
        ngpi = Cells(2, 9)
        ngpd = Cells(2, 9)
        ngtv = Cells(2, 9)
        
        ' - - - - - - - - -- - - -  -- - -  -- - - - - - - -  --  - -- - - - -  - - --
        'Start the challenge group
        
        For k = 2 To nrow_c
            ' - - - - - - - - -- - - -  -- - -  -- - - - - - - -  --  - -- - - - -  - - --
            'Search minimun percent
            If Cells(k, 11) < gpd Then
                gpd = Cells(k, 11)
                ngpd = Cells(k, 9)
            End If
            
            '- - - - - - - - -- - - -  -- - -  -- - - - - - - -  --  - -- - - - -  - - --
            'Search maximum percent
              If gpi < Cells(k, 11) Then
                gpi = Cells(k, 11)
                ngpi = Cells(k, 9)
            End If
            
            '- - - - - - - - -- - - -  -- - -  -- - - - - - - -  --  - -- - - - -  - - --
            'Search maximum volume
              If gtv < Cells(k, 12) Then
                gtv = Cells(k, 12)
                ngtv = Cells(k, 9)
            End If
               
        
        Next k
        
        
        'Lets put the names
        Cells(1, 15) = "Ticker"
        Cells(1, 16) = "Value"
        Cells(2, 14) = "Greatest % increase"
        Cells(3, 14) = "Greates % decrease"
        Cells(4, 14) = "Greatest Total Volume"
            
        'Lets fill the info
        Cells(2, 15) = ngpi
        Cells(2, 16) = Format(gpi, "Percent")
        Cells(3, 15) = ngpd
        Cells(3, 16) = Format(gpd, "Percent")
        Cells(4, 15) = ngtv
        Cells(4, 16) = gtv
        
        'fORMATING
        Range("I1:L1").Columns.AutoFit
        Range("N3").Columns.AutoFit
        Range("O1").Columns.AutoFit
        Range("P3").Columns.AutoFit
        
        
        
    Next i
    




End Sub




