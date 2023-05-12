# VBA-challenge
# Module 1
Sub Mod_2_Assign()
    Dim i As Long
    Dim j As Long
    Dim g As Long
    Dim total_vol As LongLong
    Dim y_open, y_close As Double
    Dim orig_ticker As String
    Dim rng As Range
    Dim condition_1, condition_2 As FormatCondition
    
    ' Assume there is no missing data.
    ' Assume there are at least 2 pieces of data for each company's stock
    
    j = 2
    
    For i = 2 To 759001
        orig_ticker = Cells(i, 1).Value
        
        If (orig_ticker <> Cells(i + 1, 1).Value) Then
            y_close = Cells(i, 6).Value
            Cells(j - 1, 10).Value = y_close - y_open
            Cells(j - 1, 11).Value = Round((((y_close - y_open) / y_open) * 100), 2)
            Cells(j - 1, 12).Value = total_vol + Cells(i, 7).Value
            
        ElseIf (orig_ticker = Cells(i - 1, 1)) Then
              total_vol = total_vol + Cells(i, 7).Value
        Else
            Cells(j, 9).Value = orig_ticker
            y_open = Cells(i, 3).Value
            total_vol = Cells(i, 7).Value
            j = j + 1
        End If
         
    Next i
   
   'Conditional Formatting For Yearly Change
    For g = 2 To j
        If (Cells(g, 10).Value > 0) Then
            Cells(g, 10).Interior.ColorIndex = 4
        ElseIf (Cells(g, 10).Value < 0) Then
            Cells(g, 10).Interior.ColorIndex = 3
        End If
    Next
    
   'Conditional Formatting For Percent Change
    For g = 2 To j
        If (Cells(g, 11).Value > 0) Then
            Cells(g, 11).Interior.ColorIndex = 4
        ElseIf (Cells(g, 10).Value < 0) Then
            Cells(g, 11).Interior.ColorIndex = 3
        End If
    Next
    
    For g = 2 To j
        If (g = 2) Then
            Cells(2, 16).Value = Cells(g, 9).Value
            Cells(2, 17).Value = Cells(g, 11).Value
            
            Cells(3, 16).Value = Cells(g, 9).Value
            Cells(3, 17).Value = Cells(g, 11).Value
            
            Cells(4, 16).Value = Cells(g, 9).Value
            Cells(4, 17).Value = Cells(g, 12).Value
        ElseIf (g > 2) Then
            If (Cells(g, 11).Value > Cells(2, 17).Value) Then
                Cells(2, 16).Value = Cells(g, 9).Value
                Cells(2, 17).Value = Cells(g, 11).Value
            End If
            If (Cells(g, 11).Value < Cells(3, 17).Value) Then
                Cells(3, 16).Value = Cells(g, 9).Value
                Cells(3, 17).Value = Cells(g, 11).Value
            End If
            If (Cells(g, 12).Value > Cells(4, 17).Value) Then
                Cells(4, 16).Value = Cells(g, 9).Value
                Cells(4, 17).Value = Cells(g, 12).Value
            End If
            
        End If
     Next g
     
    
End Sub

# Module 2
Sub Mod_2_Assign_Loop()

    Dim ws As Worksheet

    For Each ws In Worksheets
        ws.Activate
        Call Module1.Mod_2_Assign
    Next

End Sub
