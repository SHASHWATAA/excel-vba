Attribute VB_Name = "Module2"
Function LastRow() As Integer
    LastRow = Cells(Rows.Count, 1).End(xlUp).Row
End Function



Sub datafill()

    last_row = LastRow()
    
    active_column = ActiveCell.Column
    
    
    For i = 1 To last_row
    
        If Cells(i + 1, active_column) = "" Then
            Cells(i + 1, active_column) = Cells(i, active_column)
        End If
        
    Next i
 
End Sub
