Attribute VB_Name = "Module1"
Private Sub LockMergeCell()
For Each cell In ThisWorkbook.ActiveSheet.Range("A1:AM65").Cells
    If cell.MergeCells = True Then
        Set mrng = cell.MergeArea
        
        If mrng.Interior.ColorIndex = 36 Then
            mrng.Locked = False
            mrng.Hidden = False
        Else
            mrng.Locked = True
            mrng.Hidden = True
        End If
    Else
    
        If cell.Interior.ColorIndex = 36 Then
            cell.Locked = False
            cell.Hidden = False
        Else
            cell.Locked = True
            cell.Hidden = True
        
        End If
    
    End If
    
Next cell

End Sub

