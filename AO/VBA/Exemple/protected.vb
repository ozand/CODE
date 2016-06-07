
Sub chekF()
Dim iUnionRange As Range

ActiveSheet.Unprotect
Cells.Locked = False

For Each cell In ActiveSheet.UsedRange
 
    
If cell.HasFormula Then
    
    If iUnionRange Is Nothing Then
       Set iUnionRange = Union(cell, cell)
    Else
       Set iUnionRange = Union(iUnionRange, cell)
    End If

End If

Next cell

If Not iUnionRange Is Nothing Then iUnionRange.Select
Selection.Locked = True
Selection.FormulaHidden = False


 ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=false



End Sub
