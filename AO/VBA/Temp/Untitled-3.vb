Option Compare Text
Function GreatSheet (newSh as String) as String 

For Each Sh In ThisWorkbook.Worksheets
If Sh.Name = newSh Then
chek_name = 1
End If
Next Sh

If chek_name <> 1 Then
Set Sh = Worksheets.Add()
Sh.Name = newSh
End If

End Function
'--------------------------------------------------------------------------------------------------------- 


Sub getG()

Dim start_cell as String
Dim hr&, hc&
Dim NF as String


start_cell = "LOR mSeries Store"
hr = 3
hc = 2

'---------------------------------------------------------------------------------------------------------
NF = ActiveSheet.Name


new_rediz_Sheet = NF & "_" & "redz"



'---------------------------------------------------------------------------------------------------------
