Attribute VB_Name = "Module1"

Sub test()
    Dim ppl As People, per As Person
    Set ppl = New People
    ppl.FillFromSheet ActiveSheet

  Debug.Print "Test 4: return all People of a specific city and similar name"
  
    For Each per In ppl.FilterBySeminar("7 Шагов").FilterByFirstNameLike("Абдулина*")
    
    Debug.Print per.FirstName; vbTab; per.Seminar; vbTab; per.smr_date
    
    Next
    
    
    
End Sub
