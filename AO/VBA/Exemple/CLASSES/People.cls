VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "People"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private objPeople As Collection

Private Sub Class_Initialize()
    Set objPeople = New Collection
End Sub

Private Sub Class_Terminate()
    Set objPeople = Nothing
End Sub

Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    Set NewEnum = objPeople.[_NewEnum]
End Property

Public Sub Add(obj As Person)
    objPeople.Add obj
End Sub

Public Sub Remove(Index As Variant)
    objPeople.Remove Index
End Sub

Public Property Get Item(Index As Variant) As Person
Attribute Item.VB_UserMemId = 0
    Set Item = objPeople.Item(Index)
End Property

Property Get Count() As Long
    Count = objPeople.Count
End Property

Public Sub Clear()
    Set objPeople = New Collection
End Sub
Public Sub FillFromSheet(wks As Worksheet)
    Const cFirstRow = 2, cStreetGeo = 28, cCityGeo = 49, cSectoreCom = 40, cRegionGeo = 59
    Dim i As Long, obj As Person
    With wks
        For i = cFirstRow To .Cells(Rows.Count, 1).End(xlUp).Row
            Set obj = New Person
            obj.FirstName = .Cells(i, cFirstNameCol)
            obj.Seminar = .Cells(i, cSeminarCol)
            obj.smr_date = .Cells(i, cDateCol)

            Me.Add obj
        Next
    End With
End Sub

Public Function FilterBySeminar(str As String) As People
    Dim ppl As People, per As Person
    Set ppl = New People

    For Each per In Me
        If per.Seminar = str Then ppl.Add per
    Next

    Set FilterBySeminar = ppl
End Function

Public Function FilterByFirstNameLike(str As String) As People
    Dim ppl As People, per As Person
    Set ppl = New People

    For Each per In Me
        If per.FirstName Like str Then ppl.Add per
    Next

    Set FilterByFirstNameLike = ppl
End Function
