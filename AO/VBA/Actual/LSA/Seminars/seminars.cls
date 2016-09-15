VERSION 1.0 CLASS
BEGIN
  MultiUse = -1 
END

Attribute VB_Name = "Seminars"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private colSeminars As Collection

Private Sub Class_Initialize()
    Set colSeminars = New Collection
End Sub

Private Sub Class_Terminate()
    Set colSeminars = Nothing
End Sub

Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    Set NewEnum = colSeminars.[_NewEnum]
End Property

Public Sub Add(obj As SeminarUsers)
    colSeminars.Add obj
End Sub

Public Sub Remove(Index As Variant)
    colSeminars.Remove Index
End Sub

Public Property Get Item(Index As Variant) As SeminarUsers
    Attribute Item.VB_UserMemId = 0
    Set Item = colSeminars.Item(Index)
End Property

Property Get Count() As Long
    Count = colSeminars.Count
End Property

Public Sub Clear()
    Set colSeminars = New Collection
End Sub

Public Sub FillFromSheet(wks As Worksheet)
    Const cFirstRow = 2, cPersonName = 1, cEduDate = 12, cSeminarName = 14, cEducater = 11

    Dim i As Long, obj As SeminarUsers
    With wks
        For i = cFirstRow To .Cells(Rows.Count, 1).End(xlUp).Row
            Set obj = New SeminarUsers
            obj.PersonName = .Cells(i, cPersonName)
            obj.EduDate = .Cells(i, cEduDate)
            obj.SeminarName = LCase(.Cells(i, cSeminarName))
            obj.Educater = .Cells(i, cEducater)

            Me.Add obj
        Next
    End With
End Sub

Public Function FilterBySeminar(str As String) As SeminarUsers
    Dim smr As seminars, smr_type As SeminarUsers
    Set smr = New seminars

    For Each smr_type In Me
        If smr_type.SeminarName = str Then smr.Add smr_type
    Next

    Set FilterBySeminar = smr
End Function

Public Function FilterByNameLike(str As String) As SeminarUsers
    Dim smr As seminars, smr_type As SeminarUsers
    Set smr = New seminars

    For Each smr_type In Me
        If smr_type.Name Like str Then smr.Add smr_type
    Next

    Set FilterByNameLike = smr
End Function
