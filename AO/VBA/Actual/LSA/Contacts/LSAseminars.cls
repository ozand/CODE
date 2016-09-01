VERSION 1.0 CLASS
BEGIN
  MultiUse = -1 
END
Attribute VB_Name = "LSAseminars"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private objLSAsem As Collection

Private Sub Class_Initialize()
    Set objLSAsem = New Collection
End Sub

Private Sub Class_Terminate()
    Set objLSAsem = Nothing
End Sub

Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    Set NewEnum = objLSAsem.[_NewEnum]
End Property

Public Sub Add(obj As myLSA)
    objLSAsem.Add obj
End Sub

Public Sub Remove(Index As Variant)
    objLSAsem.Remove Index
End Sub

Public Property Get Item(Index As Variant) As myLSA
Attribute Item.VB_UserMemId = 0
    Set Item = objLSAsem.Item(Index)
End Property

Property Get Count() As Long
    Count = objLSAsem.Count
End Property

Public Sub Clear()
    Set objLSAsem = New Collection
End Sub

Public Sub FillFromSheet(wks As Worksheet)
    Const cFirstRow = 2, cName = 1, cEduDate = 12, cSeminarName = 14, cEducater = 11

    Dim i As Long, obj As myLSA
    With wks
        For i = cFirstRow To .Cells(Rows.Count, 1).End(xlUp).Row
            Set obj = New myLSA
            obj.Name = .Cells(i, cName)
            obj.EduDate = .Cells(i, cEduDate)
            obj.SeminarName = .Cells(i, cSeminarName)
            obj.Educater = .Cells(i, cEducater)

            Me.Add obj
        Next
    End With
End Sub

Public Function FilterBySeminar(str As String) As myLSA
    Dim smr As myLSA, smr_type As myLSA
    Set smr = New myLSA

    For Each smr_type In Me
        If smr_type.Seminar = str Then smr.Add smr_type
    Next

    Set FilterBySeminar = smr
End Function

Public Function FilterByNameLike(str As String) As myLSA
    Dim smr As myLSA, smr_type As myLSA
    Set smr = New myLSA

    For Each smr_type In Me
        If smr_type.Name Like str Then smr.Add smr_type
    Next

    Set FilterByNameLike = smr
End Function
