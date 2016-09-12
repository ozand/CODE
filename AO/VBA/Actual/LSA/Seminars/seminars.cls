VERSION 1.0 CLASS
BEGIN
  MultiUse = -1 
END
Attribute VB_Name = “seminars”
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
Attribute NewEnum.VB_MemberFlags = “40”
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
