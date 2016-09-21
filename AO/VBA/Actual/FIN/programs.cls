VERSION 1.0 CLASS
BEGIN
  MultiUse = -1 
END
Attribute VB_Name = "programs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private objProgram As Collection

Private Sub Class_Initialize()
    Set objProgram = New Collection
End Sub

Private Sub Class_Terminate()
    Set objProgram = Nothing
End Sub

Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    Set NewEnum = objProgram.[_NewEnum]
End Property

Public Sub Add(obj As program)
    objProgram.Add obj
End Sub

Public Sub Remove(Index As Variant)
    objProgram.Remove Index
End Sub

Public Property Get Item(Index As Variant) As program
Attribute Item.VB_UserMemId = 0
    Set Item = objProgram.Item(Index)
End Property

Property Get Count() As Long
    Count = objProgram.Count
End Property

Public Sub Clear()
    Set objProgram = New Collection
End Sub

Public Sub FillFromSheet(wks As Worksheet)
    Const cFirstRow = 1, cMonth = 1, cCd_partners = 2, cBrand = 3, cType_vl = 4, cS_group = 6, cVal = 7

    Dim i As Long, obj As program
    With wks
        For i = cFirstRow To .Cells(Rows.Count, 1).End(xlUp).Row
            Set obj = New program
            obj.month = .Cells(i, cMonth)
            obj.cd_partners = .Cells(i, cCd_partners)
            obj.brand = .Cells(i, cBrand)
            obj.type_vl = .Cells(i, cType_vl)
            obj.s_group = .Cells(i, cS_group)
            obj.val = .Cells(i, cVal)
            
            Me.Add obj
        Next
    End With
End Sub


