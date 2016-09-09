Option Explicit
Public sValue As String
Private Sub Workbook_SheetChange(ByVal Sh As Object, ByVal Target As Range)
    If Sh.Name = "LOG" Then Exit Sub
    Dim sLastValue As String
    Dim lLastRow As Long, wbLOG As Workbook
    Dim sPath as String
    Const sLOGName As String = "\LOG.txt" '"\LOG.xls"
    sPath = Application.DefaultFilePath
    Application.ScreenUpdating = False
    '==============   только для записи в текстовый файл   ======================
    If Dir(sPath & sLOGName, vbDirectory) = "" Then
        Open sPath & sLOGName For Output As #1: Close #1
    End If
    '==============   только для записи в отдельный файл Excel ======================
'    If Dir(sPath & sLOGName, vbDirectory) = "" Then
'        Set wbLOG = Workbooks.Add
'        wbLOG.SaveAs sPath & sLOGName, xlNormal
'    End If
    Set wbLOG = Workbooks.Open(sPath & sLOGName)
    '============================================================================
    With wbLOG.Sheets(1)
        lLastRow = .Cells.SpecialCells(xlLastCell).Row + 1
        If lLastRow = .Rows.Count Then Exit Sub
        Application.ScreenUpdating = False: Application.EnableEvents = False
        .Cells(lLastRow, 1) = CreateObject("wscript.network").UserName
        .Cells(lLastRow, 2) = Target.Address(0, 0)
        .Cells(lLastRow, 3) = Format(Now, "dd.mm.yyyy HH:MM:SS")
        .Cells(lLastRow, 4) = Sh.Name
        .Cells(lLastRow, 5).NumberFormat = "@"
        .Cells(lLastRow, 5) = sValue
        If Target.Count > 1 Then
            Dim rCell As Range, rRng As Range
            On Error Resume Next
            Set rRng = Intersect(Target, Sh.UsedRange): On Error GoTo 0
            If Not rRng Is Nothing Then
                For Each rCell In rRng
                    If Not IsError(Target) Then sLastValue = sLastValue & "," & rCell Else sLastValue = sLastValue & "," & "Err"
                Next rCell
                sLastValue = Mid(sLastValue, 2)
            Else
                sLastValue = ""
            End If
        Else
            If Not IsError(Target) Then sLastValue = Target.Value Else sLastValue = "Err"
        End If
        .Cells(lLastRow, 6).NumberFormat = "@"
        .Cells(lLastRow, 6) = sLastValue
    End With
    wbLOG.Close 1
    Application.ScreenUpdating = True: Application.EnableEvents = True
End Sub

Private Sub Workbook_SheetSelectionChange(ByVal Sh As Object, ByVal Target As Range)
    If Sh.Name = "LOG" Then Exit Sub
    If Target.Count > 1 Then
        Dim rCell As Range, rRng As Range
        On Error Resume Next
        Set rRng = Intersect(Target, Sh.UsedRange): On Error GoTo 0
        If rRng Is Nothing Then Exit Sub
        For Each rCell In rRng
            If Not IsError(rCell) Then sValue = sValue & "," & rCell Else sValue = sValue & "," & "Err"
        Next rCell
        sValue = Mid(sValue, 2)
    Else
        If Not IsError(Target) Then sValue = Target.Value Else sValue = "Err"
    End If
End Sub