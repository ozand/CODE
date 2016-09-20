Attribute VB_Name = "Conso"

Sub conso()

Dim ar_brand()
Dim ar_Data()
Dim act_month As Integer
Dim nm_00month$

myLib.VBA_Start

thisYear = 2016
act_month = CInt(InputBox("Month"))
nm_00month = month_form_00(act_month)
ar_Brands = Array("LP", "MX", "KR", "RD", "ES")
lnk_datafolder = "p:\DPP\Business development\Statistics Service\EDU\Base\"
NF = ActiveWorkbook.Name

iii = 1
For f_a = 0 To 4
    WSh = "KPI_" & ar_Brands(f_a) & "_" & nm_00month & "_" & thisYear
    in_data = "KPI_"
    nm_file = WSh & ".csv"
    myLib.openFileCSV(lnk_datafolder & "/" & nm_file)

    num_LastRow = ActiveSheet.UsedRange.Row + ActiveSheet.UsedRange.Rows.Count - 1
    num_LastColum = ActiveSheet.UsedRange.Column + ActiveSheet.UsedRange.Columns.Count - 1
    
    Select Case f_a
        Case 0
        str_rw = 1
        ReDim ar_Data(1 To 999999, 1 To num_LastColum + 1)
        Case Else
        str_rw = 2
    End Select
     
    For f_r = str_rw To num_LastRow
        clm_x = 1
        For f_c = 0 To num_LastColum
                If iii = 1 Then ar_Data(1, 1) = "brand"
                If f_c = 0 Then
                    val_ar_data = ar_Brands(f_a)
                    Else
                    val_ar_data = Cells(f_r, f_c)
                End If
                ar_Data(iii, clm_x) = val_ar_data: clm_x = clm_x + 1
        Next f_c: iii = iii + 1
    Next f_r
Workbooks(WSh).Close
Next f_a
Workbooks(NF).Activate
myLib.CreateSh(in_data)
Sheets(in_data).Select
ActiveSheet.Cells(1, 1).Resize(iii + 1, num_LastColum + 1) = ar_Data()
myLib.VBA_End
End Sub