Option Compare Text
'Function List
    '---------------------------------------------------------------------------------------------------------
    Function fn_Replace_symbols(ByVal txt As String) As String
        St$ = "~!@/\#$%^:?&*=|`;"""
        For f_i% = 1 To Len(St$)
            txt = Replace(txt, Mid(St$, f_i, 1), "_")
            txt = Replace(txt, Chr(10), "_")
        Next
        fn_Replace_symbols = txt
    End Function
    '---------------------------------------------------------------------------------------------------------
    Function fn_VBA_Start() As String
    With Application
        .ScreenUpdating = False
        .EnableEvents = False
        .Calculation = xlCalculationManual
        '.DisplayPageBreaks = False
        .DisplayAlerts = False
    End With
    End Function
    '---------------------------------------------------------------------------------------------------------
    Function fn_VBA_End() As String
    With Application
        .ScreenUpdating = True
        .Calculation = xlCalculationAutomatic
        .EnableEvents = True
        .DisplayStatusBar = True
        .DisplayAlerts = True
    End With
    End Function
    '---------------------------------------------------------------------------------------------------------
    Function fn_CreateSh(cr_sh As String) As String
    For Each Sh In ThisWorkbook.Worksheets
        If Sh.Name = cr_sh Then
        chek_name = 1
        End If
    Next Sh
        If chek_name <> 1 Then
        Set Sh = Worksheets.Add()
        Sh.Name = cr_sh
        End If
    End Function
    '---------------------------------------------------------------------------------------------------------
    Function fn_openFile(ByRef patch As String, nm_sh As String) As String
    Dim result$
    If Dir(patch) = "" Then
        MsgBox ("FileNotFound")
    Else
        Workbooks.Open Filename:=patch, Notify:=False
        
        result = ActiveWorkbook.Name
        Sheets(nm_sh).Select
        ActiveSheet.AutoFilterMode = False
    End If
    fn_openFile = result
    End Function
    '---------------------------------------------------------------------------------------------------------
    Function fn_openFileCSV(ByRef patch As String)
    Dim result$
    If Dir(patch) = "" Then
        MsgBox ("FileNotFound")
    Else
        Workbooks.OpenText Filename:=patch, _
            Origin:=65001, StartRow:=1, DataType:=xlDelimited, TextQualifier:= _
            xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=False, Semicolon:=True, _
            Comma:=False, Space:=False, Other:=False, TrailingMinusNumbers:=True

        
    End If
    End Function
    '---------------------------------------------------------------------------------------------------------
    Function fn_quartal(month&) As String
    Dim result As String
    result = Empty
            Select Case month
            Case 1, 2, 3
            result = "1Q"
            Case 4, 5, 6
            result = "2Q"
            Case 7, 8, 9
            result = "3Q"
            Case 10, 11, 12
            result = "4Q"
        End Select
    fn_quartal = result
    End Function
    '---------------------------------------------------------------------------------------------------------
    Function fn_month_form_00(month As Integer) As String
    Dim result As String
    result = Empty

        If month < 10 Then
            result = "0" & month
        Else
            result = month
        End If

    fn_month_form_00 = result
    End Function
    '---------------------------------------------------------------------------------------------------------
    Function fn_patch_history_TR(brand As String, year As Integer, thisMonth As Integer, ver_month As String) As String
    Dim result As String
    result = Empty

        Select Case month
            Case this_month
            result = "p:\DPP\Business development\Book commercial\" & brand & "\Top Russia Total " & year & " " & brand & ".xlsm"
            Case Else
            result = "p:\DPP\Business development\Book commercial\" & brand & "\" & year & "\History " & year & "\Top Russia Total " & year & "." & ver_month & " " & brand & ".xlsm"
        End Select

    fn_patch_history_TR = result
    End Function
    '---------------------------------------------------------------------------------------------------------
    Function fn_num_LastRow() As Integer
    Dim result As Integer
    result = Empty
        With ActiveWorkbook.ActiveSheet.UsedRange
        result = .Row + .Rows.Count - 1
        End With
    fn_num_LastRow = result
    End Function
    '---------------------------------------------------------------------------------------------------------
    Function fn_lastColumn() As Integer
    Dim result As Integer
    result = Empty
        
        With ActiveWorkbook.ActiveSheet.UsedRange
        result = .Column + .Columns.Count - 1
        End With

    fn_lastColumn = result
    End Function
    '---------------------------------------------------------------------------------------------------------
    Function fn_clnt_type(in_data$, i&)
    Dim result
    Dim ar_type_clients(1 To 4, 1 To 12)
    Dim f_sl&

    '--array
        ar_type_clients(1, 1) = "салон"
        ar_type_clients(2, 1) = "salon"
        ar_type_clients(3, 1) = "salon"
        ar_type_clients(4, 1) = "single"

        ar_type_clients(1, 2) = "сеть салонов"
        ar_type_clients(2, 2) = "chain_salons"
        ar_type_clients(3, 2) = "salon"
        ar_type_clients(4, 2) = "chain"

        ar_type_clients(1, 3) = "ч/м"
        ar_type_clients(2, 3) = "hdres"
        ar_type_clients(3, 3) = "salon"
        ar_type_clients(4, 3) = "single"

        ar_type_clients(1, 4) = "сеть магазинов"
        ar_type_clients(2, 4) = "chain_shops"
        ar_type_clients(3, 4) = "shop"
        ar_type_clients(4, 4) = "chain"

        ar_type_clients(1, 5) = "магазин"
        ar_type_clients(2, 5) = "shop"
        ar_type_clients(3, 5) = "shop"
        ar_type_clients(4, 5) = "single"

        ar_type_clients(1, 6) = "салон-маг."
        ar_type_clients(2, 6) = "salon"
        ar_type_clients(3, 6) = "salon"
        ar_type_clients(4, 6) = "single"

        ar_type_clients(1, 7) = "(пусто)"
        ar_type_clients(2, 7) = "other"
        ar_type_clients(3, 7) = "other"
        ar_type_clients(4, 7) = "single"

        ar_type_clients(1, 8) = "школа"
        ar_type_clients(2, 8) = "school"
        ar_type_clients(3, 8) = "school"
        ar_type_clients(4, 8) = "single"

        ar_type_clients(1, 9) = "другое"
        ar_type_clients(2, 9) = "other"
        ar_type_clients(3, 9) = "other"
        ar_type_clients(4, 9) = "single"

        ar_type_clients(1, 10) = "нейл-бар"
        ar_type_clients(2, 10) = "nails_bar"
        ar_type_clients(3, 10) = "nails"
        ar_type_clients(4, 10) = "single"

        ar_type_clients(1, 11) = "сеть нейл-баров"
        ar_type_clients(2, 11) = "chain_nails"
        ar_type_clients(3, 11) = "nails"
        ar_type_clients(4, 11) = "chain"

        ar_type_clients(1, 12) = "e-commerce"
        ar_type_clients(2, 12) = "e-commerce"
        ar_type_clients(3, 12) = "e-commerce"
        ar_type_clients(4, 12) = "single"

    For f_sl = 1 To 12
        
    If StrComp(ar_type_clients(1, f_sl), LCase(in_data), vbTextCompare) Then
        
        result = ar_type_clients(i, f_sl)
        Exit For
        Else
        result = Empty
    End If
    Next f_sl

    fn_clnt_type = result
    End Function
    '---------------------------------------------------------------------------------------------------------
    Function fn_mreg_ext$(in_data_mreg$, in_data_reg$)
    Dim result$
    Dim extPos&
    textPos = 0
    If LCase(in_data_mreg) = LCase("Moscou GR") Then
        textPos = InStr(in_data_reg, "MSK")
        textPos = InStr(in_data_reg, "Moscou") + textPos
            If textPos > 0 Then
            result = "Moscou"
            Else
            result = "GR"
            End If
    End If
    fn_mreg_ext = result
    End Function
    '---------------------------------------------------------------------------------------------------------
    Function fn_mreg_lat(in_data_mreg As String) As String
    Dim result$
    Dim f_mr&
    Dim ar_nmMregEN(), ar_nmMregLT()

    ar_nmMregEN = Array("MOSCOW", "GR", "NORTHWEST", "CENTER", "VOLGA", "SOUTH", "URAL", "SIBERIA", "FAR EAST")
    ar_nmMregLT = Array("Moscou", "GR", "Nord-Ouest", "Centre", "Volga-Centre", "Sud", "Oural", "Siberie", "EO")

    For f_mr = 0 To UBound(ar_nmMregLT)
    If ar_nmMregLT(f_mr) = in_data_mreg Then
    result = ar_nmMregEN(f_mr)
    Exit For
    End If
    Next f_mr

    fn_mreg_lat = result

    End Function
    '---------------------------------------------------------------------------------------------------------
    Function fn_salon_name$(in_sln_nm$, in_sln_addres$, in_city$)
    Dim result$

    result = Trim(fn_Replace_symbols(Left(in_sln_nm, 30) & ". " & Left(in_sln_addres, 50) & " " & Left(in_city, 50)))

    fn_salon_name = result
    End Function
    '---------------------------------------------------------------------------------------------------------
    Function fn_mont_num&(in_data$)
    Dim result&
    Dim f_m&, num_month&

    ar_nm_month_qnc_rus = Array("январь", "февраль", "март", "апрель", "май", "июнь", "июль", "август", "сентябрь", "октябрь", "ноябрь", "декабрь")
    result = 1
        For f_m = 0 To 11
        If ar_nm_month_qnc_rus(f_m) = in_data Then
        result = f_m + 1
        Exit For
        End If
        Next f_m
    
    fn_mont_num = result
    End Function

    '---------------------------------------------------------------------------------------------------------
    Function fn_month_eng$(month$)
    Dim result$
    Dim f_m&

    ar_month_rus = Array("январь", "февраль", "март", "апрель", "май", "июнь", "июль", "август", "сентябрь", "октябрь", "ноябрь", "декабрь")
    ar_month_eng = Array("Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec")
    
        For f_m = 0 To 11
            If ar_month_rus(f_m) = month Then
            result = ar_month_eng(f_m)
            Exit For
            End If
        Next f_m
        
    fn_month_eng = result
    End Function

    '---------------------------------------------------------------------------------------------------------
    Function fn_getYearType(in_act_year&, in_data&, i&) As Variant
    Dim result1&, result2$
        
        If Len(in_data) = 4 Then result1 = in_data Else result1 = 2008
        
            Select Case result1
                Case in_act_year
                    result2 = "TY"
                Case in_act_year - 1
                    result2 = "PY"
                Case Else
                    result2 = "PPY"
            End Select
    
    Select Case i
    Case 1
        fn_getYearType = result1
    Case 2
        fn_getYearType = result2
    Case Else
        fn_getYearType = Empty
    End Select
    End Function
    '---------------------------------------------------------------------------------------------------------

    Function fn_mag(in_min_price As Long, in_max_price As Long, in_place As Long, mag_type As String) As Variant

    Dim result As Variant
    Dim mag_avg_price&
            
    If IsNumeric(in_min_price) and IsNumeric(in_max_price) Then
        mag_avg_price = Application.WorksheetFunction.Average(in_min_price, in_max_price)
    Else
        mag_avg_price = in_min_price + in_max_price
    End If

    Select Case mag_type
        Case "avg_price"
            result = mag_avg_price

        Case "hair"
            Select Case mag_avg_price
                Case 100 To 799
                    result = "D"
                Case 800 To 1199
                    result = "C"
                Case 1200 To 2000
                    result = "B"
                Case Is > 2000
                    result = "A"
                Case Else
                    result = Empty
            End Select
        
        Case "nail"
            Select Case mag_avg_price
                Case 10 To 319
                    result = "D"
                Case 320 To 479
                    result = "C"
                Case 480 To 799
                    result = "B"
                Case Is > 800
                    result = "A"
                Case Else
                    result = Empty
            End Select
        
        Case "skin"
            Select Case mag_avg_price
                Case 100 To 799
                    result = "D"
                Case 800 To 1199
                    result = "C"
                Case 1200 To 2000
                    result = "B"
                Case Is > 2000
                    result = "A"
                Case Else
                    result = Empty
            End Select

        Case "place"
            If IsNumeric(in_place) Then
            in_place = Round(in_place, 0)
            End If
            Select Case in_place
                Case 1 To 2
                result = "1"
                Case 3 To 4
                result = "2"
                Case Is > 4
                result = "3"
                Case Else
                result = Empty
            End Select

        End Select
        
    fn_mag = result
    End Function
    '---------------------------------------------------------------------------------------------------------
    Function fn_type_business$(in_brand$)
    Dim result$
    Select Case in_brand
            Case "LP", "MX", "KR", "RD"
            result = "Hair"
            Case "ES"
            result = "Nails"
            Case "DE", "CR"
            result = "Skin"
    End Select
    fn_type_business = result
    End Function
    '---------------------------------------------------------------------------------------------------------
    Function fn_type_active_DN$(in_data&)
    Dim result$

    Select Case in_data
        Case 1
            result = "Active"
        Case 0
            result = "Closed"
    End Select
    fn_type_active_DN = result
    End Function
    '---------------------------------------------------------------------------------------------------------
    Function fn_rnd_num&(in_data)
    Dim result&
    If IsNumeric(in_data) And Len(in_data) > 0 Then
        result = Round(in_data, 0)
    Else
        result = 0
    End If
    fn_rnd_num = result
    End Function
    '---------------------------------------------------------------------------------------------------------
    Function fn_num2num0&(in_data As Variant)
    Dim result&
    If Len(in_data) > 0 And IsNumeric(in_data) Then
    result = in_data
    Else
    result = 0
    End If
    fn_num2num0 = result
    End Function
    '---------------------------------------------------------------------------------------------------------
    Function fn_num2numNull(in_data) As Variant
    Dim result As Variant
    If Len(in_data) > 0 And in_data <> 0 Then
    result = in_data
    Else
    result = Empty
    End If
    fn_num2numNull = result
    End Function
    '---------------------------------------------------------------------------------------------------------
    Function fn_getNmChainTop$(inNmChain$, inCdChain&, inNmTypeClnt$)
    Dim result$
    If Left(inCdChain, 2) = "92" And fn_clnt_type(inNmTypeClnt, 4) = "chain" Then
    result = inNmChain
    Else
    result = Empty
    End If
    fn_getNmChainTop = result
    End Function

    '---------------------------------------------------------------------------------------------------------
    Function fn_GetLTM(in_row&, inThisMonth&, typeFN$) As Variant
    Dim result$
    Dim f_a&, f_avg&, sum_CA_LTM&, AVG_CA_LTM&, frqOrder&
    Dim MinVal!, MaxVal!
    Dim val As Variant
    
    ar_DataMonthPRTN = Array(66, 67, 68, 69, 70, 71, 72, 73, 74, 75, 76, 77, 79, 80, 81, 82, 83, 84, 85, 86, 87, 88, 89, 90)
    ar_nmAVG_Order = Array(2.5, 5, 10, 15, 20, 25, 30, 50, 60, 70, 100000)

    sum_CA_LTM = 0
    frqOrder = 0
    
    For f_a = inThisMonth To inThisMonth + 11
        val = Cells(in_row, ar_DataMonthPRTN(f_a))
        If IsNumeric(val) And val > 0 Then
        frqOrder = frqOrder + 1
        sum_CA_LTM = sum_CA_LTM + val
        End If
    Next f_a
    AVG_CA_LTM = Round(sum_CA_LTM / 12 / 1000, 1)

    Select Case typeFN
    Case "avg_ca"
        If sum_CA_LTM <> 0 Then
        result = AVG_CA_LTM
        Else
        result = Empty
        End If

    Case "frqOrders"
        result = frqOrder & "\12"
        
    Case "type_avg_ca"
        MinVal = 0
        MaxVal = 0
        
            
            Select Case AVG_CA_LTM
            Case 0
                result = "0"
            Case Is >= 70
                result = ">70"
            Case Is < 70
                For f_avg = 0 To UBound(ar_nmAVG_Order)
                    MaxVal = ar_nmAVG_Order(f_avg)
                    If AVG_CA_LTM <= MaxVal And AVG_CA_LTM > MinVal Then result = "'" & MinVal & "-" & MaxVal: Exit For
                    
                    MinVal = MaxVal
                Next f_avg
            Case Else
            result = Empty
            End Select
        
    End Select
    fn_GetLTM = result
    End Function

    '---------------------------------------------------------------------------------------------------------
    Function fn_getVectoreEV$(in_data#)
    Dim result$

    If IsNumeric(in_data) Then
        Select Case in_data
        Case Is > 0
            result = "+"
        Case Is < 0
            result = "-"
        Case Else
            result = Empty
        End Select
    Else
    result = Null
    End If

    fn_getVectoreEV = result
    End Function
    '---------------------------------------------------------------------------------------------------------

    Function fn_getMonthlyCA&(in_row&, in_month&, in_thisMonth&, in_typeY$, in_typeVal$, in_type_period$)
    Dim result&, val&
    Dim typeF$
    Dim clm_PY_LOR_VAL%, clm_TY_LOR_VAL%, clm_PY_PRTN_VAL%, clm_TY_PRTN_VAL%
    Dim ar_Matrix(1 To 2, 1 To 2)

    val = Empty
    typeF = in_typeY & "_" & in_typeVal
    Select Case typeF
        Case "PY_LOR"
            clm = 106
        Case "TY_LOR"
            clm = 93
        Case "PY_PRTN"
            clm = 79
        Case "TY_PRTN"
            clm = 66
        Case Else
            Exit Function
    End Select

    Select Case in_type_period
        Case "Total"
            in_thisMonth = 12
        Case "YTD"
            in_thisMonth = in_thisMonth
    End Select

    Select Case in_month
        Case Is <= in_thisMonth
            val = fn_num2num0(Cells(in_row, clm + in_month - 1))
            If val = 0 Then val = Empty Else val = val / 1000
        Case Else
            val = Empty
    End Select

    result = val
    fn_getMonthlyCA = result
    End Function
    '---------------------------------------------------------------------------------------------------------
    Sub IsOpenTRtoClsd()
        Dim wbBook As Workbook
        For Each wbBook In Workbooks
            If wbBook.Name <> ThisWorkbook.Name Then
                If Windows(wbBook.Name).Visible Then
                    If wbBook.Name Like "Top Russia*" Then wbBook.Close: Exit For
                End If
            End If
        Next wbBook
    End Sub
    '---------------------------------------------------------------------------------------------------------
    Function fn_getCA_Cnq(in_monthQnc&)

            Case cd_ThisYear - 1
            fst_order_LOR_PY = Cells(f_i, clm_PYper_LOR_VAL + cd_month_qnc - 1) / 1000
            fst_order_PRTN_PY = Cells(f_i, clm_PYper_PRTN_VAL + cd_month_qnc - 1) / 1000
            
                If cd_month_qnc = cd_ThisMonth Then
                fst_order_LOR_M_PY = Cells(f_i, clm_PYper_LOR_VAL + cd_month_qnc - 1) / 1000
                End If
                                
            Case cd_ThisYear
            fst_order_LOR_TY = Cells(f_i, clm_TYper_LOR_VAL + cd_month_qnc - 1) / 1000
            fst_order_PRTN_TY = Cells(f_i, clm_TYper_PRTN_VAL + cd_month_qnc - 1) / 1000

                If cd_month_qnc = cd_ThisMonth Then
                fst_order_LOR_M_TY = Cells(f_i, clm_TYper_LOR_VAL + cd_month_qnc - 1) / 1000
                End If

            End Select

    End Function

    '---------------------------------------------------------------------------------------------------------
    Function fn_avgCA&(in_data&, in_month&)
    Dim result&
 
    If Not IsEmpty(in_data) And IsNumeric(in_data) Then
    result = in_data / in_month
    Else
    result = Empty
    
    End If
    fn_avgCA = result
    End Function


Sub GetDataTR()



Dim pathc2file$
Dim cd_ThisMonth&, cd_month_qnc&
Dim nm_PatchTR$, nm_brand$, num_LastRow$, nm_Mreg$, nm_Sector$, nm_Mreg_ext$, nm_month_qnc$, nm_business$, nm_Salon$, nm_Salon_addr$, nm_Salon_city$, nm_TypeClntRus$, nm_chain$, nm_Y$, nm_period$
Dim cd_ThisYear&, mag_min_price&, num_month&, mag_max_price&, mag_hd_place&, cd_year_qnc&, cd_sts_dn_cln&, num_StatusHead&, cd_chain&
Dim num_ev_ca#
Dim nm_ActTR$, in_data$
Dim f_b&, iii&, i&, x&, y&, frqOrder&, f_i&, f_y&, f_m&, val_ca_PY_YTD&, val_CA_MYTD_PY&, val_ca&, val_ca_cumul&, val_ca_quarter&, val_ca_TY_YTD&, val_CA_MYTD_TY&
Dim ar_Data(), ar_brand(), ar_nmHead()

nm_WB = ActiveWorkbook.Name
in_data = "in_TR"
fn_CreateSh (in_data)



cd_ThisMonth = CInt(InputBox("Month"))
cd_ThisYear = 2016

'colums CA PRTNN VAL for LTM
'---------------------------------------------------------------------------------------------------------


    'colums CA LOREAL VAL
    ar_month_eng = Array("Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec")
    ar_brand = Array("LP", "KR", "RD", "MX", "ES", "DE", "CR")
        num_ar_brand = UBound(ar_brand)

fn_VBA_Start
IsOpenTRtoClsd

iii = 0

For f_b = 0 To num_ar_brand

    nm_brand = ar_brand(f_b)
    nm_PatchTR = "p:\DPP\Business development\Book commercial\" & nm_brand & "\Top Russia Total " & cd_ThisYear & " " & nm_brand & ".xlsm"
    
    nm_ActTR = fn_openFile(nm_PatchTR, nm_brand)
    num_LastRow = fn_num_LastRow
    
    For f_i = 4 To num_LastRow

    Select Case iii
        Case 0
        end_clm = 200
        ReDim ar_Data(1 To 100000, 1 To end_clm)
        ReDim ar_nmHead(1 To end_clm)
        Case 1
        ReDim Preserve ar_Data(1 To 100000, 1 To num_last_colum)
        ReDim Preserve ar_nmHead(1 To num_last_colum)
    End Select

    nm_Mreg = Cells(f_i, 4)
    If Not IsEmpty(nm_Mreg) Then iii = iii + 1

    nm_Sector = Cells(f_i, 6)
    nm_Mreg_ext = fn_mreg_ext(nm_Mreg, nm_Sector)
    nm_Mreg_LT = fn_mreg_lat(nm_Mreg_ext)
    nm_REG = Cells(f_i, 5)
    nm_FLSM = Cells(f_i, 165)
    nm_SREP = Cells(f_i, 7)
    nm_Salon = Cells(f_i, 9)
    nm_Salon_addr = Cells(f_i, 12)
    nm_Salon_city = Cells(f_i, 11)
    nm_month_qnc = Cells(f_i, 64)
    cd_month_qnc = fn_mont_num(nm_month_qnc)
    cd_year_qnc = fn_getYearType(cd_ThisYear, fn_num2num0(Cells(f_i, 65)), 1)
    nm_TypeClntRus = Cells(f_i, 18)
    nm_club_type = Cells(f_i, 40)
    nm_chain = Cells(f_i, 19)
    cd_chain = fn_num2numNull(Cells(f_i, 20))
    cd_Univers = Cells(f_i, 2)
    mag_min_price = fn_rnd_num(Cells(f_i, 23))
    mag_max_price = fn_rnd_num(Cells(f_i, 25))
        mag_price = fn_mag(mag_min_price, mag_max_price, mag_hd_place, nm_business)
    mag_hd_place = fn_rnd_num(Cells(f_i, 27))
        mag_type_place = fn_mag(mag_min_price, mag_max_price, mag_hd_place, "place")
        vl_mag = mag_price & mag_type_place
        If Len(vl_mag) <> 2 Then vl_mag = Null
    cnt_AVG_HD = fn_rnd_num(Cells(f_i, 28))
    nm_business = fn_type_business(nm_brand)
    vr_TypeEmotion = Cells(f_i, 41)
    cd_sts_dn_cln = Cells(f_i, 8)
    nm_Partners = Cells(f_i, 167)
    cd_Partner = Cells(f_i, 173)
    num_ev_ca = fn_num2num0(Cells(f_i, 92))
    cd_idECAD = Cells(f_i, 29)
    EDU_ALLTIME_MSTR = Cells(f_i, 30)
    EDU_PY_MSTR = Cells(f_i, 31)
    EDU_TY_MSTR = Cells(f_i, 32)
    EDU_ALLTIME_CNTCT = Cells(f_i, 33)
    EDU_PY_CNTCT = Cells(f_i, 34)
    EDU_TY_CNTCT = Cells(f_i, 35)
    val_comKPI = Cells(f_i, 209)
    type_brand = fn_type_business(nm_brand)
    cd_brand_row = nm_brand & Cells(f_i, 1)

    n = 1:      ar_Data(iii, n) = nm_brand:         ar_nmHead(n) = "brand"
    n = n + 1:  ar_Data(iii, n) = type_brand:       ar_nmHead(n) = "bussines"
    n = n + 1:  ar_Data(iii, n) = Cells(f_i, 1):    ar_nmHead(n) = "rowTR"
    n = n + 1:  ar_Data(iii, n) = cd_brand_row:     ar_nmHead(n) = "BRAND_rowTR"
    
    
        If Len(cd_Univers) <> 9 Then
        cd_Univers = cd_brand_row
        Else: cd_Univers = cd_Univers
        End If
n = n + 1:      ar_Data(iii, n) = cd_Univers:        ar_nmHead(n) = "unvCD"
    
    n = n + 1
    ar_Data(iii, n) = nm_brand & Cells(f_i, 2)
    ar_nmHead(n) = "BRAND_unvCD"
    
    n = n + 1
    ar_Data(iii, n) = nm_Mreg
    ar_nmHead(n) = "mreg"
            
    num_colus = n + 1
    ar_Data(iii, n) = nm_Mreg_LT
    ar_nmHead(n) = "mreg_EXT"
    
    n = n + 1
    ar_Data(iii, n) = nm_REG
    ar_nmHead(n) = "REG"
    
    n = n + 1
    ar_Data(iii, n) = nm_FLSM
    ar_nmHead(n) = "FLSM"
    
    n = n + 1
    ar_Data(iii, n) = nm_Sector
    ar_nmHead(n) = "SEC"
    
    n = n + 1
    ar_Data(iii, n) = nm_SREP
    ar_nmHead(n) = "SREP"
        
    n = n + 1
    ar_Data(iii, n) = fn_salon_name(nm_Salon, nm_Salon_addr, nm_Salon_city)
    ar_nmHead(n) = "salon"
     
    n = n + 1
    ar_Data(iii, n) = nm_chain
    ar_nmHead(n) = "Chain_name"
    
    n = n + 1
    ar_Data(iii, n) = nm_Salon_city
    ar_nmHead(n) = "city"
    
    n = n + 1
    ar_Data(iii, n) = fn_clnt_type(nm_TypeClntRus, 1)
    ar_nmHead(n) = "type_SLN"
    
    n = n + 1
    ar_Data(iii, n) = fn_clnt_type(nm_TypeClntRus, 2)
    ar_nmHead(n) = "salon_type_eng"
    
    n = n + 1
    ar_Data(iii, n) = fn_clnt_type(nm_TypeClntRus, 3)
    ar_nmHead(n) = "salon_type_short_eng"
    
    n = n + 1
    ar_Data(iii, n) = fn_clnt_type(nm_TypeClntRus, 4)
    ar_nmHead(n) = "salon_type_chain_eng"
    
    n = n + 1
    ar_Data(iii, n) = cd_chain
    ar_nmHead(n) = "cd_chain"
    
    n = n + 1
    ar_Data(iii, n) = fn_getNmChainTop(nm_chain, cd_chain, nm_TypeClntRus)
    ar_nmHead(n) = "nm_Top10_chain"
     
    n = n + 1
    ar_Data(iii, n) = nm_club_type
    ar_nmHead(n) = "type_confirmed_CLUB"
       
    n = n + 1
    ar_Data(iii, n) = vr_TypeEmotion
    ar_nmHead(n) = "type_emotion"
         
    n = n + 1
    ar_Data(iii, n) = 1 & "." & cd_month_qnc & "." & cd_year_qnc
    ar_nmHead(n) = "date_CNQ_Y"
    
    n = n + 1
    ar_Data(iii, n) = cd_month_qnc
    ar_nmHead(n) = "date_month_num"
    
    n = n + 1
    ar_Data(iii, n) = fn_month_eng(nm_month_qnc)
    ar_nmHead(n) = "date_month_name"
    
    n = n + 1
    ar_Data(iii, n) = fn_getYearType(cd_ThisYear, cd_year_qnc, 1)
    ar_nmHead(n) = "date_year"
  
    n = n + 1
        nm_TypeGA_Y = fn_getYearType(cd_ThisYear, cd_year_qnc, 2)
    ar_Data(iii, n) = nm_TypeGA_Y
    ar_nmHead(n) = "nm_TypeGA_YEAR"
    
    n = n + 1
    ar_Data(iii, n) = vl_mag
    ar_nmHead(n) = "type_MAG"
    
    n = n + 1
    ar_Data(iii, n) = mag_price
    ar_nmHead(n) = "type_MAG_PRICE"
    
    n = n + 1
      ar_Data(iii, n) = mag_type_place
    ar_nmHead(n) = "type_MAG_type_place"
    
    n = n + 1
    ar_Data(iii, n) = cd_sts_dn_cln
    ar_nmHead(n) = "status_DN_num"
    
    n = n + 1
    ar_Data(iii, n) = fn_type_active_DN(cd_sts_dn_cln)
    ar_nmHead(n) = "status_DN_name"
        
    n = n + 1
    ar_Data(iii, n) = fn_GetLTM(f_i, cd_ThisMonth, "avg_ca")
    ar_nmHead(n) = "CA_AVG_LTM"
   
    n = n + 1
    ar_Data(iii, n) = fn_GetLTM(f_i, cd_ThisMonth, "type_avg_ca")
    ar_nmHead(n) = "CA_AVG_LTM_name"
    
    n = n + 1
    ar_Data(iii, n) = fn_GetLTM(f_i, cd_ThisMonth, "frqOrders")
    ar_nmHead(n) = "frq_order_LTM"

    n = n + 1
    ar_Data(iii, n) = fn_num2numNull(num_ev_ca)
    ar_nmHead(n) = "CA_ev"
   
    n = n + 1
    ar_Data(iii, n) = fn_getVectoreEV(num_ev_ca)
    ar_nmHead(n) = "CA_ev_name"
    
    n = n + 1
    ar_Data(iii, n) = cd_idECAD
    ar_nmHead(n) = "EDU_id_ECAD"
    
    n = n + 1
    ar_Data(iii, n) = fn_num2numNull(EDU_ALLTIME_MSTR)
    ar_nmHead(n) = "EDU_ALLTIME_MSTR"
    
    n = n + 1
    ar_Data(iii, n) = fn_num2numNull(EDU_PY_MSTR)
    ar_nmHead(n) = "EDU_PY_MSTR"
    
    n = n + 1
    ar_Data(iii, n) = fn_num2numNull(EDU_TY_MSTR)
    ar_nmHead(n) = "EDU_TY_MSTR"
    
    n = n + 1
    ar_Data(iii, n) = fn_num2numNull(EDU_ALLTIME_CNTCT)
    ar_nmHead(n) = "EDU_ALLTIME_CNTCT"
    
    n = n + 1
    ar_Data(iii, n) = fn_num2numNull(EDU_PY_CNTCT)
    ar_nmHead(n) = "EDU_PY_CNTCT"
    
    n = n + 1
    ar_Data(iii, n) = fn_num2numNull(EDU_TY_CNTCT)
    ar_nmHead(n) = "EDU_TY_CNTCT"
    
    n = n + 1
    ar_Data(iii, n) = fn_num2numNull(mag_hd_place)
    ar_nmHead(n) = "type_place_HD"
    
    n = n + 1
    ar_Data(iii, n) = fn_num2numNull(cnt_AVG_HD)
    ar_nmHead(n) = "type_AVG_HD"
        
    n = n + 1
    ar_Data(iii, n) = val_comKPI
    ar_nmHead(n) = "com_KPI"
    
    n = n + 1
    ar_Data(iii, n) = nm_Partners
    ar_nmHead(n) = "nm_PRTNner"
          
    n = n + 1
    ar_Data(iii, n) = cd_Partner
    ar_nmHead(n) = "cd_PRTNner"


    '---------------------------------------------------------------------------------------------------------
    'CA_MONTHLY&CUMUL&QUARTER_LOR_VAL
    '---------------------------------------------------------------------------------------------------------
    xyz = 0
    qq = 0
    n = n
    clm_ca_q = Empty
    For f_y = cd_ThisYear - 1 To cd_ThisYear
        nm_Y = fn_getYearType(cd_ThisYear, f_y, 2)
        If f_y = cd_ThisYear Then nm_period = "YTD": num_shift_clm = 12 Else nm_period = "Total": num_shift_clm = 0
      
            val_ca_cumul = Empty
            nm_ca_quarter = Empty
            val_ca_quarter = Empty

        For f_m = 1 To 12
            
            val_ca = Empty
            val_ca = fn_getMonthlyCA(f_i, f_m, cd_ThisMonth, nm_Y, "LOR", nm_period)
            val_ca_cumul = val_ca_cumul + val_ca
            
            clm_ca_m = n + f_m + num_shift_clm
                ar_Data(iii, clm_ca_m) = fn_num2numNull(val_ca)
                ar_nmHead(clm_ca_m) = "CA_" & nm_Y & "_" & "M" & f_m
            clm_ca_ytd = n + f_m + 24 + num_shift_clm
                ar_Data(iii, clm_ca_ytd) = fn_num2numNull(val_ca_cumul)
                ar_nmHead(clm_ca_ytd) = "CA_" & nm_Y & "_" & "YTD" & f_m

            Select Case f_y
                Case cd_ThisYear - 1
                    If f_m = cd_month_qnc Then val_ca_PY_YTD = val_ca_cumul: val_CA_MYTD_PY = val_ca
                Case cd_ThisYear
                    If f_m = cd_month_qnc Then val_ca_TY_YTD = val_ca_cumul: val_CA_MYTD_TY = val_ca
            End Select

            '---------------------------------------------------------------------------------------------------------
            'QUARTER
            If fn_quartal(f_m) = nm_ca_quarter Then
                val_ca_quarter = val_ca_quarter + val_ca
            Else
                nm_ca_quarter = fn_quartal(f_m)
                val_ca_quarter = val_ca
                qq = qq + 1
                clm_ca_q = n + 48 + qq
            End If

                ar_Data(iii, clm_ca_q) = fn_num2numNull(val_ca_quarter)
                ar_nmHead(clm_ca_q) = "CA_" & nm_Y & "_" & nm_ca_quarter

        Next f_m
        xyz = xyz + 1
    Next f_y

 '---------------------------------------------------------------------------------------------------------
 ' first conq order
 '---------------------------------------------------------------------------------------------------------

    n = n + 24 * xyz + 8
    
    fst_order_LOR_TY = Empty
    fst_order_PRTN_TY = Empty
    fst_order_LOR_PY = Empty
    fst_order_PRTN_PY = Empty
    fst_order_LOR_M_TY = Empty
    fst_order_LOR_M_PY = Empty
        
    Select Case nm_TypeGA_Y
        Case "PY"
            fst_order_LOR_PY = fn_getMonthlyCA(f_i, cd_ThisMonth, cd_ThisMonth, "PY", "LOR", "YTD")
            fst_order_PRTN_PY = fn_getMonthlyCA(f_i, cd_ThisMonth, cd_ThisMonth, "PY", "PRTN", "YTD")
            If cd_month_qnc = cd_ThisMonth Then fst_order_LOR_M_PY = fn_getMonthlyCA(f_i, cd_ThisMonth, cd_ThisMonth, "PY", "LOR", "YTD")
        Case "TY"
            fst_order_LOR_TY = fn_getMonthlyCA(f_i, cd_ThisMonth, cd_ThisMonth, "TY", "LOR", "YTD")
            fst_order_PRTN_TY = fn_getMonthlyCA(f_i, cd_ThisMonth, cd_ThisMonth, "TY", "PRTN", "YTD")
            If cd_month_qnc = cd_ThisMonth Then fst_order_LOR_M_TY = fn_getMonthlyCA(f_i, cd_ThisMonth, cd_ThisMonth, "TY", "LOR", "YTD")
    End Select


    n = n + 1
    ar_Data(iii, n) = fst_order_LOR_PY
    ar_nmHead(n) = "PY_CNQ_Order"
    
    n = n + 1
    ar_Data(iii, n) = fst_order_LOR_M_PY
    ar_nmHead(n) = "M_PY_CNQ_Order"
    
    n = n + 1
    ar_Data(iii, n) = fst_order_LOR_TY
    ar_nmHead(n) = "TY_CNQ_Order"
    
    n = n + 1
    ar_Data(iii, n) = fst_order_LOR_M_TY
    ar_nmHead(n) = "M_TY_CNQ_Order"
    
    
    n = n + 1
    ar_Data(iii, n) = fst_order_PRTN_PY
    ar_nmHead(n) = "PY_CNQ_Order_PRTN_CA"
    
    n = n + 1
    ar_Data(iii, n) = fst_order_PRTN_TY
    ar_nmHead(n) = "TY_CNQ_Order_PRTN_CA"

'---------------------------------------------------------------------------------------------------------
  'creat ca val loreal PYvsTY YTD
'---------------------------------------------------------------------------------------------------------
    
    n = n + 1
    ar_Data(iii, n) = val_ca_PY_YTD
    ar_nmHead(n) = "CA_PY_YTD"
     
    n = n + 1
    ar_Data(iii, n) = fn_avgCA(val_ca_PY_YTD, cd_ThisMonth)
    ar_nmHead(n) = "AVG_CA_PY_YTD"
        
    n = n + 1
    ar_Data(iii, n) = val_CA_MYTD_PY
    ar_nmHead(n) = "CA_PY_M"
    
    n = n + 1
    ar_Data(iii, n) = fn_avgCA(val_CA_MYTD_PY, cd_ThisMonth)
    ar_nmHead(n) = "AVG_CA_PY_M"
    
    n = n + 1
    ar_Data(iii, n) = val_ca_TY_YTD
    ar_nmHead(n) = "CA_TY_YTD"
    
    n = n + 1
    ar_Data(iii, n) = fn_avgCA(val_ca_TY_YTD, cd_ThisMonth)
    ar_nmHead(n) = "AVG_CA_TY_YTD"
    
    n = n + 1
    ar_Data(iii, n) = val_CA_MYTD_TY
    ar_nmHead(n) = "CA_TY_M"
    
    n = n + 1
    ar_Data(iii, n) = fn_avgCA(val_CA_MYTD_TY, cd_ThisMonth)
    ar_nmHead(n) = "CA_AVG_TY_M"
     
    
    n = n + 1
    If val_ca_PY_YTD <> 0 And val_ca_TY_YTD = 0 Then
    type_cln_react = "lost"
    val_ca_PY_YTD_lost = val_ca_PY_YTD * -1
    End If

    
    If val_ca_PY_YTD = 0 And val_ca_TY_YTD = 0 Then
    sts_clnt_act = "null"
    Else
    type_cln_react = "act"
    val_ca_PY_YTD_lost = ""
    End If
        
        
    If sts_clnt_act <> 0 Then
        type_cln_react = "lost"
        val_ca_PY_YTD_lost = val_ca_PY_YTD
        Else
        type_cln_react = "act"
        val_ca_PY_YTD_lost = ""
    End If
    ar_Data(iii, n) = type_cln_react
    ar_nmHead(n) = "type_LOST"
    
         
    n = n + 1
    ar_Data(iii, n) = val_ca_PY_YTD_lost
    ar_nmHead(n) = "CA_LOST_PY"
    
    
'---------------------------------------------------------------------------------------------------------
'dt_constante
'---------------------------------------------------------------------------------------------------------
    n = n + 1
    dt_st = 0
    
   'If val_ca_PY_YTD > 0 Then dt_st = dt_st + 1
   'If val_ca_TY_YTD > 0 Then dt_st = dt_st + 1
    If cd_sts_dn_cln = 1 Then dt_st = dt_st + 1
    If nm_TypeGA_Y = "PPY" Then dt_st = dt_st + 1

    If dt_st = 2 Then
    dt_st_nm = 1

    Else
    dt_st_nm = 0

    End If
    ar_Data(iii, n) = dt_st_nm
    ar_nmHead(n) = "LfL"
    
    n = n + 1
        
    Select Case nm_TypeGA_Y
    Case "PPY"
             
    If dt_st_nm = 1 Then
    nm_TypeGA_Y_2 = nm_TypeGA_Y & " LfL"
    Else
    nm_TypeGA_Y_2 = nm_TypeGA_Y & " not LfL"
    End If
    
    Case Else
    nm_TypeGA_Y_2 = nm_TypeGA_Y
    End Select
    
    
    
    ar_Data(iii, n) = nm_TypeGA_Y_2
    ar_nmHead(n) = "nm_TypeGA_YEAR_DT"
     
    
 
    '---------------------------------------------------------------------------------------------------------
    If dt_st_nm = 1 Then
    val_ca_PY_YTD_dt = val_ca_PY_YTD
    val_ca_TY_YTD_dt = val_ca_TY_YTD
    val_CA_MYTD_PY_dt = val_CA_MYTD_PY
    val_CA_MYTD_TY_dt = val_CA_MYTD_TY
    Else
    val_ca_PY_YTD_dt = Null
    val_ca_TY_YTD_dt = Null
    val_CA_MYTD_PY_dt = Null
    val_CA_MYTD_TY_dt = Null

    End If
    
    n = n + 1
    ar_Data(iii, n) = val_ca_PY_YTD_dt
    ar_nmHead(n) = "CA_PY_LfL"
    
    n = n + 1
    ar_Data(iii, n) = val_ca_TY_YTD_dt
    ar_nmHead(n) = "CA_TY_LfL"
    
    
    n = n + 1
    ar_Data(iii, n) = val_CA_MYTD_PY_dt
    ar_nmHead(n) = "CA_M_PY_LfL"
    
    n = n + 1
    ar_Data(iii, n) = val_CA_MYTD_TY_dt
    ar_nmHead(n) = "CA_M_TY_LfL"
    
    
    

'---------------------------------------------------------------------------------------------------------
'CA YTD split by GA


    For f_qe = 1 To 3
    


        Select Case f_qe
        Case 1
        find_nm_TypeGA_Y = "PPY"

        Case 2
        find_nm_TypeGA_Y = "CNQ_PY"

        Case 3
        find_nm_TypeGA_Y = "CNQ_TY"

        End Select

    If nm_TypeGA_Y = find_nm_TypeGA_Y Then
    val_ca_PY_YTD_GA = val_ca_PY_YTD
    val_ca_TY_YTD_GA = val_ca_TY_YTD
    val_CA_MYTD_PY_GA = val_CA_MYTD_PY
    val_CA_MYTD_TY_GA = val_CA_MYTD_TY

    Else
    val_ca_PY_YTD_GA = Null
    val_ca_TY_YTD_GA = Null
    val_CA_MYTD_PY_GA = Null
    val_CA_MYTD_TY_GA = Null
    End If
          
    If val_ca_PY_YTD_GA = 0 Then val_ca_PY_YTD_GA = Null
    If val_ca_TY_YTD_GA = 0 Then val_ca_TY_YTD_GA = Null
    If val_ca_PY_YTD_GA = 0 Then val_CA_MYTD_PY_GA = Null
    If val_ca_TY_YTD_GA = 0 Then val_CA_MYTD_TY_GA = Null
         
    n = n + 1
    ar_Data(iii, n) = val_ca_PY_YTD_GA
    ar_nmHead(n) = "CA_PY_" & find_nm_TypeGA_Y
 
    n = n + 1
    ar_Data(iii, n) = val_ca_TY_YTD_GA
    ar_nmHead(n) = "CA_TY_" & find_nm_TypeGA_Y
    
    n = n + 1
    ar_Data(iii, n) = val_CA_MYTD_PY_GA
    ar_nmHead(n) = "CA_M_PY_" & find_nm_TypeGA_Y
 
    n = n + 1
    ar_Data(iii, n) = val_CA_MYTD_TY_GA
    ar_nmHead(n) = "CA_M_TY_" & find_nm_TypeGA_Y
    

    Next f_qe


'---------------------------------------------------------------------------------------------------------

If CInt(cd_year_qnc) = cd_ThisYear And cd_month_qnc <= cd_ThisMonth Then
    nm_sts_Act_CLN = "NEW"
    Else
        If val_ca_TY_YTD <> 0 And CInt(cd_year_qnc) <> cd_ThisYear Then
            nm_sts_Act_CLN = "REACTIVATED"
            Else
                If sumCA12M <> 0 And val_ca_TY_YTD = 0 Then
                    nm_sts_Act_CLN = "NOT_REACTIVATED"
                    Else
                        If sumCA_PY2LTM <> 0 And sumCA12M = 0 Then
                        nm_sts_Act_CLN = "LOST"
                            Else
                            nm_sts_Act_CLN = "OTHER"
                        End If
                End If
        End If
End If

n = n + 1
ar_Data(iii, n) = nm_sts_Act_CLN
ar_nmHead(n) = "Status_CLNT"

'---------------------------------------------------------------------------------------------------------
'creat closed data
'---------------------------------------------------------------------------------------------------------


nm_clsd_open_month = Empty

clm_m = 0

Select Case nm_sts_Act_CLN
Case "LOST"
      
        For f_m = cd_ThisMonth To 1 Step -1
        clm_m = clm_PYper_LOR_VAL + f_m - 1
        
        If Cells(f_i, clm_m) <> 0 Then
        num_clsd_open_month = f_m
        nm_clsd_open_month = ar_month_eng(f_m - 1)
        Exit For
        End If
    
        Next f_m

Case "NEW"

nm_clsd_open_month = nmMonth

Case Else
nm_clsd_open_month = Empty
End Select

    
n = n + 1
ar_Data(iii, n) = nm_clsd_open_month
ar_nmHead(n) = "Closed_Open_month"

'---------------------------------------------------------------------------------------------------------
num_last_colum = n

Next f_i
 
If nm_ActTR <> nm_WB Then
ActiveWindow.Close
Else
MsgBox ("ERR" & ar_brand(f_b))
End If

Application.DisplayAlerts = False

Next f_b
  
    
Workbooks(nm_WB).Activate
If Sheets(in_data).Visible = False Then
Sheets(in_data).Visible = True
End If
Sheets(in_data).Activate

'clear sheet & create head & create name OR fiil data
'---------------------------------------------------------------------------------------------------------

ActiveSheet.UsedRange.Cells.ClearContents
end_POS = iii
start_POS = 2

Dim n As Name
For Each n In ThisWorkbook.Names
    On Error Resume Next
    n.Delete
    Next n

For t = 1 To n
Cells(1, t) = ar_nmHead(t)
Cells(1, t).Select
ActiveWorkbook.Names.Add Name:=ar_nmHead(t), RefersTo:="=" & ActiveSheet.Name & "!" & ActiveCell.Address()
Next t


ActiveSheet.Cells(start_POS, 1).Resize(end_POS - start_POS, n) = ar_Data()
num_StatusHead = 1


ActiveWorkbook.Names.Add Name:="SOURCE", RefersToR1C1:="=OFFSET(in_TR!R1C1,0,0,COUNTA(in_TR!R1C1:R65535C1),COUNTA(in_TR!R1C1:R1C255))"
'ActiveWorkbook.Names("SOURCE").Comment = ""

ActiveWorkbook.RefreshAll
Sheets(in_data).Visible = False
'---------------------------------------------------------------------------------------------------------

fn_VBA_End

End Sub