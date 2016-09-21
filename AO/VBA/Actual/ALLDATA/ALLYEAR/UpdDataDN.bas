Sub data_CC_in_GC()


Dim pathc2file As String
Dim ar_code_Brand(6, 1)
Dim LastRow_CC, LastColumns_CC As Integer
Dim num_month

Dim patchTR, actTR, ar_LastRow, in_data, status_head   As String
Dim ALLTIME, EDU_PY, EDU_TY, place, AVG_HD As Variant
Dim b, iii, i, x, y, frqOrder   As Integer
Dim ar_nmAVG_Order(), ar_type_clients(1 To 4, 1 To 12)

'Dim   As Worksheet
Dim ar_Data(), ar_brand(), ar_nmHead(150), ar_PYPer_PRTN_VAL, ar_TYPer_PRTN_VAL, ar_PYPer_LOR_VAL(), ar_TYPer_LOR_VAL(), ar_nm_month(), ar_nmMregEN(), ar_nmMregLT()

NF = ActiveWorkbook.Name
act_month = CInt(InputBox("Month"))
act_year = CInt(InputBox("year"))


'colums CA PRTNN VAL for LTM
'---------------------------------------------------------------------------------------------------------
ar_PYPer_PRTN_VAL = Array(0, 79, 80, 81, 82, 83, 84, 85, 86, 87, 88, 89)
ar_TYPer_PRTN_VAL = Array(0, 66, 67, 68, 69, 70, 71, 72, 73, 74, 75, 76, 77)


'colums CA LOREAL VAL
str_PYper_LOR_VAL = 106
str_TYper_LOR_VAL = 93
str_PYper_PRTN_VAL = 79
str_TYper_PRTN_VAL = 66

ar_nm_month = Array("Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec")

ar_nmMregEN = Array("MOSCOW", "GR", "NORTHWEST", "CENTER", "VOLGA", "SOUTH", "URAL", "SIBERIA", "FAR EAST")
ar_nmMregLT = Array("Moscou", "GR", "Nord-Ouest", "Centre", "Volga-Centre", "Sud", "Oural", "Siberie", "EO")
ar_nmCompetitors = Array("Estel", "Schwarzkopf", "Wella", "Londa", "Keune", "Revlon", "Goldwell", "Cutrin", "Kadus", "Indola", "Paul Mitchell", "Label", "Syoss", "Chi", "Seah", "Kydra", "Sebastian", "American Crew", "Alterna", "Other")


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

ar_nmAVG_Order = Array(0, 2.5, 5, 10, 15, 20, 25, 30, 50, 60, 70, 100000)
'---------------------------------------------------------------------------------------------------------

myLib.VBA_Start

status_head = 0

in_data = "in_TR"

Workbooks(NF).Activate
If Sheets(in_data).Visible = False Then
Sheets(in_data).Visible = True
End If
Sheets(in_data).Activate

NF_LastRow = myLib.getLastRow

For f_nf = 1 To NF_LastRow
    If Cells(f_nf, 1) = 2016 Then
        Exit For
    End If
Next f_nf

ReDim Preserve ar_Data(999999, 150)

ar_brand = Array("LP", "KR", "RD", "MX", "ES", "DE", "CR")
num_ar_brand = UBound(ar_brand)
array_row = 0

iii = 0
For f_year = act_year To 2016
    actual_TY = f_year
    actual_PY = f_year - 1

    If f_year <> 2016 Then
        act_month_Y = 12
    Else
        act_month_Y = act_month
    End If

    For b = 0 To UBound(ar_brand)
    nmBrand = ar_brand(b)
        patchTR = "p:\DPP\Business development\Book commercial\" & nmBrand & "\Top Russia Total " & f_year & " " & nmBrand & ".xlsm"
        actTR = myLib.OpenFile(patchTR, nmBrand)
        ar_LastRow = myLib.getLastRow
        array_row = array_row + ar_LastRow - 4

 
    For i = 4 To ar_LastRow
    
    cdMonth = myLib.mont_num(Cells(i, 64))
    cdYear = myLib.getYearType(Cells(i, 65), 1) 

    n = 0
    ar_Data(iii, n) = f_year
    ar_nmHead(n) = "TR_year"
    
    
    n = n + 1
    ar_Data(iii, n) = nmBrand
    ar_nmHead(n) = "brand"
    
    
    n = n + 1
    ar_Data(iii, n) = myLib.type_business(nmBrand)
    ar_nmHead(n) = "bussines"
            
    n = n + 1
    ar_Data(iii, n) = Cells(i, 1)
    ar_nmHead(n) = "rowTR"
    
    n = n + 1
    cd_brand_row = nmBrand & Cells(i, 1)
    ar_Data(iii, n) = cd_brand_row
    ar_nmHead(n) = "BRAND_rowTR"
    
    n = n + 1
    cd_Univers = Cells(i, 2)
    If Len(cd_Univers) <> 9 Then
    cd_Univers = cd_brand_row
    Else: cd_Univers = cd_Univers
    End If
    ar_Data(iii, n) = cd_Univers
    ar_nmHead(n) = "unvCD"

    
    n = n + 1
    ar_Data(iii, n) =nmBrand & Cells(i, 2)
    ar_nmHead(n) = "BRAND_unvCD"
    
    n = n + 1
    nm_Mreg = getMregWhitoutBrand(Cells(i, 4)  )  
    ar_Data(iii, n) = nm_Mreg
    ar_nmHead(n) = "mreg"
    
    n = n + 1
    ar_Data(iii, n) = myLib.mreg_ext(nm_Mreg)
    ar_nmHead(n) = "mreg_EXT"
    
    n = n + 1
    ar_Data(iii, n) = Cells(i, 5)
    ar_nmHead(n) = "REG"
    
    n = n + 1
    ar_Data(iii, n) = Cells(i, 165)
    ar_nmHead(n) = "FLSM"
    
    n = n + 1
    ar_Data(iii, n) = Cells(i, 6)
    ar_nmHead(n) = "SEC"
    
    n = n + 1
    ar_Data(iii, n) = Cells(i, 7)
    ar_nmHead(n) = "SREP"
        
    n = n + 1
    ar_Data(iii, n) = salon_name(Cells(i, 9),  Cells(i, 13), Cells(i, 11))
    ar_nmHead(n) = "salon"
 
    n = n + 1
    ar_Data(iii, n) = Cells(i, 19)
    ar_nmHead(n) = "Chain_name"
    
    n = n + 1
    ar_Data(iii, n) = Cells(i, 11)
    ar_nmHead(n) = "city"
    
    n = n + 1
    ar_Data(iii, n) = myLib.clnt_type(Cells(i, 18), 1)
    ar_nmHead(n) = "type_SLN"
        
    n = n + 1
    ar_Data(iii, n) = myLib.clnt_type(Cells(i, 18), 2)
    ar_nmHead(n) = "salon_type_eng"
    
    n = n + 1
    ar_Data(iii, n) = myLib.clnt_type(Cells(i, 18), 3)
    ar_nmHead(n) = "salon_type_short_eng"
    
    n = n + 1
    ar_Data(iii, n) = myLib.clnt_type(Cells(i, 18), 4)
    ar_nmHead(n) = "salon_type_chain_eng"
    
    n = n + 1
    ar_Data(iii, n) = Cells(i, 42)
    ar_nmHead(n) = "type_CLUB"
    
    n = n + 1
    ar_Data(iii, n) = Cells(i, 40)
    ar_nmHead(n) = "type_confirmed_CLUB"
    
    n = n + 1
    ar_Data(iii, n) = cdMonth & "-" & cdYear
    ar_nmHead(n) = "date_CNQ_Y"
    
    n = n + 1
    ar_Data(iii, n) = cdMonth
    ar_nmHead(n) = "date_month_num"
    
    n = n + 1
    ar_Data(iii, n) = myLib.getNameMonthEN(cdMonth)
    ar_nmHead(n) = "date_month_name"
    
    n = n + 1
    ar_Data(iii, n) = cdYear
    ar_nmHead(n) = "date_year"
  
    n = n + 1
    ar_Data(iii, n) = myLib.getYearType(cdYear, 2)
    ar_nmHead(n) = "GA_YEAR"
        
    n = n + 1
    vl_mag = Cells(i, 160)
    If Len(vl_mag) <> 2 Then vl_mag = Null
    ar_Data(iii, n) = vl_mag
    ar_nmHead(n) = "type_MAG"
    
    n = n + 1
    ar_Data(iii, n) = Cells(i, 158)
    ar_nmHead(n) = "type_MAG_PRICE"
    
    n = n + 1
      ar_Data(iii, n) = Cells(i, 159)
    ar_nmHead(n) = "type_MAG_type_place"
    
    n = n + 1
    st_dn_cln = Cells(i, 8)
    ar_Data(iii, n) = st_dn_cln
    ar_nmHead(n) = "status_DN_num"
    
    n = n + 1
    If Cells(i, 8) = 1 Then
    st_cln_base = "Active"

    Else
    st_cln_base = "Closed"

    End If
    ar_Data(iii, n) = st_cln_base
    ar_nmHead(n) = "status_DN_name"
       
'---------------------------------------------------------------------------------------------------------
'   calculate LTM AVG CA & FrqRate
'---------------------------------------------------------------------------------------------------------
    sumCA12M = 0
    frqOrder = 0
    
    
    For iq = act_month_Y To 11
    
    
        If isNumeric(Cells(i, ar_PYPer_PRTN_VAL(iq))) Then
        CA = Cells(i, ar_PYPer_PRTN_VAL(iq))
        Else
        CA = 0
        End If
        
        sumCA12M = sumCA12M + CA
        If Cells(i, ar_PYPer_PRTN_VAL(iq)) <> "" And Cells(i, ar_PYPer_PRTN_VAL(iq)) > 0 Then
        frqOrder = frqOrder + 1
        End If
    
    Next iq
    
    For iw = 1 To act_month_Y
    
    If isNumeric(Cells(i, ar_TYPer_PRTN_VAL(iw))) Then
        CA = Cells(i, ar_TYPer_PRTN_VAL(iw))
        Else
        CA = 0
        End If
    
    sumCA12M = sumCA12M + CA
        If Cells(i, ar_TYPer_PRTN_VAL(iw)) <> "" And Cells(i, ar_TYPer_PRTN_VAL(iw)) > 0 Then
        frqOrder = frqOrder + 1
        End If
    
    Next iw
            
        If sumCA12M <> 0 Then
        AVG_CA_LTM = Round(sumCA12M / 12 / 1000, 1)
        Else
        AVG_CA_LTM = ""
        End If
'---------------------------------------------------------------------------------------------------------
    
    n = n + 1
    ar_Data(iii, n) = AVG_CA_LTM
    ar_nmHead(n) = "CA_AVG_LTM"
   
   
   
    n = n + 1
    For f_avg = 1 To UBound(ar_nmAVG_Order())
        If AVG_CA_LTM <= ar_nmAVG_Order(f_avg) And AVG_CA_LTM > ar_nmAVG_Order(f_avg - 1) Then
        
        nm_avg_CA = "'" & ar_nmAVG_Order(f_avg - 1) & "-" & ar_nmAVG_Order(f_avg)
        Exit For
        Else
        nm_avg_CA = Null
        End If
    Next f_avg
    
        If nm_avg_CA = 100000 Then nm_avg_CA = ">70"
       
    
    ar_Data(iii, n) = nm_avg_CA
    ar_nmHead(n) = "CA_AVG_LTM_name"
    
    
    

    n = n + 1
    ar_Data(iii, n) = frqOrder & "\12" '
    ar_nmHead(n) = "frq_order_LTM"
    
    
    n = n + 1
        ev_ca = Cells(i, 92)

        If isNumeric(ev_ca) Then
        ev_ca = Round(ev_ca, 2)



        Else
        ev_ca = Null
        End If
    ar_Data(iii, n) = ev_ca
    ar_nmHead(n) = "CA_ev"
   
 ' ev CA vector
 '---------------------------------------------------------------------------------------------------------
        n = n + 1
        ev_ca = Cells(i, 92)



        If isNumeric(ev_ca) Then
        Select Case ev_ca
        Case Is > 0
        nm_ev_ca = "+"

        Case Is < 0
        nm_ev_ca = "-"

        Case Else
        nm_ev_ca = Null


        End Select
        Else
        nm_ev_ca = Null


End If
                 
    ar_Data(iii, n) = nm_ev_ca
    ar_nmHead(n) = "CA_ev_name"
 '---------------------------------------------------------------------------------------------------------
    
    
    n = n + 1
    ar_Data(iii, n) = Cells(i, 29)
    ar_nmHead(n) = "EDU_id_ECAD"
    
    n = n + 1
    EDU_ALLTIME_MSTR = Cells(i, 30)
        If isNumeric(EDU_ALLTIME_MSTR) And EDU_ALLTIME_MSTR <> 0 Then
        EDU_ALLTIME_MSTR = Round(EDU_ALLTIME_MSTR, 0)
        Else
        EDU_ALLTIME_MSTR = ""
        End If
    ar_Data(iii, n) = EDU_ALLTIME_MSTR
    ar_nmHead(n) = "EDU_ALLTIME_MSTR"
    
    n = n + 1
    EDU_PY_MSTR = Cells(i, 31)
        If isNumeric(EDU_PY_MSTR) And EDU_PY_MSTR <> 0 Then
        EDU_PY = Round(EDU_PY_MSTR, 0)
        Else
        EDU_PY = ""
        End If
    ar_Data(iii, n) = EDU_PY_MSTR
    ar_nmHead(n) = "EDU_PY_MSTR"
    
    n = n + 1
    EDU_TY_MSTR = Cells(i, 32)
        If isNumeric(EDU_TY_MSTR) And EDU_TY_MSTR <> 0 Then
        EDU_TY_MSTR = Round(EDU_TY_MSTR, 0)
        Else
        EDU_TY = ""
        End If
    ar_Data(iii, n) = EDU_TY_MSTR
    ar_nmHead(n) = "EDU_TY_MSTR"
    
    
    n = n + 1
    EDU_ALLTIME_CNTCT = Cells(i, 33)
        If isNumeric(EDU_ALLTIME_CNTCT) And EDU_ALLTIME_CNTCT <> 0 Then
        EDU_ALLTIME = Round(EDU_ALLTIME_CNTCT, 0)
        Else
        EDU_ALLTIME = ""
        End If
    ar_Data(iii, n) = EDU_ALLTIME_CNTCT
    ar_nmHead(n) = "EDU_ALLTIME_CNTCT"
    
    n = n + 1
    EDU_PY_CNTCT = Cells(i, 34)
        If isNumeric(EDU_PY_CNTCT) And EDU_PY_CNTCT <> 0 Then
        EDU_PY = Round(EDU_PY_CNTCT, 0)
        Else
        EDU_PY = ""
        End If
    ar_Data(iii, n) = EDU_PY_CNTCT
    ar_nmHead(n) = "EDU_PY_CNTCT"
    
    n = n + 1
    EDU_TY_CNTCT = Cells(i, 35)
        If isNumeric(EDU_TY_CNTCT) And EDU_TY_CNTCT <> 0 Then
        EDU_TY = Round(EDU_TY_CNTCT, 0)
        Else
        EDU_TY = ""
        End If
    ar_Data(iii, n) = EDU_TY_CNTCT
    ar_nmHead(n) = "EDU_TY_CNTCT"
    
      
    
    
    n = n + 1
    place = Cells(i, 27)
        If isNumeric(place) Then
        place = Round(place, 0)
        Else
        place = ""
        End If
    ar_Data(iii, n) = place
    ar_nmHead(n) = "type_place_HD"
    
    n = n + 1
    AVG_HD = Cells(i, 28)
        If isNumeric(AVG_HD) Then
        AVG_HD = Round(AVG_HD, 0)
        Else
        place = ""
        End If
    ar_Data(iii, n) = AVG_HD
    ar_nmHead(n) = "type_AVG_HD"
    
    
    n = n + 1
    ar_Data(iii, n) = Cells(i, 209)
    ar_nmHead(n) = "com_KPI"
    
    n = n + 1
    ar_Data(iii, n) = Cells(i, 167)
    ar_nmHead(n) = "nm_PRTNner"
          
    n = n + 1
    ar_Data(iii, n) = Cells(i, 173)
    ar_nmHead(n) = "cd_PRTNner"
'---------------------------------------------------------------------------------------------------------
'creat ca val loreal monthly
'---------------------------------------------------------------------------------------------------------

    For f_m = 0 To 11
    n = n + 1
    clm_m = str_PYper_LOR_VAL + f_m
    If Cells(i, clm_m) = 0 Then
    m_val = Null

    Else
    m_val = Cells(i, clm_m) / 1000

    End If
    ar_Data(iii, n) = m_val
    ar_nmHead(n) = "CA_PY_M" & f_m + 1

    Next f_m
    

    For f_m = 0 To 11
    n = n + 1
    clm_m = str_TYper_LOR_VAL + f_m
    If Cells(i, clm_m) = 0 Then
    m_val = Null

    Else
    m_val = Cells(i, clm_m) / 1000

    End If
    ar_Data(iii, n) = m_val
    ar_nmHead(n) = "CA_TY_M" & f_m + 1


Next f_m
 '---------------------------------------------------------------------------------------------------------
  'creat ca val loreal cumul
'---------------------------------------------------------------------------------------------------------
    
    m_valP = 0

    For f_m = 0 To 11
    n = n + 1
    clm_m = str_PYper_LOR_VAL + f_m
    m_val = (Cells(i, clm_m) / 1000) + m_valP
    
    
    If m_val = 0 Then  ' del 0 value out
    ar_Data(iii, n) = Null

    Else
    ar_Data(iii, n) = m_val

    End If
    
    ar_nmHead(n) = "CA_PY_YTD" & f_m + 1
    m_valP = m_val

    Next f_m
    
    
    m_valP = 0
    For f_m = 0 To 11  ' limit tange by actuale period
    n = n + 1
    If f_m < CInt(act_month_Y) Then
    clm_m = str_TYper_LOR_VAL + f_m
    m_val = (Cells(i, clm_m) / 1000) + m_valP

    Else
    m_val = Null

    End If
    
    If m_val = 0 Then  ' del 0 value out
    ar_Data(iii, n) = Null

    Else
    ar_Data(iii, n) = m_val

    End If
    
    ar_nmHead(n) = "CA_TY_YTD" & f_m + 1
    m_valP = m_val

    Next f_m
    
'---------------------------------------------------------------------------------------------------------
'creat  ca val loreal Quarter
'---------------------------------------------------------------------------------------------------------
 
    q_m_c = 0
    For f_q = 0 To 3
    n = n + 1
    m_val_q = 0
    m_val = 0
    
        For f_m = 0 To 2
        clm_m = str_PYper_LOR_VAL + q_m_c
        m_val = Cells(i, clm_m)
        m_val_q = m_val_q + m_val
        
        q_m_c = q_m_c + 1
        
        Next f_m
        
        If m_val_q = 0 Then
        m_val_q = Null

        Else
        m_val_q = m_val_q / 1000
        End If
           
    ar_Data(iii, n) = m_val_q
    ar_nmHead(n) = "CA_PY_Q_" & f_q + 1
    Next f_q
    
    
   q_m_c = 0
    For f_q = 0 To 3
    n = n + 1
    m_val_q = 0
    m_val = 0
    
        For f_m = 0 To 2
        clm_m = str_TYper_LOR_VAL + q_m_c
        m_val = Cells(i, clm_m)
        m_val_q = m_val_q + m_val
        
        q_m_c = q_m_c + 1
        
        Next f_m
        
        If m_val_q = 0 Then
        m_val_q = Null

        Else
        m_val_q = m_val_q / 1000
        End If
           
    ar_Data(iii, n) = m_val_q
    ar_nmHead(n) = "CA_TY_Q" & f_q + 1
    Next f_q
    
 '---------------------------------------------------------------------------------------------------------
 ' first conq order
 '---------------------------------------------------------------------------------------------------------
    

    
    num_cnq_year = CInt(cdYear)
    num_cnq_month = CInt(act_month_Y)
    
    
        On Error Resume Next
          Select Case num_cnq_year
          Case actual_PY
          fst_order_LOR_PY = Cells(i, str_PYper_LOR_VAL + cdMonth - 1) / 1000
          fst_order_PRTN_PY = Cells(i, str_PYper_PRTN_VAL + cdMonth - 1) / 1000
          fst_order_LOR_TY = Null
          fst_order_PRTN_TY = Null
          Case actual_TY
          fst_order_LOR_TY = Cells(i, str_TYper_LOR_VAL + cdMonth - 1) / 1000
          fst_order_PRTN_TY = Cells(i, str_TYper_PRTN_VAL + cdMonth - 1) / 1000
          fst_order_LOR_PY = Null
          fst_order_PRTN_PY = Null
          Case Else
          fst_order_LOR_PY = Null
          fst_order_PRTN_PY = Null
          fst_order_LOR_TY = Null
          fst_order_PRTN_TY = Null
          End Select


n = n + 1
    ar_Data(iii, n) = fst_order_LOR_PY
    ar_nmHead(n) = "PY_CNQ_Order"
    
    n = n + 1
    ar_Data(iii, n) = fst_order_LOR_TY
    ar_nmHead(n) = "TY_CNQ_Order"
    
    
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
    m_valP = 0
    For f_m = 0 To 11  ' limit tange by actuale period
    If f_m < CInt(act_month_Y) Then
    clm_m = str_PYper_LOR_VAL + f_m
    m_val = (Cells(i, clm_m) / 1000) + m_valP
    Else
    Exit For
    End If
      
    m_valP = m_val

    Next f_m
    
    If m_val = 0 Then  ' del 0 value out
    ar_Data(iii, n) = Null

    Else
    ar_Data(iii, n) = m_val

    End If
    ar_nmHead(n) = "CA_PY_YTD"
    ca_ytd_PY = m_val
    
    
    n = n + 1
    m_valP = 0
    For f_m = 0 To 11  ' limit tange by actuale period
    If f_m < CInt(act_month_Y) Then
    clm_m = str_TYper_LOR_VAL + f_m
    m_val = (Cells(i, clm_m) / 1000) + m_valP



    Else
    Exit For
    End If
    
    m_valP = m_val

    Next f_m
    
    If m_val = 0 Then  ' del 0 value out
    ar_Data(iii, n) = Null

    Else
    ar_Data(iii, n) = m_val

    End If
    ar_nmHead(n) = "CA_TY_YTD"
    ca_ytd_TY = m_val
    
    n = n + 1
    
    
    If ca_ytd_PY <> 0 And ca_ytd_TY = 0 Then
    type_cln_react = "lost"
    ca_ytd_PY_lost = ca_ytd_PY * -1
    End If

    
    If ca_ytd_PY = 0 And ca_ytd_TY = 0 Then
    sts_clnt_act = "null"
    Else
    type_cln_react = "act"
    ca_ytd_PY_lost = Empty
    End If
        
        
    If sts_clnt_act <> 0 Then
        type_cln_react = "lost"
        ca_ytd_PY_lost = ca_ytd_PY
        Else
        type_cln_react = "act"
        ca_ytd_PY_lost = Empty
    End If
    ar_Data(iii, n) = type_cln_react
    ar_nmHead(n) = "type_LOST"
    
         
    n = n + 1
    ar_Data(iii, n) = ca_ytd_PY_lost
    ar_nmHead(n) = "CA_LOST_PY"
    
    
'---------------------------------------------------------------------------------------------------------
'dt_constante
'---------------------------------------------------------------------------------------------------------
    n = n + 1
    dt_st = 0
    
    If ca_ytd_PY > 0 Then dt_st = dt_st + 1
    If ca_ytd_TY > 0 Then dt_st = dt_st + 1
    If st_dn_cln = 1 Then dt_st = dt_st + 1
    If GA_Y = "PPY" Then dt_st = dt_st + 1

    If dt_st = 4 Then
    dt_st_nm = 1

    Else
    dt_st_nm = 0

    End If
    ar_Data(iii, n) = dt_st_nm
    ar_nmHead(n) = "LfL"
 
    '---------------------------------------------------------------------------------------------------------
    If dt_st_nm = 1 Then
    ca_ytd_PY_dt = ca_ytd_PY
    ca_ytd_TY_dt = ca_ytd_TY
    Else
    ca_ytd_PY_dt = Null
    ca_ytd_TY_dt = Null

    End If
    
    n = n + 1
    ar_Data(iii, n) = ca_ytd_PY_dt
    ar_nmHead(n) = "CA_PY_LfL"
    
    n = n + 1
    ar_Data(iii, n) = ca_ytd_TY_dt
    ar_nmHead(n) = "CA_TY_LfL"

'---------------------------------------------------------------------------------------------------------
'CA YTD split by GA


    For f_qe = 1 To 3
    


        Select Case f_qe
        Case 1
        find_GA_Y = "PPY"

        Case 2
        find_GA_Y = "CNQ_PY"

        Case 3
        find_GA_Y = "CNQ_TY"

        End Select

    If GA_Y = find_GA_Y Then
    ca_ytd_PY_GA = ca_ytd_PY
    ca_ytd_TY_GA = ca_ytd_TY


Else
    ca_ytd_PY_GA = Null
    ca_ytd_TY_GA = Null

    End If
          
    If ca_ytd_PY_GA = 0 Then ca_ytd_PY_GA = Null
    If ca_ytd_TY_GA = 0 Then ca_ytd_TY_GA = Null
    
         
    n = n + 1
    ar_Data(iii, n) = ca_ytd_PY_GA
    ar_nmHead(n) = "CA_PY_" & find_GA_Y
 
    n = n + 1
    ar_Data(iii, n) = ca_ytd_TY_GA
    ar_nmHead(n) = "CA_TY_" & find_GA_Y
    

    Next f_qe

'---------------------------------------------------------------------------------------------------------
'creat closed data
'---------------------------------------------------------------------------------------------------------
sts_cls_f = False
num_clsd_month = Empty
num_clsd_year = Empty
clm_m = 0

If st_dn_cln = 0 Then


    For f_yy = 1 To 2
        
        Select Case f_yy
            Case 1
            strt_month_clm = str_TYper_LOR_VAL
            Case 2
            strt_month_clm = str_PYper_LOR_VAL
        End Select

       
        
        For f_m = 11 To 0 Step -1


        clm_m = strt_month_clm + f_m
        
If f_yy = 2 Then

End If
            
        If Cells(i, clm_m) <> 0 Then
        num_clsd_month = 1 + f_m
        
        Select Case f_yy
            Case 1
            num_clsd_year = actual_TY
            Case 2
            num_clsd_year = actual_PY
        End Select
        
        sts_cls_f = True
        Exit For
        End If
    
        Next f_m
            If sts_cls_f = True Then
            Exit For
            End If
    Next f_yy
Else

End If
    
n = n + 1
ar_Data(iii, n) = num_clsd_month
ar_nmHead(n) = "Closed_M"

n = n + 1
ar_Data(iii, n) = num_clsd_year
ar_nmHead(n) = "Closed_Y"

'---------------------------------------------------------------------------------------------------------
iii = iii + 1
Next i
array_row = iii


Workbooks(actTR).Close
Application.DisplayAlerts = False

Next b
Next f_year

  
    
Workbooks(NF).Activate
If Sheets(in_data).Visible = False Then
Sheets(in_data).Visible = True
End If
Sheets(in_data).Activate

'clear sheet & create head & create name OR fiil data
'---------------------------------------------------------------------------------------------------------

Dim n As Name
For Each n In ThisWorkbook.Names
    On Error Resume Next
    n.Delete
    Next n

For t = 0 To n
Cells(1, t + 1) = ar_nmHead(t)
Cells(1, t + 1).Select
ActiveWorkbook.Names.Add Name:=ar_nmHead(t), RefersTo:="=" & ActiveSheet.Name & "!" & ActiveCell.Address()
Next t

ActiveSheet.Cells(f_nf, 1).Resize(array_row + f_nf, n + 1) = ar_Data()
status_head = 1


ActiveWorkbook.Names.Add Name:="SOURCE", RefersToR1C1:="=OFFSET(in_TR!R1C1,0,0,COUNTA(in_TR!R1C1:R999999C1),COUNTA(in_TR!R1C1:R1C1000))"
'ActiveWorkbook.Names("SOURCE").Comment = ""
Sheets(in_data).Visible = False
ActiveWorkbook.RefreshAll

'---------------------------------------------------------------------------------------------------------

myLib.VBA_End

End Sub



