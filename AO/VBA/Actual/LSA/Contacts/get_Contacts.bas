Sub getKPIfromContact()

Dim nm_ActWb$
Dim nm_ShInData$, nm_ShOutData$, Partner$
Dim nm_ThisMonth$, nm_brand$, nm_Srep$, nm_FLSM$, nm_Sector$, nm_Reg$, nm_Mreg$, nm_Staff$, cont_email$, cont_phone$
Dim cd_Sector&, num_target_CA&
Dim ar_Brand() As Variant, ar_Data(1 To 500000, 1 To 50) As Variant, ar_nmHead(1 To 50) As Variant
Dim f_mnth%, f_brnd%, cd_ActualYear%, cd_ActualMonth%, nnnm$, patch$
Dim Experience As Variant
Dim dic_People As Variant, dic_LSA as Variant
Dim varKey As Variant, varItem As Variant
Dim sts_add2dic as Boolean
Dim objPersone As myPersone
Dim objLSA As myLSA

Set dic_People = CreateObject("Scripting.Dictionary"): dic_People.RemoveAll
Set dic_LSA = CreateObject("Scripting.Dictionary"): dic_LSA.RemoveAll

nm_ActWb        = ActiveWorkbook.Name
cd_ActualMonth  = CInt(InputBox("Month"))
cd_ActualYear   = 2016
ar_Brand        = Array("MX", "ES", "LP", "KR", "RD")

nm_ShOutData    = "Contacts"
nm_ShInData     = "Cnt_SREP"

MyLib.VBA_Start
MyLib.CreateSh (nm_ShInData)
iii = 1

For f_mnth = 1 To cd_ActualMonth
    For f_brnd = 0 To UBound(ar_Brand)
        nm_brand = ar_Brand(f_brnd)
            
            patch           = MyLib.patch_history_TR(nm_brand, cd_ActualYear, cd_ActualMonth, f_mnth)
            actTR           = MyLib.OpenFile(patch, nm_ShOutData)
            num_LastRow     = MyLib.getLastRow
            num_LastColum   = MyLib.getLastColumn
        
          
        For f_rw = 2 To num_LastRow
            nm_Mreg         = MyLib.getMregWhitoutBrand(Cells(f_rw, 10))
            nm_Reg          = Trim(Cells(f_rw, 11))
            nm_mreg_EXT     = MyLib.mreg_lat(MyLib.mreg_ext(nm_Mreg, nm_Reg))
            
            If Len(nm_mreg_EXT) > 0 Then
                     
                nm_Srep             = Trim(Cells(f_rw, 3))
                nm_FLSM             = Trim(Cells(f_rw, 6))
                nm_Sector           = Trim(Cells(f_rw, 1))
                nm_Staff            = Cells(f_rw, 4)
                cont_email          = Trim(Cells(f_rw, 8))
                cont_phone          = Trim(Cells(f_rw, 7))
                Partner             = Trim(Cells(f_rw, 9))
                Experience          = MyLib.getLast4quartal(Cells(f_rw, 12), cd_ActualMonth, cd_ActualYear)
                num_target_CA       = MyLib.num2num0(Cells(f_rw, 14))
                num_orders_SLN      = MyLib.num2num0(Cells(f_rw, 15))
                num_orders_phone    = MyLib.num2num0(Cells(f_rw, 16))
                num_visits2act      = MyLib.num2num0(Cells(f_rw, 17))
                num_visited_act     = MyLib.num2num0(Cells(f_rw, 18))
                num_visits2cnq      = MyLib.num2num0(Cells(f_rw, 19))
                num_visited_cnq     = MyLib.num2num0(Cells(f_rw, 20))
                nm_month            = MyLib.getNameMonthEN(f_mnth)
                nm_vacancy_status   = MyLib.getSREP_type(nm_Srep, nm_FLSM)
                
                For f_p = 1 To 2
                    sts_add2dic = false
                    Select Case f_p
                        Case 1: key_Persone = nm_month & nm_FLSM: sts_add2dic = true
                        Case 2: key_Persone = nm_month & nm_Srep: If nm_vacancy_status = "active" Then sts_add2dic = true
                    End Select

                    If Not dic_People.Exists(key_Persone) and sts_add2dic = true Then
                        Set objPersone = New myPersone
                        objPersone.cdDateStat    = DateSerial(cd_ActualYear, cd_ActualMonth, 1)
                        objPersone.MegaReg       = nm_mreg_EXT
                        Select Case f_p
                            Case 1
                                objPersone.Name             = nm_FLSM
                                objPersone.Role             = "FLSM"
                                objPersone.Experience       = "OLD"
                            Case 2
                                objPersone.Name             = nm_Srep
                                objPersone.Status           = nm_Staff
                                objPersone.Mail             = cont_email
                                objPersone.Experience       = Experience
                                objPersone.Role             = "SREP"
                        End Select
                        dic_People.Add key_Persone, objPersone
                    End If

                    If dic_People.Exists(key_Persone) Then 
                        Select Case nm_brand
                            Case "LP": dic_People.Item(key_Persone).Brand_LP = nm_brand
                            Case "MX": dic_People.Item(key_Persone).Brand_MX = nm_brand
                            Case "KR": dic_People.Item(key_Persone).Brand_KR = nm_brand
                            Case "RD": dic_People.Item(key_Persone).Brand_RD = nm_brand
                            Case "ES": dic_People.Item(key_Persone).Brand_ES = nm_brand
                            Case "DE": dic_People.Item(key_Persone).Brand_DE = nm_brand
                            Case "CR": dic_People.Item(key_Persone).Brand_CR = nm_brand
                        End Select
                    End If
                Next f_p

                n = 0 + 1: ar_nmHead(n) = "months":         ar_Data(iii, n) = nm_month
                n = n + 1: ar_nmHead(n) = "num_months":     ar_Data(iii, n) = f_mnth
                n = n + 1: ar_nmHead(n) = "brand":          ar_Data(iii, n) = nm_brand
                n = n + 1: ar_nmHead(n) = "mreg":           ar_Data(iii, n) = nm_Mreg
                n = n + 1: ar_nmHead(n) = "mreg_EXT":       ar_Data(iii, n) = nm_mreg_EXT
                n = n + 1: ar_nmHead(n) = "REG":            ar_Data(iii, n) = nm_Reg
                n = n + 1: ar_nmHead(n) = "FLSM":           ar_Data(iii, n) = nm_FLSM
                n = n + 1: ar_nmHead(n) = "SEC":            ar_Data(iii, n) = nm_Sector
                n = n + 1: ar_nmHead(n) = "SREP":           ar_Data(iii, n) = nm_Srep
                n = n + 1: ar_nmHead(n) = "staff":          ar_Data(iii, n) = nm_Staff
                n = n + 1: ar_nmHead(n) = "cont_email":     ar_Data(iii, n) = cont_email
                n = n + 1: ar_nmHead(n) = "cont_phone":     ar_Data(iii, n) = cont_phone
                n = n + 1: ar_nmHead(n) = "partner":        ar_Data(iii, n) = Partner
                n = n + 1: ar_nmHead(n) = "experience":     ar_Data(iii, n) = Experience
                n = n + 1: ar_nmHead(n) = "vacancy_status": ar_Data(iii, n) = nm_vacancy_status
                n = n + 1: ar_nmHead(n) = "target_CA":      ar_Data(iii, n) = num_target_CA
                n = n + 1: ar_nmHead(n) = "orders_SLN":     ar_Data(iii, n) = num_orders_SLN
                n = n + 1: ar_nmHead(n) = "orders_phone":   ar_Data(iii, n) = num_orders_phone
                n = n + 1: ar_nmHead(n) = "visits2act":     ar_Data(iii, n) = num_visits2act
                n = n + 1: ar_nmHead(n) = "visited_act":    ar_Data(iii, n) = num_visited_act
                n = n + 1: ar_nmHead(n) = "visits2cnq":     ar_Data(iii, n) = num_visits2cnq
                n = n + 1: ar_nmHead(n) = "visited_cnq":    ar_Data(iii, n) = num_visited_cnq
            
            iii = iii + 1
            End If

        Next f_rw
        
    Next f_brnd
Next f_mnth

Workbooks(nm_ActWb).Activate
MyLib.sheetActivateCleer(nm_ShInData)

For f_head = 1 To n
    If IsEmpty(ar_nmHead(f_head)) Then
        head_clmn_name = f_head
        Else
        head_clmn_name = ar_nmHead(f_head)
    End If
    Cells(1, f_head) = head_clmn_name
Next f_head

ActiveSheet.Cells(2, 1).Resize(iii, n) = ar_Data()
Cells(1, 1).Select

ActiveWorkbook.RefreshAll

ActiveWindow.FreezePanes = False
Cells(2, 16).Select
ActiveWindow.FreezePanes = True
ActiveWindow.DisplayGridlines = False

Dim LSADataPatch as String$, ShLSAoutData$
ShLSAoutData = "eduT"
LSADataPatch = "p:\DPP\Business development\LSA\DATA\EduT.xlsm"
MyLib.OpenFile(LSADataPatch, ShLSAoutData)
For f_c = 2 to MyLib.getLastRow
    



nm_ShInUniqPersone = "Dict"
MyLib.CreateSh (nm_ShInUniqPersone)
MyLib.sheetActivateCleer(nm_ShInUniqPersone)

For Each myPersone In dic_People.Items
    i = i + 1
    n = 0
    n = n + 1: Cells(i, n) = MyLib.getNameMonthEN(Month(myPersone.cdDateStat))
    n = n + 1: Cells(i, n) = Year(myPersone.cdDateStat)
    n = n + 1: Cells(i, n) = myPersone.Name
    n = n + 1: Cells(i, n) = myPersone.Role
    n = n + 1: Cells(i, n) = myPersone.Status
    n = n + 1: Cells(i, n) = myPersone.Experience
    n = n + 1: Cells(i, n) = myPersone.Brand_LP
    n = n + 1: Cells(i, n) = myPersone.Brand_MX
    n = n + 1: Cells(i, n) = myPersone.Brand_KR
    n = n + 1: Cells(i, n) = myPersone.Brand_RD
    n = n + 1: Cells(i, n) = myPersone.Brand_ES
    n = n + 1: Cells(i, n) = myPersone.Brand_DE
    n = n + 1: Cells(i, n) = myPersone.Brand_CR
Next

MyLib.VBA_End
End Sub