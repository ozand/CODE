Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Sub getKPIfromContact()

Dim nm_ActWb$
Dim nm_ShInData$, nm_ShOutData$, Partner$
Dim nm_ThisMonth$, nm_brand$, nm_Srep$, nm_FLSM$, nm_Sector$, nm_Reg$, nm_Mreg$, nm_Staff$, cont_email$, cont_phone$
Dim cd_Sector&, num_target_CA&
Dim ar_Brand() As Variant, ar_Data(1 To 500000, 1 To 50) As Variant, ar_nmHead(1 To 50) As Variant
Dim f_mnth%, f_brnd%, cd_ActualYear%, cd_ActualMonth%, nnnm$, patch$, endMonth%, cd_StartYear%, f_year%
Dim Experience As Variant
Dim dic_People As Variant, dic_SeminarName As Variant
Dim varKey As Variant, varItem As Variant
Dim sts_add2dic As Boolean
Dim objUser As UserData

Set dic_People = CreateObject("Scripting.Dictionary"): dic_People.RemoveAll
Set dic_SeminarName = CreateObject("Scripting.Dictionary"): dic_SeminarName.RemoveAll
Set dic_UserEducated = CreateObject("Scripting.Dictionary"): dic_UserEducated.RemoveAll

nm_ActWb = ActiveWorkbook.Name
cd_ActualMonth = CInt(InputBox("Month"))
cd_StartYear = CInt(InputBox("YearStart"))
cd_ActualYear = CInt(InputBox("YearEnd"))

ar_Brand = Array("LP", "MX", "KR", "RD", "ES")

nm_ShOutData = "Contacts"
nm_ShInData = "Cnt_Persone"

myLib.VBA_Start
myLib.CreateSh (nm_ShInData)
iii = 1

For f_year = cd_StartYear to cd_ActualYear
    Select Case f_year
        Case  cd_ActualYear: endMonth = cd_ActualMonth
        Case Else: endMonth = 12
    End Select

    For f_mnth = 1 To endMonth
        For f_brnd = 0 To UBound(ar_Brand)
            nm_brand = ar_Brand(f_brnd)
                
                patch = myLib.patch_history_TR(nm_brand, cd_ActualYear, f_year, cd_ActualMonth, f_mnth)
                actTR = myLib.OpenFile(patch, nm_ShOutData)
                num_LastRow = myLib.getLastRow
                num_LastColum = myLib.getLastColumn
            
            
            For f_rw = 2 To num_LastRow
                nm_Mreg = myLib.getMregWhitoutBrand(myLib.fixError(Cells(f_rw, 10)))
                nm_Reg = Trim(myLib.fixError(Cells(f_rw, 11)))
                nm_mreg_EXT = myLib.mreg_lat(myLib.mreg_ext(nm_Mreg, nm_Reg))
                
                If Len(nm_mreg_EXT) > 0 Then
                        
                    nm_Srep = Trim(Cells(f_rw, 3))
                    nm_FLSM = Trim(Cells(f_rw, 6))
                    nm_Sector = Trim(Cells(f_rw, 1))
                    nm_Staff = myLib.getStatus(Cells(f_rw, 4))
                    cont_email = Trim(Cells(f_rw, 8))
                    cont_phone = Trim(Cells(f_rw, 7))
                    Partner = Trim(Cells(f_rw, 9))
                    Experience = myLib.getLast4quartal(Cells(f_rw, 12), f_mnth, f_year)
                    num_target_CA = myLib.num2num0(Cells(f_rw, 14))
                    num_orders_SLN = myLib.num2num0(Cells(f_rw, 15))
                    num_orders_phone = myLib.num2num0(Cells(f_rw, 16))
                    num_visits2act = myLib.num2num0(Cells(f_rw, 17))
                    num_visited_act = myLib.num2num0(Cells(f_rw, 18))
                    num_visits2cnq = myLib.num2num0(Cells(f_rw, 19))
                    num_visited_cnq = myLib.num2num0(Cells(f_rw, 20))
                    nm_month = myLib.getNameMonthEN(f_mnth)
                    nm_vacancy_status = myLib.getSREP_type(nm_Srep, nm_FLSM)
                    
                    For f_p = 1 To 2
                        sts_add2dic = False
                        Select Case f_p
                            Case 1: keyUser = f_year & nm_month & nm_FLSM: sts_add2dic = True
                            Case 2: keyUser = f_year & nm_month & nm_Srep: If nm_vacancy_status = "active" Then sts_add2dic = True
                        End Select

                        If Not dic_People.Exists(keyUser) And sts_add2dic = True Then
                            Set objUser = New UserData
                            objUser.cdDateStat = DateSerial(f_year, f_mnth, 1)
                            objUser.MegaReg = nm_mreg_EXT
                            Select Case f_p
                                Case 1
                                    objUser.PersonName = nm_FLSM
                                    objUser.Role = "FLSM"
                                    objUser.Experience = "OLD"
                                Case 2
                                    objUser.PersonName = nm_Srep
                                    objUser.Status = nm_Staff
                                    objUser.Mail = cont_email
                                    objUser.Experience = Experience
                                    objUser.Role = "SREP"
                            End Select
                            dic_People.Add keyUser, objUser
                        End If

                        If dic_People.Exists(keyUser) Then
                        With dic_People
                            Select Case nm_brand
                                Case "LP": .Item(keyUser).Brand_LP = nm_brand
                                Case "MX": .Item(keyUser).Brand_MX = nm_brand
                                Case "KR": .Item(keyUser).Brand_KR = nm_brand
                                Case "RD": .Item(keyUser).Brand_RD = nm_brand
                                Case "ES": .Item(keyUser).Brand_ES = nm_brand
                                Case "DE": .Item(keyUser).Brand_DE = nm_brand
                                Case "CR": .Item(keyUser).Brand_CR = nm_brand
                            End Select
                        End With
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
        myLib.CloseNoMotherBook (nm_ActWb)
        Next f_brnd
    Next f_mnth
Next f_year
Workbooks(nm_ActWb).Activate
myLib.sheetActivateCleer (nm_ShInData)

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


nm_ShInUniqPersone = "Users"
myLib.CreateSh (nm_ShInUniqPersone)
myLib.sheetActivateCleer (nm_ShInUniqPersone)

i = 1
For Each UserData In dic_People.Items
    i = i + 1
    n = 0
    n = n + 1: Cells(i, n) = myLib.getNameMonthEN(month(UserData.cdDateStat))
    n = n + 1: Cells(i, n) = year(UserData.cdDateStat)
    n = n + 1: Cells(i, n) = UserData.PersonName
    n = n + 1: Cells(i, n) = UserData.Role
    n = n + 1: Cells(i, n) = UserData.Status
    n = n + 1: Cells(i, n) = UserData.Experience
    n = n + 1: Cells(i, n) = UserData.Brand_LP
    n = n + 1: Cells(i, n) = UserData.Brand_MX
    n = n + 1: Cells(i, n) = UserData.Brand_KR
    n = n + 1: Cells(i, n) = UserData.Brand_RD
    n = n + 1: Cells(i, n) = UserData.Brand_ES
    n = n + 1: Cells(i, n) = UserData.Brand_DE
    n = n + 1: Cells(i, n) = UserData.Brand_CR
Next

Dim LSADataPatch$, ShLSAoutData$
Dim ShIn$
Dim smr As Seminars, smu As SeminarUsers

ShLSAoutData = "eduT"
LSADataPatch = "p:\DPP\Business development\LSA\DATA\EduT.xlsm"
nmWbLSA = myLib.OpenFile(LSADataPatch, ShLSAoutData)
Workbooks(nmWbLSA).Activate
Set smr = New Seminars
ShIn = "eduT"
ShOut = "Education"

Sheets(ShIn).Select
smr.FillFromSheet ActiveSheet
Workbooks(nmWbLSA).Close
Workbooks(nm_ActWb).Activate 
myLib.CreateSh (ShOut)
myLib.sheetActivateCleer (ShOut)

i = 1
For Each smu In smr
i = i + 1
    With smu
        n = 1: Cells(i, n) = .PersonName
        n = n + 1: Cells(i, n) = .EduDate
        n = n + 1: Cells(i, n) = .SeminarName
        n = n + 1: Cells(i, n) = .Educater
    End With

    With dic_SeminarName
        If Not .Exists(smu.SeminarName) Then .Add smu.SeminarName, .count + 1
    End With  

Next

Dim objEduUSR As UserEducated

For Each UserData In dic_People.Items
    With UserData
        keyEduUSR = .PersonName & .cdDateStat
    End With
    If Not dic_UserEducated.Exists(keyEduUSR) Then
        Set objEduUSR = New UserEducated
        With UserData
            objEduUSR.cdDateStat        = .cdDateStat
            objEduUSR.PersonName        = .PersonName
            objEduUSR.Role              = .Role
            objEduUSR.Status            = .Status  
            objEduUSR.Mobile            = .Mobile 
            objEduUSR.Mail              = .Mail       
            objEduUSR.Partner           = .Partner 
            objEduUSR.Experience        = .Experience 
            objEduUSR.Territory         = .Territory
            objEduUSR.ParentTerritory   = .ParentTerritory 
            objEduUSR.Brand_LP          = .Brand_LP          
            objEduUSR.Brand_MX          = .Brand_MX
            objEduUSR.Brand_KR          = .Brand_KR
            objEduUSR.Brand_RD          = .Brand_RD
            objEduUSR.Brand_ES          = .Brand_ES
            objEduUSR.Brand_DE          = .Brand_DE
            objEduUSR.Brand_CR          = .Brand_CR
            objEduUSR.MegaReg           = .MegaReg
            objEduUSR.TeamType          = 1
            objEduUSR.EduDate           = Empty
        End With
        dic_UserEducated.Add keyEduUSR, objEduUSR
    End if
Next

For f_year = cd_StartYear to cd_ActualYear
    For Each smu in smr
        With smu
            cdDateStat = DateSerial(Year(.EduDate), Month(.EduDate), 1)
            keyEduUSR = .PersonName & cdDateStat
            If Not dic_UserEducated.Exists(keyEduUSR) and Year(.EduDate) = f_year  Then
                Set objEduUSR = New UserEducated

                objEduUSR.cdDateStat        = cdDateStat
                objEduUSR.PersonName        =.PersonName
                objEduUSR.Role              = "UNKNOW"
                objEduUSR.Brand_Other       = 1
                objEduUSR.Experience        = "OLD"
                objEduUSR.TeamType          = Empty
                objEduUSR.EduDate           = .EduDate
                objEduUSR.diffEduDate       = 0
                objEduUSR.EducatedStatus    = Empty
                dic_UserEducated.Add keyEduUSR, objEduUSR
            End If
        End With
    Next
Next f_year



For Each smu in smr
    With smu
        dateEduUsr =  DateSerial(Year(.EduDate), Month(.EduDate), 1)
        numSMR = dic_SeminarName.Item(.SeminarName)
    End With
    For f_y = 0 to 35
        PeriodY3EduUsr = DateAdd("m", f_y,  dateEduUsr)
        keyEduUSR = smu.PersonName & PeriodY3EduUsr
        Select Case f_y
            Case 0: diffEduDate = f_y
            Case Else: diffEduDate = f_y * -1
        End Select

        With dic_UserEducated
            If .Exists(keyEduUSR) Then
                Select Case numSMR
                    Case 1: .Item(keyEduUSR).Seminar1 = 1
                    Case 2: .Item(keyEduUSR).Seminar2 = 1
                    Case 3: .Item(keyEduUSR).Seminar3 = 1
                    Case 4: .Item(keyEduUSR).Seminar4 = 1
                    Case 5: .Item(keyEduUSR).Seminar5 = 1
                    Case 6: .Item(keyEduUSR).Seminar6 = 1
                    Case 7: .Item(keyEduUSR).Seminar7 = 1
                    Case 8: .Item(keyEduUSR).Seminar8 = 1
                    Case 9: .Item(keyEduUSR).Seminar9 = 1
                    Case 10: .Item(keyEduUSR).Seminar10 = 1
                    Case 11: .Item(keyEduUSR).Seminar11 = 1
                    Case 12: .Item(keyEduUSR).Seminar12 = 1
                    Case 13: .Item(keyEduUSR).Seminar13 = 1
                    Case 14: .Item(keyEduUSR).Seminar14 = 1
                    Case 15: .Item(keyEduUSR).Seminar15 = 1
                    Case 16: .Item(keyEduUSR).Seminar16 = 1
                    Case 17: .Item(keyEduUSR).Seminar17 = 1
                    Case 18: .Item(keyEduUSR).Seminar18 = 1
                    Case 19: .Item(keyEduUSR).Seminar19 = 1
                    Case 20: .Item(keyEduUSR).Seminar20 = 1
                End Select


                If .Item(keyEduUSR).EduDate > smu.EduDate or IsEmpty(smu.EduDate) Then 
                    .Item(keyEduUSR).EduDate = smu.EduDate
                    .Item(keyEduUSR).diffEduDate = diffEduDate
                End If
                .Item(keyEduUSR).EducatedStatus = 1
            End If
        End With
    Next f_y
Next

nm_ShEducatedUsers = "LSA"
myLib.CreateSh (nm_ShEducatedUsers)
myLib.sheetActivateCleer (nm_ShEducatedUsers)        
Sheets(nm_ShEducatedUsers).Select

i = 0   
SmrCnt = dic_SeminarName.Count
For Each UserEducated In dic_UserEducated.Items
    i = i + 1
    n = 0
    With UserEducated
        n = n + 1: Cells(i, n) = IIF(i = 1, "Month", myLib.getNameMonthEN(month(.cdDateStat)))
        n = n + 1: Cells(i, n) = IIF(i = 1, "Year", year(.cdDateStat))
        n = n + 1: Cells(i, n) = IIF(i = 1, "PersonName", .PersonName)
        n = n + 1: Cells(i, n) = IIF(i = 1, "Role", .Role)
        n = n + 1: Cells(i, n) = IIF(i = 1, "Status", .Status)
        n = n + 1: Cells(i, n) = IIF(i = 1, "Mobile", .Mobile)
        n = n + 1: Cells(i, n) = IIF(i = 1, "Mail", .Mail)
        n = n + 1: Cells(i, n) = IIF(i = 1, "Partner", .Partner)
        n = n + 1: Cells(i, n) = IIF(i = 1, "Experience", .Experience)
        n = n + 1: Cells(i, n) = IIF(i = 1, "Territory", .Territory)
        n = n + 1: Cells(i, n) = IIF(i = 1, "ParentTerritory", .ParentTerritory)
        n = n + 1: Cells(i, n) = IIF(i = 1, "LP", .Brand_LP)
        n = n + 1: Cells(i, n) = IIF(i = 1, "MX", .Brand_MX)
        n = n + 1: Cells(i, n) = IIF(i = 1, "KR", .Brand_KR)
        n = n + 1: Cells(i, n) = IIF(i = 1, "RD", .Brand_RD)
        n = n + 1: Cells(i, n) = IIF(i = 1, "ES", .Brand_ES)
        n = n + 1: Cells(i, n) = IIF(i = 1, "DE", .Brand_DE)
        n = n + 1: Cells(i, n) = IIF(i = 1, "CR", .Brand_CR)
        n = n + 1: Cells(i, n) = IIF(i = 1, "Other", .Brand_Other)
        n = n + 1: Cells(i, n) = IIF(i = 1, "TeamType", .TeamType)
        n = n + 1: Cells(i, n) = IIF(i = 1, "MegaReg", .MegaReg)
        n = n + 1: Cells(i, n) = IIF(i = 1, "EduDate", .EduDate)
        n = n + 1: Cells(i, n) = IIF(i = 1, "diffEduDate", .diffEduDate)
        n = n + 1: Cells(i, n) = IIF(i = 1, "Educater", .Educater)
        n = n + 1: Cells(i, n) = IIF(i = 1, "EducatedStatus", .EducatedStatus)
        On Error Resume Next
        n = n + 1: Cells(i, n) = IIF(i = 1, dic_SeminarName.Keys()(0), .Seminar1)
        n = n + 1: Cells(i, n) = IIF(i = 1, dic_SeminarName.Keys()(1), .Seminar2)
        n = n + 1: Cells(i, n) = IIF(i = 1, dic_SeminarName.Keys()(2), .Seminar3)
        n = n + 1: Cells(i, n) = IIF(i = 1, dic_SeminarName.Keys()(3), .Seminar4)
        n = n + 1: Cells(i, n) = IIF(i = 1, dic_SeminarName.Keys()(4), .Seminar5)
        n = n + 1: Cells(i, n) = IIF(i = 1, dic_SeminarName.Keys()(5), .Seminar6)
        n = n + 1: Cells(i, n) = IIF(i = 1, dic_SeminarName.Keys()(6), .Seminar7)
        n = n + 1: Cells(i, n) = IIF(i = 1, dic_SeminarName.Keys()(7), .Seminar8)
        n = n + 1: Cells(i, n) = IIF(i = 1, dic_SeminarName.Keys()(8), .Seminar9)
        n = n + 1: Cells(i, n) = IIF(i = 1, dic_SeminarName.Keys()(9), .Seminar10)
        n = n + 1: Cells(i, n) = IIF(i = 1, dic_SeminarName.Keys()(10), .Seminar11)
        n = n + 1: Cells(i, n) = IIF(i = 1, dic_SeminarName.Keys()(11), .Seminar12)
        n = n + 1: Cells(i, n) = IIF(i = 1, dic_SeminarName.Keys()(12), .Seminar13)
        n = n + 1: Cells(i, n) = IIF(i = 1, dic_SeminarName.Keys()(13), .Seminar14)
        n = n + 1: Cells(i, n) = IIF(i = 1, dic_SeminarName.Keys()(14), .Seminar15)
        n = n + 1: Cells(i, n) = IIF(i = 1, dic_SeminarName.Keys()(15), .Seminar16)
        n = n + 1: Cells(i, n) = IIF(i = 1, dic_SeminarName.Keys()(16), .Seminar17)
        n = n + 1: Cells(i, n) = IIF(i = 1, dic_SeminarName.Keys()(17), .Seminar18)
        n = n + 1: Cells(i, n) = IIF(i = 1, dic_SeminarName.Keys()(18), .Seminar19)
        n = n + 1: Cells(i, n) = IIF(i = 1, dic_SeminarName.Keys()(19), .Seminar20)
    End  With
Next


myLib.VBA_End
End Sub
