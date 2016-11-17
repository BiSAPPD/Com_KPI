
Attribute VB_Name = "getUsers"
Sub getPersone(ByRef wks As Worksheet, ByVal cStatYear As Integer, ByVal cStatMonth As Integer, ByVal cBrand As String)

Dim LastColum as Long
Dim nm_Srep as String, nm_FLSM as String

LastColum = myLib.getLastColumn

For f_rw = 2 To LastColum
    nm_Mreg = myLib.getMregWhitoutBrand(myLib.fixError(Cells(f_rw, 10)))

    If Len(nm_Mreg) > 0 Then

        nm_Reg = Trim(myLib.fixError(Cells(f_rw, 11)))
        nm_mreg_EXT = myLib.mreg_lat(myLib.mreg_ext(nm_Mreg, nm_Reg))
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
