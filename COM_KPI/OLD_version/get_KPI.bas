Option Compare Text

Sub data_TR_add_Y()

Dim brand, ar_Colmn(), patchTR, nm_FLSM, nm_Mreg, nm_SREP, type_CLNT As String
Dim comp_colo, comp_rev, st_cmp, nmFile, disk, map_search, actTR, ActWb As String
Dim num_ar_Brand, num_ar_Colmn, ar_LastRow(), f_brnd, iii, f_rw, nc, ee, cdMonth, cdYear, CA1 As Integer
Dim eee, num_colums, CA, Q1, Q2, Q3, Q4, a, z, dogovor, club_2015, st_pot_club, clnt_err, st_club, f As Integer
Dim Type_bonus As Double
Dim in_data, Sh As Worksheet
Dim ar_Data_TR(), ar_Data_CNTCT(), ar_CA_PART_VAL(), ar_nmHead_TR(150), ar_nmAVG_Order(), ar_nmHead_CNTCT(50), ar_Data_COACH(), ar_nmHead_COACH()
Dim discount, koef As Double

ActWb = ActiveWorkbook.Name
cd_ActualMonth = CInt(InputBox("Month"))
cd_ActualYear = CInt(InputBox("YearEnd"))

ar_Brand = Array("LP", "MX", "KR", "RD", "ES")

ReDim ar_Data_TR(500000, 200) ' num_colums)
ReDim ar_Data_CNTCT(500000, 50)
ReDim ar_Data_COACH(500000, 50) ' num_colums)
ReDim ar_nmHead_COACH(50)


status_head = 0

nm_ShInDataTR = "DPP"

myLib.VBA_start
myLib.CreateSh (nm_ShInData)



iii = 0
xxx = 0
yyy = 0


'---------------------------------------------------------------------------------------------------------
For f_mnth = 1 To cd_ActualMonth
    For f_brnd = 0 To UBound(ar_brand)
        nm_brand = ar_brand(f_brnd)
        nm_ShOutDataTR = nm_brand
        patch = myLib.patch_history_TR(nm_brand, cd_ActualYear, cd_ActualYear, cd_ActualMonth, f_mnth)
        actTR = myLib.OpenFile(patch, nm_ShOutDataTR)
        LastRow = myLib.getLastRow
        LastColum = myLib.getLastColumn

        Application.StatusBar = "OpeActWbFile: " & actTR & " lastRow : " & LastRow & " iii: " & iii & "  "
    
        For f_rw = 4 To LastRow

            num_colums_TR = 0
            nm_short_month = ar_nm_short_month(f_mnth - 1)
            ar_Data_TR(iii, num_colums_TR) = nm_short_month
            If iii = 0 Then ar_nmHead_TR(num_colums_TR) = "month"
            
            num_colums_TR = num_colums_TR + 1
            ar_Data_TR(iii, num_colums_TR) = nm_brand
            If iii = 0 Then ar_nmHead_TR(num_colums_TR) = "brand"
                
            
            num_colums_TR = num_colums_TR + 1
            nm_Mreg = Cells(f_rw, 4)
            On Error Resume Next
            If Left(nm_Mreg, 2) = nm_brand Then
            nm_Mreg = Right(Cells(f_rw, 4), Len(Cells(f_rw, 4).Value) - 3)
            End If
            ar_Data_TR(iii, num_colums_TR) = nm_Mreg
            If iii = 0 Then ar_nmHead_TR(num_colums_TR) = "mreg"
            
            num_colums_TR = num_colums_TR + 1
            ar_Data_TR(iii, num_colums_TR) = nm_Mreg
            If iii = 0 Then ar_nmHead_TR(num_colums_TR) = "#mreg"

            
            'Mreg LT-> EN + split Moscou GR
            '---------------------------------------------------------------------------------------------------------
                    
            num_colums_TR = num_colums_TR + 1
            textPos = 0
            
            
            If nm_Mreg = "Moscou GR" Then
            nm_reg = Cells(f_rw, 5)
            textPos = InStr(nm_reg, "MSK")
            textPos = InStr(nm_reg, "Moscou") + textPos
                If textPos > 0 Then
                nm_Mreg = "Moscou"
                Else
                nm_Mreg = "GR"

                End If
            End If
            
            For f_mr = 0 To UBound(ar_nmMregLT)
            If ar_nmMregLT(f_mr) = nm_Mreg Then
            nm_Mreg = ar_nmMregEN(f_mr)
            End If
            Next f_mr
            
            ar_Data_TR(iii, num_colums_TR) = nm_Mreg
            If iii = 0 Then ar_nmHead_TR(num_colums_TR) = "mreg_EXT"
            
            '--------------------------------------------------------------

            num_colums_TR = num_colums_TR + 1
            ar_Data_TR(iii, num_colums_TR) = Cells(f_rw, 165)
            If iii = 0 Then ar_nmHead_TR(num_colums_TR) = "FLSM"
            
            num_colums_TR = num_colums_TR + 1
            ar_Data_TR(iii, num_colums_TR) = Cells(f_rw, 165)
            If iii = 0 Then ar_nmHead_TR(num_colums_TR) = nm_short_month & "#FLSM" & nm_brand
            
            num_colums_TR = num_colums_TR + 1
            ar_Data_TR(iii, num_colums_TR) = Cells(f_rw, 6)
            If iii = 0 Then ar_nmHead_TR(num_colums_TR) = "SEC"
            
            num_colums_TR = num_colums_TR + 1
            ar_Data_TR(iii, num_colums_TR) = Cells(f_rw, 7)
            If iii = 0 Then ar_nmHead_TR(num_colums_TR) = "SREP"
            
            num_colums_TR = num_colums_TR + 1
            ar_Data_TR(iii, num_colums_TR) = nm_short_month & Cells(f_rw, 7) & nm_brand
            If iii = 0 Then ar_nmHead_TR(num_colums_TR) = "#SREP"
            
            
            
            '---------------------------------------------------------------------------------------------------------
            '---------------------------------------------------------------------------------------------------------
            'open month
            '---------------------------------------------------------------------------------------------------------

                If Cells(f_rw, 161) <> "" Then cdMonth = fn_mont_num&(Cells(f_rw, 64).Value)
                If Len(Cells(f_rw, 65)) = 4 Then cdYear = Cells(f_rw, 65) Else cdYear = 2008

                For f_m = 0 To 11
                If cdMonth - 1 = f_m Then
                nmMonth = ar_nm_short_month(f_m)
                Exit For
                End If
                Next f_m
            '---------------------------------------------------------------------------------------------------------

            
        
            
            num_colums_TR = num_colums_TR + 1
            ar_Data_TR(iii, num_colums_TR) = cdMonth
            If iii = 0 Then ar_nmHead_TR(num_colums_TR) = "date_month_num"
            
            
            num_colums_TR = num_colums_TR + 1


            
            ar_Data_TR(iii, num_colums_TR) = nmMonth
            If iii = 0 Then ar_nmHead_TR(num_colums_TR) = "date_month_name"
            
            num_colums_TR = num_colums_TR + 1
            ar_Data_TR(iii, num_colums_TR) = cdYear
            If iii = 0 Then ar_nmHead_TR(num_colums_TR) = "date_year"
            
            '--------------------------------------
                
            num_colums_TR = num_colums_TR + 1
            st_dn_cln = Cells(f_rw, 8)
            ar_Data_TR(iii, num_colums_TR) = st_dn_cln
            If iii = 0 Then ar_nmHead_TR(num_colums_TR) = "status_DN_num"
            
            
            '---------------------------------------------------------------------------------------------------------
            'creat ca val loreal monthly
            '---------------------------------------------------------------------------------------------------------

            
            num_colums_TR = num_colums_TR + 1
            clm_m = str_PYper_LOR_VAL + f_mnth - 1
            
                If Cells(f_rw, clm_m) = 0 Then
                m_val = Empty
                Else
                m_val = Cells(f_rw, clm_m) / 1000
                End If
                
            ar_Data_TR(iii, num_colums_TR) = m_val
            If iii = 0 Then ar_nmHead_TR(num_colums_TR) = "CA_PY_M"
                
            num_colums_TR = num_colums_TR + 1
            If cdYear = cd_ActualYear - 1 And cdMonth = f_mnth Then
                ar_Data_TR(iii, num_colums_TR) = m_val
            End If
            If iii = 0 Then ar_nmHead_TR(num_colums_TR) = "CA_CNQ_PY_1st_order"
            
            

            num_colums_TR = num_colums_TR + 1
            clm_m = str_TYper_LOR_VAL + f_mnth - 1
                If Cells(f_rw, clm_m) = 0 Then
                m_val = Empty
                Else
                m_val = Cells(f_rw, clm_m) / 1000
                End If
            ar_Data_TR(iii, num_colums_TR) = m_val
            If iii = 0 Then ar_nmHead_TR(num_colums_TR) = "CA_TY_M"
                
            num_colums_TR = num_colums_TR + 1
            If cdYear = cd_ActualYear And cdMonth = f_mnth Then
                ar_Data_TR(iii, num_colums_TR) = m_val
            End If
            
            If iii = 0 Then ar_nmHead_TR(num_colums_TR) = "CA_CNQ_TY_1st_order"
                
                
            num_colums_TR = num_colums_TR + 1
                If cdYear = cd_ActualYear Or m_val = 0 Then
                m_val = Empty
                Else
                ar_Data_TR(iii, num_colums_TR) = m_val
                End If
            If iii = 0 Then ar_nmHead_TR(num_colums_TR) = "CPS_CA_TY_M"
                
                
                
                
                

            
        '---------------------------------------------------------------------------------------------------------
        'creat ca val loreal cumul
            '---------------------------------------------------------------------------------------------------------
                
            m_val = Empty
            m_val_ytd = Empty
            m_val_ty = Empty
            num_colums_TR = num_colums_TR + 1
            
            For f_m = 0 To 11
                clm_m = str_PYper_LOR_VAL + f_m
                m_val = (Cells(f_rw, clm_m) / 1000) + m_val
                
                If f_m = CInt(f_mnth) - 1 Then m_val_ytd = m_val
                If f_m = 11 Then m_val_ty = m_val
                
            Next f_m
                
                If m_val_ytd = 0 Then  ' del 0 value out
                ar_Data_TR(iii, num_colums_TR) = Empty
                Else
                ar_Data_TR(iii, num_colums_TR) = m_val_ytd
                End If
            If iii = 0 Then ar_nmHead_TR(num_colums_TR) = "CA_PY_YTD"
            

            
            num_colums_TR = num_colums_TR + 1
            If m_val_ty = 0 Then  ' del 0 value out
                ar_Data_TR(iii, num_colums_TR) = Empty
                Else
                ar_Data_TR(iii, num_colums_TR) = m_val_ty
                End If
            If iii = 0 Then ar_nmHead_TR(num_colums_TR) = "CA_TPY"
        
            
            num_colums_TR = num_colums_TR + 1
            m_val = Empty
                
            For f_m = 0 To 11
                If f_m < CInt(f_mnth) Then
                clm_m = str_TYper_LOR_VAL + f_m
                m_val = (Cells(f_rw, clm_m) / 1000) + m_val
                End If
            Next f_m
                    
            If m_val = 0 Then
            ar_Data_TR(iii, num_colums_TR) = Empty
            Else
            ar_Data_TR(iii, num_colums_TR) = m_val
            End If
            
            If iii = 0 Then ar_nmHead_TR(num_colums_TR) = "CA_TY_YTD"
            
            
            num_colums_TR = num_colums_TR + 1
            If cdYear = cd_ActualYear Or m_val = 0 Then
            m_val = Empty
            Else
            ar_Data_TR(iii, num_colums_TR) = m_val
            End If
            If iii = 0 Then ar_nmHead_TR(num_colums_TR) = "CPS_CA_TY_YTD"
                    
            
            iii = iii + 1
 
        Next f_rw


'---------------------------------------------------------------------------------------------------------
'_GET_CNTCT
'---------------------------------------------------------------------------------------------------------

'---------------------------------------------------------------------------------------------------------
'check Sheets and if not - add
'---------------------------------------------------------------------------------------------------------
sh_in_data_CNTC = "Cnt_SREP"

For Each Sh In ThisWorkbook.Worksheets
If Sh.Name = sh_in_data_CNTC Then
chek_name = 1
End If
Next Sh

If chek_name <> 1 Then
Set Sh = Worksheets.Add()
Sh.Name = sh_in_data_CNTC
End If

'---------------------------------------------------------------------------------------------------------

nm_sh_work_CNTCT = "Contacts"


Sheets(nm_sh_work_CNTCT).Select
ActiveSheet.AutoFilterMode = False
LastRow = ActiveSheet.UsedRange.Row + ActiveSheet.UsedRange.Rows.Count - 1  ' ????????? ??????
LastColum = ActiveSheet.UsedRange.Column - 1 + ActiveSheet.UsedRange.Columns.Count
'-----------------------------------------------------------
  
Dim dic_idmSeriesTR: Set dic_idmSeriesTR = CreateObject("Scripting.Dictionary")
dic_idmSeriesTR.RemoveAll
Dim dic_idmSeriesWSOT: Set dic_idmSeriesWSOT = CreateObject("Scripting.Dictionary")
dic_idmSeriesWSOT.RemoveAll
Dim dic_id_not_mSeriesTR: Set dic_id_not_mSeriesTR = CreateObject("Scripting.Dictionary")
dic_id_not_mSeriesTR.RemoveAll
    
    
   
    For f_rw = 2 To LastRow
    st_next = 0
        
            nm_SREP = Trim(Cells(f_rw, 3))
            nm_FLSM = Trim(Cells(f_rw, 6))
            nm_sector = Trim(Cells(f_rw, 1))
            nm_reg = Trim(Cells(f_rw, 11))
            nm_Mreg = Trim(Cells(f_rw, 10))
            nm_sector = Trim(Cells(f_rw, 1))
        
    num_colums_CNTC = 0

    nm_short_month = ar_nm_short_month(f_mnth - 1)
   
    ar_Data_CNTCT(xxx, num_colums_CNTC) = ar_nm_short_month(f_mnth - 1)
    ar_nmHead_CNTCT(num_colums_CNTC) = "months"
    
    num_colums_CNTC = num_colums_CNTC + 1
    ar_Data_CNTCT(xxx, num_colums_CNTC) = f_mnth
    ar_nmHead_CNTCT(num_colums_CNTC) = "num_months"
    
    
    num_colums_CNTC = num_colums_CNTC + 1
    ar_Data_CNTCT(xxx, num_colums_CNTC) = nm_brand
    ar_nmHead_CNTCT(num_colums_CNTC) = "brand"
     
    
    num_colums_CNTC = num_colums_CNTC + 1
    If Len(Cells(f_rw, 10)) < 1 Then
        nm_Mreg = Empty
        Else
    nm_Mreg = Cells(f_rw, 10)
    End If
    
        
    If Left(nm_Mreg, 2) = nm_brand Then
    nm_Mreg = Right(Cells(f_rw, 10), Len(Cells(f_rw, 10).Value) - 3)
    End If
    ar_Data_CNTCT(xxx, num_colums_CNTC) = nm_Mreg
    ar_nmHead_CNTCT(num_colums_CNTC) = "mreg"
    
    
        
    num_colums_CNTC = num_colums_CNTC + 1
    ar_Data_CNTCT(xxx, num_colums_CNTC) = nm_Mreg
    If xxx = 0 Then ar_nmHead_CNTCT(num_colums_CNTC) = "#mreg"
   
'Mreg LT-> EN + split Moscou GR
'---------------------------------------------------------------------------------------------------------
            
    num_colums_CNTC = num_colums_CNTC + 1
    textPos = 0
    
    If nm_Mreg = "Moscou GR" Then
    nm_sec = Cells(f_rw, 1)
    textPos = InStr(nm_sec, "MSK")
    textPos = InStr(nm_sec, "Moscou") + textPos
        If textPos > 0 Then
        nm_Mreg = "Moscou"
        Else
        nm_Mreg = "GR"

        End If
    End If
    
    For f_mr = 0 To UBound(ar_nmMregLT)
    If ar_nmMregLT(f_mr) = nm_Mreg Then
    nm_Mreg = ar_nmMregEN(f_mr)
    End If
    Next f_mr
    
    ar_Data_CNTCT(xxx, num_colums_CNTC) = nm_Mreg
    ar_nmHead_CNTCT(num_colums_CNTC) = "mreg_EXT"
 
 '---------------------------------------------------------------------------------------------------------
     
    num_colums_CNTC = num_colums_CNTC + 1
    ar_Data_CNTCT(xxx, num_colums_CNTC) = nm_reg
    ar_nmHead_CNTCT(num_colums_CNTC) = "REG"
    
    num_colums_CNTC = num_colums_CNTC + 1
    ar_Data_CNTCT(xxx, num_colums_CNTC) = nm_FLSM
    ar_nmHead_CNTCT(num_colums_CNTC) = "FLSM"
    
    
    num_colums_CNTC = num_colums_CNTC + 1
    ar_Data_CNTCT(xxx, num_colums_CNTC) = nm_short_month & " |" & nm_brand & " |" & nm_FLSM
    ar_nmHead_CNTCT(num_colums_CNTC) = "#FLSM"
    
    num_colums_CNTC = num_colums_CNTC + 1
    ar_Data_CNTCT(xxx, num_colums_CNTC) = nm_sector
    ar_nmHead_CNTCT(num_colums_CNTC) = "SEC"
    
    num_colums_CNTC = num_colums_CNTC + 1
    ar_Data_CNTCT(xxx, num_colums_CNTC) = Cells(f_rw, 2)
    ar_nmHead_CNTCT(num_colums_CNTC) = "cd_SEC"
    
    num_colums_CNTC = num_colums_CNTC + 1
    nm_SREP = Trim(Cells(f_rw, 3))
    ar_Data_CNTCT(xxx, num_colums_CNTC) = nm_SREP
    ar_nmHead_CNTCT(num_colums_CNTC) = "SREP"

    num_colums_CNTC = num_colums_CNTC + 1
    nm_SREP = Trim(Cells(f_rw, 3))
    ar_Data_CNTCT(xxx, num_colums_CNTC) = nm_short_month & nm_SREP & nm_brand
    ar_nmHead_CNTCT(num_colums_CNTC) = "#SREP"
    
    num_colums_CNTC = num_colums_CNTC + 1
    ar_Data_CNTCT(xxx, num_colums_CNTC) = Cells(f_rw, 4)
    ar_nmHead_CNTCT(num_colums_CNTC) = "staff"
    
    num_colums_CNTC = num_colums_CNTC + 1
    ar_Data_CNTCT(xxx, num_colums_CNTC) = Trim(Cells(f_rw, 8))
    ar_nmHead_CNTCT(num_colums_CNTC) = "cont_email"
        
    num_colums_CNTC = num_colums_CNTC + 1
    ar_Data_CNTCT(xxx, num_colums_CNTC) = Trim(Cells(f_rw, 7))
    ar_nmHead_CNTCT(num_colums_CNTC) = "cont_phone"
    
    num_colums_CNTC = num_colums_CNTC + 1
    ar_Data_CNTCT(xxx, num_colums_CNTC) = Trim(Cells(f_rw, 10))
    ar_nmHead_CNTCT(num_colums_CNTC) = "partner"
        
    num_colums_CNTC = num_colums_CNTC + 1
    ar_Data_CNTCT(xxx, num_colums_CNTC) = Trim(Cells(f_rw, 12))
    ar_nmHead_CNTCT(num_colums_CNTC) = "experience"
    
    
        num_colums_CNTC = num_colums_CNTC + 1
    
    testpos = Empty
    testpos = InStr(1, nm_SREP, "âàêàí", vbTextCompare)
    
    If testpos <> 0 Then
       st_vacancy = "vacancy"
       st_next = 1
        Else
        st_vacancy = "active"
     End If
    
    If nm_SREP = nm_FLSM Then st_vacancy = "FLSMasSREP"
    st_next = 1
    
    ar_Data_CNTCT(xxx, num_colums_CNTC) = st_vacancy
    ar_nmHead_CNTCT(num_colums_CNTC) = "vacancy_status"
    
    
            
    num_colums_CNTC = num_colums_CNTC + 1
    val_target_CA = Cells(f_rw, 14)
    If val_target_CA = 0 Then val_target_CA = Empty
    ar_Data_CNTCT(xxx, num_colums_CNTC) = val_target_CA
    ar_nmHead_CNTCT(num_colums_CNTC) = "target_CA"
    
            
    num_colums_CNTC = num_colums_CNTC + 1
    frsr_clm = num_colums_CNTC
    val_orders_SLN = Cells(f_rw, 15)
    If val_orders_SLN = 0 Then val_orders_SLN = Empty
    ar_Data_CNTCT(xxx, num_colums_CNTC) = val_orders_SLN
    ar_nmHead_CNTCT(num_colums_CNTC) = "orders_SLN"
    
            
    num_colums_CNTC = num_colums_CNTC + 1
    val_orders_phone = Cells(f_rw, 16)
    If val_orders_phone = 0 Then val_orders_phone = Empty
    ar_Data_CNTCT(xxx, num_colums_CNTC) = val_orders_phone
    ar_nmHead_CNTCT(num_colums_CNTC) = "orders_phone"
    
            
    num_colums_CNTC = num_colums_CNTC + 1
    val_visits2act = Cells(f_rw, 17)
    If val_visits2act = 0 Then val_visits2act = Empty
    ar_Data_CNTCT(xxx, num_colums_CNTC) = val_visits2act
    ar_nmHead_CNTCT(num_colums_CNTC) = "visits2act"
    
            
    num_colums_CNTC = num_colums_CNTC + 1
    val_visited_act = Cells(f_rw, 18)
    If val_visited_act = 0 Then val_visited_act = Empty
    ar_Data_CNTCT(xxx, num_colums_CNTC) = val_visited_act
    ar_nmHead_CNTCT(num_colums_CNTC) = "visited_act"
    
    num_colums_CNTC = num_colums_CNTC + 1
    val_visits2cnq = Cells(f_rw, 19)
    If val_visits2cnq = 0 Then val_visits2cnq = Empty
    ar_Data_CNTCT(xxx, num_colums_CNTC) = val_visits2cnq
    ar_nmHead_CNTCT(num_colums_CNTC) = "visits2cnq"
    
    num_colums_CNTC = num_colums_CNTC + 1
    end_clm = num_colums_CNTC
    val_visited_cnq = Cells(f_rw, 20)
    If val_visited_cnq = 0 Then val_visited_cnq = Empty
    ar_Data_CNTCT(xxx, num_colums_CNTC) = val_visited_cnq
    ar_nmHead_CNTCT(num_colums_CNTC) = "visited_cnq"

    sts_val_KPI = 0
    For f_x = frsr_clm To end_clm
        val_KPI = ar_Data_CNTCT(xxx, f_x)
        If Len(val_KPI) >= 1 Then
        sts_val_KPI = sts_val_KPI + 1
        End If
    Next f_x

    num_colums_CNTC = num_colums_CNTC + 1
    ar_Data_CNTCT(xxx, num_colums_CNTC) = sts_val_KPI
    ar_nmHead_CNTCT(num_colums_CNTC) = "sts_val_KPI"
        
 
       
If Len(nm_FLSM) > 0 Or InStr(1, nm_Mreg, "E-Commerce", vbTextCompare) Then xxx = xxx + 1
    
Next f_rw

'-----------------------------------------------------------------------
'Get_Coaching
'-----------------------------------------------------------------------


'---------------------------------------------------------------------------------------------------------
'check Sheets and if not - add
'---------------------------------------------------------------------------------------------------------
chek_name = 0
sh_in_data_COACH = "data_COACH"
nm_sh_work_COCH = "Coaching"

For Each Sh In ThisWorkbook.Worksheets
If Sh.Name = sh_in_data_COACH Then
chek_name = chek_name + 1
Else
chek_name = chek_name + 0
End If

Next Sh

If chek_name = 0 Then
Set Sh = Worksheets.Add()
Sh.Name = sh_in_data_COACH
End If


Sheets(nm_sh_work_COCH).Select
ActiveSheet.AutoFilterMode = False
LastRow = ActiveSheet.UsedRange.Row + ActiveSheet.UsedRange.Rows.Count - 1
LastColum = ActiveSheet.UsedRange.Column - 1 + ActiveSheet.UsedRange.Columns.Count


    
    
    For i = 5 To LastRow
    
    num_colums_COACH = 0
    nm_FLSM = Trim(Cells(i, 1))
    ar_Data_COACH(yyy, num_colums_COACH) = nm_FLSM
    ar_nmHead_COACH(num_colums_COACH) = "FLSM"
    
    num_colums_COACH = num_colums_COACH + 1
    ar_Data_COACH(yyy, num_colums_COACH) = Trim(Cells(i, 2))
    ar_nmHead_COACH(num_colums_COACH) = "SREP"
     
    num_colums_COACH = num_colums_COACH + 1
    nm_month = Cells(i, 3)

For f_m = 0 To 11
        If ar_nm_month_rus(f_m) = nm_month Then
        nm_month = ar_nm_short_month(f_m)
        Exit For
        End If
        Next f_m


    ar_Data_COACH(yyy, num_colums_COACH) = nm_month
    ar_nmHead_COACH(num_colums_COACH) = "month"
    
    num_colums_COACH = num_colums_COACH + 1
    ar_Data_COACH(yyy, num_colums_COACH) = Cells(i, 4)
    ar_nmHead_COACH(num_colums_COACH) = "day"

    num_colums_COACH = num_colums_COACH + 1
    ar_Data_COACH(yyy, num_colums_COACH) = Cells(i, 5)
    ar_nmHead_COACH(num_colums_COACH) = "visites_DN"
    
    num_colums_COACH = num_colums_COACH + 1
    ar_Data_COACH(yyy, num_colums_COACH) = Cells(i, 6)
    ar_nmHead_COACH(num_colums_COACH) = "of_them_potenc."
    
    num_colums_COACH = num_colums_COACH + 1
    ar_Data_COACH(yyy, num_colums_COACH) = Cells(i, 7)
    ar_nmHead_COACH(num_colums_COACH) = "#_orders"
    
    num_colums_COACH = num_colums_COACH + 1
    str_rating = num_colums_COACH
    ar_Data_COACH(yyy, num_colums_COACH) = Cells(i, 8)
    ar_nmHead_COACH(num_colums_COACH) = "preparation"
    
    num_colums_COACH = num_colums_COACH + 1
    ar_Data_COACH(yyy, num_colums_COACH) = Cells(i, 9)
    ar_nmHead_COACH(num_colums_COACH) = "contact"
    
    num_colums_COACH = num_colums_COACH + 1
    ar_Data_COACH(yyy, num_colums_COACH) = Cells(i, 10)
    ar_nmHead_COACH(num_colums_COACH) = "interest"
    num_colums_COACH = num_colums_COACH + 1
    ar_Data_COACH(yyy, num_colums_COACH) = Cells(i, 11)
    ar_nmHead_COACH(num_colums_COACH) = "desire"

    num_colums_COACH = num_colums_COACH + 1
    ar_Data_COACH(yyy, num_colums_COACH) = Cells(i, 12)
    ar_nmHead_COACH(num_colums_COACH) = "over.objections"
 
    num_colums_COACH = num_colums_COACH + 1
    ar_Data_COACH(yyy, num_colums_COACH) = Cells(i, 13)
    ar_nmHead_COACH(num_colums_COACH) = "gain.agreement"
    
    num_colums_COACH = num_colums_COACH + 1
    ar_Data_COACH(yyy, num_colums_COACH) = Cells(i, 14)
    ar_nmHead_COACH(num_colums_COACH) = "coActWbirmation"
    
    num_colums_COACH = num_colums_COACH + 1
    ar_Data_COACH(yyy, num_colums_COACH) = Cells(i, 15)
    ar_nmHead_COACH(num_colums_COACH) = "comments"

    num_colums_COACH = num_colums_COACH + 1
    ar_Data_COACH(yyy, num_colums_COACH) = Cells(i, 16)
    ar_nmHead_COACH(num_colums_COACH) = "brand"

    num_colums_COACH = num_colums_COACH + 1
    ar_Data_COACH(yyy, num_colums_COACH) = Cells(i, 17)
    ar_nmHead_COACH(num_colums_COACH) = "AVG"
    
    num_colums_COACH = num_colums_COACH + 1
    ar_Data_COACH(yyy, num_colums_COACH) = Cells(i, 18)
    ar_nmHead_COACH(num_colums_COACH) = "null"
        
    num_colums_COACH = num_colums_COACH + 1
    ar_Data_COACH(yyy, num_colums_COACH) = Cells(i, 19)
    ar_nmHead_COACH(num_colums_COACH) = "week"
            
    num_colums_COACH = num_colums_COACH + 1
    ar_Data_COACH(yyy, num_colums_COACH) = Cells(i, 20)
    ar_nmHead_COACH(num_colums_COACH) = "nm_day"
    
    num_colums_COACH = num_colums_COACH + 1
    
    If Len(Cells(i, 22)) > 0 Then
                
        nm_Mreg = Cells(i, 22)
        If Left(nm_Mreg, 2) = ar_brand(b) Then
        nm_Mreg = Right(Cells(i, 22), Len(Cells(i, 22)) - 3)
        End If
       
        Else
        nm_Mreg = 0
    End If
        
    
    ar_Data_COACH(yyy, num_colums_COACH) = nm_Mreg
    ar_nmHead_COACH(num_colums_COACH) = "mreg"
   
    If Len(nm_FLSM) > 0 Then yyy = yyy + 1
    
Next i

    
Workbooks(actTR).Activate
If ActiveWorkbook.Name <> ActWb Then
   ActiveWindow.Close
End If
Application.DisplayAlerts = False


Next f_brnd

Next f_mnth

   
Workbooks(ActWb).Activate
Sheets(ShInDataTR).Activate

ActiveSheet.UsedRange.Cells.ClearContents
end_POS = iii + 1
start_POS = 2

For t = 0 To num_colums_TR
Cells(1, t + 1) = ar_nmHead_TR(t)
Cells(1, t + 1).Select
Next t
 
ActiveSheet.Cells(start_POS, 1).Resize(end_POS + 1, num_colums_TR + 1) = ar_Data_TR()

Sheets(sh_in_data_CNTC).Activate

ActiveSheet.UsedRange.Cells.ClearContents
end_POS = xxx + 1
start_POS = 2

For t = 0 To num_colums_CNTC
Cells(1, t + 1) = ar_nmHead_CNTCT(t)
Cells(1, t + 1).Select
Next t
 
ActiveSheet.Cells(start_POS, 1).Resize(end_POS + 1, num_colums_CNTC + 1) = ar_Data_CNTCT()


Sheets(sh_in_data_COACH).Activate

ActiveSheet.UsedRange.Cells.ClearContents
end_POS = yyy + 1
start_POS = 2

For t = 0 To num_colums_COACH
Cells(1, t + 1) = ar_nmHead_COACH(t)
Cells(1, t + 1).Select
Next t
 
ActiveSheet.Cells(start_POS, 1).Resize(end_POS + 1, num_colums_COACH + 1) = ar_Data_COACH()


ActiveWorkbook.RefreshAll

'---------------------------------------------------------------------------------------------------------

With Application
.ScreenUpdating = True
.Calculation = xlCalculationAutomatic
.EnableEvents = True
.DisplayStatusBar = True
.DisplayAlerts = True
End With

End Sub

Function fn_mont_num&(in_data$)
    Dim result&
    Dim f_m&, num_month&
    ar_nm_month_qnc_rus = Array("ÿíâàðü", "ôåâðàëü", "ìàðò", "àïðåëü", "ìàé", "èþíü", "èþëü", "àâãóñò", "ñåíòáÿðü", "îêòÿáðü", "íîÿáðü", "äåêàáðü")
    result = 1
    For f_m = 0 To 11
        If ar_nm_month_qnc_rus(f_m) = in_data Then
            result = f_m + 1
            Exit For
        End If
    Next f_m
    fn_mont_num = result
End Function

