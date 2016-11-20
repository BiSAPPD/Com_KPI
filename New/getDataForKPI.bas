Attribute VB_Name = "getDataForKPI"
Sub get_data()
Dim nm_brand As String, patch As String, ShInTR As String

Dim cd_ActualYear As Integer, cd_ActualMonth As Integer
Dim LastRow As Long
Dim f_mnth As Integer
Dim ShOutTRCntct As String
Dim n As Long
Dim hh As Boolean
Dim i As Long
Dim diag As New ProgressDialogue
Dim tempArray()
Dim key_srep As String


Dim clnts As clsClients, clnt As clsClient
Set clnts = New clsClients

Set dic_flsm = CreateObject("Scripting.Dictionary"): dic_flsm.RemoveAll

Dim users As clsUsers
Set users = New clsUsers

Dim kpis As clsKPIs, kpi As clsKPI
Set kpis = New clsKPIs

Dim coachs As clsCoachDays, coach As clsCoachDay
Set coachs = New clsCoachDays

nm_ActWb = ActiveWorkbook.name
cd_ActualMonth = CInt(InputBox("Month"))
cd_ActualYear = CInt(InputBox("YearEnd"))

ar_brand = Array("KR")
myLib.VBA_Start


For f_mnth = 1 To cd_ActualMonth
    For f_brnd = 0 To UBound(ar_brand)
        nm_brand = ar_brand(f_brnd)
        
        ShOutTRCntct = "Contacts"
        patch = myLib.GetPatchHistTR(nm_brand, cd_ActualYear, cd_ActualYear, cd_ActualMonth, f_mnth)
        WbTR = myLib.OpenFile(patch, ShOutTRCntct)
        Workbooks(WbTR).Activate
        
        ShOutTRCntc = "Cnt_Persone"
        Sheets(ShOutTRCntct).Select
        users.FillFromSheet ActiveSheet, cd_ActualYear, f_mnth, nm_brand
        kpis.FillFromSheet ActiveSheet, cd_ActualYear, f_mnth, nm_brand
                
        ShOutTRCoach = "Coaching"
        Sheets(ShOutTRCoach).Select
        coachs.FillFromSheet ActiveSheet, cd_ActualYear, f_mnth, nm_brand
                
                
        Workbooks(WbTR).Activate
        ShOutTR = nm_brand
        ShInTR = "TR_KPI"
        Sheets(ShOutTR).Select
        clnts.FillFromSheet ActiveSheet, cd_ActualYear, f_mnth, nm_brand
        
        Workbooks(WbTR).Close
        Workbooks(nm_ActWb).Activate
    Next f_brnd
Next f_mnth

myLib.CreateSh (ShOutTRCntc)
myLib.sheetActivateCleer (ShOutTRCntc)
i = 1

'Dim clsUser As Variant 'Variant
Dim usr As Variant

diag.Configure "Wasting Time", "Now wasting your time...", 1, users.dic_People.Count
diag.Show


hh = True
ReDim tempArray(1 To clnts.Count + users.dic_People.Count, 1 To 100)
For Each usr In users.dic_People.Items
    
    st = st + 1
    diag.SetValue st
    diag.SetStatus "Now wasting your time... " & st
    If diag.cancelIsPressed Then Exit For

    n = 0
    
    With usr
        n = n + 1: tempArray(i, n) = i: myLib.letHead hh, n, "#": clm_first = n
        n = n + 1: tempArray(i, n) = .PersonName: myLib.letHead hh, n, "#srep":
        n = n + 1: tempArray(i, n) = .status: myLib.letHead hh, n, "status": clm_status = n
        n = n + 1: tempArray(i, n) = .cdDateStat: myLib.letHead hh, n, "datastat"
                
        n = n + 1
        If Not dic_flsm.Exists(.shef & .cdDateStat) Then
            dic_flsm.Add .shef & .cdDateStat, n
            tempArray(i, n) = .shef
        End If
        myLib.letHead hh, n, "#FLSM"
        
    End With
    
        strt_i = i
        strt_n = n
    Dim rw_arr As Long

    sum_OrdersSLN = 0
    sum_OrdersPhone = 0
    sum_Visits2Act = 0
    VisitedAct = 0
    status_kpi = 0

    key_srep = usr.BrandStat & usr.cdDateStat & usr.PersonName
    For Each clnt In clnts.FilterByKeyRep(key_srep)
        i = i + 1
        tempArray(i, clm_first) = i
        tempArray(i, clm_status) = usr.status
        With clnt
            n = strt_n

            n = n + 1: tempArray(i, n) = .BrandName:                     myLib.letHead hh, n, "BrandName": tempArray(strt_i, n) = .BrandName
            n = n + 1: tempArray(i, n) = .cdDateStat:                    myLib.letHead hh, n, "StatYear": tempArray(strt_i, n) = .cdDateStat
            n = n + 1: tempArray(i, n) = "":                             myLib.letHead hh, n, "StatMonth": tempArray(strt_i, n) = ""
            n = n + 1: tempArray(i, n) = .UniverseCode:                  myLib.letHead hh, n, "UniverseCode"
            n = n + 1: tempArray(i, n) = .ExtMregName:                   myLib.letHead hh, n, "ExtMregName": tempArray(strt_i, n) = .ExtMregName
            n = n + 1: tempArray(i, n) = .RegName:                       myLib.letHead hh, n, "RegName": tempArray(strt_i, n) = .RegName
            n = n + 1: tempArray(i, n) = .FlsmName:                      myLib.letHead hh, n, "FlsmName": tempArray(strt_i, n) = .FlsmName
            n = n + 1: tempArray(i, n) = .SecName:                       myLib.letHead hh, n, "SecName": tempArray(strt_i, n) = .SecName
            n = n + 1: tempArray(i, n) = .SrepName:                      myLib.letHead hh, n, "SrepName": tempArray(strt_i, n) = .SrepName
            n = n + 1: tempArray(i, n) = .ClientName:                    myLib.letHead hh, n, "ClientName": tempArray(strt_i, n) = "#KPIs_Data"
            n = n + 1: tempArray(i, n) = .ChainName:                     myLib.letHead hh, n, "ChainName"
            n = n + 1: tempArray(i, n) = .ClientTypeRus:                 myLib.letHead hh, n, "ClientTypeRus"
            n = n + 1: tempArray(i, n) = .ClubStatus:                    myLib.letHead hh, n, "ClubStatus"
            n = n + 1: tempArray(i, n) = .EmotionStatus:                 myLib.letHead hh, n, "EmotionStatus"
            n = n + 1: tempArray(i, n) = .CnqFullDate:                   myLib.letHead hh, n, "CnqFullDate"
            n = n + 1: tempArray(i, n) = .CnqYearDate:                   myLib.letHead hh, n, "CnqYear"
            n = n + 1: tempArray(i, n) = .CnqYearGA:                     myLib.letHead hh, n, "CnqGA"
            n = n + 1: tempArray(i, n) = .LtmAvgCaName:                  myLib.letHead hh, n, "LtmAvgCaName"
            n = n + 1: tempArray(i, n) = .LtmFrqOrders:                  myLib.letHead hh, n, "LtmFrqOrders"
            n = n + 1: tempArray(i, n) = .ClientEcadCode:                myLib.letHead hh, n, "ClientEcadCode"
            n = n + 1: tempArray(i, n) = .MastersEducatedAllY:           myLib.letHead hh, n, "MastersEducatedAllY"
            n = n + 1: tempArray(i, n) = .MastersEducatedPY:             myLib.letHead hh, n, "MastersEducatedPY"
            n = n + 1: tempArray(i, n) = .MastersEducatedTY:             myLib.letHead hh, n, "MastersEducatedTY"
            n = n + 1: tempArray(i, n) = .HairdressersNum:               myLib.letHead hh, n, "HairdressersNum"
            n = n + 1: tempArray(i, n) = .HairdressersWorkPlace:         myLib.letHead hh, n, "HairdressersWorkPlace"
               
            n = n + 1: tempArray(i, n) = IIF(.CA_PY_YTD.Item(12) <> 0, 1, Empty):                                               myLib.letHead hh, n, "DN_PY_T"
            n = n + 1: tempArray(i, n) = IIF(.CA_TY_YTD.Item(month(.cdDateStat)) <> 0, 1, Empty):                               myLib.letHead hh, n, "DN_YTD"
            n = n + 1: tempArray(i, n) = IIF(.CA_TY_M.Item(month(.cdDateStat)) <> 0, 1, Empty):                                 myLib.letHead hh, n, "DN_TY_M"
            n = n + 1: tempArray(i, n) = IIF(.CA_TY_YTD.Item(month(.cdDateStat)) <> 0 And .CnqYearGA <> "CNQ_TY", 1, Empty):    myLib.letHead hh, n, "DN_TY_YTD_CPS"
            n = n + 1: tempArray(i, n) = IIF(.CA_TY_M.Item(month(.cdDateStat)) <> 0 And .CnqYearGA <> "CNQ_TY", 1, Empty):      myLib.letHead hh, n, "DN_TY_M_CPS"
            n = n + 1: tempArray(i, n) = myLib.getNumInThrousend(.CA_TY_M.Item(month(.cdDateStat))):                            myLib.letHead hh, n, "CA_TY_M"
            n = n + 1: tempArray(i, n) = myLib.getNumInThrousend(.CA_PY_M.Item(month(.cdDateStat))):                            myLib.letHead hh, n, "CA_PY_M"
            n = n + 1: tempArray(i, n) = myLib.getNumInThrousend(.CA_TY_YTD.Item(month(.cdDateStat))):                          myLib.letHead hh, n, "CA_TY_YTD"
            n = n + 1: tempArray(i, n) = myLib.getNumInThrousend(.CA_PY_YTD.Item(month(.cdDateStat))):                          myLib.letHead hh, n, "CA_PY_YTD"
            ' = n + 1: Cells(i , n) = .isClosed:                                                                                myLib.letHead hh, n, "LostClientsLTM"
            n = n + 1: tempArray(i, n) = "":                                                                                    myLib.letHead hh, n, "WinClientsLTM"
            n = n + 1: tempArray(i, n) = .OrdersSLN:        myLib.letHead hh, n, "OrdersSLN": sum_OrdersSLN = sum_OrdersSLN + .OrdersSLN: clm_orfdersInSLN = n
            n = n + 1: tempArray(i, n) = .OrdersPhone:      myLib.letHead hh, n, "OrdersPhone": sum_OrdersPhone = sum_OrdersPhone + .OrdersPhone: clm_orderByPhone = n
            n = n + 1: tempArray(i, n) = .Visits2Act:       myLib.letHead hh, n, "Visits2Act": sum_Visits2Act = sum_Visits2Act + .Visits2Act: clm_visit2act = n

                numVisitedAct = IIF(.Visits2Act <> 0, 1, 0)
            n = n + 1: tempArray(i, n) = numVisitedAct: myLib.letHead hh, n, "VisitedAct": sum_VisitedAct = sum_VisitedAct + numVisitedAct: clm_numVisitedAct = n

  
        End With
        
    Next
    If kpis.dic_KPI.Exists(key_srep) Then
        With kpis.dic_KPI.Item(key_srep)
            status_kpi = status_kpi + IIF(sum_OrdersSLN + .OrdersSLN = 0, 0, 1)
            tempArray(strt_i, clm_orfdersInSLN) = IIF(sum_OrdersSLN = 0, .OrdersSLN, 0)
            
            status_kpi = status_kpi + IIF(sum_OrdersPhone + .OrdersPhone = 0, 0, 1)
            tempArray(strt_i, clm_orderByPhone) = IIF(sum_OrdersPhone = 0, .OrdersPhone, 0)
            
            status_kpi = status_kpi + IIF(sum_Visits2Act + .Visits2Act = 0, 0, 1)
            tempArray(strt_i, clm_visit2act) = IIF(sum_Visits2Act = 0, .Visits2Act, 0)
            
            status_kpi = status_kpi + IIF(sum_VisitedAct + .VisitedAct = 0, 0, 1)
            tempArray(strt_i, clm_numVisitedAct) = IIF(sum_VisitedAct = 0, .VisitedAct, 0)
                   
            status_kpi = status_kpi + IIF(.OrdersSLN = 0, 0, 1)
            n = n + 1: tempArray(strt_i, n) = .OrdersSLN:       myLib.letHead hh, n, "Visits2cnq"
            
            status_kpi = status_kpi + IIF(.OrdersPhone = 0, 0, 1)
            n = n + 1: tempArray(strt_i, n) = .OrdersPhone:       myLib.letHead hh, n, "VisitedCnq"
            
            n = n + 1: tempArray(strt_i, n) = .TargetCA:       myLib.letHead hh, n, "TargetCA"
            
            n = n + 1: tempArray(strt_i, n) = IIF(.WDays = 0, 20, .WDays):         myLib.letHead hh, n, "WDays"
            n = n + 1: tempArray(strt_i, n) = status_kpi:  myLib.letHead hh, n, "StatusDataKPI": clm_status_kpi = n
        End With
    End If
    For ii = strt_i + 1 To i
        tempArray(ii, clm_status_kpi) = status_kpi
    Next ii

    n = n + 1: tempArray(strt_i, n) = coachs.FilterByKeyRep(key_srep).Count: myLib.letHead hh, n, "CoachDays"

    hh = False
    i = i + 1
Next
 diag.Hide
 ActiveSheet.Cells(2, 1).Resize(i, n) = tempArray()
myLib.VBA_End
End Sub
    



