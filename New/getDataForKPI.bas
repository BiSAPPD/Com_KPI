Attribute VB_Name = "getDataForKPI"
Sub get_data()
Dim nm_brand As String, patch As String, ShIn As String

Dim cd_ActualYear As Integer, cd_ActualMonth As Integer
Dim LastRow As Long
Dim f_mnth As Integer
Dim ShOutCntct As String
Dim n As Long
Dim hh As Boolean
Dim i As Long
Dim diag As New ProgressDialogue
Dim tempArray()

Dim clnts As clsClients, clnt As clsClientInfo
Set clnts = New clsClients

Set dic_flsm = CreateObject("Scripting.Dictionary"): dic_flsm.RemoveAll

Dim users As clsUsers
Set users = New clsUsers


nm_ActWb = ActiveWorkbook.name
cd_ActualMonth = CInt(InputBox("Month"))
cd_ActualYear = CInt(InputBox("YearEnd"))

ar_brand = Array("KR")
myLib.VBA_Start


For f_mnth = 1 To cd_ActualMonth
    For f_brnd = 0 To UBound(ar_brand)
        nm_brand = ar_brand(f_brnd)
        
        ShOutCntct = "Contacts"
        patch = myLib.GetPatchHistTR(nm_brand, cd_ActualYear, cd_ActualYear, cd_ActualMonth, f_mnth)
        WbTR = myLib.OpenFile(patch, ShOutCntct)
        Workbooks(WbTR).Activate
        
        ShOutCntc = "Cnt_Persone"
        
        Sheets(ShOutCntct).Select
        users.FillFromSheet ActiveSheet, cd_ActualYear, f_mnth, nm_brand
        
        Workbooks(WbTR).Activate
        ShOut = nm_brand
        ShIn = "TR_KPI"
        Sheets(ShOut).Select
        clnts.FillFromSheet ActiveSheet, cd_ActualYear, f_mnth, nm_brand
        
        Workbooks(WbTR).Close
        Workbooks(nm_ActWb).Activate
    Next f_brnd
Next f_mnth

myLib.CreateSh (ShOutCntc)
myLib.sheetActivateCleer (ShOutCntc)
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
        If Not dic_flsm.Exists(.Shef & .cdDateStat) Then
            dic_flsm.Add .Shef & .cdDateStat, n
            tempArray(i, n) = .Shef
        End If
        myLib.letHead hh, n, "#FLSM"
        
    End With
    
        strt_i = i
        strt_n = n
    Dim rw_arr As Long
    
    
    For Each clnt In clnts.FilterByKeyRep(usr.brandStat & usr.cdDateStat & usr.PersonName)
       
    
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
            n = n + 1: tempArray(i, n) = .ClientName:                    myLib.letHead hh, n, "ClientName"
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
               
            
            n = n + 1: tempArray(i, n) = IIf(.CA_PY_YTD.Item(12) <> 0, 1, Empty):                                        myLib.letHead hh, n, "DN_PY_T"
            n = n + 1: tempArray(i, n) = IIf(.CA_TY_YTD.Item(month(.cdDateStat)) <> 0, 1, Empty):                                myLib.letHead hh, n, "DN_YTD"
            n = n + 1: tempArray(i, n) = IIf(.CA_TY_M.Item(month(.cdDateStat)) <> 0, 1, Empty):                                  myLib.letHead hh, n, "DN_TY_M"
            n = n + 1: tempArray(i, n) = IIf(.CA_TY_YTD.Item(month(.cdDateStat)) <> 0 And .CnqYearGA <> "CNQ_TY", 1, Empty):     myLib.letHead hh, n, "DN_TY_YTD_CPS"
            n = n + 1: tempArray(i, n) = IIf(.CA_TY_M.Item(month(.cdDateStat)) <> 0 And .CnqYearGA <> "CNQ_TY", 1, Empty):       myLib.letHead hh, n, "DN_TY_M_CPS"
            n = n + 1: tempArray(i, n) = myLib.getNumInThrousend(.CA_TY_M.Item(month(.cdDateStat))):                             myLib.letHead hh, n, "CA_TY_M"
            n = n + 1: tempArray(i, n) = myLib.getNumInThrousend(.CA_PY_M.Item(month(.cdDateStat))):                             myLib.letHead hh, n, "CA_PY_M"
            n = n + 1: tempArray(i, n) = myLib.getNumInThrousend(.CA_TY_YTD.Item(month(.cdDateStat))):                           myLib.letHead hh, n, "CA_TY_YTD"
            n = n + 1: tempArray(i, n) = myLib.getNumInThrousend(.CA_PY_YTD.Item(month(.cdDateStat))):                           myLib.letHead hh, n, "CA_PY_YTD"
            ' = n + 1: Cells(i , n) = .isClosed:                                                                     myLib.letHead hh, n, "LostClientsLTM"
            n = n + 1: tempArray(i, n) = "":                                                                             myLib.letHead hh, n, "WinClientsLTM"
        End With
        hh = False
        
    Next
    
    i = i + 1
Next
 diag.Hide
 ActiveSheet.Cells(2, 1).Resize(i, n) = tempArray()
myLib.VBA_End
End Sub
    



