Attribute VB_Name = "getDataForKPI"
Sub get_data()
Dim nm_brand As String, patch As String, ShIn As String

Dim cd_ActualYear As Integer, cd_ActualMonth As Integer
Dim LastRow As Long
Dim f_mnth As Integer
Dim ShOutCntct As String

Dim clnts As clsClients, clnt As clsClientInfo
Set clnts = New clsClients

Set dic_People = CreateObject("Scripting.Dictionary"): dic_People.RemoveAll

nm_ActWb = ActiveWorkbook.Name
cd_ActualMonth = CInt(InputBox("Month"))
cd_ActualYear = CInt(InputBox("YearEnd"))

ar_brand = Array("LP")
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
i=0
For Each usr In users
    i = i + 1
    n = 0
    With usr
        n = n + 1: Cells(i + 1, n) = .PersonName:                    If i = 1 Then Cells(1, n) = "BrandName"
    End With
Next


myLib.CreateSh (ShOut)
myLib.sheetActivateCleer (ShOut)

i = 0
For Each clnt In clnts
    i = i + 1
    n = 0
    With clnt
        n = n + 1: Cells(i + 1, n) = .BrandName:                    If i = 1 Then Cells(1, n) = "BrandName"
        n = n + 1: Cells(i + 1, n) = .StatYear:                     If i = 1 Then Cells(1, n) = "StatYear"
        n = n + 1: Cells(i + 1, n) = .StatMonth:                    If i = 1 Then Cells(1, n) = "StatMonth"
        n = n + 1: Cells(i + 1, n) = .UniverseCode:                 If i = 1 Then Cells(1, n) = "UniverseCode"
        n = n + 1: Cells(i + 1, n) = .ExtMregName:                  If i = 1 Then Cells(1, n) = "ExtMregName"
        n = n + 1: Cells(i + 1, n) = .RegName:                      If i = 1 Then Cells(1, n) = "RegName"
        n = n + 1: Cells(i + 1, n) = .FlsmName:                     If i = 1 Then Cells(1, n) = "FlsmName"
        n = n + 1: Cells(i + 1, n) = .SecName:                      If i = 1 Then Cells(1, n) = "SecName"
        n = n + 1: Cells(i + 1, n) = .SrepName:                     If i = 1 Then Cells(1, n) = "SrepName"
        n = n + 1: Cells(i + 1, n) = .ClientName:                   If i = 1 Then Cells(1, n) = "ClientName"
        n = n + 1: Cells(i + 1, n) = .ChainName:                    If i = 1 Then Cells(1, n) = "ChainName"
        n = n + 1: Cells(i + 1, n) = .ClientTypeRus:                If i = 1 Then Cells(1, n) = "ClientTypeRus"
        n = n + 1: Cells(i + 1, n) = .ClubStatus:                   If i = 1 Then Cells(1, n) = "ClubStatus"
        n = n + 1: Cells(i + 1, n) = .EmotionStatus:                If i = 1 Then Cells(1, n) = "EmotionStatus"
        n = n + 1: Cells(i + 1, n) = .CnqFullDate:                  If i = 1 Then Cells(1, n) = "CnqFullDate"
        n = n + 1: Cells(i + 1, n) = .CnqYearDate:                  If i = 1 Then Cells(1, n) = "CnqYear"
        n = n + 1: Cells(i + 1, n) = .CnqYearGA:                    If i = 1 Then Cells(1, n) = "CnqGA"
        n = n + 1: Cells(i + 1, n) = .LtmAvgCaName:                 If i = 1 Then Cells(1, n) = "LtmAvgCaName"
        n = n + 1: Cells(i + 1, n) = .LtmFrqOrders:                 If i = 1 Then Cells(1, n) = "LtmFrqOrders"
        n = n + 1: Cells(i + 1, n) = .ClientEcadCode:               If i = 1 Then Cells(1, n) = "ClientEcadCode"
        n = n + 1: Cells(i + 1, n) = .MastersEducatedAllY:          If i = 1 Then Cells(1, n) = "MastersEducatedAllY"
        n = n + 1: Cells(i + 1, n) = .MastersEducatedPY:            If i = 1 Then Cells(1, n) = "MastersEducatedPY"
        n = n + 1: Cells(i + 1, n) = .MastersEducatedTY:            If i = 1 Then Cells(1, n) = "MastersEducatedTY"
        n = n + 1: Cells(i + 1, n) = .HairdressersNum:              If i = 1 Then Cells(1, n) = "HairdressersNum"
        n = n + 1: Cells(i + 1, n) = .HairdressersWorkPlace:        If i = 1 Then Cells(1, n) = "HairdressersWorkPlace"
        
        n = n + 1: Cells(i + 1, n) = IIF(.CA_PY_YTD.Item(12) <> 0, 1, Empty): If i = 1 Then Cells(1, n) = "DN_PY_T"
        n = n + 1: Cells(i + 1, n) = IIF(.CA_TY_YTD.Item(.StatMonth) <> 0, 1, Empty): If i = 1 Then Cells(1, n) = "DN_YTD"
        n = n + 1: Cells(i + 1, n) = IIF(.CA_TY_M.Item(.StatMonth) <> 0, 1, Empty): If i = 1 Then Cells(1, n) = "DN_TY_M"
        n = n + 1: Cells(i + 1, n) = IIF(.CA_TY_YTD.Item(.StatMonth) <> 0 And .CnqYearGA <> "CNQ_TY", 1, Empty): If i = 1 Then Cells(1, n) = "DN_TY_YTD_CPS"
        n = n + 1: Cells(i + 1, n) = IIF(.CA_TY_M.Item(.StatMonth) <> 0 And .CnqYearGA <> "CNQ_TY", 1, Empty): If i = 1 Then Cells(1, n) = "DN_TY_M_CPS"
        n = n + 1: Cells(i + 1, n) = myLib.getNumInThrousend(.CA_TY_M.Item(.StatMonth)): If i = 1 Then Cells(1, n) = "CA_TY_M"
        n = n + 1: Cells(i + 1, n) = myLib.getNumInThrousend(.CA_PY_M.Item(.StatMonth)): If i = 1 Then Cells(1, n) = "CA_PY_M"
        n = n + 1: Cells(i + 1, n) = myLib.getNumInThrousend(.CA_TY_YTD.Item(.StatMonth)): If i = 1 Then Cells(1, n) = "CA_TY_YTD"
        n = n + 1: Cells(i + 1, n) = myLib.getNumInThrousend(.CA_PY_YTD.Item(.StatMonth)): If i = 1 Then Cells(1, n) = "CA_PY_YTD"
        ' = n + 1: Cells(i + 1, n) = .isClosed: If i = 1 Then Cells(1, n) = "LostClientsLTM"
        n = n + 1: Cells(i + 1, n) = "": If i = 1 Then Cells(1, n) = "WinClientsLTM"
    End With

Next
myLib.VBA_End
End Sub
    



