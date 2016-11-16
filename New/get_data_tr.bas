Attribute VB_Name = "get_data_tr"
Sub get_data()
Dim nm_brand As String, patch As String, ShIn As String

Dim ThisYear As Integer, cd_ActualMonth As Integer
Dim LastRow As Long
Dim f_mnth As Integer


myLib.VBA_Start

nm_ActWb = ActiveWorkbook.Name
cd_ActualMonth = CInt(InputBox("Month"))
ThisYear = CInt(InputBox("YearEnd"))

ar_brand = Array("KR", "RD")
myLib.VBA_Start

Dim clnts As clsClients, clnt As clsClientInfo
Set clnts = New clsClients

For f_mnth = 1 To cd_ActualMonth
    For f_brnd = 0 To UBound(ar_brand)
        nm_brand = ar_brand(f_brnd)
        ShIn = nm_brand
        ShOut = "TR"
        patch = myLib.GetPatchHistTR(nm_brand, ThisYear, ThisYear, cd_ActualMonth, f_mnth)
        WbTR = myLib.OpenFile(patch, ShIn)
        Workbooks(WbTR).Activate
        Sheets(ShIn).Select
        clnts.FillFromSheet ActiveSheet, ThisYear, f_mnth, nm_brand
        
        Workbooks(WbTR).Close
        Workbooks(nm_ActWb).Activate
    Next f_brnd
Next f_mnth

myLib.CreateSh (ShOut)
myLib.sheetActivateCleer (ShOut)

i = 0
For Each clnt In clnts
    i = i + 1
    n = 0
    With clnt
        n = n + 1: Cells(i + 1, n) = .BrandName:                    If i = 1 Then Cells(1, n) = "BrandName"
        n = n + 1: Cells(i + 1, n) = .DatabaseClientNum:            If i = 1 Then Cells(1, n) = "DatabaseClientNum"
        n = n + 1: Cells(i + 1, n) = .StatYear:                     If i = 1 Then Cells(1, n) = "StatYear"
        n = n + 1: Cells(i + 1, n) = .StatMonth:                    If i = 1 Then Cells(1, n) = "StatMonth"
        n = n + 1: Cells(i + 1, n) = .BrandName:                    If i = 1 Then Cells(1, n) = "BrandName"
        n = n + 1: Cells(i + 1, n) = .TypeBusiness:                 If i = 1 Then Cells(1, n) = "TypeBusiness"
        n = n + 1: Cells(i + 1, n) = .DatabaseClientNum:            If i = 1 Then Cells(1, n) = "DatabaseClientNum"
        n = n + 1: Cells(i + 1, n) = .DatabaseClientAndBrandNum:    If i = 1 Then Cells(1, n) = "DatabaseClientAndBrandNum"
        n = n + 1: Cells(i + 1, n) = .UniverseCode:                 If i = 1 Then Cells(1, n) = "UniverseCode"
        n = n + 1: Cells(i + 1, n) = .UniversCodeAndBrand:          If i = 1 Then Cells(1, n) = "UniversCodeAndBrand"
        n = n + 1: Cells(i + 1, n) = .MregName:                     If i = 1 Then Cells(1, n) = "MregName"
        n = n + 1: Cells(i + 1, n) = .ExtMregName:                  If i = 1 Then Cells(1, n) = "ExtMregName"
        n = n + 1: Cells(i + 1, n) = .RegName:                      If i = 1 Then Cells(1, n) = "RegName"
        n = n + 1: Cells(i + 1, n) = .FlsmName:                     If i = 1 Then Cells(1, n) = "FlsmName"
        n = n + 1: Cells(i + 1, n) = .SecName:                      If i = 1 Then Cells(1, n) = "SecName"
        n = n + 1: Cells(i + 1, n) = .SrepName:                     If i = 1 Then Cells(1, n) = "SrepName"
        n = n + 1: Cells(i + 1, n) = .ClientName:                   If i = 1 Then Cells(1, n) = "ClientName"
        n = n + 1: Cells(i + 1, n) = .ChainName:                    If i = 1 Then Cells(1, n) = "ChainName"
        n = n + 1: Cells(i + 1, n) = .ChainCode:                    If i = 1 Then Cells(1, n) = "ChainCode"
        n = n + 1: Cells(i + 1, n) = .GeoCity:                      If i = 1 Then Cells(1, n) = "GeoCity"
        n = n + 1: Cells(i + 1, n) = .GeoReg:                       If i = 1 Then Cells(1, n) = "GeoReg"
        n = n + 1: Cells(i + 1, n) = .ClientTypeRus:                If i = 1 Then Cells(1, n) = "ClientTypeRus"
        n = n + 1: Cells(i + 1, n) = .ClientTypeEng:                If i = 1 Then Cells(1, n) = "ClientTypeEng"
        n = n + 1: Cells(i + 1, n) = .ClientTypeEngShort:           If i = 1 Then Cells(1, n) = "ClientTypeEngChort"
        n = n + 1: Cells(i + 1, n) = .ClientTypeEngChain:           If i = 1 Then Cells(1, n) = "ClientTypeEngChain"
        n = n + 1: Cells(i + 1, n) = .ClubStatus:                   If i = 1 Then Cells(1, n) = "ClubStatus"
        n = n + 1: Cells(i + 1, n) = .EmotionStatus:                If i = 1 Then Cells(1, n) = "EmotionStatus"
        n = n + 1: Cells(i + 1, n) = .CnqFullDate:                  If i = 1 Then Cells(1, n) = "CnqFullDate"
        n = n + 1: Cells(i + 1, n) = .CnqYearDate:                  If i = 1 Then Cells(1, n) = "CnqYearGA"
        n = n + 1: Cells(i + 1, n) = .CnqYearGA:                    If i = 1 Then Cells(1, n) = "CnqYearGA"
        n = n + 1: Cells(i + 1, n) = .CnqMonthNum:                  If i = 1 Then Cells(1, n) = "CnqMonthNum"
        n = n + 1: Cells(i + 1, n) = .CnqMonthNameRus:              If i = 1 Then Cells(1, n) = "CnqMonthNameRus"
        n = n + 1: Cells(i + 1, n) = .CnqMonthNameEng:              If i = 1 Then Cells(1, n) = "CnqMonthNameEng"
        n = n + 1: Cells(i + 1, n) = .MagType:                      If i = 1 Then Cells(1, n) = "MagType"
        n = n + 1: Cells(i + 1, n) = .MagTypePrice:                 If i = 1 Then Cells(1, n) = "MagTypePrice"
        n = n + 1: Cells(i + 1, n) = .MagTypePlace:                 If i = 1 Then Cells(1, n) = "MagTypePlace"
        n = n + 1: Cells(i + 1, n) = .WorkStatusNum:                If i = 1 Then Cells(1, n) = "WorkStatusNum"
        n = n + 1: Cells(i + 1, n) = .WorkStatusName:               If i = 1 Then Cells(1, n) = "WorkStatusName"
        n = n + 1: Cells(i + 1, n) = .LtmAvgCaVal:                  If i = 1 Then Cells(1, n) = "LtmAvgCaVal"
        n = n + 1: Cells(i + 1, n) = .LtmAvgCaName:                 If i = 1 Then Cells(1, n) = "LtmAvgCaName"
        n = n + 1: Cells(i + 1, n) = .LtmFrqOrders:                 If i = 1 Then Cells(1, n) = "LtmFrqOrders"
        n = n + 1: Cells(i + 1, n) = .ClientEvVal:                  If i = 1 Then Cells(1, n) = "ClientEvVal"
        n = n + 1: Cells(i + 1, n) = .ClientEvName:                 If i = 1 Then Cells(1, n) = "ClientEvName"
        n = n + 1: Cells(i + 1, n) = .ClientEcadCode:               If i = 1 Then Cells(1, n) = "ClientEcadCode"
        n = n + 1: Cells(i + 1, n) = .MastersEducatedAllY:          If i = 1 Then Cells(1, n) = "MastersEducatedAllY"
        n = n + 1: Cells(i + 1, n) = .MastersEducatedPY:            If i = 1 Then Cells(1, n) = "MastersEducatedPY"
        n = n + 1: Cells(i + 1, n) = .MastersEducatedTY:            If i = 1 Then Cells(1, n) = "MastersEducatedTY"
        n = n + 1: Cells(i + 1, n) = .HairdressersNum:              If i = 1 Then Cells(1, n) = "HairdressersNum"
        n = n + 1: Cells(i + 1, n) = .HairdressersWorkPlace:        If i = 1 Then Cells(1, n) = "HairdressersWorkPlace"
        n = n + 1: Cells(i + 1, n) = .PartnerName:                  If i = 1 Then Cells(1, n) = "PartnerName"
        n = n + 1: Cells(i + 1, n) = .PartnerCode:                  If i = 1 Then Cells(1, n) = "PartnerCode"
    
        For f_m = 1 To 12
            n1 = n + 1: Cells(i + 1, n1) = myLib.getNumInThrousend(.CA_TY_M.Item(f_m)): If i = 1 Then Cells(1, n1) = "CA_TY_M" & f_m
            n2 = n + 11 + 1: Cells(i + 1, n2) = myLib.getNumInThrousend(.CA_PY_M.Item(f_m)): If i = 1 Then Cells(1, n2) = "CA_PY_M" & f_m
            n3 = n + 23 + 1: Cells(i + 1, n3) = myLib.getNumInThrousend(.CA_TY_YTD.Item(f_m)): If i = 1 Then Cells(1, n3) = "CA_TY_YTD" & f_m
            n4 = n + 35 + 1: Cells(i + 1, n4) = myLib.getNumInThrousend(.CA_PY_YTD.Item(f_m)): If i = 1 Then Cells(1, n4) = "CA_PY_YTD" & f_m
        Next f_m

        
        

    End With

Next
myLib.VBA_End
End Sub
    



