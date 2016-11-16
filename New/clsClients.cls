VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsClients"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Explicit

Private colClients As Collection

Private Sub Class_Initialize()
    Set colClients = New Collection
End Sub

Private Sub Class_Terminate()
    Set colClients = Nothing
End Sub

Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    Set NewEnum = colClients.[_NewEnum]
End Property

Public Sub Add(obj As clsClientInfo)
    colClients.Add obj
End Sub

Public Sub Remove(Index As Variant)
    colClients.Remove Index
End Sub

Public Property Get Item(Index As Variant) As clsClientInfo
Attribute Item.VB_UserMemId = 0
    Set Item = colClients.Item(Index)
End Property

Property Get Count() As Long
    Count = colClients.Count
End Property

Public Sub Clear()
    Set colClients = New Collection
End Sub


Public Sub FillFromSheet(ByRef wks As Worksheet, ByVal cStatYear As Integer, ByVal cStatMonth As Integer, ByVal cBrand As String)
    Const cFirstRow = 4
    Const clmTYLor = 93
    Const clmPYLor = 106
  
    Dim f_i As Long, end_f As Long, obj As clsClientInfo, n As Long, cumulTY As Double, cumulPY As Double

    
    
    end_f = myLib.GetLastRow
    For f_i = cFirstRow To myLib.GetLastRow
    Application.StatusBar = cBrand & " Year: " & cStatYear & " Month: " & cStatMonth & " Progress: " & f_i & " of " & end_f & ": " & Format(f_i / end_f, "0%"): DoEvents



        Set obj = New clsClientInfo
        With obj
            .StatYear = cStatYear
            .StatMonth = cStatMonth
            .BrandName = cBrand
            .TypeBusiness = myLib.GetTypeBusiness(cBrand)
            .DatabaseClientNum = wks.Cells(f_i, 1)
            .DatabaseClientAndBrandNum = cBrand & .DatabaseClientNum
            .UniverseCode = myLib.getUniversCode(cBrand, f_i, wks.Cells(f_i, 2))
            .UniversCodeAndBrand = cBrand & .UniverseCode
            .MregName = myLib.fixError(wks.Cells(f_i, 4))
            .RegName = myLib.fixError(wks.Cells(f_i, 5))
            .ExtMregName = myLib.GetMregExt(.MregName, .RegName)
            .FlsmName = myLib.fixError(wks.Cells(f_i, 165))
            .SecName = myLib.fixError(wks.Cells(f_i, 6))
            .SrepName = myLib.fixError(wks.Cells(f_i, 7))
            .Client = wks.Cells(f_i, 9)
            .ClientAddres = wks.Cells(f_i, 12)
            .GeoCity = wks.Cells(f_i, 11)
            .ClientName = myLib.GetClientName(.Client, .ClientAddres, .GeoCity)
            .ChainName = wks.Cells(f_i, 19)
            .ChainCode = myLib.GetNum2numNull(Cells(f_i, 20))
            .ChainNameTop = IIf(Left(.ChainCode, 2) = "93", .ChainName, Empty)
            .GeoReg = wks.Cells(f_i, 13)
            .ClientType = wks.Cells(f_i, 18)
            .ClientTypeRus = myLib.GetClntType(.ClientType, 1)
            .ClientTypeEng = myLib.GetClntType(.ClientType, 2)
            .ClientTypeEngShort = myLib.GetClntType(.ClientType, 3)
            .ClientTypeEngChain = myLib.GetClntType(.ClientType, 4)
            .ClientTypeGA = myLib.GetYearType(cStatYear, myLib.GetNum2num0(wks.Cells(f_i, 65)), 2)
            .ClubStatus = wks.Cells(f_i, 40)
            .EmotionStatus = Empty
            .CnqYearDate = myLib.GetYearType(cStatYear, myLib.GetNum2num0(wks.Cells(f_i, 65)), 1)
            .CnqYearGA = myLib.GetYearType(cStatYear, myLib.GetNum2num0(wks.Cells(f_i, 65)), 3)
            .CnqFullDate = DateSerial(.CnqYearDate, .CnqMonthNum, 1)
            .CnqMonthNameRus = wks.Cells(f_i, 64)
            .CnqMonthNum = myLib.GetMonthNumeric(.CnqMonthNameRus)
            .CnqMonthNameEng = myLib.GetNameMonthEN(.CnqMonthNum)
            .MagMinPrice = myLib.GetNum2num0(wks.Cells(f_i, 23))
            .MagMaxPrice = myLib.GetNum2num0(wks.Cells(f_i, 25))
            .HairDressersPlaces = myLib.GetNum2num0(wks.Cells(f_i, 27))
            .MagTypePrice = myLib.GetMag(.MagMinPrice, .MagMaxPrice, .HairDressersPlaces, .TypeBusiness)
            .MagTypePlace = myLib.GetMag(.MagMinPrice, .MagMaxPrice, .HairDressersPlaces, "place")
            .Mag = .MagTypePrice & .MagTypePlace
            .MagType = IIf(Len(.Mag) = 2, .Mag, Empty)
            .WorkStatusNum = wks.Cells(f_i, 8)
            .WorkStatusName = myLib.GetTypeDN(.WorkStatusNum)
            .ClientEcadCode = wks.Cells(f_i, 29)
            .HairDressers = myLib.GetNum2num0(wks.Cells(f_i, 28))
            .PartnerCode = wks.Cells(f_i, 173)
            .PartnerName = wks.Cells(f_i, 167)
            .LtmAvgCaVal = myLib.GetLTM(wks, f_i, cStatMonth, "avg_ca")
            .LtmAvgCaName = myLib.GetLTM(wks, f_i, cStatMonth, "type_avg_ca")
            .LtmFrqOrders = myLib.GetLTM(wks, f_i, cStatMonth, "frqOrders")
            .ClientEvVal = myLib.GetNum2num0(wks.Cells(f_i, 92))
            .ClientEvName = myLib.GetVectoreEV(.ClientEvVal)
            .MastersEducatedAllY = myLib.GetNum2numNull(wks.Cells(f_i, 30))
            .MastersEducatedPY = myLib.GetNum2numNull(wks.Cells(f_i, 31))
            .MastersEducatedTY = myLib.GetNum2numNull(wks.Cells(f_i, 32))
            .IsLost = IFF(.LtmAvgCaVal = 0 , 1 , Empty)
           
            cumulTY = Empty
            cumulPY = Empty
            For n = 1 To 12
                .CA_TY_M.Add IIf(cStatMonth >= n, CLng(wks.Cells(f_i, clmTYLor + n - 1)), Empty), CStr(n)
                .CA_PY_M.Add IIf(cStatMonth >= 1, CLng(wks.Cells(f_i, clmPYLor + n - 1)), Empty), CStr(n)
                    cumulTY = cumulTY + .CA_TY_M.Item(n)
                    cumulPY = cumulPY + .CA_PY_M.Item(n)
                .CA_TY_YTD.Add IIf(cStatMonth >= n, cumulTY, Empty), CStr(n)
                .CA_PY_YTD.Add IIf(cStatMonth >= 1, cumulPY, Empty), CStr(n)
            Next n
            
            For n = 1 to 12
                
            


        End With
        
        Me.Add obj
    Next
End Sub



