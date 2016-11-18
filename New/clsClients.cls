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
  
    Dim f_i As Long, end_f As Long, obj As clsClientInfo, n As Long, cumulTY As Double, cumulPY As Double, sv As Integer
    
    Dim prg As Integer
       
    Dim diag As New ProgressDialogue
    
    end_f = myLib.GetLastRow
     
    diag.Configure "Upload data", "Now wasting your time...", cFirstRow, end_f
    diag.Show
    
    
    
    For f_i = cFirstRow To end_f
        sv = sv + 1
        If sv = 500 Then
            diag.SetValue f_i
            diag.SetStatus wks.name & " " & cStatYear & " " & cStatMonth & " from " & f_i & " to " & end_f
            sv = 0
        End If
        If diag.cancelIsPressed Then Exit For
    
  

        Set obj = New clsClientInfo
        With obj
            
            .cdDateStat = DateSerial(cStatYear, cStatMonth, 1)
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
            .key_srep = .BrandName & .cdDateStat & .SrepName
            
            
           
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
        
        End With
        
        Me.Add obj
    Next
    diag.Hide
    End Sub

    Public Function FilterByKeyRep(str As String) As clsClients
    Dim clnts As clsClients, clnt As clsClientInfo
    Set clnts = New clsClients

    For Each clnt In Me
        If clnt.key_srep = str Then clnts.Add clnt
    Next

    Set FilterByKeyRep = clnts

    End Function

