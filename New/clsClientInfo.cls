VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsClientInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public StatYear As Integer
Public StatMonth As Integer
Public BrandName As String
Public TypeBusiness As String
Public DatabaseClientNum As Integer
Public DatabaseClientAndBrandNum As String
Public UniverseCode As Variant
Public UniversCodeAndBrand As String
Public MregName As String
Public ExtMregName As String
Public RegName As String
Public FlsmName As String
Public SecName As String
Public SrepName As String
Public Client As String
Public ClientAddres As String
Public ClientName As String
Public ChainName As String
Public ChainNameTop As String
Public ChainCode As Variant
Public GeoCity As String
Public GeoReg As String
Public ClientType As String
Public ClientTypeRus As String
Public ClientTypeEng As String
Public ClientTypeEngShort As String
Public ClientTypeEngChain As String
Public ClientTypeGA As String
Public ClubStatus As String
Public EmotionStatus As String
Public CnqFullDate As Date
Public CnqYearDate As Date
Public CnqYearGA As String
Public CnqMonthNum As Integer
Public CnqMonthNameRus As String
Public CnqMonthNameEng As String
Public MagMinPrice As Long
Public MagMaxPrice As Long
Public HairDressersPlaces As Long
Public MagType As String
Public MagTypePrice As String
Public MagTypePlace As String
Public Mag As String
Public WorkStatusNum As Integer
Public WorkStatusName As String
Public HairDressers As Integer
Public LtmAvgCaVal As Variant
Public LtmAvgCaName As String
Public LtmFrqOrders As String
Public ClientEvVal As Double
Public ClientEvName As String
Public ClientEcadCode As Variant
Public MastersEducatedAllY As Integer
Public MastersEducatedPY As Variant
Public MastersEducatedTY As Variant
Public HairdressersNum As Variant
Public HairdressersWorkPlace As Integer
Public PartnerName As String
Public PartnerCode As Variant
Public isLost as Integrer

Public CA_TY_M As Collection
Public CA_PY_M As Collection
Public CA_TY_YTD As Collection
Public CA_PY_YTD As Collection

Public DN_TY_M As Collection
Public DN_PY_M As Collection
Public DN_TY_YTD As Collection
Public DN_PY_YTD As Collection

Private Sub Class_Initialize()
  Set CA_TY_M = New Collection
  Set CA_PY_M = New Collection
  Set CA_TY_YTD = New Collection
  Set CA_PY_YTD = New Collection

  Set DN_TY_M = New Collection
  Set DN_PY_M = New Collection
  Set DN_TY_YTD = New Collection
  Set DN_PY_YTD = New Collection
End Sub



