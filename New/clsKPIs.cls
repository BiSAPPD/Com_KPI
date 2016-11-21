VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsKPIs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Explicit

Public dic_KPI As Dictionary

Private Sub Class_Initialize()
    Set dic_KPI = New Dictionary
End Sub

Private Sub Class_Terminate()
    Set dic_KPI = Nothing
End Sub

Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    Set NewEnum = dic_KPI.[_NewEnum]
End Property

Public Sub Add(key As Variant, obj As clsKPI)
    dic_KPI.Add key, obj
End Sub

Public Sub FillFromSheet(ByRef wks As Worksheet, ByVal cStatYear As Integer, ByVal cStatMonth As Integer, ByVal cBrand As String)
    Const cFirstRow = 2
    Const cSrep = 3
    Const cWDays = 13
    Const cTargetCA = 14
    Const cOrdersSLN = 15
    Const cOrdersPhone = 16
    Const cVisits2Act = 17
    Const cVisitedAct = 18
    Const cVisits2cnq = 19
    Const cVisitedCnq = 20
  
    Dim f_rw As Long, end_f As Long, obj As clsKPI, n As Long, nmSrep As String
    Dim keyUser As String
            
    end_f = myLib.GetLastRow
    For f_rw = cFirstRow To end_f
            
        nmSrep = Trim(wks.Cells(f_rw, cSrep))
        keyUser = cBrand & DateSerial(cStatYear, cStatMonth, 1) & LCase(Trim(nmSrep))
        
        If Not dic_KPI.Exists(keyUser) And Len(nmSrep) > 0 Then
            Set obj = New clsKPI
            With obj
                .cdDateStat = DateSerial(cStatYear, cStatMonth, 1)
                .PersonName = nmSrep
                .BrandStat = cBrand
                .WDays = myLib.GetNum2num0(Cells(f_rw, cWDays))
                .TargetCA = myLib.GetNum2num0(Cells(f_rw, cTargetCA))
                .OrdersSLN = myLib.GetNum2num0(Cells(f_rw, cOrdersSLN))
                .OrdersPhone = myLib.GetNum2num0(Cells(f_rw, cOrdersPhone))
                .Visits2Act = myLib.GetNum2num0(Cells(f_rw, cVisits2Act))
                .VisitedAct = myLib.GetNum2num0(Cells(f_rw, cVisitedAct))
                .Visits2cnq = myLib.GetNum2num0(Cells(f_rw, cVisits2cnq))
                .VisitedCnq = myLib.GetNum2num0(Cells(f_rw, cVisitedCnq))
            End With
            Me.Add keyUser, obj
        ElseIf Len(nmSrep) > 0 Then
            With dic_KPI.Item(keyUser)
                .TargetCA = .TargetCA + myLib.GetNum2num0(Cells(f_rw, cTargetCA))
                .OrdersSLN = .OrdersSLN + myLib.GetNum2num0(Cells(f_rw, cOrdersSLN))
                .OrdersPhone = .OrdersPhone + myLib.GetNum2num0(Cells(f_rw, cOrdersPhone))
                .Visits2Act = .Visits2Act + myLib.GetNum2num0(Cells(f_rw, cVisits2Act))
                .VisitedAct = .VisitedAct + myLib.GetNum2num0(Cells(f_rw, cVisitedAct))
                .Visits2cnq = .Visits2cnq + myLib.GetNum2num0(Cells(f_rw, cVisits2cnq))
                .VisitedCnq = .VisitedCnq + myLib.GetNum2num0(Cells(f_rw, cVisitedCnq))
            End With
        End If
    Next
End Sub



  
