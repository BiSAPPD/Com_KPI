VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsCoachDays"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Explicit

Private colCoaching As Collection

Private Sub Class_Initialize()
    Set colCoaching = New Collection
End Sub

Private Sub Class_Terminate()
    Set colCoaching = Nothing
End Sub

Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    Set NewEnum = colCoaching.[_NewEnum]
End Property

Public Sub Add(obj As clsCoachDay)
 colCoaching.Add obj
End Sub

Public Sub Remove(Index As Variant)
 colCoaching.Remove Index
End Sub

Public Property Get Item(Index As Variant) As clsCoachDay
Attribute Item.VB_UserMemId = 0
    Set Item = colCoaching.Item(Index)
End Property

Property Get Count() As Long
    Count = colCoaching.Count
End Property

Public Sub Clear()
    Set colCoaching = New Collection
End Sub

Public Sub FillFromSheet(ByRef wks As Worksheet, ByVal cStatYear As Integer, ByVal cStatMonth As Integer, ByVal cBrand As String)

    Const cFirstRow = 5
    Const cSrep = 2
    Const cFlsm = 1
    Const cMonth_coach = 3
    Const cDay_coch = 4
    Const cVisits2Act = 5
    Const cVisits2cnq = 6
    Const cOrdersInSLN = 7
    Const cStep1 = 8
    Const cStep2 = 9
    Const cStep3 = 10
    Const cStep4 = 11
    Const cStep5 = 12
    Const cStep6 = 13
    Const cStep7 = 14
    Const cComment = 15

    Dim f_rw As Long, end_f As Long, obj As clsCoachDay, n As Long, nmSrep As String
    Dim keyUser As String
            
    end_f = myLib.GetLastRow
    For f_rw = cFirstRow To end_f
           
        nmSrep = Trim(wks.Cells(f_rw, cSrep))
        keyUser = cBrand & DateSerial(cStatYear, cStatMonth, 1) & LCase(Trim(nmSrep))
    
        Set obj = New clsCoachDay
        With obj
            .key_srep = keyUser
            .cdDateStat = DateSerial(cStatYear, cStatMonth, 1)
            .person = nmSrep
            .BrandStat = cBrand
            .shef = wks.Cells(f_rw, cFlsm)
            .cdDateCoaching = DateSerial(cStatYear, myLib.GetMonthNumeric(wks.Cells(f_rw, cMonth_coach)), wks.Cells(f_rw, cDay_coch))
            .visitedTotalCLNs = myLib.GetNum2num0(wks.Cells(f_rw, cVisits2Act))
            .visitedCnqCLNs = myLib.GetNum2num0(wks.Cells(f_rw, cVisits2cnq))
            .orders = myLib.GetNum2num0(wks.Cells(f_rw, cOrdersInSLN))
            .step1 = myLib.GetNum2num0(wks.Cells(f_rw, cStep1))
            .step2 = myLib.GetNum2num0(wks.Cells(f_rw, cStep2))
            .step3 = myLib.GetNum2num0(wks.Cells(f_rw, cStep3))
            .step4 = myLib.GetNum2num0(wks.Cells(f_rw, cStep4))
            .step5 = myLib.GetNum2num0(wks.Cells(f_rw, cStep5))
            .step6 = myLib.GetNum2num0(wks.Cells(f_rw, cStep6))
            .step7 = myLib.GetNum2num0(wks.Cells(f_rw, cStep7))
            .comment = wks.Cells(f_rw, cComment)
        End With

        Me.Add obj
    Next
End Sub

Public Function FilterByKeyRep(str As String) As clsCoachDays
Dim coachs As clsCoachDays, coach As clsCoachDay
Set coachs = New clsCoachDays

For Each coach In Me
    If coach.key_srep = str Then coachs.Add coach
Next

Set FilterByKeyRep = coachs

End Function

  
