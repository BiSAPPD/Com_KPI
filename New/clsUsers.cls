VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsUsers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Explicit

Public dic_People As Dictionary

Private Sub Class_Initialize()
    Set dic_People = New Dictionary
End Sub

Private Sub Class_Terminate()
    Set dic_People = Nothing
End Sub

Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    Set NewEnum = dic_People.[_NewEnum]
End Property


Public Sub Add(key As Variant, obj As clsUser)
    dic_People.Add key, obj
End Sub

Public Sub FillFromSheet(ByRef wks As Worksheet, ByVal cStatYear As Integer, ByVal cStatMonth As Integer, ByVal cBrand As String)
    Const cFirstRow = 2
    Const cSrep = 3, cFlsm = 6
  
    Dim f_rw As Long, end_f As Long, obj As clsUser, n As Long, cumulTY As Double, cumulPY As Double, nm_Srep As String
    Dim keyUser As String, nm_FLSM As String, nm_vacancy_status As String
    Dim prg As Integer
    
            
    end_f = myLib.GetLastRow
    For f_rw = cFirstRow To end_f
    
        
        nm_Srep = Trim(wks.Cells(f_rw, cSrep))
        nm_FLSM = Trim(Cells(f_rw, cFlsm))
        nm_vacancy_status = myLib.GetSREP_type(nm_Srep, nm_FLSM)
        
        keyUser = cStatYear & cStatMonth & cBrand & nm_Srep
        
        If Not dic_People.Exists(keyUser) And Len(nm_Srep) > 0 Then
            Set obj = New clsUser
            With obj
                .cdDateStat = DateSerial(cStatYear, cStatMonth, 1)
                .PersonName = nm_Srep
                .shef = nm_FLSM
                .status = nm_vacancy_status
                .BrandStat = cBrand
                
                                
            End With
            Me.Add keyUser, obj
        End If
    Next
End Sub



  
