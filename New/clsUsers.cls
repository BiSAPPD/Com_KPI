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

Private dic_People As Object

Private Sub Class_Initialize()
    Set dic_People = CreateObject("Scripting.Dictionary")
End Sub

Private Sub Class_Terminate()
    Set dic_People = Nothing
End Sub


Public Sub Add(key As Variant, obj As clsUser)
    dic_People.Add key, obj
End Sub

Public Sub FillFromSheet(ByRef wks As Worksheet, ByVal cStatYear As Integer, ByVal cStatMonth As Integer, ByVal cBrand As String)
    Const cFirstRow = 2
    Const cSrep = 3, cFLSM = 6
  
    Dim f_rw As Long, end_f As Long, obj As clsUser, n As Long, cumulTY As Double, cumulPY As Double, nm_Srep As String
    Dim keyUser As String, nm_FLSM As String, nm_vacancy_status As String
      
            
    end_f = myLib.GetLastRow
    For f_rw = cFirstRow To end_f
    Application.StatusBar = cBrand & " Year: " & cStatYear & " Month: " & cStatMonth & " Progress: " & f_rw & " of " & end_f & ": " & Format(f_rw / end_f, "0%"): DoEvents

    nm_Srep = Trim(wks.Cells(f_rw, cSrep))
    nm_FLSM = Trim(Cells(f_rw, cFLSM))
    nm_vacancy_status = myLib.GetSREP_type(nm_Srep, nm_FLSM)
    
    keyUser = cStatYear & cStatMonth & cBrand & nm_Srep
    
    If Not dic_People.Exists(keyUser) Then
        Set obj = New clsUser
        With obj
            .PersonName = nm_Srep
        End With
        Me.Add keyUser, obj
    End If
    Next
End Sub



  
