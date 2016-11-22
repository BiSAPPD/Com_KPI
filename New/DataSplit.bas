

Declare PtrSafe Function SHCreateDirectoryEx Lib "shell32" Alias "SHCreateDirectoryExA" (ByVal hwnd As Long, ByVal pszPath As String, ByVal psa As Any) As Long
                                      
Sub CreateFolderWithSubfolders(ByVal PatchCreateFolder$)
 
   If Len(Dir(PatchCreateFolder$, vbDirectory)) = 0 Then
       SHCreateDirectoryEx Application.hwnd, PatchCreateFolder$, ByVal 0&
   End If
End Sub

Function Replace_symbols(ByVal txt As String) As String
    St$ = "~!@/\#$%^&*=|`-"""
    For i% = 1 To Len(St$)
        txt = Replace(txt, Mid(St$, i, 1), "_")
    Next
    Replace_symbols = txt
End Function

Sub KPI_DATA_SPLIT()

Dim arr1()
Dim arr2()
Dim num_row(2), ar_Split(), DynMas(), new_DynMas(), arr_nm_SREP(), ar_ShData()
Dim j As Long
Dim i As Long
Dim iColumns As Integer
Dim f_ar_a, f_b, f_c, f_d, f_ar_i, f_r, LastRow, colum As Integer
Dim lRangeDel As Range
Dim lRangeH As Range
Dim nm_Mreg, nm_findVAl, NF, val_Cell_This, val_Cell_Previous As String
Dim arData() , arTempData()

Dim NFW, lPath, nMreg, nmPatch, LastRowFM, nEmployer, nmCreatPatch2, nmCreatPatch, nMonth, nYear, LP, KR, RD, MX, SHEF, FN As String
Dim oWbk As Workbook

Dim dic_split: Set dic_split = CreateObject("Scripting.Dictionary")

clmMregKPI = 10
clmFLSMsKPI = 12

myLib.Vba_start

NF = ActiveWorkbook.Name
act_month = CInt(InputBox("Month"))
act_year = CInt(InputBox("Year"))

ShData = "OutKPI"

Sheets(ShData).Select
ActiveSheet.AutoFilterMode = False
LastRow = myLib.getLastRow
LastColmn = myLib.getLastColumn  

ReDim arData(1 To LastRow, 1 To LastColmn)
ReDim arTempData(1 To LastRow, 1 To LastColmn)
iii = 0
For xRow = 2 To LastRow
    For y_clm = 1 To LastColmn
        iii = iii + 1
        arData(iii, y_clm) = Cells(xRow, y_clm)
    Next y_clm
    
    For f_split = 1 to 2 
        Select Case f_split
            Case 1: uniqKey = Cells(i, clmMregKPI): clmSplit = clmMregKPI
            Case 2: uniqKey = Cells(i, clmFLSMsKPI): clmSplit = clmFLSMsKPI
        End Select

        If Not dic_split.Exists(uniqKey) And Not IsEmpty(uniqKey) Then
            dic_split.Add uniqKey, clmSplit
        End If   
    Next f_split
Next xRow


With dic_split
For varKey in .keys
    yyy = 0
    Sheets(ShData).Select           
                                                   
    For f_ar = 1 to LastRow 
        If varKey = arData(f_ar, .Item(varKey)) Then
            yyy = yyy + 1
            For f_clm_arr = 1 To LastColmn
                arTempData(yyy, f_clm_arr) = arData(f_ar, f_clm_arr)
            Next f_clm_arr
        End If
    Next f_ar

    Sheets(ShData).Select        
    ActiveSheet.Cells(2, 1).Resize(myLib.LastRow, myLib.LastColmn).Cells.ClearContents
    ActiveSheet.Cells(2, 1).Resize(yyy , LastColmn) = arTempData()
            
        Select Case .Item(varKey)
        Case 10
            nmPath = "\\Rucorpruwks0665\cards\" & nm_Mreg & "\" & "KPI\"
            nmFile = "KPI_" & "2016_" & act_month & "_" & nm_Mreg & ".xlsm"
        Case 12
            nmPath = "\\Rucorpruwks0665\cards\" & nm_Mreg & "\" & nm_FLSM & "\" & "KPI\"
            nmFile = "KPI_" & "2016_" & act_month & "_" & nm_FLSM & ".xlsm"
        End Select
   
        For Each PivotCache In ActiveWorkbook.PivotCaches
            PivotCache.Refresh
        Next
                Sheets("KPI").Select
        
        CreateFolderWithSubfolders nmPath
        ActiveWorkbook.SaveAs Filename:=nmPath & nmFile, _
        FileFormat:=xlOpenXMLWorkbookMacroEnabled, CreateBackup:=False
 
Next
End With
    myLib.Vba_end
End Sub




