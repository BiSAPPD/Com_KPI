Sub get_TR()

    Dim ar_Brand as Variant
    Dim ShInData as String
    Dim cd_ActualMonth as Integer
    Dim cd_ActualYear as Integer
    Dim nm_ActWb as String
    Dim f_mnth as Integer
    Dim f_brnd as Integer
    Dim nm_brand as String
    Dim patch as String
    Dim ActTR as String
    Dim LastColum as Long
    Dim LastColum as Long


nm_ActWb = ActiveWorkbook.Name
cd_ActualMonth = CInt(InputBox("Month"))
cd_ActualYear = CInt(InputBox("YearEnd"))


ar_Brand = Array( "KR", "RD")
ShInData = TR

myLib.VBA_Start
myLib.CreateSh (ShInData)

For f_mnth = 1 To cd_ActualMonth
    For f_brnd = 0 To UBound(ar_Brand)
        nm_brand = ar_Brand(f_brnd)
            
        patch = myLib.patch_history_TR(nm_brand, cd_ActualYear, f_year, cd_ActualMonth, f_mnth)
        ActTR = myLib.OpenFile(patch, nm_ShOutData)
        LastRow = myLib.getLastRow
        LastColum = myLib.getLastColumn
        
        For f_rw = 2 To LastRow

End Sub