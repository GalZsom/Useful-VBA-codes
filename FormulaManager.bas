Attribute VB_Name = "FormulaManager"
Sub FormulaManager(ByVal StColNum As Long, Optional ByVal TestArray As Variant)
    Dim laast_row As Long
    Dim Lrow As Long
    Dim c As Range
    Dim lNumElements As Long
    Dim TargetSheet As Worksheet
    Dim ColumnLetter As String
    Dim StColNum As Long
    
    Set TargetSheet = ThisWorkbook.Sheets("Sheet1")
    laast_row = ThisWorkbook.Worksheets("Sheet1").Cells(Rows.Count, 1).End(xlUp).Row 'Store last row as num
    
    StColNum = 1 'Starting Row of Formulas
    Dim strFormulas(1 To 2) As Variant 'X = the number of formulas planned to be used
    
    With TargetSheet 'Az al�bbi sor(okat) toldja meg ezzel, jelenleg itt megadhat� a c�lt�bla amelybe a k�pletek m�soland�k
'        ====================================
'        T�mb felt�lt�se: strFormulas(sorsz�m) = "K�plet", amennyiben a k�pletben id�z�jel szerepel ott dupl�n kell azt kitenni
'        ====================================
        strFormulas(1) = ("=Sheet1!A1")
        strFormulas(2) = "=HA(A1="""";""URESCELLA"";HAHIBA(K�Z�P(A1;SZ�VEG.KERES(""("";A1) +1;SZ�VEG.KERES("")"";A1)-SZ�VEG.KERES(""("";A1)-1);""Sample""))" 'Foreign Formula searching for Sample

        
        lNumElements = UBound(strFormulas) - LBound(strFormulas) + 1 'Returns the number of elements
        ColumnLetter = Split(Cells(StColNum, StColNum + lNumElements - 1).Address, "$")(1) 'turns number into letter
        .Range("A" & StColNum & ":" & ColumnLetter & StColNum).FormulaLocal = strFormulas: '.Range, FormulaLocal is used because of foreign formulas
        
        If laast_row > 1 Then
            .Range("A" & StColNum & ":" & ColumnLetter & laast_row).FillDown ' Fills down the formula until the last row
            .Range("B2").BorderAround LineStyle:=Continous, _
            Weight:=xlThick
        End If
        
    End With
    With Range("A" & StColNum & ":" & ColumnLetter & laast_row).Borders
        .LineStyle = xlContinuous
        .Color = vbBlack
        .Weight = xlThin ' Styling
    End With
    
    
End Sub

