Attribute VB_Name = "Module1"
Sub ProduceIncentivization()

'''Match EE & SP

Dim xrow As Integer, xrow2 As Integer, SSNA As String, SSND As String, NCells As Integer

xrow = 2
xrow2 = 2

Worksheets("Census").Activate

NCells = Application.WorksheetFunction.CountA(Sheets("Census").Columns("A"))

Do While xrow <= NCells

    SSNA = Cells(xrow, 1)
    SSND = Cells(xrow2, 4)

    If SSNA = SSND Then
            
        Cells(xrow, 13) = "NA"
        xrow2 = 2
        
        Do While xrow2 <= NCells
        
            SSND = Cells(xrow2, 4)
                    
            If xrow2 <> xrow And Cells(xrow2, 1) = SSNA Then
            
                Cells(xrow, 13) = SSND
                xrow2 = NCells
            
            End If
            
            xrow2 = xrow2 + 1
        
        Loop
        
    Else
    
        Cells(xrow, 13) = SSNA
        
    End If
    
    xrow = xrow + 1
    xrow2 = xrow

Loop

'Name Correction

Cells(2, 11) = "=UPPER(E2)"
Range("'Census'!K2:'Census'!K2").AutoFill Destination:=Range("'Census'!K2:'Census'!K" & NCells)

Columns("K:K").Select
Application.CutCopyMode = False
Selection.Copy
Columns("E:E").Select
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
    :=False, Transpose:=False
        
Cells(2, 11) = "=UPPER(F2)"
Range("'Census'!K2:'Census'!K2").AutoFill Destination:=Range("'Census'!K2:'Census'!K" & NCells)

Columns("K:K").Select
Application.CutCopyMode = False
Selection.Copy
Columns("F:F").Select
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
    :=False, Transpose:=False
    
Columns("E:F").Select

Selection.Replace What:=".", Replacement:="", LookAt:=xlPart, SearchOrder _
    :=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False _
    , FormulaVersion:=xlReplaceFormula2

Selection.Replace What:=" ", Replacement:="", LookAt:=xlPart, _
    SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
    ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    
Selection.Replace What:="-", Replacement:="", LookAt:=xlPart, _
    SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
    ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    
Selection.Replace What:=",", Replacement:="", LookAt:=xlPart, _
    SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
    ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    
'Gender Correction

Cells(2, 11) = "=IFS(H2 = ""MALE"", ""MALE"", H2 = ""Male"", ""MALE"", H2 = ""M"", ""MALE"", 1=1, ""FEMALE"")"
Range("'Census'!K2:'Census'!K2").AutoFill Destination:=Range("'Census'!K2:'Census'!K" & NCells)

Columns("K:K").Select
Application.CutCopyMode = False
Selection.Copy
Columns("H:H").Select
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
    :=False, Transpose:=False
    
'Generate IDs

Cells(2, 11) = "=CONCAT(F2, E2, H2, RIGHT(D2, 4), YEAR(G2))"
Range("'Census'!K2:'Census'!K2").AutoFill Destination:=Range("'Census'!K2:'Census'!K" & NCells)

'''Compliance Sheet Repeat

Worksheets("Compliance").Activate

'Name Correction

NCells = Application.WorksheetFunction.CountA(Sheets("Compliance").Columns("A"))

Cells(2, 11) = "=UPPER(A2)"
Range("'Compliance'!K2:'Compliance'!K2").AutoFill Destination:=Range("'Compliance'!K2:'Compliance'!K" & NCells)

Columns("K:K").Select
Application.CutCopyMode = False
Selection.Copy
Columns("A:A").Select
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
    :=False, Transpose:=False
        
Cells(2, 11) = "=UPPER(B2)"
Range("'Compliance'!K2:'Compliance'!K2").AutoFill Destination:=Range("'Compliance'!K2:'Compliance'!K" & NCells)

Columns("K:K").Select
Application.CutCopyMode = False
Selection.Copy
Columns("B:B").Select
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
    :=False, Transpose:=False
    
Columns("A:B").Select

Selection.Replace What:=".", Replacement:="", LookAt:=xlPart, SearchOrder _
    :=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False _
    , FormulaVersion:=xlReplaceFormula2

Selection.Replace What:=" ", Replacement:="", LookAt:=xlPart, _
    SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
    ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    
Selection.Replace What:="-", Replacement:="", LookAt:=xlPart, _
    SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
    ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    
Selection.Replace What:=",", Replacement:="", LookAt:=xlPart, _
    SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
    ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    
'Gender Correction

Cells(2, 11) = "=IFS(G2 = ""MALE"", ""MALE"", G2 = ""Male"", ""MALE"", G2 = ""M"", ""MALE"", 1=1, ""FEMALE"")"
Range("'Compliance'!K2:'Compliance'!K2").AutoFill Destination:=Range("'Compliance'!K2:'Compliance'!K" & NCells)

Columns("K:K").Select
Application.CutCopyMode = False
Selection.Copy
Columns("G:G").Select
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
    :=False, Transpose:=False
    
'SSN Correction

Dim ShortLen As Boolean

Columns("C:C").NumberFormat = "@"

xrow = 2

Do While xrow <= NCells

    ShortLen = True

    Do While ShortLen = True
    
        If Len(Cells(xrow, 3)) < 4 Then
            Cells(xrow, 3) = "0" & Cells(xrow, 3)
        End If
        
        If Len(Cells(xrow, 3)) >= 4 Then
            ShortLen = False
        End If
        
    Loop

    xrow = xrow + 1

Loop

Columns("F:F").NumberFormat = "@"

xrow = 2

Do While xrow <= NCells

    ShortLen = True

    Do While ShortLen = True
    
        If Len(Cells(xrow, 6)) < 4 Then
            Cells(xrow, 6) = "0" & Cells(xrow, 6)
        End If
        
        If Len(Cells(xrow, 6)) >= 4 Then
            ShortLen = False
        End If
        
    Loop

    xrow = xrow + 1

Loop
    
'Generate IDs

Cells(2, 11) = "=CONCAT(A2, B2, G2, C2, YEAR(D2))"
Range("'Compliance'!K2:'Compliance'!K2").AutoFill Destination:=Range("'Compliance'!K2:'Compliance'!K" & NCells)

'Copy Columns for organization

Columns("K:K").Select
Selection.Copy
Columns("I:I").Select
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
    :=False, Transpose:=False
    
Columns("K:K").Select
Selection.ClearContents

Columns("H:H").Select
Selection.Copy
Columns("J:J").Select
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
    :=False, Transpose:=False
    
'''Compliance Matching

Worksheets("Census").Activate

NCells = Application.WorksheetFunction.CountA(Sheets("Census").Columns("A"))

'VLookup Compliance

Cells(2, 12) = "=VLOOKUP(K2, Compliance!I:J, 2, 0)"
Range("'Census'!L2:'Census'!L2").AutoFill Destination:=Range("'Census'!L2:'Census'!L" & NCells)

'VLookup SP IDs

Cells(2, 14) = "=VLOOKUP(M2, D:K, 8, 0)"
Range("'Census'!N2:'Census'!N2").AutoFill Destination:=Range("'Census'!N2:'Census'!N" & NCells)

'VLookup SP Compliance

Cells(2, 15) = "=VLOOKUP(N2, Compliance!I:J, 2, 0)"
Range("'Census'!O2:'Census'!O2").AutoFill Destination:=Range("'Census'!O2:'Census'!O" & NCells)

'Add Non-Participant as Compliance option

Columns("L:L").Select
Selection.SpecialCells(xlCellTypeFormulas, 16).Select
Selection.ClearContents
Selection.Replace What:="", Replacement:="NP", LookAt:=xlPart, _
    SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
    ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    
Columns("N:N").Select
Selection.SpecialCells(xlCellTypeFormulas, 16).Select
Selection.ClearContents
Selection.Replace What:="", Replacement:="NA", LookAt:=xlPart, _
    SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
    ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2

Columns("O:O").Select
Selection.SpecialCells(xlCellTypeFormulas, 16).Select
Selection.ClearContents

Cells(2, 16) = "=IFS(O2 = ""YES"", ""YES"", O2= ""NO"", ""NO"", N2= ""NA"", ""NA"", 1=1, ""NP"")"
Range("'Census'!P2:'Census'!P2").AutoFill Destination:=Range("'Census'!P2:'Census'!P" & NCells)

Columns("P:P").Select
Application.CutCopyMode = False
Selection.Copy
Columns("O:O").Select
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
    :=False, Transpose:=False
    
Columns("P:P").Select
Selection.ClearContents

Cells(2, 16) = "=IFS(AND(L2=""NP"", OR(O2=""NP"", O2=""NA"")), ""0"", OR(AND(L2=""NP"", O2=""NO""), AND(L2=""NO"", O2=""NP"")), ""1"", OR(AND(L2=""NO"", OR(O2=""NA"", O2=""NO"")), AND(L2=""YES"", O2=""NP""), AND(L2=""NP"", O2=""YES"")), ""2"", OR(AND(L2=""YES"",O2=""NO""), AND(L2=""NO"", O2=""YES"")), ""3"", OR(AND(L2=""YES"", OR(O2=""NA"", O2=""YES""))), ""4"")"
Range("'Census'!P2:'Census'!P2").AutoFill Destination:=Range("'Census'!P2:'Census'!P" & NCells)

Cells(1, 5) = "Last Name"
Cells(1, 6) = "First Name"
Cells(1, 8) = "Gender"
Cells(1, 11) = "ID (Macro Filled)"
Cells(1, 12) = "Compliance (Macro Filled)"
Cells(1, 13) = "SP SSN (Macro Filled)"
Cells(1, 14) = "SP ID (Macro Filled)"
Cells(1, 15) = "SP Compliance (Macro Filled)"
Cells(1, 16) = "Incentive Level (Macro Filled)"

Worksheets("Compliance").Activate

Cells(1, 1) = "First Name"
Cells(1, 2) = "Last Name"
Cells(1, 7) = "Gender"
Cells(1, 9) = "ID (Macro Filled)"
Cells(1, 10) = "Compliant Copy (Macro Filled)"

Worksheets("Census").Activate

'''Add incentive levels

Cells(1, 17) = "Incentive Rate (Macro Filled)"
Worksheets("Instructions").Activate

'Plan Storage

Dim Plan1 As String, Plan2 As String

Plan1 = Range("E14")
Plan2 = Range("E22")

'Incentive assignment

'By xrow, assign value for tier, level, and plan

Dim PlanNum As Integer, TierNum As Integer, LevelNum As Integer

Worksheets("Census").Activate

xrow = 2

Do While xrow <= NCells

    'Plan Assignment

    If Cells(xrow, 3) = Plan1 Then
    
        PlanNum = 16
        
    End If
    
    If Cells(xrow, 3) = Plan2 Then
    
        PlanNum = 24
    
    End If
    
    If Cells(xrow, 3) = "" Then
    
        PlanNum = 1000
        
    End If
    
    'Tier Assignment
    
    If Cells(xrow, 9) = Sheets("Instructions").Cells(16, 4) Then

        TierNum = 0
        
    End If
    
    If Cells(xrow, 9) = Sheets("Instructions").Cells(17, 4) Then

        TierNum = 1
        
    End If
    
    If Cells(xrow, 9) = Sheets("Instructions").Cells(18, 4) Then

        TierNum = 2
        
    End If
    
    If Cells(xrow, 9) = Sheets("Instructions").Cells(19, 4) Then

        TierNum = 3
        
    End If
    
    If Cells(xrow, 9) = Sheets("Instructions").Cells(20, 4) Then

        TierNum = 4
        
    End If
        
    'Level Assignment
    
    LevelNum = Cells(xrow, 16) + 5

    'Incentive Assignment
    
    Sheets("Census").Cells(xrow, 17) = Sheets("Instructions").Cells(PlanNum + TierNum, LevelNum)
    
    xrow = xrow + 1

Loop




End Sub
