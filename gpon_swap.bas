' date: 2026_05M_28 11:11:37
' ------------------------------------------------------------

Option Explicit

Sub Macro_01_FTTx_022_sort_i_OLT_PORT()
    ' Call Function:
    No_01_FTTx_022_SortBy_ormar_polica_SpliterSlot_PortName
    No_02_FTTx_022_OLT_PORT_Q
End Sub

Sub Macro_02_FTTx_019_sort_i_OLT_PORT_i_Lookup_4Cols()
    ' Call Function:
    No_03_FTTx_019_SortBy_OLT_Then_FirstAndLastNumberIn_PortName
    No_04_FTTx_019_OLT_PORT_Q
    No_05_FTTx_019_Lookup_4Cols_to_FTTx_022
End Sub

Sub Macro_03_GPON_Rep_001_sort_i_OLT_PORT_i_KOR_i_ER_ONT)
    ' Call Function:
    No_06_GPON_Report_001_SortBy_OLT_Then_1And2_NumberIn_Port
    No_07_GPON_Report_001_OLT_PORT_V
    No_08_GPON_Report_001_KOR_i_ER_ONT
    No_09_GPON_Report_001_Lookup_2Cols_to_FTTx_022
End Sub


Function No_01_FTTx_022_SortBy_ormar_polica_SpliterSlot_PortName()
    ' FTTx_022

    ' D-polica (D): SR_..._04_D_2
    ' -->
    ' 1. Ormar: 04
    ' 2. Polica: 2

    ' 3. Spliter - slot (F): 02
    ' 4. Port Name (K): 001

    Dim ws As Worksheet

    Dim hasHeader As Boolean
    Dim lastRow As Long, lastCol As Long
    Dim startRow As Long
    Dim helperFirstCol As Long, helperSecondCol As Long, helperLastCol As Long
    Dim r As Long
    Dim s As String
    Dim parts As Variant
    Dim firstNum As Long, secondNum As Long, lastNum As Long
    Dim lastCell As Range, lastColCell As Range

    ' Call Function:
    No_00_inser_No

    '===SETTINGS===
    ' Set ws=ActiveSheet          ' Or:
    Set ws = ThisWorkbook.Worksheets("FTTx_022")
    hasHeader = True                      ' Set to False if there is NO header row
    '=================

    ' Find last used row & column robustly
    Set lastCell = ws.Cells.Find(What:="*", LookIn:=xlFormulas, LookAt:=xlPart, _
                                 SearchOrder:=xlByRows, SearchDirection:=xlPrevious)
    ' If lastCell Is Nothing Then Exit Sub
    If lastCell Is Nothing Then Exit Function
    lastRow = lastCell.Row

    Set lastColCell = ws.Cells.Find(What:="*", LookIn:=xlFormulas, LookAt:=xlPart, _
                                    SearchOrder:=xlByColumns, SearchDirection:=xlPrevious)
    lastCol = lastColCell.Column
    If lastCol < 4 Then lastCol = 4   ' Ensure we at least include column D

    startRow = IIf(hasHeader, 2, 1)
    ' If lastRow < startRow Then Exit Sub
    If lastRow < startRow Then Exit Function

    ' Add two helper columns at the far right
    helperFirstCol = lastCol + 1
    helperSecondCol = lastCol + 2

    ' helperLastCol = lastCol + 3
    helperLastCol = helperSecondCol

    ws.Cells(1, helperFirstCol).Value = "__FirstNum_D__"
    ws.Cells(1, helperSecondCol).Value = "__SecondNum_D__"
    ' ws.Cells(1, helperLastCol).Value = "__LastNum_D__"

    ' Fill helpers:
    '-__FirstNum_D__: number after the FIRST dash in D(e.g., "ON-1-0-10"-> 1)
    '-__SecondNum_D__: number after the SECOND dash in D(e.g., "ON-1-0-10"-> 0)
    '-__LastNum_D__:  number after the LAST dash in D(e.g., "ON-1-0-10"-> 10)

    ' D-polica (D): SR_..._04_D_2
    ' -->
    ' 1. Ormar: 04
    ' 2. Polica: 2

    For r = startRow To lastRow
        s = CStr(ws.Cells(r, "D").Value)
        If Len(s) > 0 Then
            parts = Split(s, "_") ' First numeric after the first dash

            If UBound(parts) >= 1 And IsNumeric(parts(3)) Then
                firstNum = CLng(parts(3))
            Else
                firstNum = 0  ' Use 0; change to 999999 if you want non-numeric to sink to bottom
            End If

            If UBound(parts) >= 1 And IsNumeric(parts(5)) Then
                secondNum = CLng(parts(5))
            Else
                secondNum = 0  ' Use 0; change to 999999 if you want non-numeric to sink to bottom
            End If

            ' Last numeric after the last dash
            ' If UBound(parts) >= 0 And IsNumeric(parts(UBound(parts))) Then
            '     lastNum = CLng(parts(UBound(parts)))
            ' Else
            '     lastNum = 0
            ' End If
        Else
            firstNum = 0
            secondNum = 0
            ' lastNum = 0
        End If

        ws.Cells(r, helperFirstCol).Value = firstNum
        ws.Cells(r, helperSecondCol).Value = secondNum
        ' ws.Cells(r, helperLastCol).Value = lastNum
    Next r

    ' text to numeric values
    With ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, helperLastCol)) ' include helpers
        .Value = .Value
    End With

    ' Sort by Column A, then helperFirst, then helperLast(all ascending)
    ws.Sort.SortFields.Clear


    ws.Sort.SortFields.Add key:=ws.Range(ws.Cells(startRow, "B"), ws.Cells(lastRow, "B")), _
                           SortOn:=xlSortOnValues, _
                           Order:=xlAscending, _
                           DataOption:=xlSortNormal

    ' D-polica (D): SR_..._04_D_2
    ' -->
    ' 1. Ormar: 04
    ws.Sort.SortFields.Add key:=ws.Range(ws.Cells(startRow, helperFirstCol), ws.Cells(lastRow, helperFirstCol)), _
                           SortOn:=xlSortOnValues, _
                           Order:=xlAscending, _
                           DataOption:=xlSortNormal

    ' D-polica (D): SR_..._04_D_2
    ' -->
    ' 2. Polica: 2
    ws.Sort.SortFields.Add key:=ws.Range(ws.Cells(startRow, helperSecondCol), ws.Cells(lastRow, helperSecondCol)), _
                           SortOn:=xlSortOnValues, _
                           Order:=xlAscending, _
                           DataOption:=xlSortNormal

    ' ws.Sort.SortFields.Add key:=ws.Range(ws.Cells(startRow, helperLastCol), ws.Cells(lastRow, helperLastCol)), _
    '                        SortOn:=xlSortOnValues, _
    '                        Order:=xlAscending, _
    '                        DataOption:=xlSortNormal

    ' 3. Spliter - slot (F): 02
    ws.Sort.SortFields.Add key:=ws.Range(ws.Cells(startRow, "F"), ws.Cells(lastRow, "F")), _
                           SortOn:=xlSortOnValues, _
                           Order:=xlAscending, _
                           DataOption:=xlSortNormal

    ' 4. Port Name (K): 001
    ws.Sort.SortFields.Add key:=ws.Range(ws.Cells(startRow, "K"), ws.Cells(lastRow, "K")), _
                           SortOn:=xlSortOnValues, _
                           Order:=xlAscending, _
                           DataOption:=xlSortNormal

    With ws.Sort
        .SetRange ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, helperLastCol)) ' include helpers
        .Header = IIf(hasHeader, xlYes, xlNo)
        .MatchCase = False
        .Orientation = xlTopToBottom
        .Apply
    End With

    ' Clean up: remove both helper columns at once(prevents index shifting)
    ws.Range(ws.Columns(helperFirstCol), ws.Columns(helperLastCol)).Delete

End Function


Function No_00_inser_No()
'
' Macro1 Macro
'

'
    Columns("A:A").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("A1").Select
    ActiveCell.FormulaR1C1 = "No."
    Range("A1").Select
    With Selection
        .HorizontalAlignment = xlGeneral
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Selection.Font.Bold = True
End Function


Function No_02_FTTx_022_OLT_PORT_Q()
    '
    ' date: 2026_03M_08 20:20:55
    ' Column P is "always exactly" OL-x-y
    '
    ' OLT (O): EquipmentID
    ' OLT - port (P): port-9-7
    '
    ' OLT PORT (Q): EquipmentID_9/7

    Dim ws As Worksheet, lastRow As Long, r As Long
    Dim parts As Variant

    ' Set ws=ActiveSheet ' Or:
    Set ws = ThisWorkbook.Worksheets("FTTx_022")

    ' Determine last used row across O and P
    lastRow = ws.Cells(ws.Rows.Count, "O").End(xlUp).Row

    If ws.Cells(ws.Rows.Count, "P").End(xlUp).Row > lastRow Then
        lastRow = ws.Cells(ws.Rows.Count, "P").End(xlUp).Row
    End If

    ' If your first row is headers, change the loop to: For r = 2 To lastRow.
    ' For r = 1 To lastRow
    For r = 2 To lastRow
        If Len(ws.Cells(r, "O").Value) > 0 And Len(ws.Cells(r, "P").Value) > 0 Then
            parts = Split(ws.Cells(r, "P").Value, "-") ' port-9-7
            If UBound(parts) >= 2 Then
                ws.Cells(r, "Q").Value = ws.Cells(r, "O").Value & "_" & parts(1) & "/" & parts(2)
            Else
                ws.Cells(r, "Q").Value = ws.Cells(r, "O").Value & "_?/?"
            End If
        End If
    Next r

End Function


Function No_03_FTTx_019_SortBy_OLT_Then_FirstAndLastNumberIn_PortName()
    '
    ' Equipment ID (A): EquipmentID
    ' Port name (D): port-9-7

    Dim ws As Worksheet

    Dim hasHeader As Boolean
    Dim lastRow As Long, lastCol As Long
    Dim startRow As Long
    Dim helperFirstCol As Long, helperLastCol As Long
    Dim r As Long
    Dim s As String
    Dim parts As Variant
    Dim firstNum As Long, lastNum As Long
    Dim lastCell As Range, lastColCell As Range

    '===SETTINGS===
    ' Set ws=ActiveSheet          ' Or:
    ' Set ws = ThisWorkbook.Worksheets("FTTx_019")
    Set ws = Workbooks("FTTx_019.xlsx").Worksheets("report1")
    hasHeader = True                      ' Set to False if there is NO header row
    '=================

    ' Find last used row & column robustly
    Set lastCell = ws.Cells.Find(What:="*", LookIn:=xlFormulas, LookAt:=xlPart, _
                                 SearchOrder:=xlByRows, SearchDirection:=xlPrevious)
    ' If lastCell Is Nothing Then Exit Sub
    If lastCell Is Nothing Then Exit Function
    lastRow = lastCell.Row

    Set lastColCell = ws.Cells.Find(What:="*", LookIn:=xlFormulas, LookAt:=xlPart, _
                                    SearchOrder:=xlByColumns, SearchDirection:=xlPrevious)
    lastCol = lastColCell.Column
    If lastCol < 4 Then lastCol = 4   ' Ensure we at least include column D

    startRow = IIf(hasHeader, 2, 1)
    ' If lastRow < startRow Then Exit Sub
    If lastRow < startRow Then Exit Function

    ' Add two helper columns at the far right
    helperFirstCol = lastCol + 1
    helperLastCol = lastCol + 2

    ws.Cells(1, helperFirstCol).Value = "__FirstNum_D__"
    ws.Cells(1, helperLastCol).Value = "__LastNum_D__"

    ' Fill helpers:
    '-__FirstNum_D__: number after the FIRST dash in D(e.g., "OL-12-34"-> 12)
    '-__LastNum_D__:  number after the LAST dash in D(e.g., "OL-12-34"-> 34)

    For r = startRow To lastRow
        s = CStr(ws.Cells(r, "D").Value)
        If Len(s) > 0 Then
            parts = Split(s, "-") ' First numeric after the first dash

            If UBound(parts) >= 1 And IsNumeric(parts(1)) Then
                firstNum = CLng(parts(1))
            Else
                firstNum = 0  ' Use 0; change to 999999 if you want non-numeric to sink to bottom
            End If

            ' Last numeric after the last dash
            If UBound(parts) >= 0 And IsNumeric(parts(UBound(parts))) Then
                lastNum = CLng(parts(UBound(parts)))
            Else
                lastNum = 0
            End If
        Else
            firstNum = 0
            lastNum = 0
        End If

        ws.Cells(r, helperFirstCol).Value = firstNum
        ws.Cells(r, helperLastCol).Value = lastNum
    Next r

    ' Sort by Column A, then helperFirst, then helperLast(all ascending)
    ws.Sort.SortFields.Clear

    ws.Sort.SortFields.Add key:=ws.Range(ws.Cells(startRow, "A"), ws.Cells(lastRow, "A")), _
                           SortOn:=xlSortOnValues, _
                           Order:=xlAscending, _
                           DataOption:=xlSortNormal

    ws.Sort.SortFields.Add key:=ws.Range(ws.Cells(startRow, helperFirstCol), ws.Cells(lastRow, helperFirstCol)), _
                           SortOn:=xlSortOnValues, _
                           Order:=xlAscending, _
                           DataOption:=xlSortNormal

    ws.Sort.SortFields.Add key:=ws.Range(ws.Cells(startRow, helperLastCol), ws.Cells(lastRow, helperLastCol)), _
                           SortOn:=xlSortOnValues, _
                           Order:=xlAscending, _
                           DataOption:=xlSortNormal

    With ws.Sort
        .SetRange ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, helperLastCol)) ' include helpers
        .Header = IIf(hasHeader, xlYes, xlNo)
        .MatchCase = False
        .Orientation = xlTopToBottom
        .Apply
    End With

    ' Clean up: remove both helper columns at once(prevents index shifting)
    ws.Range(ws.Columns(helperFirstCol), ws.Columns(helperLastCol)).Delete
End Function


Function No_04_FTTx_019_OLT_PORT_Q()
    '
    ' date: 2026_03M_08 20:20:55
    ' Column D is "always exactly" OL-x-y
    '
    ' Equipment ID (A): EquipmentID
    ' Port name (D): port-9-7
    '
    ' OLT PORT (Q): EquipmentID_9/7

    Dim ws As Worksheet, lastRow As Long, r As Long
    Dim parts As Variant

    ' Set ws=ActiveSheet ' Or:
    ' Set ws = ThisWorkbook.Worksheets("FTTx_019")
    Set ws = Workbooks("FTTx_019.xlsx").Worksheets("report1")

    ' Determine last used row across O and P
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

    If ws.Cells(ws.Rows.Count, "D").End(xlUp).Row > lastRow Then
        lastRow = ws.Cells(ws.Rows.Count, "D").End(xlUp).Row
    End If

    ' If your first row is headers, change the loop to: For r = 2 To lastRow.
    ' For r = 1 To lastRow
    For r = 2 To lastRow
        If Len(ws.Cells(r, "A").Value) > 0 And Len(ws.Cells(r, "D").Value) > 0 Then
            parts = Split(ws.Cells(r, "D").Value, "-") ' port-9-7
            If UBound(parts) >= 2 Then
                ws.Cells(r, "Q").Value = ws.Cells(r, "A").Value & "_" & parts(1) & "/" & parts(2)
            Else
                ws.Cells(r, "Q").Value = ws.Cells(r, "A").Value & "_?/?"
            End If
        End If
    Next r
End Function


Function No_05_FTTx_019_Lookup_4Cols_to_FTTx_022()

    Dim prvaSH As Worksheet, drugaSH As Worksheet
    Set prvaSH = ThisWorkbook.Worksheets("FTTx_022")

    ' Set drugaSH = ThisWorkbook.Worksheets("FTTx_019")
    Set drugaSH = Workbooks("FTTx_019.xlsx").Worksheets("report1")

    Dim lastQprva As Long, lastQdruga As Long
    lastQprva = prvaSH.Cells(prvaSH.Rows.Count, "Q").End(xlUp).Row
    lastQdruga = drugaSH.Cells(drugaSH.Rows.Count, "Q").End(xlUp).Row

    Dim dict As Object: Set dict = CreateObject("Scripting.Dictionary")
    Dim arrQdruga As Variant, arrEtoH As Variant
    Dim r As Long, key As String

    '--- Load lookup table from drugaSH (B -> E,F,G,H)
    arrQdruga = drugaSH.Range("Q2:Q" & lastQdruga).Value
    arrEtoH = drugaSH.Range("E2:H" & lastQdruga).Value   ' RETURN COLUMNS

    For r = 1 To UBound(arrQdruga, 1)
        key = Trim$(CStr(arrQdruga(r, 1)))
        If Len(key) > 0 Then
            If Not dict.Exists(key) Then
                dict.Add key, Array(arrEtoH(r, 1), arrEtoH(r, 2), arrEtoH(r, 3), arrEtoH(r, 4))
            End If
        End If
    Next r

    '--- Read prvaSH column Q
    Dim arrPprva As Variant, outArr As Variant
    arrPprva = prvaSH.Range("Q2:Q" & lastQprva).Value
    ReDim outArr(1 To UBound(arrPprva, 1), 1 To 4)

    '--- Lookup loop
    For r = 1 To UBound(arrPprva, 1)
        key = Trim$(CStr(arrPprva(r, 1)))

        If Len(key) = 0 Then
            ' blank: leave outputs empty
            outArr(r, 1) = ""
            outArr(r, 2) = ""
            outArr(r, 3) = ""
            outArr(r, 4) = ""
        ElseIf dict.Exists(key) Then
            Dim item4
            item4 = dict(key)
            outArr(r, 1) = item4(0)   ' E -> R
            outArr(r, 2) = item4(1)   ' F -> S
            outArr(r, 3) = item4(2)   ' G -> T
            outArr(r, 4) = item4(3)   ' H -> U
        Else
            outArr(r, 1) = ""
            outArr(r, 2) = ""
            outArr(r, 3) = ""
            outArr(r, 4) = ""
        End If
    Next r

    '--- Write results into R:S:T:U
    prvaSH.Range("R2").Resize(UBound(outArr, 1), 4).Value = outArr

    ActiveWorkbook.Save
    Workbooks("FTTx_019.xlsx").Save

End Function


Function No_06_GPON_Report_001_SortBy_OLT_Then_1And2_NumberIn_Port()
    '
    ' OLT (I): EquipmentID
    ' PORT (M): ON-1-0-10

    Dim ws As Worksheet

    Dim hasHeader As Boolean
    Dim lastRow As Long, lastCol As Long
    Dim startRow As Long
    Dim helperFirstCol As Long, helperSecondCol As Long, helperLastCol As Long
    Dim r As Long
    Dim s As String
    Dim parts As Variant
    Dim firstNum As Long, secondNum As Long, lastNum As Long
    Dim lastCell As Range, lastColCell As Range

    '===SETTINGS===
    ' Set ws=ActiveSheet          ' Or:
    ' Set ws = ThisWorkbook.Worksheets("GPON_Report_001")
    Set ws = Workbooks("GPON_Report_001.xlsx").Worksheets("report5")
    hasHeader = True                      ' Set to False if there is NO header row
    '=================

    ' Find last used row & column robustly
    Set lastCell = ws.Cells.Find(What:="*", LookIn:=xlFormulas, LookAt:=xlPart, _
                                 SearchOrder:=xlByRows, SearchDirection:=xlPrevious)
    ' If lastCell Is Nothing Then Exit Sub
    If lastCell Is Nothing Then Exit Function
    lastRow = lastCell.Row

    Set lastColCell = ws.Cells.Find(What:="*", LookIn:=xlFormulas, LookAt:=xlPart, _
                                    SearchOrder:=xlByColumns, SearchDirection:=xlPrevious)
    lastCol = lastColCell.Column
    If lastCol < 4 Then lastCol = 4   ' Ensure we at least include column D

    startRow = IIf(hasHeader, 2, 1)
    ' If lastRow < startRow Then Exit Sub
    If lastRow < startRow Then Exit Function

    ' Add two helper columns at the far right
    helperFirstCol = lastCol + 1
    helperSecondCol = lastCol + 2
    helperLastCol = lastCol + 3

    ws.Cells(1, helperFirstCol).Value = "__FirstNum_D__"
    ws.Cells(1, helperSecondCol).Value = "__SecondNum_D__"
    ws.Cells(1, helperLastCol).Value = "__LastNum_D__"

    ' Fill helpers:
    '-__FirstNum_D__: number after the FIRST dash in D(e.g., "ON-1-0-10"-> 1)
    '-__SecondNum_D__: number after the SECOND dash in D(e.g., "ON-1-0-10"-> 0)
    '-__LastNum_D__:  number after the LAST dash in D(e.g., "ON-1-0-10"-> 10)

    For r = startRow To lastRow
        s = CStr(ws.Cells(r, "M").Value)
        If Len(s) > 0 Then
            parts = Split(s, "-") ' First numeric after the first dash

            If UBound(parts) >= 1 And IsNumeric(parts(1)) Then
                firstNum = CLng(parts(1))
            Else
                firstNum = 0  ' Use 0; change to 999999 if you want non-numeric to sink to bottom
            End If

            If UBound(parts) >= 1 And IsNumeric(parts(1)) Then
                secondNum = CLng(parts(2))
            Else
                secondNum = 0  ' Use 0; change to 999999 if you want non-numeric to sink to bottom
            End If

            ' Last numeric after the last dash
            If UBound(parts) >= 0 And IsNumeric(parts(UBound(parts))) Then
                lastNum = CLng(parts(UBound(parts)))
            Else
                lastNum = 0
            End If
        Else
            firstNum = 0
            lastNum = 0
        End If

        ws.Cells(r, helperFirstCol).Value = firstNum
        ws.Cells(r, helperSecondCol).Value = secondNum
        ws.Cells(r, helperLastCol).Value = lastNum
    Next r

    ' Sort by Column A, then helperFirst, then helperLast(all ascending)
    ws.Sort.SortFields.Clear

    ws.Sort.SortFields.Add key:=ws.Range(ws.Cells(startRow, "I"), ws.Cells(lastRow, "I")), _
                           SortOn:=xlSortOnValues, _
                           Order:=xlAscending, _
                           DataOption:=xlSortNormal

    ws.Sort.SortFields.Add key:=ws.Range(ws.Cells(startRow, helperFirstCol), ws.Cells(lastRow, helperFirstCol)), _
                           SortOn:=xlSortOnValues, _
                           Order:=xlAscending, _
                           DataOption:=xlSortNormal

    ws.Sort.SortFields.Add key:=ws.Range(ws.Cells(startRow, helperSecondCol), ws.Cells(lastRow, helperSecondCol)), _
                           SortOn:=xlSortOnValues, _
                           Order:=xlAscending, _
                           DataOption:=xlSortNormal

    ws.Sort.SortFields.Add key:=ws.Range(ws.Cells(startRow, helperLastCol), ws.Cells(lastRow, helperLastCol)), _
                           SortOn:=xlSortOnValues, _
                           Order:=xlAscending, _
                           DataOption:=xlSortNormal

    With ws.Sort
        .SetRange ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, helperLastCol)) ' include helpers
        .Header = IIf(hasHeader, xlYes, xlNo)
        .MatchCase = False
        .Orientation = xlTopToBottom
        .Apply
    End With

    ' Clean up: remove both helper columns at once(prevents index shifting)
    ws.Range(ws.Columns(helperFirstCol), ws.Columns(helperLastCol)).Delete
End Function


Function No_07_GPON_Report_001_OLT_PORT_V()
    '
    ' date: 2026_03M_09 15:24:05
    ' Column M is "always exactly" ON-x-y-z
    '
    ' OLT (I): EquipmentID
    ' PORT (M): ON-1-0-10
    '
    ' OLT PORT (V): EquipmentID_9/7

    Dim ws As Worksheet, lastRow As Long, r As Long
    Dim parts As Variant

    ' Set ws=ActiveSheet ' Or:
    ' Set ws = ThisWorkbook.Worksheets("GPON_Report_001")
    Set ws = Workbooks("GPON_Report_001.xlsx").Worksheets("report5")

    ' Determine last used row across O and P
    lastRow = ws.Cells(ws.Rows.Count, "I").End(xlUp).Row

    If ws.Cells(ws.Rows.Count, "M").End(xlUp).Row > lastRow Then
        lastRow = ws.Cells(ws.Rows.Count, "D").End(xlUp).Row
    End If

    ws.Cells(1, "V").Value = "OLT PORT"

    ' If your first row is headers, change the loop to: For r = 2 To lastRow.
    ' For r = 1 To lastRow
    For r = 2 To lastRow
        If Len(ws.Cells(r, "I").Value) > 0 And Len(ws.Cells(r, "M").Value) > 0 Then
            parts = Split(ws.Cells(r, "M").Value, "-") ' port-9-7
            If UBound(parts) >= 2 Then
                ws.Cells(r, "V").Value = ws.Cells(r, "I").Value & "_" & parts(1) & "/" & parts(2)
            Else
                ws.Cells(r, "V").Value = ws.Cells(r, "I").Value & "_?/?"
            End If
        End If
    Next r
End Function


Function No_08_GPON_Report_001_KOR_i_ER_ONT)
    ' Count: KOR i ER_ONT
    ' date: 2026_03M_20 21:13:09
    '
    ' OLT PORT (V): EquipmentID_9/7
    '
    ' EDIT: 2026_03M_23 15:26:26

    ' Below is a simple and clean Excel VBA macro
    ' that counts how many times each name appears in a list
    ' like the one you posted (Marko, Marko, Janko…).

    Dim ws As Worksheet, wsNew As Worksheet
    Dim wb As Workbook

    Dim dict As Object, dictONT As Object
    Dim lastRow As Long
    Dim i As Long
    Dim sOLT_PORT As String
    Dim sER_ONT_HU_OL As String

    Dim iBrojac As Integer

    Set wb = Workbooks("GPON_Report_001.xlsx")

    ' Set wsNew = wb.Worksheets.Add(Before:=Worksheets("report5"))
    ' Set wsNew = wb.Worksheets.Add(After:=wb.Worksheets(Worksheets.Count))

    Set wsNew = wb.Worksheets.Add(After:=wb.Worksheets("report5"))
    wsNew.Name = "KOR_i_ER_ONT"
    ' Set wsNew = wb.Worksheets("KOR_i_ER_ONT")

    ' Set ws=ActiveSheet ' Or:
    ' Set ws = ThisWorkbook.Worksheets("report5")
    ' Set ws = Workbooks("GPON_Report_001.xlsx").Worksheets("report5")
    Set ws = wb.Worksheets("report5")

    ' Create dictionary
    Set dict = CreateObject("Scripting.Dictionary")
    Set dictONT = CreateObject("Scripting.Dictionary")

    ' Find last row in column V
    lastRow = ws.Cells(ws.Rows.Count, "V").End(xlUp).Row

    iBrojac = 0

    ' Loop through names
    ' For i = 1 To lastRow
    For i = 2 To lastRow
        sOLT_PORT = Trim(ws.Cells(i, "V").Value)
        sER_ONT_HU_OL = Trim(ws.Cells(i, "K").Value)

        ' Count KOR
        If dict.Exists(sOLT_PORT) Then
            dict(sOLT_PORT) = dict(sOLT_PORT) + 1
        Else
            dict.Add sOLT_PORT, 1
        End If

        ' Count ER_ONT_HU_OL
        If dictONT.Exists(sOLT_PORT) Then
            If sER_ONT_HU_OL = "DA" Then
                iBrojac = iBrojac + 1
                dictONT(sOLT_PORT) = iBrojac
            End If
        Else
            iBrojac = 0
            If sER_ONT_HU_OL = "DA" Then
                iBrojac = iBrojac + 1
            End If
            dictONT.Add sOLT_PORT, iBrojac
        End If

    Next i

    ' Output results to columns A, B and C
    Dim rowOut As Long
    rowOut = 1

    wsNew.Cells(1, "A").Value = "OLT PORT (V)"
    wsNew.Cells(1, "B").Value = "KOR"
    wsNew.Cells(1, "C").Value = "ER_ONT"

    Dim key As Variant

    For Each key In dict.Keys
        wsNew.Cells(rowOut + 1, "A").Value = key
        wsNew.Cells(rowOut + 1, "B").Value = dict(key)

        ' wsNew.Cells(rowOut + 1, "D").Value = key
        wsNew.Cells(rowOut + 1, "C").Value = dictONT(key)

        rowOut = rowOut + 1
    Next key

    MsgBox "Done! Names counted successfully."

End Function


Function No_09_GPON_Report_001_Lookup_2Cols_to_FTTx_022()

    Dim prvaSH As Worksheet, drugaSH As Worksheet
    Set prvaSH = ThisWorkbook.Worksheets("FTTx_022")

    ' Set drugaSH = ThisWorkbook.Worksheets("GPON_Report_001")
    Set drugaSH = Workbooks("GPON_Report_001.xlsx").Worksheets("KOR_i_ER_ONT")
    ' Set drugaSH = Workbooks("GPON_Report_001.xlsx").Worksheets("report5")

    Dim lastQprva As Long, lastQdruga As Long
    lastQprva = prvaSH.Cells(prvaSH.Rows.Count, "Q").End(xlUp).Row
    lastQdruga = drugaSH.Cells(drugaSH.Rows.Count, "A").End(xlUp).Row

    Dim dict As Object: Set dict = CreateObject("Scripting.Dictionary")
    Dim arrQdruga As Variant, arrBtoC As Variant
    Dim r As Long, key As String

    '--- Load lookup table from drugaSH (A -> B,C)
    arrQdruga = drugaSH.Range("A5:A" & lastQdruga).Value
    arrBtoC = drugaSH.Range("B5:C" & lastQdruga).Value   ' RETURN COLUMNS

    For r = 1 To UBound(arrQdruga, 1)
        key = Trim$(CStr(arrQdruga(r, 1)))
        If Len(key) > 0 Then
            If Not dict.Exists(key) Then
                dict.Add key, Array(arrBtoC(r, 1), arrBtoC(r, 2))
            End If
        End If
    Next r

    '--- Read prvaSH column Q
    Dim arrPprva As Variant, outArr As Variant
    arrPprva = prvaSH.Range("Q2:Q" & lastQprva).Value
    ReDim outArr(1 To UBound(arrPprva, 1), 1 To 2)

    '--- Lookup loop
    For r = 1 To UBound(arrPprva, 1)
        key = Trim$(CStr(arrPprva(r, 1)))

        If Len(key) = 0 Then
            ' blank: leave outputs empty
            outArr(r, 1) = ""
            outArr(r, 2) = ""
        ElseIf dict.Exists(key) Then
            Dim item2
            item2 = dict(key)
            outArr(r, 1) = item2(0)   ' B -> V
            outArr(r, 2) = item2(1)   ' C -> W
        Else
            outArr(r, 1) = ""
            outArr(r, 2) = ""
        End If
    Next r

    '--- Write results into V:W
    prvaSH.Range("V2").Resize(UBound(outArr, 1), 2).Value = outArr

    ActiveWorkbook.Save
    Workbooks("GPON_Report_001.xlsx").Save

End Function


Function No_10_insert_Mreza3()
'
' date: 2026_03M_23 20:42:25
'
    Columns("O:O").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove

    Range("O1").Select
'    ActiveCell.FormulaR1C1 = "Mreza3"
    ActiveCell.FormulaR1C1 = "identifikator (prazno=0/kruta=1/fleksibilna=100)"

    Range("O1").Select

    With Selection
        .HorizontalAlignment = xlGeneral
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With

    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With

    Selection.Font.Bold = True

End Function

Function No_10_Red1()
' Macro1
'
    ThisWorkbook.Worksheets("RED1").Rows("1:1").Copy

    ThisWorkbook.Worksheets("FTTx_022").Range("1:1").PasteSpecial Paste:=xlPasteColumnWidths, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False

    ThisWorkbook.Worksheets("FTTx_022").Range("1:1").PasteSpecial Paste:=xlPasteAll

    Range("A1").Select

End Function

Sub No_10_identifikator_za_PivotTable()
    ' identifikator (O) - za kreiranje PivotTable
    ' date: 2026_03M_11 13:00:29
    '
    ' EDIT: date: 2026_05M_28 11:47:53

    ' ------------------------------
    ' Call Function:
    No_10_insert_Mreza3

    Dim prvaSH As Worksheet

    Dim lastRowPrva As Long
    Dim i As Long
    Dim sMreza As String, sMreza3 As String

    Set prvaSH = ThisWorkbook.Worksheets("FTTx_022")

    lastRowPrva = prvaSH.Cells(prvaSH.Rows.Count, "G").End(xlUp).Row

    For i = 2 To lastRowPrva
        sMreza = prvaSH.Range("N" & i).Value
        sMreza3 = Left(sMreza, 3)
        If sMreza = "kruta" Then
            prvaSH.Range("O" & i).Value = "1"
        ElseIf sMreza3 = "" Then
            prvaSH.Range("O" & i).Value = "0"
        ElseIf sMreza3 = "ODC" Then
            prvaSH.Range("O" & i).Value = "100"
        End If
    Next i

    ' Call Function:
    No_10_Red1

    ActiveWorkbook.Save
End Sub


Sub No_11_Create_PivotTable()
    ' VBA Code to Create a Pivot Table
    ' This code assumes you already have a worksheet with tabular data (with headers) and
    ' want to create a Pivot Table on a new sheet.
    '
    ' How It Works
    ' 1.    Source Data:
    '   o   The macro uses CurrentRegion from cell A1 in the "Data" sheet to automatically detect the full table.
    '   o   Ensure your data has headers in the first row.
    '
    ' 2.    Pivot Cache:
    '   o   Stores the data for the Pivot Table.
    '
    ' 3.    Pivot Table Creation:
    '   o   Creates a new sheet "PivotTableSheet" if it doesn’t exist.
    '   o   Places the Pivot Table starting at cell A3.
    '
    ' 4.    Field Setup:
    '   o   "Category" is set as a Row Field.
    '   o   "Region" is set as a Column Field.
    '   o   "Sales" is summarized as Sum.

    ' Replace "Category", "Region", and "Sales" with your actual column headers.

    Dim wsData As Worksheet
    Dim wsPivot As Worksheet

    Dim pc As PivotCache
    Dim pt As PivotTable

    Dim srcData As String
    Dim pivotName As String

    On Error Resume Next
    Set wsData = ThisWorkbook.Sheets("FTTx_022")       ' <-- CHANGE if needed
    On Error GoTo 0

    If wsData Is Nothing Then
        MsgBox "Data sheet not found!", vbCritical
        Exit Sub
    End If

    '--- Create or get pivot sheet ---
    On Error Resume Next
    Set wsPivot = ThisWorkbook.Sheets("PivotTableSheet")
    If wsPivot Is Nothing Then
        Set wsPivot = ThisWorkbook.Sheets.Add
        wsPivot.Name = "PivotTableSheet"
    End If
    On Error GoTo 0

    '--- DYNAMIC SOURCE RANGE (bulletproof) ---
    srcData = "'" & wsData.Name & "'!" & _
              wsData.Range("A1").CurrentRegion.Address(ReferenceStyle:=xlR1C1)

    ' VBA editor, Immediate Window: Debug.Print srcData --> 'FTTx_022'!R1C1:R259C18
    ' OLT PORT (R) nije bio upisan!!!

    '--- Create pivot cache ---
    Set pc = ThisWorkbook.PivotCaches.Create( _
        SourceType:=xlDatabase, _
        SourceData:=srcData)

    pivotName = "MyPivotTable"

    '--- Delete existing pivot with same name ---
    On Error Resume Next
    wsPivot.PivotTables(pivotName).TableRange2.Clear
    On Error GoTo 0

    '--- Create the pivot table ---
    Set pt = pc.CreatePivotTable( _
        TableDestination:=wsPivot.Range("A3"), _
        TableName:=pivotName)

    '--- Add fields ---
    With pt
        .PivotFields("Spliter - modul").Orientation = xlRowField
        .PivotFields("Spliter - modul").Position = 1
        .PivotFields("Spliter - modul").Subtotals(1) = False
        .PivotFields("Spliter - modul").LayoutForm = xlTabular

        .PivotFields("Spliter - tip").Orientation = xlRowField
        .PivotFields("Spliter - tip").Position = 2
        .PivotFields("Spliter - tip").Subtotals(1) = False
        .PivotFields("Spliter - tip").LayoutForm = xlTabular

        .PivotFields("Vendor").Orientation = xlRowField
        .PivotFields("Vendor").Position = 3
        .PivotFields("Vendor").Subtotals(1) = False
        .PivotFields("Vendor").LayoutForm = xlTabular

        .AddDataField .PivotFields("Spliter"), "Count of Spliter", xlCount
        .AddDataField .PivotFields("identifikator (prazno=0/kruta=1/fleksibilna=100)"), "Sum of identifikator (prazno=0/kruta=1/fleksibilna=100)", xlSum
    End With

    ' MsgBox "Pivot Table created successfully!", vbInformation

    wsPivot.Range("F3").Value = "Ožičenje modula v2"
    ' wsPivot.Range("F4").Formula = "=IF(AND(E4=0;D4=0);""nije instaliran"";IF(AND(E4=0;D4>0);""nije ožičen"";IF(E4=D4;""kruti"";IF(E4<D4;""kruti nepotpun"";IF(E4=100*D4;""fleksibilni"";IF(MOD(E4;10)=0;""flexibilni nepotpun"";IF(ROUNDDOWN(E4/100;0)+E4-(ROUNDDOWN(E4/100;0)*100)=D4;""hibridni"";""hibridni nepotpun"")))))))"

    wsPivot.Range("F4").Formula = _
    "=IF(AND(E4=0,D4=0)," & _
    """nije instaliran""," & _
    "IF(AND(E4=0,D4>0)," & _
    """nije ožičen""," & _
    "IF(E4=D4," & _
    """kruti""," & _
    "IF(E4<D4," & _
    """kruti nepotpun""," & _
    "IF(E4=100*D4," & _
    """fleksibilni""," & _
    "IF(MOD(E4,10)=0," & _
    """fleksibilni nepotpun""," & _
    "IF(ROUNDDOWN(E4/100,0)+E4-(ROUNDDOWN(E4/100,0)*100)=D4," & _
    """hibridni""," & _
    """hibridni nepotpun"")))))))"

    Exit Sub

    ' ------------------------------------------------------------
    ' Commonly Used Pivot Field Properties
    '
    ' Property / Method Description

    ' .NumberFormat Sets the display format for numbers.
    ' .Subtotals(index) Enables/disables specific subtotal types (1 = Automatic).
    ' .LayoutForm   Controls layout (xlCompactRow, xlOutline, xlTabular).
    ' .RepeatLabels Repeats item labels in tabular form.
    ' .AutoSort Sorts field items ascending/descending by a data field.
    ' .Function Changes summary calculation (xlSum, xlAverage, xlCount, etc.).
    ' .Name Renames the field in the Pivot Table.

    ' ------------------------------------------------------------
'     ' Example 1: Change Number Format for a Data Field
'     pt.DataFields("Total Sales").NumberFormat = "#,##0.00" ' Two decimal places
'
'     ' Example 2: Change Row Field Layout
'     Set pf = pt.PivotFields("Category")
'     pf.Subtotals(1) = False ' Remove automatic subtotals
'     pf.LayoutForm = xlTabular ' Show in tabular form
'     pf.RepeatLabels = True    ' Repeat item labels
'
'     ' Example 3: Sort Row Field
'     pf.AutoSort xlDescending, "Total Sales"
'
'     ' Example 4: Change Summary Function
'     pt.DataFields("Total Sales").Function = xlAverage
'     pt.DataFields("Total Sales").Name = "Average Sales"
'
   ' ------------------------------------------------------------

ErrHandler:
    MsgBox "Error: " & Err.Description, vbCritical

    ActiveWorkbook.Save

End Sub


Sub No_12_CopyFrom_PivotTable_Spliter_modul_nije_instaliran()
    ' Ožičenje modula v2
    ' Spliter - modul: nije instaliran
    ' date: 2026_03M_11 12:13:28
    '
    ' EDIT: date: 2026_05M_04 16:07:24

    ' ------------------------------
    Dim prvaSH As Worksheet, drugaSH As Worksheet

    Dim lastRowPrva As Long, lastRowDruga As Long
    Dim i As Long, j As Long
    Dim s1SpliterModul As String, s2SpliterModul As String
    Dim sPrvaF As String

    Set prvaSH = ThisWorkbook.Worksheets("PivotTableSheet")
    Set drugaSH = ThisWorkbook.Worksheets("FTTx_022")
    ' Set drugaSH = Workbooks("GPON_Report_001.xlsx").Worksheets("count")

    lastRowPrva = prvaSH.Cells(prvaSH.Rows.Count, "A").End(xlUp).Row
    lastRowDruga = drugaSH.Cells(drugaSH.Rows.Count, "G").End(xlUp).Row

    For i = 4 To lastRowPrva
        ' "Spliter - modul"
        s1SpliterModul = prvaSH.Cells(i, 1).Value
        sPrvaF = prvaSH.Range("F" & i).Value

        ' If (sPrvaF = "nije instaliran") Or (sPrvaF = "nije ožičen") Then
        '     For j = 2 To lastRowDruga
        '         s2SpliterModul = drugaSH.Cells(j, 7).Value
        '         If s2SpliterModul = s1SpliterModul Then
        '             drugaSH.Range("Y" & j).Value = sPrvaF
        '             Exit For
        '         End If
        '     Next j
        ' End If

        For j = 2 To lastRowDruga
            s2SpliterModul = drugaSH.Cells(j, 7).Value
            If s2SpliterModul = s1SpliterModul Then
                drugaSH.Range("Y" & j).Value = sPrvaF
                Exit For
            End If
        Next j

    Next i

    ActiveWorkbook.Save

End Sub


Sub No_13_FORMAT_Mark_new_D_polica()
    ' date: 2026_03M_10 20:27:39
    ' EDIT:

    ' ------------------------------
    ' Call Function:
    Mark_new_Spliter_slot

    Dim FinalRow As Long
    FinalRow = ActiveSheet.Cells(ActiveSheet.Rows.Count, 1).End(xlUp).Row

    Dim i As Long
    Dim s1Value As String, s1ValueOLD As String

    Dim Rng As Range

    s1ValueOLD = ActiveSheet.Cells(2, 4).Value

    For i = 3 To FinalRow
        ' D-polica (D)
        s1Value = ActiveSheet.Cells(i, 4).Value

        If s1Value <> s1ValueOLD Then
           ' Set Rng = ActiveSheet.Cells(i, j)
           Set Rng = ActiveSheet.Range("A" & i & ":Y" & i)
           With Rng
                With .Borders(xlEdgeTop)
                   .LineStyle = xlContinuous
                   ' .Weight = xlThin
                   .Weight = xlThick
                   ' .ColorIndex = 1   ' black
                   ' .ColorIndex = 3   ' red
                   ' .ColorIndex = 4   ' green
                   .ColorIndex = 5   ' blue
                End With
            End With
        End If
        s1ValueOLD = s1Value
    Next i

    Set Rng = Nothing

    ActiveWorkbook.Save
End Sub


Function Mark_new_Spliter_slot()
    ' date: 2026_03M_10 21:26:08
    ' EDIT:

    ' ------------------------------
    Dim FinalRow As Long
    FinalRow = ActiveSheet.Cells(ActiveSheet.Rows.Count, 1).End(xlUp).Row

    Dim i As Long
    Dim s1Slot As String, s1SlotOLD As String

    Dim Rng As Range

    s1SlotOLD = ActiveSheet.Cells(1, 6).Value

    For i = 2 To FinalRow
        s1Slot = ActiveSheet.Cells(i, 6).Value

        If s1Slot <> s1SlotOLD Then
           ' Set Rng = ActiveSheet.Cells(i, j)
           Set Rng = ActiveSheet.Range("A" & i & ":Y" & i)
           With Rng
                With .Borders(xlEdgeTop)
                   .LineStyle = xlContinuous
                   ' .Weight = xlThin
                   .Weight = xlThick
                   .ColorIndex = 1   ' black
                   ' .ColorIndex = 3   ' red
                   ' .ColorIndex = 4   ' green
                   ' .ColorIndex = 5   ' blue
                End With
            End With
        End If
        s1SlotOLD = s1Slot
    Next i

    Set Rng = Nothing

End Function


Sub No_30_Novi_spliter_WriteDSeries()
    ' Excel VBA macro that writes the sequence
    ' date: 2026_03M_11 15:51:32
    '
    ' "2 novi spliter"

    Dim ws As Worksheet
    Dim i As Long
    Dim rowPtr As Long

    Set ws = ActiveSheet        ' or
    ' Set ws = Sheets("Sheet1")
    rowPtr = 1                  ' starting row

    For i = 1001 To 1999
        ws.Cells(rowPtr, 1).Value = "D" & i
        ws.Cells(rowPtr + 1, 1).Value = "D" & i
        rowPtr = rowPtr + 2
    Next i

End Sub


Sub No_31_Novi_slot_modul_WriteSeries_H_S()
    ' Excel VBA macro that writes the sequence
    ' date: 2026_03M_11 20:09:30
    '
    ' "1 novi slot-modul"

    Dim ws As Worksheet

    Set ws = ActiveSheet        ' or
    ' Set ws = Sheets("Sheet1")

    Dim i As Long, j As Long
    Dim rowPtr As Long

    rowPtr = 3                  ' starting row

    For i = 0 To 7
        For j = 1 To 4
            ActiveSheet.Range("A" & rowPtr).Value = 9 + i
            rowPtr = rowPtr + 1
        Next j
    Next i

    rowPtr = 3                  ' starting row

    For i = 0 To 7
        For j = 1 To 6
            ActiveSheet.Range("B" & rowPtr).Value = 9 + i
            rowPtr = rowPtr + 1
        Next j
    Next i

End Sub


Sub No_32_Novi_slot_modul_WriteSeries_ADC()
    ' Excel VBA macro that writes the sequence
    ' date: 2026_03M_11 20:09:30
    '
    ' "1 novi slot-modul"
    '
    ' EDIT: date: 2026_05M_08 15:52:49

    Dim ws As Worksheet

    Set ws = ActiveSheet        ' or
    ' Set ws = Sheets("Sheet1")

    Dim i As Long, j As Long
    Dim rowPtr As Long

    rowPtr = 3                  ' starting row

    For i = 0 To 11
        For j = 1 To 2
            ActiveSheet.Range("A" & rowPtr).Value = 13 + i
            rowPtr = rowPtr + 1
        Next j
    Next i

    rowPtr = 3                  ' starting row

    For i = 0 To 11
        For j = 1 To 4
            ActiveSheet.Range("B" & rowPtr).Value = 13 + i
            rowPtr = rowPtr + 1
        Next j
    Next i

    rowPtr = 3                  ' starting row

    For i = 0 To 11
        ActiveSheet.Range("C" & rowPtr).Value = 13 + i
        rowPtr = rowPtr + 1
    Next i

End Sub

Sub No_33_WriteOLTSeries()
    ' Excel VBA macro that writes the sequence
    ' date: 2026_03M_11 15:51:32
    '
    ' "3 Novi OLT port"
    '
    ' EDIT: date: 2026_05M_08 13:52:24

    Dim ws As Worksheet

    Dim olt As Long
    Dim port As Long
    Dim rowPtr1 As Long, rowPtr2 As Long

    Set ws = ActiveSheet      ' Or:
    ' Set ws=Sheets("Sheet1")
    rowPtr1 = 1                ' Starting row
    rowPtr2 = 1                ' Starting row

    For olt = 1 To 7          ' OL-1 to OL-7
        For port = 1 To 16    ' 1 to 16 for each OLT
            ws.Cells(rowPtr1, 1).Value = "OL-" & olt & "-" & port
            rowPtr1 = rowPtr1 + 1

            ws.Cells(rowPtr2, 2).Value = "OL-" & olt & "-" & port
            rowPtr2 = rowPtr2 + 2
        Next port
    Next olt

End Sub
