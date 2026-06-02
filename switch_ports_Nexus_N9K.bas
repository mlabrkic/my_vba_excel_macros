' date: 2026_06M_02 21:16:02
' ------------------------------------------------------------

' On this way...
'
' data base, Web browser
' router, all ports view, open
' Select all ports, Copy
'
' Paste to "ports.csv"
' Import "ports.csv" to Excel (tab delimiter).
' ------------------------------------------------------------

Function Fun_01_Delete_empty_rows()
' Attribute Fun_01_xxx_W.VB_ProcData.VB_Invoke_Func = "k\n14"
' CTRL-K
' mlabrkic, date: 2024_09M_25

' EDIT:

' C:\02_PROG\10_Windows\05_VBA\PacktPublishing\vba_Excel_section-11-iteration_2023_05M_11.bas

' ------------------------------
'--- BETTER
' Dim myWB As Workbook
' Set myWB = ThisWorkbook

' Dim radnaSH As Worksheet
' Set radnaSH = myWB.Sheets("Sheet2")
' Set radnaSH = myWB.Worksheets("Sheet2")

' radnaSH.Cells(1, 1).Value = Now()
' radnaSH.Cells(2, 1).Value = "Hello"  'will place "Hello" in A2

' Dim myCell As Range
' Set myCell = radnaSH.Range("D3")
' myCell.Value = 3.1415

' ------------------------------
    Dim myWB As Workbook
    Set myWB = ThisWorkbook

    Dim radnaSH As Worksheet
    Set radnaSH = myWB.Worksheets("PORTOVI")

' ------------------------------
    ' Delete the first row:
    Rows("1:1").Select
    Selection.Delete Shift:=xlUp

    Dim FinalRow As Long
    FinalRow = radnaSH.Cells(radnaSH.Rows.Count, 1).End(xlUp).Row

    Dim i As Long
    Dim s1Slot As String, s2PortName As String

    For i = FinalRow To 1 Step -1
        s1Slot = radnaSH.Cells(i, 1).Value
        s2PortName = radnaSH.Cells(i, 2).Value

        If (s1Slot = "") And (s2PortName = "") Then
            radnaSH.Rows(i).EntireRow.Delete
        ElseIf (s1Slot = "--") And (s2PortName = "--") Then
            radnaSH.Rows(i).EntireRow.Delete
        ElseIf (s1Slot = "-1") Then
            radnaSH.Rows(i).EntireRow.Delete
        End If
    Next i

    Set radnaSH = Nothing
    Set myWB = Nothing

End Function

Function Fun_02_Copy_user_name_and_address()
' mlabrkic, date: 2024_09M_25

' EDIT:

' ------------------------------
    Dim myWB As Workbook

    Set myWB = ThisWorkbook
    ' ThisWorkbook refers to the workbook containing the running VBA code

    Dim radnaSH As Worksheet
    Set radnaSH = myWB.Worksheets("PORTOVI")

' ------------------------------
    Dim FinalRow As Long
    FinalRow = radnaSH.Cells(radnaSH.Rows.Count, 1).End(xlUp).Row

    Dim i As Long
    Dim j As Integer
    Dim s2PortName As String

    j = 0

'    radnaSH.Range("K" & 1).Value = "KORISNIK"
    radnaSH.Range("K1").Value = "KORISNIK"


    For i = 1 To FinalRow Step 1
        s2PortName = radnaSH.Cells(i, 2).Value
        If s2PortName = "" Then
            j = j + 1
            radnaSH.Cells(i - j, 10 + j).Value = radnaSH.Cells(i, 1).Value  ' K, L
        Else
            j = 0
        End If
    Next i

    Set radnaSH = Nothing
    Set myWB = Nothing

End Function

Function Fun_03_Delete_rows_without_Port_Name()
' mlabrkic, date: 2024_09M_25

' EDIT:

' ------------------------------
    Dim myWB As Workbook
    Set myWB = ThisWorkbook

    Dim radnaSH As Worksheet
    Set radnaSH = myWB.Worksheets("PORTOVI")

' ------------------------------
    Dim FinalRow As Long
    FinalRow = radnaSH.Cells(radnaSH.Rows.Count, 1).End(xlUp).Row

    Dim i As Long
    Dim s2PortName As String

    For i = FinalRow To 1 Step -1
        s2PortName = radnaSH.Cells(i, 2).Value

        If (s2PortName = "") Then
            radnaSH.Rows(i).EntireRow.Delete
        End If
    Next i

    Set radnaSH = Nothing
    Set myWB = Nothing

End Function

Function Fun_04_Find_and_insert_No_porta()
' mlabrkic, date: 2024_09M_25

' EDIT:

' ------------------------------
    Dim myWB As Workbook
    Set myWB = ThisWorkbook

    Dim radnaSH As Worksheet
    Set radnaSH = myWB.Worksheets("PORTOVI")

' ------------------------------
    Dim FinalRow As Long
    FinalRow = radnaSH.Cells(radnaSH.Rows.Count, 1).End(xlUp).Row

    Dim i As Long
    Dim s2PortName As String, s2PortNo As String

    radnaSH.Range("M1").Value = "No."  ' col 13

    For i = 2 To FinalRow Step 1
        s2PortName = radnaSH.Cells(i, 2).Value
        s2PortNo = Right(s2PortName, 2)

        If (Left(s2PortNo, 1) = "/") Then
            s2PortNo = Mid(s2PortNo, 2)
        End If
        radnaSH.Cells(i, 13).Value = s2PortNo
    Next i

    Set radnaSH = Nothing
    Set myWB = Nothing

End Function

Function Fun_05_Sort_Slot_Port()
' mlabrkic, date: 2024_09M_25

' EDIT:

' ------------------------------
'    Cells.Select
    ThisWorkbook.Worksheets("PORTOVI").Sort.SortFields.Clear

    Dim FinalRow As Long
    FinalRow = ThisWorkbook.Worksheets("PORTOVI").Cells(ThisWorkbook.Worksheets("PORTOVI").Rows.Count, 1).End(xlUp).Row

    ThisWorkbook.Worksheets("PORTOVI").Sort.SortFields.Add2 Key:=Range( _
        "A2:A" & FinalRow), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal

    ThisWorkbook.Worksheets("PORTOVI").Sort.SortFields.Add2 Key:=Range( _
        "M2:M" & FinalRow), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal

    With ThisWorkbook.Worksheets("PORTOVI").Sort
        .SetRange Range("A1:M" & FinalRow)
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Range("G2").Select

End Function

Function Fun_06_Delete_UI_and_ME_A()
' mlabrkic, date: 2024_09M_25

' EDIT:

' ------------------------------
    Dim myWB As Workbook
    Set myWB = ThisWorkbook

    Dim radnaSH As Worksheet
    Set radnaSH = myWB.Worksheets("PORTOVI")

' ------------------------------
    Dim FinalRow As Long
    FinalRow = radnaSH.Cells(radnaSH.Rows.Count, 1).End(xlUp).Row

    Dim i As Long
    Dim iPosUI As Integer, iPosME_ACC As Integer
    Dim s8Path As String, s8PathNo As String

    For i = 2 To FinalRow Step 1
        s8Path = radnaSH.Cells(i, 8).Value

'        https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/instr-function
'        vbTextCompare   1   Performs a textual comparison
        ' TODO: enter the correct PTH cat.
        iPosUI = InStr(1, s8Path, " - Aktivan - PTH_D_U", 1) ' Pozicija "." nakon 1 karaktera
        iPosME_ACC = InStr(1, s8Path, " - Aktivan - PTH_D_ME_A", 1)

        If (iPosUI > 0) Then
            s8PathNo = Mid(s8Path, 1, iPosUI - 1)
        ElseIf (iPosME_ACC > 0) Then
            s8PathNo = Mid(s8Path, 1, iPosME_ACC - 1)
        Else
            s8PathNo = s8Path
        End If

        radnaSH.Cells(i, 8).Value = s8PathNo
    Next i

    Set radnaSH = Nothing
    Set myWB = Nothing

End Function

Function Fun_07_Format_Iskljucen_Rezerviran()
' mlabrkic, date: 2024_09M_26

' EDIT:

' ------------------------------
    Dim myWB As Workbook
    Set myWB = ThisWorkbook

    Dim radnaSH As Worksheet
    Set radnaSH = myWB.Worksheets("PORTOVI")

' ------------------------------
    Dim FinalRow As Long
    FinalRow = radnaSH.Cells(radnaSH.Rows.Count, 1).End(xlUp).Row

    Dim i As Long
    Dim s3Status As String

    For i = FinalRow To 1 Step -1
        s3Status = radnaSH.Cells(i, 3).Value

        Set Rng = radnaSH.Range("A" & i & ":M" & i)
        If s3Status = "Isključen" Then
            ' Set Rng = radnaSH.Cells(i, 3)
            With Rng
                With .Font
    '                  .ColorIndex = 1   ' black
                    .ColorIndex = 3   ' red
    '                  .ColorIndex = 4   ' green
    '                  .ColorIndex = 5   ' blue
                    .Bold = True
                End With
            End With

            Set Rng = radnaSH.Cells(i, 14)
            With Rng
                With .Font
                    .ColorIndex = 3   ' red
                    .Bold = True
                End With
            End With
        ElseIf s3Status = "Rezerviran" Then
'            radnaSH.Cells(i, 3).Font.ColorIndex = 5   ' blue
'            radnaSH.Cells(i, 3).Font.Bold = True
'            radnaSH.Cells(i, 9).Font.ColorIndex = 5   ' blue
'            radnaSH.Cells(i, 9).Font.Bold = True
'            https://learn.microsoft.com/en-us/office/vba/api/excel.cellformat
            Rng.Interior.ColorIndex = 36 ' yellow
        End If
    Next i

    Set Rng = Nothing

    Set radnaSH = Nothing
    Set myWB = Nothing

End Function

Function Fun_08_Column_Width()
' mlabrkic, date: 2024_09M_26

' EDIT:

' ------------------------------
    Columns("A:C").Select
    Columns("A:C").EntireColumn.AutoFit

    Columns("D:D").ColumnWidth = 5.57
    Columns("E:F").EntireColumn.AutoFit

    Columns("G:G").ColumnWidth = 2.86
    Columns("H:H").ColumnWidth = 12 ' 2025_11M_11

    Columns("I:J").Select
    Selection.ColumnWidth = 5.57

    Columns("K:K").ColumnWidth = 55 ' 2025_11M_11

    Columns("L:L").ColumnWidth = 5

    Columns("M:M").ColumnWidth = 2.86
    Range("G2").Select

'    Columns("A:A").Select
'    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
'    Columns("A:A").ColumnWidth = 5.2
'    ThisWorkbook.Worksheets("PORTOVI").Range("A1").Value = "Slot 0"

End Function

Sub Main_01_fun()
' Sheet: PORTOVI

    Fun_01_Delete_empty_rows
    Fun_02_Copy_user_name_and_address
    Fun_03_Delete_rows_without_Port_Name

    Fun_04_Find_and_insert_No_porta
    Fun_05_Sort_Slot_Port
    Fun_06_Delete_UI_and_ME_A

    Fun_07_Format_Iskljucen_Rezerviran
    Fun_08_Column_Width

End Sub


Function Mark_new_10()
' mlabrkic, date: 2024_09M_26

' EDIT:

' ------------------------------
'https://learn.microsoft.com/en-us/office/vba/api/excel.borders

    Dim FinalRow As Long
    FinalRow = ActiveSheet.Cells(ActiveSheet.Rows.Count, 1).End(xlUp).Row

    Dim i As Long
    Dim sSORT As String, sSORT_0 As String

    For i = 2 To FinalRow
        sSORT = ActiveSheet.Range("M" & i).Value  ' 13
        sSORT_0 = Right(sSORT, 1)

        If (sSORT_0 = "0") Then
           ' Set Rng = ActiveSheet.Cells(i, j)
           Set Rng = ActiveSheet.Range("A" & i & ":T" & i)
           With Rng
'                With .Borders(xlEdgeTop)
                With .Borders(xlEdgeBottom)
                   .LineStyle = xlContinuous
'                   .Weight = xlThin
                    .Weight = xlThick
'                   .ColorIndex = 1   ' black
'                    .ColorIndex = 3   ' red
'                    .ColorIndex = 4   ' green
                    .ColorIndex = 5   ' blue
                End With
            End With
        End If

    Next i

    Set Rng = Nothing

    ActiveWorkbook.Save
End Function

Function No_02_Mark_new_Slot()
' mlabrkic, date: 2024_09M_26

' EDIT:

' ------------------------------
    ' Call Function:
    Mark_new_10

    Dim FinalRow As Long
    FinalRow = ActiveSheet.Cells(ActiveSheet.Rows.Count, 1).End(xlUp).Row

    Dim i As Long
    Dim s1Slot As String, s1SlotOLD As String

    s1SlotOLD = ActiveSheet.Cells(1, 1).Value

    ActiveSheet.Range("M1").Value = "No."
    ActiveSheet.Range("M1").Font.Bold = True

    For i = 2 To FinalRow
        s1Slot = ActiveSheet.Cells(i, 1).Value

        If s1Slot <> s1SlotOLD Then
           ' Set Rng = ActiveSheet.Cells(i, j)
           Set Rng = ActiveSheet.Range("A" & i & ":T" & i)
           With Rng
                With .Borders(xlEdgeTop)
                   .LineStyle = xlContinuous
                   ' .Weight = xlThin
                   .Weight = xlThick
                   ' .ColorIndex = 1   ' black
                    .ColorIndex = 3   ' red
                   ' .ColorIndex = 4   ' green
'                   .ColorIndex = 5   ' blue
                End With
            End With
        End If
        s1SlotOLD = s1Slot
    Next i

    Set Rng = Nothing

    ActiveWorkbook.Save
End Function


Sub No_02_Mark_groups_of_4_ports()
' mlabrkic, date: 2025_04M_23

' EDIT:

' ------------------------------
    Dim i As Long
    Dim j As Integer

    j = 0

    For i = 2 To 148
        j = j + 1
        If j = 4 Then
            j = 0
            ' Set Rng = ActiveSheet.Cells(i, j)
            Set Rng = ActiveSheet.Range("A" & i & ":O" & i)
            With Rng
                ' With .Borders(xlEdgeTop)
                With .Borders(xlEdgeBottom)
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
    Next i

    Set Rng = Nothing

    ActiveWorkbook.Save
End Sub


Sub No_03_Sakrij_stupce()
' mlabrkic, date: 2025_04M_23

' EDIT:

' ------------------------------
    Columns("G:G").Select
    Selection.EntireColumn.Hidden = True
    Columns("J:J").Select
    Selection.EntireColumn.Hidden = True
    Columns("L:M").Select
    Selection.EntireColumn.Hidden = True

    Columns("H:H").ColumnWidth = 12
    Columns("N:N").ColumnWidth = 13.29
    Columns("O:O").ColumnWidth = 65.71

    Range("N1").Value = "Rack"
    Range("O1").Value = "Rezervacija"

    Range("N1:O1").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 65535
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    With Selection.Font
        .Name = "Arial"
        .Size = 11
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ThemeColor = xlThemeColorLight1
        .TintAndShade = 0
        .ThemeFont = xlThemeFontNone
    End With
    Selection.Font.Bold = True

End Sub


Sub No_04_Format_Aktivan()
' mlabrkic, date: 2024-11M-05

' EDIT:
' date: 2025-02M-27 18:40:38

' ------------------------------
    Dim myWB As Workbook
    Set myWB = ThisWorkbook

    Dim radnaSH As Worksheet
    ' Set radnaSH = myWB.Worksheets("PORTOVI")
    Set radnaSH = myWB.ActiveSheet

' ------------------------------
    Dim FinalRow As Long
    FinalRow = radnaSH.Cells(radnaSH.Rows.Count, 1).End(xlUp).Row

    Dim i As Long
    Dim s3Status As String

    For i = 2 To FinalRow Step 1
        ' s3Status = radnaSH.Cells(i, 21).Value
        s3Status = Trim(radnaSH.Range("C" & i).Value) ' U, 21
        Set Rng = radnaSH.Range("C" & i & ":K" & i)

        If s3Status = "Aktivan" Then
            ' Rng.Font.ColorIndex = 4   ' green
            Rng.Font.Bold = True

            ' https://learn.microsoft.com/en-us/office/vba/api/excel.cellformat
            ' Rng.Interior.ColorIndex = 36 ' yellow
            ' Rng.Interior.ColorIndex = 4   ' green

            ' https://learn.microsoft.com/en-us/office/vba/api/excel.cellformat.interior
            With Rng.Interior
                .Pattern = xlSolid
                .PatternColorIndex = xlAutomatic
                ' .Color = 5296274
                ' .TintAndShade = 0
                .ThemeColor = xlThemeColorAccent6
                .TintAndShade = 0.599993896298105
                .PatternTintAndShade = 0
            End With

        End If
    Next i

    Set Rng = Nothing

    Set radnaSH = Nothing
    Set myWB = Nothing

End Sub


Sub No_05_Other()
' mlabrkic, date: 2025_11M_11

' EDIT:
' date: 2025_11M_11 22:03:52

' ------------------------------
    Dim myWB As Workbook
    Set myWB = ThisWorkbook

    Dim radnaSH As Worksheet
    ' Set radnaSH = myWB.Worksheets("PORTOVI")
    Set radnaSH = myWB.ActiveSheet

' ------------------------------
    Dim FinalRow As Long
    FinalRow = radnaSH.Cells(radnaSH.Rows.Count, 1).End(xlUp).Row

    Dim i As Long
    Dim s6Connector As String

    For i = 2 To FinalRow Step 1
        ' s6Connector = radnaSH.Cells(i, 21).Value
        s6Connector = Trim(radnaSH.Range("F" & i).Value) ' U, 21
        Set Rng = radnaSH.Range("F" & i)

        If s6Connector = "Other" Then
            ' https://learn.microsoft.com/en-us/office/vba/api/excel.font.colorindex
            ' Rng.Font.ColorIndex = 4   ' green
            Rng.Font.Bold = True

            Rng.Font.ColorIndex = 1   ' black
            ' Rng.Font.ColorIndex = 3   ' red
            ' Rng.Font.ColorIndex = 4   ' green
            ' Rng.Font.ColorIndex = 5   ' blue

            ' https://learn.microsoft.com/en-us/office/vba/api/excel.cellformat
            ' Rng.Interior.ColorIndex = 36 ' yellow
            ' Rng.Interior.ColorIndex = 4   ' green

            ' https://learn.microsoft.com/en-us/office/vba/api/excel.cellformat.interior
            Rng.Interior.ColorIndex = 36 ' yellow
            ' Rng.Interior.ColorIndex = 4   ' green

            ' With Rng.Interior
            '     .Pattern = xlSolid
            '     .PatternColorIndex = xlAutomatic
            '     ' .Color = 5296274
            '     ' .TintAndShade = 0
            '     .ThemeColor = xlThemeColorAccent6
            '     .TintAndShade = 0.599993896298105
            '     .PatternTintAndShade = 0
            ' End With

        End If
    Next i

    Set Rng = Nothing

    Set radnaSH = Nothing
    Set myWB = Nothing

End Sub
