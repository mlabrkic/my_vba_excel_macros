' date: 2024-12M-12 18:54:09

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

Function Fun_06_Delete_UI_and_ME_ACCESS()
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
        iPosUI = InStr(1, s8Path, " - Aktivan - PTH_DATA_UI", 1) ' Pozicija "." nakon 1 karaktera
        iPosME_ACC = InStr(1, s8Path, " - Aktivan - PTH_DATA_ME_ACCESS", 1)

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
    Columns("H:H").ColumnWidth = 58.57

    Columns("I:J").Select
    Selection.ColumnWidth = 5.57

    Columns("K:K").ColumnWidth = 35

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
    Fun_06_Delete_UI_and_ME_ACCESS

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

Function Fun2_00_Mark_new_Slot()
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

Function Fun2_01_Copy_from_INT_DESCRIPTION()
' mlabrkic, date: 2024_09M_27

' EDIT:

' ------------------------------
    Dim myWB As Workbook
    Set myWB = ThisWorkbook

    Dim portSH As Worksheet, descSH As Worksheet
    Set portSH = myWB.Worksheets("PORTOVI")
    Set descSH = myWB.Worksheets("INT_DESCRIPTION")

' ------------------------------
    Dim finalRowPORT As Integer, finalRowDESC As Integer

    finalRowPORT = portSH.Cells(portSH.Rows.Count, 1).End(xlUp).Row
    finalRowDESC = descSH.Cells(descSH.Rows.Count, 1).End(xlUp).Row

    Dim i As Integer, j As Integer
    Dim sPrva As String, sPrvaRight As String, sDruga As String

    portSH.Range("O1").Value = "DESCRIPTION"

    For i = 2 To finalRowPORT
        sPrva = portSH.Range("B" & i).Value
        sPrvaRight = Mid(sPrva, 9)
        sPrva = "Eth" + sPrvaRight

        For j = 2 To finalRowDESC
            sDruga = Trim(descSH.Range("A" & j).Value)
            If sDruga = sPrva Then   '  Vrijednost iz portSH našao je u descSH
                portSH.Range("O" & i) = descSH.Range("D" & j)
                Exit For
            End If
        Next j
    Next i

    Set portSH = Nothing
    Set descSH = Nothing

    Set myWB = Nothing

End Function

Function Fun2_02_Copy_from_INT_STATUS()
' mlabrkic, date: 2024_09M_27

' EDIT:

' ------------------------------
    Dim myWB As Workbook
    Set myWB = ThisWorkbook

    Dim portSH As Worksheet, statusSH As Worksheet
    Set portSH = myWB.Worksheets("PORTOVI")
    Set statusSH = myWB.Worksheets("INT_STATUS")

' ------------------------------
    Dim finalRowPORT As Integer, finalRowSTATUS As Integer

    finalRowPORT = portSH.Cells(portSH.Rows.Count, 1).End(xlUp).Row
    finalRowSTATUS = statusSH.Cells(statusSH.Rows.Count, 1).End(xlUp).Row

    Dim i As Integer, j As Integer
    Dim sPrva As String, sPrvaRight As String, sDruga As String

    portSH.Range("P1").Value = "Status"
    portSH.Range("Q1").Value = "Vlan"
    portSH.Range("R1").Value = "Duplex"
    portSH.Range("S1").Value = "Speed"
    portSH.Range("T1").Value = "Type"

    For i = 2 To finalRowPORT
        sPrva = portSH.Range("B" & i).Value  ' Port Name
        sPrvaRight = Mid(sPrva, 9)
        sPrva = "Eth" + sPrvaRight

        For j = 2 To finalRowSTATUS
            sDruga = Trim(statusSH.Range("A" & j).Value)
            If sDruga = sPrva Then   '  Vrijednost iz portSH našao je u statusSH
                portSH.Range("P" & i) = statusSH.Range("C" & j) ' Status
                portSH.Range("Q" & i) = statusSH.Range("D" & j) ' Vlan
                portSH.Range("R" & i) = statusSH.Range("E" & j) ' Duplex
                portSH.Range("S" & i) = statusSH.Range("F" & j) ' Speed
                portSH.Range("T" & i) = statusSH.Range("G" & j) ' Type
                Exit For
            End If
        Next j
    Next i

    Set portSH = Nothing
    Set statusSH = Nothing

    Set myWB = Nothing

End Function

Function Fun2_03_Format_Iskljucen_Rezerviran()
' mlabrkic, date: 2024_09M_27

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
    Dim s16Status As String

    radnaSH.Range("U1").Value = "SORT"

    For i = 2 To FinalRow Step 1
        ' s16Status = radnaSH.Cells(i, 16).Value
        s16Status = Trim(radnaSH.Range("P" & i).Value) ' P, 16
        Set Rng = radnaSH.Range("P" & i & ":U" & i)

        If s16Status = "disabled" Then
            radnaSH.Range("U" & i).Value = "2"
            With Rng
                With .Font
                    ' https://learn.microsoft.com/en-us/office/vba/api/excel.font.colorindex
    '                  .ColorIndex = 1   ' black
                    .ColorIndex = 3   ' red
    '                  .ColorIndex = 4   ' green
    '                  .ColorIndex = 5   ' blue
                    .Bold = True
                End With
            End With
        ElseIf s16Status = "notconnec" Then
            radnaSH.Range("U" & i).Value = "3"
            Rng.Font.ColorIndex = 5   ' blue
            Rng.Font.Bold = True
        ElseIf s16Status = "sfpAbsent" Then  ' odsutan
            radnaSH.Range("U" & i).Value = "1"
            Rng.Font.ColorIndex = 4   ' green
            Rng.Font.Bold = True
        End If
    Next i

    Set Rng = Nothing

    Set radnaSH = Nothing
    Set myWB = Nothing

End Function

Sub Main_02_fun()
' Sheet: PORTOVI

    Fun2_00_Mark_new_Slot

    Fun2_01_Copy_from_INT_DESCRIPTION
    Fun2_02_Copy_from_INT_STATUS
    Fun2_03_Format_Iskljucen_Rezerviran

End Sub


Sub Format_01_Mark_new_Slot()
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
End Sub


Sub Format_02_Format_Aktivan()
' mlabrkic, date: 2024-11M-05

' EDIT:
' date: 2024-11M-05 13:53:10

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
            Rng.Interior.ColorIndex = 4   ' green

            ' With Rng.Interior
            '     .Pattern = xlSolid
            '     .PatternColorIndex = xlAutomatic
            '     .Color = 5296274
            '     .TintAndShade = 0
            '     .PatternTintAndShade = 0
            ' End With

        End If
    Next i

    Set Rng = Nothing

    Set radnaSH = Nothing
    Set myWB = Nothing

End Sub




Sub No_01_KOR_I_USL_Find_and_insert_No_porta()
' mlabrkic, date: 2024_09M_25

' EDIT:
' date: 2024-10M-28 19:07:00
' date: 2024-10M-28 22:01:45

' ------------------------------
    Dim myWB As Workbook
    Set myWB = ThisWorkbook

    Dim radnaSH As Worksheet
    Set radnaSH = myWB.ActiveSheet

' ------------------------------
    Dim FinalRow As Long
    FinalRow = radnaSH.Cells(radnaSH.Rows.Count, 1).End(xlUp).Row

    Dim i As Long
    Dim s3PortName As String, s1PortNo As String

    radnaSH.Range("A1").Value = "No."  ' col 1

    For i = 2 To FinalRow Step 1
        s3PortName = radnaSH.Cells(i, 3).Value
        s1PortNo = Right(s3PortName, 2)
        ' s1PortNo = Mid(s3PortName, 11)

        If (Left(s1PortNo, 1) = "/") Then
            s1PortNo = Mid(s1PortNo, 2)
        End If
        radnaSH.Cells(i, 1).Value = s1PortNo
'        radnaSH.Range("H" & i).Value = radnaSH.Range("V" & i).Value ' Data Access Id
'        radnaSH.Range("K" & i).Value = radnaSH.Range("X" & i).Value ' Vlan Id
    Next i

    Set radnaSH = Nothing
    Set myWB = Nothing

End Sub

Sub No_02_KOR_I_USL_Delete_OTHERS_rows()
' mlabrkic, date: 2024_10M_29

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
    Set radnaSH = myWB.ActiveSheet

' ------------------------------
    Dim FinalRow As Long
    FinalRow = radnaSH.Cells(radnaSH.Rows.Count, 1).End(xlUp).Row

    Dim i As Long
    Dim sME_ACCESS As String

    For i = FinalRow To 2 Step -1
        sME_ACCESS = radnaSH.Range("H" & i).Value ' Type

        If (sME_ACCESS <> "PTH_DATA_ME_ACCESS") Then
            radnaSH.Rows(i).EntireRow.Delete
        End If
    Next i

    Set radnaSH = Nothing
    Set myWB = Nothing

End Sub

Function Fun_No_03_Sort_SlotPort()
' date: 2024-10M-30

' Edit:
' date: 2024-10M-30 14:54:39

' ------------------------------
    Dim myWB As Workbook
    Set myWB = ThisWorkbook

    Dim radnaSH As Worksheet
    Set radnaSH = myWB.ActiveSheet

' ------------------------------
    Cells.Select
    radnaSH.Sort.SortFields.Clear

    Dim FinalRow As Long
    FinalRow = radnaSH.Cells(radnaSH.Rows.Count, 1).End(xlUp).Row

    radnaSH.Sort.SortFields.Add2 Key:=Range( _
        "B2:B" & FinalRow), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal

    radnaSH.Sort.SortFields.Add2 Key:=Range( _
        "A2:A" & FinalRow), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal

    With radnaSH.Sort
        .SetRange Range("A1:G" & FinalRow)
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With

    Range("D2").Select

    Set radnaSH = Nothing
    Set myWB = Nothing

End Function

Sub No_03_KOR_I_USL_Delete_SOME_columns()
'
' Macro1 Macro
'
' date: 2024-10M-29 15:27:44

    Columns("W:AA").Select
    Selection.Delete Shift:=xlToLeft

    Columns("N:U").Select
    Selection.Delete Shift:=xlToLeft

    Columns("J:L").Select
    Selection.Delete Shift:=xlToLeft

    Columns("H:H").Select
    Selection.Delete Shift:=xlToLeft

    Columns("D:F").Select
    Selection.Delete Shift:=xlToLeft

    Range("D2").Select
    ActiveWorkbook.Save

    ' Call Function:
    Fun_No_03_Sort_SlotPort

End Sub

Sub No_04_KOR_I_USL_Copy_Data_Access_Id()
' mlabrkic, date: 2024_10M_29

' EDIT:
' date: 2024-10M-29 18:56:50

' ------------------------------
    Dim myWB As Workbook

    Set myWB = ThisWorkbook
    ' ThisWorkbook refers to the workbook containing the running VBA code

    Dim radnaSH As Worksheet
    Set radnaSH = myWB.ActiveSheet

' ------------------------------
    Dim FinalRow As Long
    FinalRow = radnaSH.Cells(radnaSH.Rows.Count, 1).End(xlUp).Row

    Dim i As Long
    Dim j As Integer
    Dim s3PortNameOld As String, s3PortNameNew As String

    j = 0

    ' radnaSH.Range("K" & 1).Value = "KORISNIK"
    radnaSH.Range("K1").Value = "Path Bandwidth"
    radnaSH.Range("O1").Value = "DIS ID PATHA"

    ' radnaSH.Range("G2").Value = radnaSH.Range("G2").Value ' 7, G
    radnaSH.Range("K2").Value = radnaSH.Range("E2").Value ' 11, K
    radnaSH.Range("O2").Value = radnaSH.Range("D2").Value ' 15, O

    s3PortNameOld = radnaSH.Range("C2").Value ' col 3

    For i = 3 To FinalRow Step 1
        ' radnaSH.Range("T" & i).Value = radnaSH.Range("N" & i).Value
        s3PortNameNew = radnaSH.Range("C" & i).Value ' col 3
        If (s3PortNameNew = s3PortNameOld) Then
            j = j + 1
            radnaSH.Cells(i - j, 7 + j).Value = radnaSH.Range("G" & i).Value ' 7, G
            radnaSH.Cells(i - j, 11 + j).Value = radnaSH.Range("E" & i).Value ' 11, K
            radnaSH.Cells(i - j, 15 + j).Value = radnaSH.Range("D" & i).Value ' 15, O
        Else
            radnaSH.Cells(i, 11).Value = radnaSH.Range("E" & i).Value ' 11, K
            radnaSH.Cells(i, 15).Value = radnaSH.Range("D" & i).Value ' 15, O
            s3PortNameOld = s3PortNameNew
            j = 0
        End If
    Next i

    Set radnaSH = Nothing
    Set myWB = Nothing

End Sub

Sub No_05_KOR_I_USL_Delete_rows_ME_ACCESS()
' mlabrkic, date: 2024_10M_29

' EDIT:
' date: 2024-10M-29 19:56:30

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
    Set radnaSH = myWB.ActiveSheet

' ------------------------------
    Dim FinalRow As Long
    FinalRow = radnaSH.Cells(radnaSH.Rows.Count, 1).End(xlUp).Row

    Dim i As Long
    Dim sME_ACCESS As String

    For i = FinalRow To 2 Step -1
        sME_ACCESS = Trim(radnaSH.Range("O" & i).Value) ' DIS ID PATHA

        If (sME_ACCESS = "") Then
            radnaSH.Rows(i).EntireRow.Delete
        End If
    Next i

    Set radnaSH = Nothing
    Set myWB = Nothing

End Sub

Sub No2_06_Copy_from_KOR_i_USL()
' Copy from: KOR_i_USL
' mlabrkic, date: 2024-10M-24

' EDIT:
' date: 2024-10M-29 20:06:35

' ------------------------------
    Dim myWB As Workbook
    Set myWB = ThisWorkbook

    Dim portSH As Worksheet, statusSH As Worksheet
    Set portSH = myWB.Worksheets("PORTOVI")
    Set statusSH = myWB.Worksheets("KOR_I_USL")

' ------------------------------
    Dim finalRowPORT As Integer, finalRowSTATUS As Integer

    finalRowPORT = portSH.Cells(portSH.Rows.Count, 1).End(xlUp).Row
    finalRowSTATUS = statusSH.Cells(statusSH.Rows.Count, 1).End(xlUp).Row

    Dim i As Integer, j As Integer
    Dim sPrvaPortName As String, sPrvaPortNameRight As String, sDrugaPortName As String

    For j = 2 To finalRowSTATUS
        sDrugaPortName = Trim(statusSH.Range("C" & j).Value)  ' Port Hum Id

        For i = 2 To finalRowPORT
            sPrvaPortName = Trim(portSH.Range("B" & i).Value)  ' Port Hum Id
            If sDrugaPortName = sPrvaPortName Then   '  Vrijednost iz portSH našao je u statusSH
                portSH.Range("AH" & i & ":AS" & i).Value = statusSH.Range("G" & j & ":R" & j).Value ' 12
                Exit For
            End If
        Next i
    Next j

    Set portSH = Nothing
    Set statusSH = Nothing

    Set myWB = Nothing

End Sub

Sub No2_07_Copy_ME_ACCESS()
' mlabrkic, date: 2024-10M-24

' EDIT:
' date: 2024-11M-13 16:03:59

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
    Dim s8Path As String, s8ME As String
    Dim sAPpath As String

    For i = 2 To FinalRow Step 1
        s8Path = Trim(radnaSH.Range("H" & i).Value)
        ' s8ME = Left(s8Path, 10) ' ME_ACCESS_
        s8ME = Left(s8Path, 3) ' ME_

        sAPpath = radnaSH.Range("AP" & i).Value

        ' If (s8ME = "ME_ACCESS_") Then
        If (s8ME = "ME_") And (sAPpath = "") Then
            radnaSH.Range("AP" & i).Value = s8Path
            ' radnaSH.Range("AT" & i).Value = s8Path
        End If
    Next i

    Set radnaSH = Nothing
    Set myWB = Nothing

End Sub

Sub No2_08_Copy_KOR_from_KOR_i_USL()
' Copy from: KOR_i_USL
' mlabrkic, date: 2024-10M-19

' EDIT:
' date: 2024-11M-19 09:42:09

' ------------------------------
    Dim myWB As Workbook
    Set myWB = ThisWorkbook

    Dim portSH As Worksheet, statusSH As Worksheet
    Set portSH = myWB.Worksheets("PORTOVI")
    Set statusSH = myWB.Worksheets("KOR_I_USL")

' ------------------------------
    Dim finalRowPORT As Integer, finalRowSTATUS As Integer

    finalRowPORT = portSH.Cells(portSH.Rows.Count, 1).End(xlUp).Row
    finalRowSTATUS = statusSH.Cells(statusSH.Rows.Count, 1).End(xlUp).Row

    Dim i As Integer, j As Integer
    Dim sPrvaPortName As String, sPrvaPortNameRight As String, sDrugaPortName As String

    For j = 2 To finalRowSTATUS
        sDrugaPortName = Trim(statusSH.Range("C" & j).Value)  ' Port Hum Id

        For i = 2 To finalRowPORT
            sPrvaPortName = Trim(portSH.Range("C" & i).Value)  ' Port Name
            If sDrugaPortName = sPrvaPortName Then   '  Vrijednost iz portSH našao je u statusSH
                ' portSH.Range("AH" & i & ":AS" & i).Value = statusSH.Range("G" & j & ":R" & j).Value ' 12
                portSH.Range("H" & i).Value = statusSH.Range("F" & j).Value
                Exit For
            End If
        Next i
    Next j

    Set portSH = Nothing
    Set statusSH = Nothing

    Set myWB = Nothing

End Sub

