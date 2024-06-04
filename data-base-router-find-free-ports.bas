' On this way...
'
' data base, Web browser
' router, all ports view, open
' Select all ports, Copy
'
' Paste to "ports.csv"
' Import "ports.csv" to Excel (tab delimiter).
' ------------------------------------------------------------

Sub No_00_Main()
    No_01_Copy_user_address_K
    No_02_Sort_Slot_Port
    No_03_Delete_some_rows
    No_04_Port_Name
    No_05_Delete_rows_VLAN
    No_06_Format_cells
    No_07_ColumnWidth

End Sub


Sub No_01_Copy_user_address_K()
' Attribute No_01_xxx_W.VB_ProcData.VB_Invoke_Func = "k\n14"
' CTRL-K
' mlabrkic, date: 2023-11M-12

' EDIT:
' date: 2024-05M-20 10:20:44

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
    ' ThisWorkbook refers to the workbook containing the running VBA code

    Dim radnaSH As Worksheet
    Set radnaSH = myWB.Worksheets("PORTOVI")

' ------------------------------
    Dim FinalRow As Long
    FinalRow = radnaSH.Cells(radnaSH.Rows.Count, 1).End(xlUp).Row

    Dim i As Long
    Dim j As Integer
    Dim s2PortName As String

    For i = 1 To FinalRow Step 1
        s2PortName = radnaSH.Cells(i, 2).Value
        If s2PortName = "" Then
            j = j + 1
            radnaSH.Cells(i - j, 10 + j).Value = radnaSH.Cells(i, 1).Value
        Else
            j = 0
        End If
    Next i

    Set radnaSH = Nothing
    Set myWB = Nothing

End Sub


Sub No_02_Sort_Slot_Port()
' mlabrkic, date: 2024-06M-04

' EDIT:

'    Cells.Select
'    ActiveWorkbook.Worksheets("PORTOVI").Sort.SortFields.Clear

    Dim FinalRow As Long
    FinalRow = ThisWorkbook.Worksheets("PORTOVI").Cells(ThisWorkbook.Worksheets("PORTOVI").Rows.Count, 1).End(xlUp).Row

    ThisWorkbook.Worksheets("PORTOVI").Sort.SortFields.Add2 Key:=Range( _
        "A2:A" & FinalRow), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal

    ThisWorkbook.Worksheets("PORTOVI").Sort.SortFields.Add2 Key:=Range( _
        "B2:B" & FinalRow), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal

    With ThisWorkbook.Worksheets("PORTOVI").Sort
        .SetRange Range("A1:N" & FinalRow)
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Range("L2").Select

End Sub


Sub No_03_Delete_some_rows()
' mlabrkic, date: 2024-05M-20

' C:\02_PROG\10_Windows\05_VBA\PacktPublishing\vba_Excel_section-11-iteration_2023_05M_11.bas
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
    Dim s1Slot As String, s2PortName As String

    For i = FinalRow To 1 Step -1
        s1Slot = radnaSH.Cells(i, 1).Value
        s2PortName = radnaSH.Cells(i, 2).Value

        If s2PortName = "" Then
            radnaSH.Cells(i, 2).EntireRow.Delete
        ElseIf (s1Slot = "-1") Or (s1Slot = "--") Then
            radnaSH.Cells(i, 2).EntireRow.Delete
        End If
    Next i

    Set radnaSH = Nothing
    Set myWB = Nothing

End Sub


Sub No_04_Port_Name()
' mlabrkic, date: 2024-05M-20

' EDIT:
' date: 2024-06M-03
' ------------------------------
    Dim myWB As Workbook
    Set myWB = ThisWorkbook

    Dim radnaSH As Worksheet
    Set radnaSH = myWB.Worksheets("PORTOVI")

' ------------------------------
    Dim FinalRow As Long
    FinalRow = radnaSH.Cells(radnaSH.Rows.Count, 1).End(xlUp).Row

    Dim i As Long
    Dim iPoint As Integer, j As Integer

    Dim s2Temp As String, s2TempOld As String
    Dim sPortName As String, sPortNameOld As String

    Dim i2TempLen As Integer, i2TempOldLen As Integer
    Dim sVLAN As String

    j = 0

    s2TempOld = radnaSH.Cells(2, 2).Value
    iPoint = InStr(1, s2TempOld, ".", 1) ' Pozicija "." nakon 1 karaktera

    If iPoint > 0 Then
        sPortNameOld = Left(s2TempOld, iPoint - 1)
    Else
        sPortNameOld = s2TempOld
    End If
    i2TempOldLen = Len(s2TempOld)

    For i = 3 To FinalRow Step 1
        s2Temp = radnaSH.Cells(i, 2).Value
        iPoint = InStr(1, s2Temp, ".", 1) ' Pozicija "." nakon 1 karaktera

        If iPoint > 0 Then
            sPortName = Left(s2Temp, iPoint - 1)
        Else
            sPortName = s2Temp
        End If
        i2TempLen = Len(s2Temp)
'        radnaSH.Cells(i, 17).Value = s2PortName

        If sPortName = sPortNameOld Then
            j = j + 1
            sVLAN = Mid(s2Temp, iPoint + 1, i2TempLen - iPoint)
            If (sVLAN = "16386") Or (sVLAN = "32767") Then
            Else
                If j = 1 Then
                    radnaSH.Cells(i - j, 13).NumberFormat = "@" ' Text
                    radnaSH.Cells(i - j, 13).Value = sVLAN
                Else
                    radnaSH.Cells(i - j, 13).Value = radnaSH.Cells(i - j, 13).Value + "," + sVLAN
                End If
            End If
            radnaSH.Cells(i, 13).Value = "vlan"
        Else
            j = 0
            sPortNameOld = sPortName
        End If

    Next i

    Set radnaSH = Nothing
    Set myWB = Nothing

End Sub


Sub No_05_Delete_rows_VLAN()
' mlabrkic, date: 2024-05M-20

' EDIT:
' date: 2024-05M-20 13:13:08

' ------------------------------
    Dim myWB As Workbook
    Set myWB = ThisWorkbook

    Dim radnaSH As Worksheet
    Set radnaSH = myWB.Worksheets("PORTOVI")

' ------------------------------
    Dim FinalRow As Long
    FinalRow = radnaSH.Cells(radnaSH.Rows.Count, 1).End(xlUp).Row

    Dim i As Long
    Dim s3Status As String, s13Mvlan As String

    For i = FinalRow To 1 Step -1
        s3Status = radnaSH.Cells(i, 3).Value
        s13Mvlan = radnaSH.Cells(i, 13).Value

        If s13Mvlan = "vlan" Then
            radnaSH.Cells(i, 13).EntireRow.Delete
        ElseIf s13Mvlan = "" Then
            radnaSH.Cells(i, 14).Value = "NEMA VLAN"
        End If
    Next i

    Set radnaSH = Nothing
    Set myWB = Nothing

End Sub


Sub No_06_Format_cells()
' mlabrkic, date: 2024-05M-20

' EDIT:
' date: 2024-05M-20 13:13:08

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

        If s3Status = "Isključen" Then
            ' Set Rng = radnaSH.Cells(i, 3)
            Set Rng = radnaSH.Range("A" & i & ":C" & i)
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
                    .ColorIndex = 3   ' Red
                    .Bold = True
                End With
            End With
        ElseIf s3Status = "Rezerviran" Then
            radnaSH.Cells(i, 3).Font.ColorIndex = 5   ' blue
            radnaSH.Cells(i, 3).Font.Bold = True

            radnaSH.Cells(i, 9).Font.ColorIndex = 5   ' blue
            radnaSH.Cells(i, 9).Font.Bold = True
        End If
    Next i

    Set Rng = Nothing

    Set radnaSH = Nothing
    Set myWB = Nothing

End Sub


Sub No_07_ColumnWidth()
' mlabrkic, date: 2024-06M-04
' EDIT:

    Columns("B:B").ColumnWidth = 9.57
    Columns("C:C").ColumnWidth = 9.43
    Columns("D:D").ColumnWidth = 5.86
    Columns("E:E").ColumnWidth = 31
    Columns("F:F").ColumnWidth = 12
    Columns("G:G").ColumnWidth = 10.43
    Columns("H:H").ColumnWidth = 21.14
    Columns("I:I").ColumnWidth = 14
    Columns("J:J").ColumnWidth = 5.86
    Columns("K:K").ColumnWidth = 27.86
    Columns("L:L").ColumnWidth = 38.14
    Columns("M:M").ColumnWidth = 25.29
    Columns("N:N").ColumnWidth = 11.86
    ActiveWorkbook.Save
End Sub


Sub ZZ_Port_Name()
' mlabrkic, date: 2024-06M-03

' EDIT:
' date:
' ------------------------------
    Dim myWB As Workbook
    Set myWB = ThisWorkbook

    Dim radnaSH As Worksheet
    Set radnaSH = myWB.Worksheets("PORTOVI")

' ------------------------------
    Dim FinalRow As Long
    FinalRow = radnaSH.Cells(radnaSH.Rows.Count, 1).End(xlUp).Row

    Dim i As Long
    Dim iPoint As Integer
    Dim s2Temp As String, s2PortName As String

    For i = 1 To FinalRow Step 1
        s2Temp = radnaSH.Cells(i, 2).Value
        iPoint = InStr(1, s2Temp, ".", 1) ' Pozicija "." nakon 1 karaktera

        If iPoint > 0 Then
            s2PortName = Left(s2Temp, iPoint - 1)
        Else
            s2PortName = s2Temp
        End If
        radnaSH.Cells(i, 17).Value = s2PortName
    Next i

    Set radnaSH = Nothing
    Set myWB = Nothing

End Sub

