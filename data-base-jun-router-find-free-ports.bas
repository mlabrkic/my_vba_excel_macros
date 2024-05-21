Sub No_01_Copy_user_address_K()
' Attribute No_01_xxx_W.VB_ProcData.VB_Invoke_Func = "k\n14"
' CTRL-K
' Description: WebX
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
    ' Dim s8UI As String

    For i = 1 To FinalRow Step 1
        s2PortName = radnaSH.Cells(i, 2).Value
        ' s8UI = radnaSH.Cells(i, 8).Value
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


' C:\02_PROG\10_Windows\05_VBA\PacktPublishing\vba_Excel_section-11-iteration_2023_05M_11.bas
Sub No_02_Delete_some_rows()
' mlabrkic, date: 2024-05M-20

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


Sub No_03_Copy_VLAN()
' mlabrkic, date: 2024-05M-20

' EDIT:
' date: 2024-05M-20 11:19:01

' ------------------------------
    Dim myWB As Workbook
    Set myWB = ThisWorkbook

    Dim radnaSH As Worksheet
    Set radnaSH = myWB.Worksheets("PORTOVI")

    ' ------------------------------
    Dim FinalRow As Long
    FinalRow = radnaSH.Cells(radnaSH.Rows.Count, 1).End(xlUp).Row

    Dim i As Long
    Dim j As Integer, iPointVLAN As Integer
    Dim i2PNameL As Integer, i2PNameOldL As Integer

    Dim s1SlotOLD As String, s2PortNameOLD As String
    Dim s1Slot As String, s2PortName As String

    Dim s2PortNameLeft As String, sVLAN As String

    j = 0

    s1SlotOLD = radnaSH.Cells(2, 1).Value
    s2PortNameOLD = radnaSH.Cells(2, 2).Value
    i2PNameOldL = Len(s2PortNameOLD)

    For i = 3 To FinalRow Step 1
        s1Slot = radnaSH.Cells(i, 1).Value
        s2PortName = radnaSH.Cells(i, 2).Value

        i2PNameL = Len(s2PortName)
        s2PortNameLeft = Left(s2PortName, i2PNameOldL)

        If s1Slot = s1SlotOLD Then
            If s2PortNameLeft = s2PortNameOLD Then
                j = j + 1
                iPointVLAN = InStr(1, s2PortName, ".", 1) ' Pozicija "." nakon 1 karaktera
                sVLAN = Mid(s2PortName, iPointVLAN + 1, i2PNameL - iPointVLAN)
                If (sVLAN = "16386") Or (sVLAN = "32767") Then
                Else
                    radnaSH.Cells(i - j, 13).Value = radnaSH.Cells(i - j, 13).Value + "," + sVLAN
                End If
                radnaSH.Cells(i, 13).Value = "vlan"
                s2PortNameOLD = s2PortNameLeft
            Else
                j = 0
                s2PortNameOLD = s2PortName
                i2PNameOldL = Len(s2PortNameOLD)
            End If
        Else
            j = 0
            s1SlotOLD = s1Slot
            s2PortNameOLD = s2PortName
            i2PNameOldL = Len(s2PortNameOLD)
        End If
    Next i

    Set radnaSH = Nothing
    Set myWB = Nothing

End Sub


Sub No_04_Delete_rows_VLAN()
' mlabrkic, date: 2024-05M-20

' EDIT:
' date: 2024-05M-20 13:13:08

' ------------------------------
    Dim myWB As Workbook
    Set myWB = ThisWorkbook

    Dim radnaSH As Worksheet
    Set radnaSH = myWB.Worksheets("PORTOVI")

    ' radnaSH.Range(radnaSH.Cells(1, 2), radnaSH.Cells(2, 3)).Copy
    ' radnaSH.Cells(1, 1).Value = Now()

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


Sub No_05_Format_cells()
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


