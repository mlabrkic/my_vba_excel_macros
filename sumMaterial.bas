' date: 2026_07M_23
' ------------------------------------------------------------
' grupirati sve redove po vrijednosti iz kolone C
' zbrojiti vrijednosti iz kolone E
' kreirati (ili obrisati postojeći) sheet "OPC"
' upisati rezultat u sheet OPC

' Material (C)
' Material Description (D)
' Total Unrestricted (zbroj iz E)

' Debug.Print Unrestricted
' Unr  => C-Space
' Debug.Print "[" & wsSrc.Cells(i, "E").Text & "]"
' Debug.Print "[" & wsSrc.Cells(3, "E").Text & "]"

' ------------------------------------------------------------
' Da, preporučujem varijantu s dva Dictionary objekta.
' Jedan čuva opis materijala, a drugi zbroj.
'
' Kod je jednostavniji i pouzdaniji:

Sub SumMaterial()

    Dim wsSrc As Worksheet
    Dim wsOut As Worksheet
    Dim LastRow As Long
    Dim i As Long
    Dim RowOut As Long

    Dim Material As String
    Dim MatDesc As String
    Dim Unrestricted As Double

    Dim SumDict As Object
    Dim DescDict As Object
    Dim Key As Variant

    Set wsSrc = ActiveSheet

    Set SumDict = CreateObject("Scripting.Dictionary")
    Set DescDict = CreateObject("Scripting.Dictionary")

    LastRow = wsSrc.Cells(wsSrc.Rows.Count, "C").End(xlUp).Row

    ' INFO:
    ' ' text to numeric values
    ' With wsSrc.Range(wsSrc.Cells(1, 1), wsSrc.Cells(LastRow, 6))
    '     .Value = .Value
    ' End With

    ' Grupiranje i zbrajanje
    For i = 2 To LastRow

        Material = Trim(CStr(wsSrc.Cells(i, "C").Value))
        MatDesc = Trim(wsSrc.Cells(i, "D").Value)

        If IsNumeric(wsSrc.Cells(i, "E").Value) Then
            Unrestricted = CDbl(Trim(wsSrc.Cells(i, "E").Value))
        Else
            Unrestricted = 0
        End If

        If Material <> "" Then
            If SumDict.Exists(Material) Then
                SumDict(Material) = SumDict(Material) + Unrestricted
            Else
                SumDict.Add Material, Unrestricted
                DescDict.Add Material, MatDesc
            End If
        End If

    Next i

    ' Obriši postojeći OPC sheet
    Application.DisplayAlerts = False
    On Error Resume Next
    Worksheets("OPC").Delete
    On Error GoTo 0
    Application.DisplayAlerts = True

    ' Kreiraj novi OPC sheet
    Set wsOut = Worksheets.Add
    wsOut.Name = "OPC"

    ' Zaglavlja
    wsOut.Range("A1").Value = "Material"
    wsOut.Range("B1").Value = "Material Description"
    wsOut.Range("C1").Value = "Total Unrestricted"

    RowOut = 2

    ' Upis rezultata
    For Each Key In SumDict.Keys
        wsOut.Cells(RowOut, 1).Value = Key
        wsOut.Cells(RowOut, 2).Value = DescDict(Key)
        wsOut.Cells(RowOut, 3).Value = SumDict(Key)

        RowOut = RowOut + 1
    Next Key

    wsOut.Columns("A:C").AutoFit

    MsgBox "Rezultati su upisani u sheet 'OPC'.", vbInformation

End Sub
