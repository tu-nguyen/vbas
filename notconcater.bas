Attribute VB_Name = "notconcater"
'**********************************************************************************************
'* notconcater
'* Created By: Tu Nguyen
'* Created On: 07-17-2018 03:30
'* Modified: 07-17-2018 21:35
'* Purpose: Small tool to join 2 column of data for reports
'**********************************************************************************************
Sub notconcater()

Dim i As Integer
Dim j As Integer

Dim k As Integer
k = 1

Dim lastA As Integer
lastA = Range("A1").End(xlDown).Row
Dim lastB As Integer
lastB = Range("B1").End(xlDown).Row

For i = 1 To lastA
    For j = 1 To lastB
        Cells(k, "D").Value = Cells(i, "A").Value
        Cells(k, "E").Value = Cells(j, "B").Value
        k = k + 1
    Next j
Next i

End Sub
