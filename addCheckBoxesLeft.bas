Attribute VB_Name = "addCheckBoxesLeft"
'**********************************************************************************************
'* addCheckBoxesLeft
'* Created By: Tu Nguyen
'* Created On: 07-17-2018 21:30
'* Modified: 07-17-2018 21:43
'* Purpose: Creates new column on the lefthand side then
'*          fills the columns with checkboxes
'*          when checked, the row will be filled black
'**********************************************************************************************
Option Explicit
Sub addCheckBoxesLeft()
Attribute addCheckBoxesLeft.VB_ProcData.VB_Invoke_Func = "l\n14"

Dim LastBlankRow As Long

Dim myCBX As CheckBox
Dim myCell As Range
    With ActiveSheet
    .CheckBoxes.Delete ' nice for testing

    LastBlankRow = Cells(Rows.Count, 1).End(xlUp).Row
    
    ' Shifts AFTER last row number is assigned
    Columns("A:A").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    
    'MsgBox "last row: " & LastBlankRow & _
    '        "test: "
            
    For Each myCell In ActiveSheet.Range("A1:A" & LastBlankRow).Cells
        With myCell
            ' Positions the checkbox nicely within the cell
            Set myCBX = .Parent.CheckBoxes.Add _
                            (Top:=.Top, Width:=.Width, _
                            Left:=.Left, Height:=.Height)
            ' links the check box with the cell it's on top off
            With myCBX
                .LinkedCell = myCell.Address
                .Caption = "" 'Or whatever you want

                End With
                .NumberFormat = ";;;"
            End With
            
            ' Selects the entire row that the checkbox is on
            Rows(myCell.Row).Select
            ' Creates the condition and the effects
            Selection.FormatConditions.Add Type:=xlExpression, Formula1:="=" & myCell.Address & "=TRUE"
            Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
            With Selection.FormatConditions(1).Interior
                .PatternColorIndex = xlAutomatic
                .ThemeColor = xlThemeColorLight1
                .TintAndShade = 0
            End With
            Selection.FormatConditions(1).StopIfTrue = False
            
        Next myCell
    End With
End Sub

