Attribute VB_Name = "addCheckBoxesRight"
'**********************************************************************************************
'* addCheckBoxesRight
'* Created By: Tu Nguyen
'* Created On: 07-17-2018 03:30
'* Modified: 07-17-2018 21:35
'* Purpose: Creates checkboxes on the first empty column of a sheet
'*          when checked, the row will be filled black
'**********************************************************************************************
Option Explicit
Sub addCheckBoxesRight()
Attribute addCheckBoxesRight.VB_ProcData.VB_Invoke_Func = "r\n14"

Dim LastBlankCol As Long
Dim LastBlankRow As Long
Dim Start As String
Dim StartCol As String

Dim myCBX As CheckBox
Dim myCell As Range
    With ActiveSheet
    .CheckBoxes.Delete ' nice for testing
    
    LastBlankCol = Cells(1, Columns.Count).End(xlToLeft).Column
    LastBlankRow = Cells(Rows.Count, 1).End(xlUp).Row
    
    Start = Cells(1, LastBlankCol + 1).Address
    StartCol = Split(Cells(1, LastBlankCol + 1).Address(1, 0), "$")(0)
    
    'MsgBox "test last col: " & LastBlankCol & _
    '        "last row: " & LastBlankRow & _
    '        "test: " & StartCol
            
    For Each myCell In ActiveSheet.Range(StartCol & "1:" & StartCol & LastBlankRow).Cells
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
                                                                ' Example =$H1=TRUE

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

