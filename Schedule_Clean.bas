Attribute VB_Name = "Module2"
Public Sub DeleteBlankRows()
    Dim SourceRange As Range
    Dim EntireRow As Range
 
        For I = Cells.Rows.Count To 1 Step -1
            Set EntireRow = Cells(I, 1).EntireRow
            If Application.WorksheetFunction.CountA(EntireRow) = 0 Then
                EntireRow.Delete
            End If
        Next
 
        Application.ScreenUpdating = True
End Sub

Public Sub Schedule_Clean()
Attribute Schedule_Clean.VB_Description = "Cleans Up Schedule to pull data"
Attribute Schedule_Clean.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Schedule_Clean Macro
' Cleans Up Schedule to pull data
'

'
Dim last_row As Long
Dim last_row2 As Long
Dim last_row3 As Long
Dim last_col As Long
Dim EntireRow As Range
Dim ColumnLetter As String


last_row = Cells(Rows.Count, 1).End(xlUp).Row

    For I = last_row To 1 Step -1
        Set EntireRow = Cells(I, 1).EntireRow
        If Application.WorksheetFunction.CountA(EntireRow) = 0 Then
            EntireRow.Delete
        End If
    Next
    
    Application.ScreenUpdating = True
    
last_row2 = Cells(Rows.Count, 1).End(xlUp).Row

    For I = last_row2 To 1 Step -1
    Set EntireRow = Cells(I, 1).EntireRow
    If IsNumeric(Cells(I, 1).Value) Then
    Else
        EntireRow.Delete
    End If
    
    Next
    
last_row3 = Cells(Rows.Count, 1).End(xlUp).Row
    
    For I = last_row3 To 1 Step -1
    last_col = Cells(I, Columns.Count).End(xlToLeft).Column
    ColumnLetter = Split(Cells(1, (last_col - 2)).Address, "$")(1)
    If last_col > 4 Then
        Range("C" & I & ":" & ColumnLetter & I).Delete Shift:=xlToLeft
    Else
    End If
    Next

End Sub
