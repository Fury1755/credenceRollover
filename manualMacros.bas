Attribute VB_Name = "manualMacros"
Sub SelectUsedCellsInColumn()
Attribute SelectUsedCellsInColumn.VB_ProcData.VB_Invoke_Func = "S\n14"

    Dim myRange As Range
    Dim cell As Range
    Dim targetCell As Range
    Dim targetOffset As Range
    Dim searchArea As Range
    
    Set searchArea = Intersect(ActiveSheet.usedRange, Selection.EntireColumn)
    Dim ceiling As Range
    
    For Each cell In searchArea
        If InStr(1, cell, "2023") > 0 Or InStr(1, cell, "2024") > 0 Or InStr(1, cell, "2025") > 0 Then
            Set ceiling = cell
            Exit For
        End If
    Next cell

    
    Dim floor As Range
    If Not ceiling Is Nothing Then
    Set floor = Cells(Rows.Count, ceiling.column).End(xlUp) 'Rows.count is the very bottom row. End(xlUp is equivalent to ctrl + Up arrow
    Set myRange = Range(floor, ceiling)
    myRange.Select
     
     End If

End Sub

Sub FormatAsAccounting()
Attribute FormatAsAccounting.VB_ProcData.VB_Invoke_Func = "F\n14"
    Selection.NumberFormat = "_(* #,##0_);_(* (#,##0);_(* ""-""??_);_(@_)"
End Sub

Sub FormatAsGeneral()
Attribute FormatAsGeneral.VB_ProcData.VB_Invoke_Func = "G\n14"
    Selection.NumberFormat = "General"
End Sub

Sub FormatAsAccountingWithZeros()
    Selection.NumberFormat = "_(* #,##0.00_);_(* (#,##0.00);_(* ""-""??_);_(@_)"
End Sub
Sub MoveFormulasFromLeftToRight()
Attribute MoveFormulasFromLeftToRight.VB_ProcData.VB_Invoke_Func = "X\n14"

    Dim myRange As Range
    Dim cell As Range
    Dim targetCell As Range
    Dim targetOffset As Long
    Dim searchArea As Range
    
    Set searchArea = Intersect(ActiveSheet.usedRange, Selection.EntireColumn)
    Dim ceiling As Range
    
    For Each cell In searchArea
        If cell.Value = 2023 Or cell.Value = 2024 Or cell.Value = 2025 Then
            Set ceiling = cell
            Exit For
        End If
    Next cell

    ' --- Guard: no year found ---
    If ceiling Is Nothing Then
        MsgBox "No year header (2023/2024/2025) found in the selected column.", vbExclamation
        Exit Sub
    End If
    
    Dim floor As Range
    Set floor = Cells(Rows.Count, ceiling.column).End(xlUp) 'Rows.count is the very bottom row. End(xlUp is equivalent to ctrl + Up arrow
    
    Set myRange = Range(ceiling, floor) ' (Optionally ceiling.Offset(1,0) to skip the header row)
    
    targetOffset = val(InputBox("ColumnBox to the right:", "shift column right"))
    'InputBox returns a string. you should Val it to make it an integer
    
    Dim report As String
    Dim isExternal As Boolean
    Dim originalFormula As String
    Dim originalValue As Variant
    
    report = ""
    
    For Each cell In myRange
        Set targetCell = cell.Offset(0, targetOffset)
        
        If cell.HasFormula Then
            ' 1. Check for External Links
            isExternal = (InStr(1, cell.Formula, "[", vbTextCompare) > 0 And _
                          InStr(1, cell.Formula, "!", vbTextCompare) > 0)
            
            If Not isExternal Then
                ' Capture the state BEFORE changing the formula ---
                originalValue = targetCell.Value2
                originalFormula = vbNullString
                If targetCell.HasFormula Then originalFormula = targetCell.Formula
                
                On Error Resume Next
                ' 2. Apply the relative formula
                targetCell.FormulaR1C1 = cell.FormulaR1C1
                targetCell.Calculate
                On Error GoTo 0
                
                ' 3. Verification Logic
                If Not targetCell.HasFormula Then
                    ' Revert if the set failed
                    GoTo RevertCell
                Else
                    ' Use your existing Private Function ValuesEqual
                    If ValuesEqual(targetCell.Value2, originalValue) Then
                        report = report & targetCell.Address & ": set formula (value unchanged)" & vbCrLf
                    Else
                        ' Math shifted to the wrong data, must revert to keep integrity
                        GoTo RevertCell
                    End If
                End If
                
            Else
                report = report & targetCell.Address & ": skipped (external link)" & vbCrLf
            End If
        End If
        GoTo NextCell

RevertCell:
        ' Revert to the state we captured at the start of the loop
        If originalFormula <> vbNullString Then
            targetCell.Formula = originalFormula
        Else
            targetCell.Value2 = originalValue
        End If
        report = report & targetCell.Address & ": skipped (values differ or error)" & vbCrLf

NextCell:
    Next cell
    
    MsgBox report, vbInformation, "Final Results"
    
End Sub

Private Function ValuesEqual(ByVal a As Variant, ByVal b As Variant) As Boolean
    Const EPS As Double = 0.0000001
    On Error GoTo Bail
    
    ' 1. Check for Errors First (Fastest exit)
    If IsError(a) Or IsError(b) Then
        ValuesEqual = (IsError(a) And IsError(b)) And (CStr(a) = CStr(b))
        Exit Function
    End If
    
    ' 2. Numeric Comparison (Handles 90% of Excel data)
    ' Use VarType check instead of IsNumeric to avoid slow string-to-number checks
    Dim vtA As Integer: vtA = VarType(a)
    Dim vtB As Integer: vtB = VarType(b)
    
    If (vtA >= 2 And vtA <= 7) And (vtB >= 2 And vtB <= 7) Then
        ValuesEqual = (Abs(a - b) < EPS)
        Exit Function
    End If
    
    ' 3. String/Empty Comparison
    ' CStr(Empty) results in "", so this covers Null/Empty/Strings in one shot
    ValuesEqual = (CStr(a) = CStr(b))
    Exit Function

Bail:
    ValuesEqual = False
End Function
Sub ClearNumbersAndFormat()
Attribute ClearNumbersAndFormat.VB_ProcData.VB_Invoke_Func = "D\n14"
    Dim cell As Range
    Dim targetRange As Range
    Set targetRange = Selection
    Application.ScreenUpdating = False

    For Each cell In targetRange
        If cellHasHardValues(cell) = True Then
            
            ' Exclude the specific years
            Select Case cell.Value
                Case 2023, 2024, 2025
                    ' These stay exactly as they are
                Case Else
                    cell.ClearContents

                    cell.Value = 0
                    cell.NumberFormat = "_(* #,##0_);_(* (#,##0);_(* "" - ""??_);_(@_)"
            End Select
        End If
    Next cell

    Application.ScreenUpdating = True
End Sub

Sub UpdateYearsInSelection()
Attribute UpdateYearsInSelection.VB_ProcData.VB_Invoke_Func = "Y\n14"
    Dim cell As Range
    Dim cellFormula As String
    
    If TypeName(Selection) <> "Range" Then Exit Sub

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

For Each cell In Selection
    ' Skip error cells to avoid Type Mismatch
    If Not IsError(cell.Value) Then
        Dim val As String
        val = CStr(cell.Value)
        'YA
        Dim posYA As Long, pos2025 As Long
        posYA = InStr(1, val, "YA", vbTextCompare)
        pos2025 = InStr(1, val, "2025", vbTextCompare)
        
        If posYA > 0 And pos2025 > posYA Then
            val = Replace(val, "2025", "2026")
        End If
        
        If InStr(val, "2024") > 0 Then val = Replace(val, "2024", "2025")
        If InStr(val, "2023") > 0 Then val = Replace(val, "2023", "2024")
        If InStr(val, "2022") > 0 Then val = Replace(val, "2022", "2023")
        
        ' Update the cell once at the end
        cell.Value = val
    End If
Next cell

    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    
End Sub
