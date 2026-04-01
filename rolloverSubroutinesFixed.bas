Attribute VB_Name = "rolloverSubroutinesFixed"
Public Type SOCE_cell
    RowIndex As Long
    ColIndex As Long
    contents As String
End Type
Private Function ValuesEqual(a As Variant, b As Variant) As Boolean
    ' Treat numerics (including dates) with tolerance; otherwise string compare.
    Const EPS As Double = 0.0000001
    On Error GoTo Bail
    
    If IsError(a) Or IsError(b) Then
        If IsError(a) And IsError(b) Then
            ValuesEqual = (CLng(a) = CLng(b))
        Else
            ValuesEqual = False
        End If
        Exit Function
    End If
    
    
    ' Normalize Empty/Null/zero-length
    If (IsEmpty(a) Or IsNull(a) Or (VarType(a) = vbString And Len(a) = 0)) _
       And (IsEmpty(b) Or IsNull(b) Or (VarType(b) = vbString And Len(b) = 0)) Then
        ValuesEqual = True
        Exit Function
    End If
    
    If IsNumeric(a) And IsNumeric(b) Then
        ValuesEqual = (Abs(CDbl(a) - CDbl(b)) < EPS)
        Exit Function
    End If
    
    ' Exact text compare; change to vbTextCompare if you want case-insensitive
    ValuesEqual = (CStr(a) = CStr(b))
    Exit Function
Bail:
    ValuesEqual = False
End Function

Public Sub MoveFormulasFromLeftToRightCustom(ByVal targetOffset As Long, _
                                            ByVal sourceCol As Range, _
                                            ByVal src As Worksheet, _
                                            ByVal tgt As Worksheet)
                                            
    Dim r As Long
    Dim sCell As Range, tCell As Range
    Dim fText As String, fTextA1 As String, fTextR1C1 As String, isExternal As Boolean, isAbsolute As Boolean
    Dim srcVal As Variant, tgtVal As Variant
    
    Debug.Print "--- START: " & src.Name & " -> " & tgt.Name & " ---"
    Debug.Print "SourceCol: " & sourceCol.Address & " | Rows: " & sourceCol.Rows.Count
    
    For r = 1 To sourceCol.Rows.Count
        
        Set sCell = sourceCol.Cells(r, 1)
        Set tCell = tgt.Cells(sCell.Row, sCell.Column + targetOffset)
        If sCell.HasFormula Then
            fTextR1C1 = sCell.FormulaR1C1
            fTextA1 = sCell.Formula
            isExternal = False
            If InStr(1, fTextA1, "!") > 0 Then
                isExternal = True
            ElseIf InStr(1, fTextA1, "[") > 0 And InStr(1, fTextA1, ".xl") > 0 Or _
                         (InStr(1, fTextA1, "!", vbTextCompare) > 0) Or _
                         (InStr(1, fTextR1C1, "!", vbTextCompare) > 0) Then
                isExternal = True
            End If
            
            isAbsolute = (fTextR1C1 Like "*R[0-9]*" And Not fTextR1C1 Like "*R[[]*") Or _
                         (fTextR1C1 Like "*C[0-9]*" And Not fTextR1C1 Like "*C[[]*")

            ' --- DEBUG FILTERS ---
            If isExternal Then
                ' Debug.Print "Row " & sCell.Row & ": Skip (External)"
            ElseIf isAbsolute Then
                ' Debug.Print "Row " & sCell.Row & ": Skip (Absolute)"
            Else
                If (InStr(1, fTextR1C1, "SUM", vbTextCompare) > 0 Or _
                    InStr(1, fTextR1C1, "+", vbTextCompare) > 0 Or _
                    InStr(1, fTextR1C1, "ROUND", vbTextCompare) > 0) Or _
                    ((Trim(fTextA1) Like "=[a-zA-Z][1-9]" Or Trim(fTextA1) Like "=[a-zA-Z][1-9]#") And Mid(fTextA1, 2, 1) = Left(sCell.Address(0, 0), 1)) Then
                    
                    ' Logic Step C: The Value Integrity Check
                    srcVal = sCell.Value2
                    tgtVal = tCell.Value2
                    
                    If ValuesEqual(srcVal, tgtVal) Then
                        On Error Resume Next
                        tCell.FormulaR1C1 = fTextR1C1
                        On Error GoTo 0
                    Else
                    End If
                    
                End If
            End If
        End If
    Next r

End Sub
Function GetFilePath(title As String) As String
    Dim fileName As Variant
    fileName = Application.GetOpenFilename("Excel Files (*.xls*), *.xls*", , title)
    
    If fileName = False Then
        GetFilePath = ""
    Else
        GetFilePath = fileName
    End If
End Function

Sub BuildClassTwins(srcWb As Workbook, tgtWb As Workbook, twinList As Collection)
Dim srcSheet As Worksheet, tgtSheet As Worksheet
    Dim wsCheck As Worksheet
    Dim twinObj As ClsSheetTwin
    
    For Each srcSheet In srcWb.Worksheets
        Set tgtSheet = Nothing
        
        For Each wsCheck In tgtWb.Worksheets
            If UCase(Trim(wsCheck.Name)) = UCase(Trim(srcSheet.Name)) And InStr(1, wsCheck.Name, "2023") = 0 And InStr(1, wsCheck.Name, "2022") = 0 Then
                Set tgtSheet = wsCheck
                Exit For
            End If
        Next wsCheck
        
        If Not tgtSheet Is Nothing Then
            Set twinObj = New ClsSheetTwin
            twinObj.Init srcSheet, tgtSheet
            
            twinList.Add twinObj
            Set twinObj = Nothing
        End If
    Next srcSheet
    
End Sub

Sub shiftColumnsInTwin(source As Worksheet, target As Worksheet, twinObj As ClsSheetTwin)
    On Error GoTo ErrorHandler
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    
    If InStr(1, source.Name, "2022", vbTextCompare) > 0 Or _
       InStr(1, source.Name, "2023", vbTextCompare) > 0 Then GoTo CleanUp
    
    If Not isFS(twinObj) Then GoTo CleanUp
    If IsValidSOCE(twinObj) Then GoTo CleanUp
    
    Dim searchRange As Range
    Set searchRange = Application.Intersect(source.usedRange, source.Range("1:100"))
    If searchRange Is Nothing Then GoTo CleanUp

    Dim cell As Range
    Dim cell2 As Range
    Dim isPeriod As Boolean
    Dim searchCol As Range
    
    For Each cell In searchRange
    isPeriod = False
        If Not IsError(cell.Value) Then
            If CStr(cell.Value) = "2024" Then
                If checkLeftForNote(cell) = True Then
                    ProcessHeaderShift source, target, cell, target.Range(cell.Address), twinObj
                End If
            End If
        End If
        
        If Not IsError(cell.Value) Then
            If CStr(cell.Value) Like "*#.####*" = True And CStr(cell.Offset(1, 0).Value) Like "*to*" = True And CStr(cell.Offset(2, 0).Value) Like "*#.##.####*" = True Then
                'check if the word "period" exists in the column
                Set searchCol = Application.Intersect(cell.EntireColumn, source.usedRange)
                For Each cell2 In searchCol
                    If InStr(1, cell2.Value, "period", vbTextCompare) > 0 Then
                        isPeriod = True
                    End If
                Next cell2
                
                For Each cell2 In searchCol.Offset
                    If InStr(1, cell2.Value, "period", vbTextCompare) > 0 Then
                        isPeriod = True
                    End If
                Next cell2
                
                If isPeriod = False Then
                    ProcessHeaderShift source, target, cell.Offset(2, 0), target.Range(cell.Offset(2, 0).Address), twinObj
                End If
            
            End If
        End If
                
    Next cell

CleanUp:
    Set cell = Nothing
    Set cell2 = Nothing
    Set searchRange = Nothing
    Set searchCol = Nothing
    Exit Sub

ErrorHandler:
    Debug.Print "!!! Error in shiftColumnsInTwin (" & source.Name & "): " & Err.Description
    Resume CleanUp
End Sub

Private Sub ProcessHeaderShift(source As Worksheet, target As Worksheet, srcHeader As Range, tgtHeader As Range, twinObj As ClsSheetTwin)
    Dim colOffset As Long
    colOffset = 0
    
    If Not IsError(srcHeader.Offset(0, 1).Value) Then
        If Trim(CStr(srcHeader.Offset(0, 1).Value)) Like "*2023*" Then colOffset = 1
    End If
    
    If colOffset = 0 And Not IsError(srcHeader.Offset(0, 2).Value) Then
        If Trim(CStr(srcHeader.Offset(0, 2).Value)) Like "*2023*" Then colOffset = 2
    End If
    
    
    Dim cell As Range
    Dim searchCol As Range
    Dim isPeriod As Boolean
    isPeriod = False
    Set cell = srcHeader
    Set searchCol = Application.Intersect(cell.EntireColumn, source.usedRange)
    For Each cell2 In searchCol.Offset(0, colOffset)
        If InStr(1, cell2.Value, "period", vbTextCompare) > 0 Then
            isPeriod = True
        End If
    Next cell2
    If isPeriod = True Then GoTo CleanUp ' verify that the column on the right of the header does not contain "period". We already checked that the header's column does not contain "period" in shiftColumnsInTwin.
    

    If colOffset > 0 Then
        Dim sourceCol As Range
        Set sourceCol = getColumn(srcHeader)

        If Not sourceCol Is Nothing Then
            Dim colData As Variant
            colData = sourceCol.Value2
            
            Set twinObj.tgtColOld = getColumn(tgtHeader)
            
            If Not twinObj.tgtColOld Is Nothing Then
                Dim newCol As Range
                Set newCol = twinObj.tgtColOld.Offset(0, colOffset).Resize(sourceCol.Rows.Count, 1)
                newCol.Value2 = colData
                DoEvents
                
                Call MoveFormulasFromLeftToRightCustom(colOffset, sourceCol, source, target)
                Call ClearNumbersAndFormatCustom(twinObj.tgtColOld)
                Call FormatAsAccountingCustom(newCol)
                Call FormatAsAccountingCustom(twinObj.tgtColOld)
                
                tgtHeader.Value = 2025
                'clear unusual headers, if any
                If CStr(srcHeader.Offset(-1, 0)) Like "*to*" Then
                    tgtHeader.Offset(-1, 0).ClearContents
                    tgtHeader.Offset(-2, 0).ClearContents
                    tgtHeader.Offset(-1, colOffset).ClearContents
                    tgtHeader.Offset(-2, colOffset).ClearContents
                    tgtHeader.Offset(0, colOffset).Value2 = "2024"
                    tgtHeader.Offset(0, colOffset).NumberFormat = "General"
                    tgtHeader.NumberFormat = "General"
                End If
                
            End If
        End If
    End If
    
CleanUp:
If Not IsEmpty(colData) Then Erase colData
Set cell = Nothing
Set searchCol = Nothing
Set newCol = Nothing


End Sub
Function checkLeftForNote(cell As Range) As Boolean
    Dim cellLeft As Range
    Dim startCol As Long, endCol As Long
    
    checkLeftForNote = False
    startCol = cell.Column
    endCol = startCol - 5
    If endCol < 1 Then endCol = 1
    
    With cell.Worksheet
        For Each cellLeft In .Range(.Cells(cell.Row, endCol), cell)
            If Not IsError(cellLeft.Value) Then
                If InStr(1, CStr(cellLeft.Value), "Note", vbTextCompare) > 0 Then
                    checkLeftForNote = True
                ElseIf InStr(1, CStr(cellLeft.Offset(1, 0).Value), "Note", vbTextCompare) > 0 Then
                    checkLeftForNote = True
                    Exit For
                End If
            End If
        Next cellLeft
    End With
End Function

Function getColumn(ceiling As Range) As Range
    Dim ws As Worksheet
    Set ws = ceiling.Worksheet
    Dim floor As Range
    
    ' 1. Find the last cell with data in that column
    Set floor = ws.Cells(ws.Rows.Count, ceiling.Column).End(xlUp)
    
    ' 2. Logic Check: If floor is above or equal to ceiling, there's no data below
    If floor.Row <= ceiling.Row Then
        Set getColumn = Nothing
        Exit Function
    End If
    
    ' 3. Safety Check: If the range is suspiciously large (e.g., > 5000 rows)
    ' on a sheet that isn't supposed to be that big, cap it at the UsedRange.
    If (floor.Row - ceiling.Row) > 5000 Then
        Dim lastUsedRow As Long
        lastUsedRow = ws.usedRange.Rows(ws.usedRange.Rows.Count).Row
        Set floor = ws.Cells(lastUsedRow, ceiling.Column)
    End If

    Set getColumn = ws.Range(ceiling, floor)
End Function

Sub ForceFullRecalc()
    Application.Calculation = xlCalculationAutomatic
    Application.CalculateFullRebuild
    Application.Calculate
End Sub

Public Function cellHasHardValues(inputCell As Range)
    If IsEmpty(inputCell.Value) Then
        cellHasHardValues = False
        Exit Function
    End If
    If Not inputCell.HasFormula And IsNumeric(inputCell.Value) And Not IsEmpty(inputCell.Value) Then
    cellHasHardValues = True
    
    ElseIf inputCell.HasFormula Then
        If Not CStr(inputCell.Formula) Like "*[A-Za-z]*" Then
            cellHasHardValues = True
        Else
            cellHasHardValues = False
        End If
    End If
End Function
Public Function IsValidSOCE(ByRef twinObj As ClsSheetTwin) As Boolean
    If twinObj.source.usedRange Is Nothing Then
        IsValidSOCE = False
        Exit Function
    End If
    If InStr(1, twinObj.source.Name, "SOCE") > 1 Then
        IsValidSOCE = True
        Exit Function
    End If
    Dim isSOCE As Boolean
    Dim usedRange As Range
    Dim cell As Range
    Dim leftCol As Long, matchCount As Long
    Set usedRange = Application.Intersect(twinObj.source.usedRange, twinObj.source.Range("A1:V300"))
    Dim identifierArray As Variant
    identifierArray = Array("At 1 January 2023", "Total comprehensive income for the year", "At 31 December 2023 and 1 January 2024", "Total comprehensive loss for the year", "At 31 December 2024")
    matchCount = 0
    For Each cell In usedRange
        For i = LBound(identifierArray) To UBound(identifierArray)
        If InStr(1, CStr(cell.Value), identifierArray(i), vbTextCompare) Then
            matchCount = matchCount + 1
            If (leftCol = 0) Then
                leftCol = cell.Column
            ElseIf (leftCol <> cell.Column) Then
                MsgBox "Error in sheet '" & twinObj.source.Name & "'. '" & identifierArray(i) & "' is in column " & cell.Column & " instead of column " & leftCol & ". (This is a message for finding SOCE sheets) Program will skip that sheet."
                isSOCE = False
                IsValidSOCE = isSOCE
                Exit Function
            End If
        Exit For 'exit i loop because match is valid
        End If
        Next i
    Next cell
    If matchCount > 2 Then isSOCE = True
    IsValidSOCE = isSOCE
End Function
Public Sub SOCE_Identifier(twinObj As ClsSheetTwin)
    If twinObj.source.usedRange Is Nothing Then Exit Sub
    On Error GoTo CleanUp

    If IsValidSOCE(twinObj) = False Then
        GoTo CleanUp
        ' Debug.Print "No SOCE found for " & twinObj.source.Name
    ElseIf IsValidSOCE(twinObj) = True Then Debug.Print "FOUND SOCE for " & twinObj.source.Name
    End If
    
    'past this point it should be SOCE
    ' now we want to identify how many SOCE tables there are
    Dim tableManager As ClsSOCETable
    Set tableManager = New ClsSOCETable
    tableManager.Init twinObj
    ' creates a ClsSOCETable with ClsSOCEInstances
    Set tableManager = Nothing 'free the memory
                    

CleanUp:
    If Err.Number <> 0 Then MsgBox "Error " & Err.Number & " : " & Err.Description, vbCritical, "An Error Occurred"
End Sub

Public Function rowColumnIntersection(row1 As Range, column1 As Range)
    ' returns the intersection range of your row stubs and column headers
    If (row1.Rows.Count > 1) Then MsgBox "Error: " & row1.Rows.Count & " rows passed to rowColumnIntersection (only 1 is allowed)"
    If (column1.Columns.Count > 1) Then MsgBox "Error: " & column1.Columns.Count & " columns passed to rowColumnIntersection (only 1 is allowed)"
    Set rowColumnIntersection = Intersect(row1.EntireColumn, column1.EntireRow)
End Function

Sub ClearNumbersAndFormatCustom(targetRange As Range)
    Dim cell As Range


    For Each cell In targetRange
        If cellHasHardValues(cell) = True Then
            Select Case cell.Value
                Case 2023, 2024, 2025
                    ' Keep
                Case Else
                    cell.ClearContents
                    cell.Value = 0
                    cell.NumberFormat = "_(* #,##0_);_(* (#,##0);_(* "" - ""??_);_(@_)"
            End Select
        End If
    Next cell
End Sub

Sub UpdateYearsInArray(ByRef dataArr As Variant)
    If Not IsArray(dataArr) Then Exit Sub
    Dim val As String
    Dim r As Long, c As Long
    For r = LBound(dataArr, 1) To UBound(dataArr, 1) 'rows dimension
        For c = LBound(dataArr, 2) To UBound(dataArr, 2) 'columns dimension
            If VarType(dataArr(r, c)) = vbString Then
                val = CStr(dataArr(r, c))
                Dim posYA As Long, pos2025 As Long
                posYA = InStr(1, val, "YA", vbTextCompare)
                pos2025 = InStr(1, val, "2025", vbTextCompare)
                If posYA > 0 And pos2025 > posYA Then
                    val = Replace(val, "2025", "2026")
                End If
                If InStr(val, "2024") > 0 Then val = Replace(val, "2024", "2025")
                If InStr(val, "2023") > 0 Then val = Replace(val, "2023", "2024")
                If InStr(val, "2022") > 0 Then val = Replace(val, "2022", "2023")
                'save data into the array
                dataArr(r, c) = val
            End If
        Next c
    Next r
End Sub

Public Function isFS(twinObj As ClsSheetTwin) As Boolean
    If twinObj.isFinancialStatement = True Then
        isFS = True
        Exit Function
    End If
    
    Dim target As Worksheet, source As Worksheet
    Dim cell As Range, searchRange As Range
    Dim gridlinesOff As Boolean
    
    Set target = twinObj.target
    Set source = twinObj.source
    
    If InStr(1, target.Name, "2022", vbTextCompare) > 0 Or _
       InStr(1, target.Name, "2023", vbTextCompare) > 0 Then
        isFS = False
        Exit Function
    End If
    
    gridlinesOff = Not target.Parent.Windows(1).SheetViews(target.Name).DisplayGridlines
    
    If Not gridlinesOff Or IsValidSOCE(twinObj) Then
        isFS = False 'we take SOCE != FS
        Exit Function
    ElseIf gridlinesOff And Not IsValidSOCE(twinObj) Then
        isFS = True
        Exit Function
    End If
    
    Set searchRange = Application.Intersect(source.usedRange, source.Range("A1:Z300"))
    If searchRange Is Nothing Then isFS = False: Exit Function

    For Each cell In searchRange
        If Not IsError(cell.Value) Then
            If InStr(1, CStr(cell.Value), "2024") > 0 Then
                If checkLeftForNote(cell) Then
                    isFS = True
                    twinObj.isFinancialStatement = True
                    Exit Function
                End If
            End If
        End If
    Next cell
    
    isFS = False
End Function

Sub FormatAsAccountingCustom(targetRange As Range)
    Dim cell As Range
    Dim v As Variant

    For Each cell In targetRange
        v = cell.Value
        If v = 2023 Or v = 2024 Or v = 2025 Or v = 2026 Or _
           v = "2023" Or v = "2024" Or v = "2025" Or v = "2026" Then
            cell.NumberFormat = "@"
        Else
            cell.NumberFormat = "_(* #,##0_);_(* (#,##0);_(* ""-""??_);_(@_)"
        End If
    Next cell
End Sub

Sub updateYearsInSheet(twinObj As ClsSheetTwin)
    Dim source As Worksheet
    Set source = twinObj.source
    Dim sourceRange As Range
    Set sourceRange = Application.Intersect(source.usedRange, source.Range("A1", "Z200"))
    Dim cell As Range
    For Each cell In sourceRange
        If CStr(cell.Value) Like "*31 December 2024*" Then
            With twinObj.target.Cells(cell.Row, cell.Column)
                .NumberFormat = "@" ' Sets format to "Text"
                .Value = Replace(CStr(cell.Value), "31 December 2024", "31 December 2025")
            End With
        End If
    Next cell
End Sub


