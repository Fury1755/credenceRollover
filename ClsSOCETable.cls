VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClsSOCETable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public rowStubs As Range
Public tableHeaders As Range
Public prevYearValues As Range
Public currentYearValues As Range
Public Instances As Collection
Public Parent As ClsSheetTwin ' small warning: circular reference


Public Function Init(twinObj As ClsSheetTwin)
' initializes and holds a collection of your tables, where each table is a ClsSOCEInstance.
    ' called by SOCE_Identifier which is the master sub.
    Set Me.Instances = New Collection
    Set Me.Parent = twinObj
    Dim cell As Range
    Dim source As Worksheet
    Dim target As Worksheet
    Set source = twinObj.source
    Set target = twinObj.target
    Dim rowStubsBegin As Range
    Dim rowStubsMiddle As Range
    Dim rowStubsEnd As Range
    Dim tableHeadersBegin As Range
    Dim tableHeadersEnd As Range
    Dim headerCollection As Variant
    headerCollection = Array("capital", "earnings", "Total")
    Dim rowStubsCollection As Variant
    rowStubsCollection = Array("At 1 January 2023", "At 31 December 2023 and 1 January 2024", "At 31 December 2024")
    Dim tableValues As Range
    Dim rowStubIsNothing As Long
    
    If twinObj.usedRange Is Nothing Then GoTo CleanUp
    For Each cell In twinObj.usedRange
        For i = LBound(headerCollection) To UBound(headerCollection)
            If InStr(1, CStr(cell.Value), headerCollection(i)) > 0 Then
                If InStr(1, CStr(cell.Offset(1, 0).Value), "S$") > 0 Or InStr(1, CStr(cell.Offset(1, 0).Value), "US$") > 0 Then
                    If headerCollection(i) = "capital" Then
                        Set tableHeadersBegin = cell
                    ElseIf UCase(Trim(headerCollection(i))) = "TOTAL" Then
                        Set tableHeadersEnd = cell
                    End If
                End If
            End If
        Next i

        For j = LBound(rowStubsCollection) To UBound(rowStubsCollection)
            If InStr(1, CStr(cell.Value), rowStubsCollection(j)) > 0 Then
                If rowStubsCollection(j) = "1 January 2023" Or rowStubsCollection(j) = "1 January 2023" Then
                    Set rowStubsBegin = cell
                ElseIf rowStubsCollection(j) = "At 31 December 2023 and 1 January 2024" Then
                    Set rowStubsMiddle = cell
                ElseIf rowStubsCollection(j) = "At 31 December 2024" Then
                    Set rowStubsEnd = cell
                End If
            End If
        Next j
        
        'now that we have both pairs, we create a new instance
        If (Not rowStubsEnd Is Nothing) And (Not tableHeadersEnd Is Nothing) And Not rowStubsMiddle Is Nothing And Not rowStubsBegin Is Nothing Then
                Set Me.rowStubs = source.Range(rowStubsBegin, rowStubsEnd)
                Set Me.tableHeaders = source.Range(tableHeadersBegin, tableHeadersEnd)
                Set tableValues = Application.Intersect(Me.rowStubs.EntireRow, Me.tableHeaders.EntireColumn)
                If Me.rowStubs.Count Mod 2 = 1 Then
                    Set Me.prevYearValues = Application.Intersect(source.Range(rowStubsBegin, rowStubsMiddle.Offset(-1, 0)).EntireRow, Me.tableHeaders.EntireColumn)
                    Set Me.currentYearValues = Application.Intersect(source.Range(rowStubsMiddle, rowStubsEnd.Offset(-1, 0)).EntireRow, Me.tableHeaders.EntireColumn)
                ElseIf Me.rowStubs.Count Mod 2 = 0 Then
                    MsgBox "SOCE table in " & source.Name & " has even number of rows. Unable to process, check carefully."
                    GoTo CleanUp
                'incomplete logic. if this happens, it means SOCE has different number of rows for both years (like issuance of shares or something)
                End If
                ' Debug.Print "currentYearValues: " & currentYearValues.Address
                ' Debug.Print "prevYearValues: " & prevYearValues.Address
                
                Dim newIns As ClsSOCEInstance
                Set newIns = New ClsSOCEInstance
                newIns.Init Me, Me.currentYearValues, Me.prevYearValues
                Instances.Add newIns
                
                Set rowStubsBegin = Nothing
                Set rowStubsMiddle = Nothing
                Set rowStubsEnd = Nothing
                Set tableHeadersBegin = Nothing
                Set tableHeadersEnd = Nothing
            End If
    Next cell
    
    
    
If Me.Instances.Count > 0 Then
    For i = 1 To Me.Instances.Count
        Me.Instances(i).processValues Me
    Next i
    
    Else
        Call identifyIrregularRowStubs(twinObj)
End If

CleanUp:
    If Err.Number <> 0 Then MsgBox "Error " & Err.Number & " in ClsSOCETable doing " & twinObj.source.Name & ": " & Err.Description, vbCritical
    Set cell = Nothing
    Set source = Nothing
    Set target = Nothing
    Set rowStubsBegin = Nothing
    Set rowStubsMiddle = Nothing
    Set rowStubsEnd = Nothing
    Set tableHeadersBegin = Nothing
    Set tableHeadersEnd = Nothing
    Set tableValues = Nothing

End Function

Public Function identifyIrregularRowStubs(twinObjarg As ClsSheetTwin)
    ' writing everything for SOCE from scratch, because everything before is already working
    ' pray i don't need to refactor this one more time.
    Dim searchRange As Range
    Dim source As Worksheet
    Dim target As Worksheet
    Dim twinObj As ClsSheetTwin
    Set twinObj = twinObjarg
    Set source = twinObj.source
    Set target = twinObj.target
    
    Set searchRange = Application.Intersect(target.usedRange, target.Range("A1", "Z100"))
    
    Dim cell As Range
    Dim rowStubsArray() As SOCE_cell
    Dim colHeadersArray() As SOCE_cell
    ReDim Preserve rowStubsArray(1 To 999)
    ReDim Preserve colHeadersArray(1 To 999)
    
    Dim searchRangeValue As Variant
    Dim searchRangeFormula As Variant
    searchRangeValue = searchRange.Value
    searchRangeFormula = searchRange.Formula
    
    Dim r As Long, c As Long, i As Long, j As Long, k As Long
    Dim months As Variant
    months = Array("January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December")
    
    Dim colBegin As SOCE_cell, colEnd As SOCE_cell
    
    j = 1
    k = 1
    
    For r = 1 To UBound(searchRangeFormula, 1)
        For c = 1 To UBound(searchRangeFormula, 2)
        
            For i = LBound(months) To UBound(months)
            
                If (CStr(searchRangeFormula(r, c)) Like "*# " & months(i) & " ####*" Or CStr(searchRangeValue(r, c)) Like "*# " & months(i) & " ####*") And _
                k > 1 Then
                ' identifies dates e.g. DD June YYYY
                ' rowstubs will only be identified after colHeaders have been identified
                
                
                     rowStubsArray(j).RowIndex = r + searchRange.Row - 1
                     rowStubsArray(j).ColIndex = c + searchRange.Column - 1
                     rowStubsArray(j).contents = CStr(searchRangeFormula(r, c))
                    j = j + 1
                    Exit For
                End If
            Next i
            
            If searchRangeFormula(r, c) Like "*capital*" Then
                colHeadersArray(k).RowIndex = r + searchRange.Row - 1
                colHeadersArray(k).ColIndex = c + searchRange.Column - 1
                colHeadersArray(k).contents = CStr(searchRangeFormula(r, c))
                k = k + 1
                
            End If
            
            If Trim(CStr(searchRangeFormula(r, c))) Like "Total" Then
                colHeadersArray(k).RowIndex = r + searchRange.Row - 1
                colHeadersArray(k).ColIndex = c + searchRange.Column - 1
                colHeadersArray(k).contents = CStr(searchRangeFormula(r, c))
                k = k + 1
            End If
                
        Next c
    Next r
    
    If j > 1 Then ReDim Preserve rowStubsArray(1 To j - 1) ' trim rowStubsArray to have exactly j elements
    If k > 1 Then ReDim Preserve colHeadersArray(1 To k - 1) ' same for this
    
    ' check the results of the search and some basic error handling
    If Not UBound(rowStubsArray) Mod 3 = 0 Then
        MsgBox source.Name & " is an irregular SOCE and the dates in the left column found are not in groups of 3. Skipping."
        GoTo CleanUp
    End If
    
    If Not UBound(colHeadersArray) Mod 2 = 0 Then
        MsgBox source.Name & " is an irregular SOCE and the column headers are not in groups of 2. Skipping."
        GoTo CleanUp
    End If
    
    If Not UBound(rowStubsArray) Mod 3 = UBound(colHeadersArray) Mod 2 Then
        MsgBox source.Name & " is an irregular SOCE and the script failed to group the table headers with the column on the left properly. Skipping. Sorry!"
        GoTo CleanUp
    End If
    
    ' now we identify whether current year has more or less rows
    Dim rowStubsBegin As SOCE_cell, rowStubsMiddle As SOCE_cell, rowStubsEnd As SOCE_cell
    
    Dim oldRowStubsArray() As SOCE_cell, oldcolHeadersArray() As SOCE_cell
    oldRowStubsArray = identifyRowStubsInTarget(source)
    oldcolHeadersArray = identifyColHeadersInTarget(source)
            
    For j = 1 To UBound(rowStubsArray) / 3
    
        rowStubsBegin = rowStubsArray(3 * j - 2)
        rowStubsMiddle = rowStubsArray(3 * j - 1)
        rowStubsEnd = rowStubsArray(3 * j)
        
        i = rowStubsMiddle.RowIndex - rowStubsBegin.RowIndex
        k = rowStubsEnd.RowIndex - rowStubsMiddle.RowIndex
            ' k we are gonna do this the easy way
            
            Dim x As Long, y As Long, z As Long, rowsAdded As Long
            
            rowsAdded = k - i
            Dim targetRowStubs() As SOCE_cell
            targetRowStubs = identifyRowStubsInTarget(target)
            
            If (UBound(targetRowStubs) >= 3 * j) Then
            
                If targetRowStubs(3 * j).RowIndex - targetRowStubs(3 * j - 1).RowIndex = _
                   targetRowStubs(3 * j - 1).RowIndex - targetRowStubs(3 * j - 2).RowIndex Then
                    GoTo skipInsert
                End If
            
            Else
                MsgBox "targetRowStubs found " & UBound(targetRowStubs) & " (supposed to be a multiple of 3)"
            End If
            
            
            If Not rowsAdded = 0 Then
                If i < k Then
                    target.Rows(rowStubsMiddle.RowIndex).Resize(k - i).Insert ' inserts (k-i) number of rows
                End If
                
                If i > k Then
                    target.Rows(rowStubsMiddle.RowIndex).Offset(k - i, 0).Resize(-(k - i)).EntireRow.Delete ' deletes (k-i) number of rows above rowStubsMiddle
                End If
                    ' update the corresponding addresses in arrays
                                        
                For x = 1 To UBound(rowStubsArray)
                        If rowStubsArray(x).RowIndex >= rowStubsMiddle.RowIndex Then
                        rowStubsArray(x).RowIndex = rowStubsArray(x).RowIndex + rowsAdded
                    End If
                Next x
                
                For x = 1 To UBound(colHeadersArray)
                    If colHeadersArray(x).RowIndex >= rowStubsMiddle.RowIndex Then
                        colHeadersArray(x).RowIndex = colHeadersArray(x).RowIndex + rowsAdded
                    End If
                Next x
            
                Set searchRange = Application.Intersect(target.usedRange, target.Range("A1", "Z100"))
                searchRangeValue = searchRange.Value ' writing here?
                searchRangeFormula = searchRange.Formula
                
                rowStubsBegin = rowStubsArray(3 * j - 2)
                rowStubsMiddle = rowStubsArray(3 * j - 1)
                rowStubsEnd = rowStubsArray(3 * j)
            End If
            
skipInsert:
            
            'currYearVal denotes current year values for this specific function
            Dim currYear As Range, prevYear As Range
            ' the rows are guaranteed to be equal so we dont need to check
            Set prevYear = Application.Intersect(target.Range(target.Cells(rowStubsBegin.RowIndex, rowStubsBegin.ColIndex), target.Cells(rowStubsMiddle.RowIndex, rowStubsMiddle.ColIndex)).EntireRow, target.Range(target.Cells(colHeadersArray(2 * j - 1).RowIndex, colHeadersArray(2 * j - 1).ColIndex), target.Cells(colHeadersArray(2 * j).RowIndex, colHeadersArray(2 * j).ColIndex)).EntireColumn)
            Set prevYear = prevYear.Resize(prevYear.Rows.Count - 1, prevYear.Columns.Count)
            Set currYear = prevYear.Offset(rowStubsMiddle.RowIndex - rowStubsBegin.RowIndex)
            

            Dim rowIdxBase As Long, colIdxBase As Long
            Dim startRow As Range, endRow As Range
            Dim startHeader As Range, endHeader As Range
            Dim sourcePrevYear As Range, sourceCurrYear As Range
            Dim rowSlice As Range, colSlice As Range
            
            rowIdxBase = (j - 1) * 3
            colIdxBase = (j - 1) * 2
            
            If (rowIdxBase + 2) <= UBound(oldRowStubsArray) And (colIdxBase + 2) <= UBound(oldcolHeadersArray) Then
            
                Set startRow = source.Cells(oldRowStubsArray(rowIdxBase + 2).RowIndex, oldRowStubsArray(rowIdxBase + 2).ColIndex)
                Set endRow = source.Cells(oldRowStubsArray(rowIdxBase + 3).RowIndex, oldRowStubsArray(rowIdxBase + 3).ColIndex)
            
                Set startHeader = source.Cells(oldcolHeadersArray(colIdxBase + 1).RowIndex, oldcolHeadersArray(colIdxBase + 1).ColIndex)
                Set endHeader = source.Cells(oldcolHeadersArray(colIdxBase + 2).RowIndex, oldcolHeadersArray(colIdxBase + 2).ColIndex)
            
                Set rowSlice = source.Range(source.Rows(startRow.Row), source.Rows(endRow.Row))
                Set colSlice = source.Range(source.Columns(startHeader.Column), source.Columns(endHeader.Column))
                
                Set sourcePrevYear = Application.Intersect(rowSlice, colSlice)
                Set sourcePrevYear = sourcePrevYear.Resize(sourcePrevYear.Rows.Count - 1, sourcePrevYear.Columns.Count)
            
            Else
                MsgBox "Source sheet ran out of headers/stubs at group " & j
                GoTo CleanUp
            End If
                        
prevYear.Borders.LineStyle = xlNone
    If prevYear.Row > 1 Then prevYear.Rows(prevYear.Rows.Count).Offset(1).Borders(xlEdgeBottom).LineStyle = xlNone
    Dim edges As Variant
    edges = Array(xlEdgeLeft, xlEdgeTop, xlEdgeBottom, xlEdgeRight)

    For r = 1 To prevYear.Rows.Count
        For c = 1 To prevYear.Columns.Count
            Dim e As Variant
            For Each e In edges
                If r = prevYear.Rows.Count And e = xlEdgeBottom Then
                    With prevYear.Cells(r, c).Borders(xlEdgeBottom)
                        .LineStyle = xlContinuous
                        .Weight = xlThin
                        .ColorIndex = xlAutomatic
                    End With
                ElseIf r > 2 Then
                    If currYear.Cells(r, c).Borders(e).LineStyle <> xlNone Then
                        With prevYear.Cells(r, c).Borders(e)
                            .LineStyle = currYear.Cells(r, c).Borders(e).LineStyle
                            .Color = currYear.Cells(r, c).Borders(e).Color
                            .Weight = currYear.Cells(r, c).Borders(e).Weight
                        End With
                    End If
                End If
            Next e
        Next c
    Next r
            ' replace the row stubs
            target.Range(target.Cells(rowStubsBegin.RowIndex, rowStubsBegin.ColIndex), target.Cells(rowStubsMiddle.RowIndex - 1, rowStubsMiddle.ColIndex)).Value = target.Range(target.Cells(rowStubsMiddle.RowIndex, rowStubsMiddle.ColIndex), target.Cells(rowStubsEnd.RowIndex - 1, rowStubsEnd.ColIndex)).Value
            rowStubsBegin.contents = "At 1 January 2024"
            target.Cells(rowStubsBegin.RowIndex, rowStubsBegin.ColIndex).Value = rowStubsBegin.contents 'overwrite rowStubsBegin on target sheet
            rowStubsMiddle.contents = "At 31 December 2024 and 1 January 2025"
            target.Cells(rowStubsMiddle.RowIndex, rowStubsBegin.ColIndex).Value = rowStubsMiddle.contents
            rowStubsEnd.contents = "At 31 December 2025"
            target.Cells(rowStubsEnd.RowIndex, rowStubsEnd.ColIndex).Value = rowStubsEnd.contents
            
            'now update and replace that values
            prevYear.Value2 = sourcePrevYear.Value2
            Call FormatAsAccountingCustom(prevYear)
            'resize currYear and prevYear to include the last line, so we can paste formulas
            Set prevYear = Application.Intersect(target.Range(target.Cells(rowStubsBegin.RowIndex, rowStubsBegin.ColIndex), target.Cells(rowStubsMiddle.RowIndex, rowStubsMiddle.ColIndex)).EntireRow, target.Range(target.Cells(colHeadersArray(2 * j - 1).RowIndex, colHeadersArray(2 * j - 1).ColIndex), target.Cells(colHeadersArray(2 * j).RowIndex, colHeadersArray(2 * j).ColIndex)).EntireColumn)
            Set prevYear = prevYear.Offset(1, 0).Resize(prevYear.Rows.Count - 1, prevYear.Columns.Count)
            Set currYear = prevYear.Offset(rowStubsMiddle.RowIndex - rowStubsBegin.RowIndex)
            
            For r = 1 To prevYear.Rows.Count
                For c = 1 To prevYear.Columns.Count
                    If currYear.Cells(r, c).HasFormula = True And CStr(currYear.Cells(r, c).Formula) Like "*!*" = False Then
                        If CStr(currYear.Cells(r, c).Formula) Like "=SUM*" Then
                            prevYear.Cells(r, c).FormulaR1C1 = currYear.Cells(r, c).FormulaR1C1
                        End If
                        
                        If CStr(currYear.Cells(r, c).Formula) Like "*+*" Then
                            prevYear.Cells(r, c).FormulaR1C1 = currYear.Cells(r, c).FormulaR1C1
                        End If
                    End If
                Next c
            Next r
            
            r = 1
            c = prevYear.Columns.Count
            'r,c now gives coordinates for top right corner
            prevYear.Cells(r - 1, c).FormulaR1C1 = "=SUM(RC" & colHeadersArray(2 * j - 1).ColIndex & ":RC[-1])"
            prevYear.Cells(r - 1, c).Formula = Replace(CStr(prevYear.Cells(r - 1, c).Formula), "$", "")
    Next j
            
            
            
            
            
CleanUp:
    ' remember only objects require cleanup
    Set searchRange = Nothing
    Set source = Nothing
    Set target = Nothing
    Set twinObj = Nothing
    Set startRow = Nothing
    Set endRow = Nothing
    Set startHeader = Nothing
    Set endHeader = Nothing
    Set sourcePrevRange = Nothing
    Set rowSlice = Nothing
    Set colSlice = Nothing
    
    Erase rowStubsArray
    Erase searchRangeValue
    Erase searchRangeFormula
    Erase months
    
    If IsArray(currYearFormula) Then
        Erase currYearFormula
    Else
        currYearFormula = Empty
    End If
    
    If IsArray(currYearAddress) Then
        Erase currYearAddress
    Else
        currYearAddress = Empty
    End If
    
    If IsArray(prevYearFormula) Then
        Erase prevYearFormula
    Else
        prevYearFormula = Empty
    End If
    
    If IsArray(prevYearAddress) Then
        Erase prevYearAddress
    Else
        prevYearAddress = Empty
    End If

End Function

Function identifyRowStubsInTarget(target As Worksheet) As SOCE_cell()
    Dim rawRange As Range
    Dim formulaArray As Variant, valueArray As Variant
    Dim output() As SOCE_cell
    Dim r As Long, c As Long, j As Long, i As Long
    
    Set rawRange = Application.Intersect(target.usedRange, target.Range("A1:Z100"))
    If rawRange Is Nothing Then Exit Function
    
    formulaArray = rawRange.Formula
    valueArray = rawRange.Value
    
    ReDim output(1 To 99)
    j = 1
    
    Dim months As Variant
    months = Array("January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December")
    
    Dim sameCol As Long
    sameCol = 0
    
    For r = 1 To UBound(formulaArray, 1)
        For c = 1 To UBound(formulaArray, 2)
            For i = LBound(months) To UBound(months)
                Dim exists: exists = (c < UBound(valueArray, 2))
                Dim hasData As Boolean: hasData = False
                If exists Then
                    If c + 1 <= UBound(valueArray, 2) Then
                        If Not IsEmpty(valueArray(r, c + 1)) Then hasData = True
                    End If
                    
                    ' If c+1 was empty, check c+2
                    If Not hasData And c + 2 <= UBound(valueArray, 2) Then
                        If Not IsEmpty(valueArray(r, c + 2)) Then hasData = True
                    End If
                    ' incredible. this sheet has the headers 3 columns away from the row stubs. "Highland Group 2024"
                    If Not hasData And c + 3 <= UBound(valueArray, 2) Then
                        If Not IsEmpty(valueArray(r, c + 3)) Then hasData = True
                    End If
                    
                End If
                
                If (CStr(formulaArray(r, c)) Like "*" & months(i) & " [12][0-9][0-9][0-9]*" Or _
                   CStr(valueArray(r, c)) Like "*" & months(i) & " [12][0-9][0-9][0-9]*") And hasData Then
                    If sameCol <> 0 And sameCol <> c Then
                         ReDim output(1 To 99)
                         sameCol = 0
                         j = 1
                        output(j).RowIndex = r + rawRange.Row - 1
                          output(j).ColIndex = c + rawRange.Column - 1
                          output(j).contents = CStr(valueArray(r, c))
                          j = j + 1
                     Else
                          output(j).RowIndex = r + rawRange.Row - 1
                          output(j).ColIndex = c + rawRange.Column - 1
                          output(j).contents = CStr(valueArray(r, c))
                          
                          If sameCol = 0 Then sameCol = c
                          j = j + 1
                          Exit For
                     End If
                End If
            Next i
        Next c
    Next r
    
    If j > 1 Then
        ReDim Preserve output(1 To j - 1)
        identifyRowStubsInTarget = output
    End If
    Erase months
    
End Function

Function identifyColHeadersInTarget(target As Worksheet) As SOCE_cell()
    Dim rawRange As Range
    Dim valueArray As Variant
    Dim output() As SOCE_cell
    Dim r As Long, c As Long, j As Long, i As Long, k As Long, x As Long
    Dim headers As Variant, currencies As Variant
    
    headers = Array("capital", "total")
    currencies = Array("US$", "S$", "$")
    
    Set rawRange = Application.Intersect(target.usedRange, target.Range("A1:Z100"))
    If rawRange Is Nothing Then Exit Function
    
    valueArray = rawRange.Value
    If IsEmpty(valueArray) Then Exit Function
    
    ReDim output(1 To 500)
    j = 1
    
    For r = LBound(valueArray, 1) To UBound(valueArray, 1) - 1
        For c = LBound(valueArray, 2) To UBound(valueArray, 2)
        Dim cellAlreadyFound As Boolean
        cellAlreadyFound = False
            For i = LBound(headers) To UBound(headers)
                
                If InStr(1, CStr(valueArray(r, c)), headers(i), vbTextCompare) > 0 Then
                    For k = LBound(currencies) To UBound(currencies)
                        For x = 1 To 2 'this x is in case I encounter one with "total" two cells above the currency sign (see KKF Asset Global, 2024)
                            If InStr(1, CStr(valueArray(r + x, c)), currencies(k), vbTextCompare) > 0 Then
                                output(j).RowIndex = r + rawRange.Row - 1
                                output(j).ColIndex = c + rawRange.Column - 1
                                output(j).contents = CStr(valueArray(r, c))
                                
                                j = j + 1
                                If j > UBound(output) Then
                                    ReDim Preserve output(1 To j + 100)
                                End If
                                GoTo NextColumn
                            End If
                        Next x
                    Next k
                    If j > 1 Then If output(j - 1).RowIndex = r + rawRange.Row - 1 And output(j - 1).ColIndex = c + rawRange.Column - 1 Then Exit For
                End If
                
            Next i
NextColumn:
        Next c
    Next r
    
    If j > 1 Then
        ReDim Preserve output(1 To j - 1)
        identifyColHeadersInTarget = output
    End If
End Function

Private Sub Class_Terminate()
    Set rowStubs = Nothing
    Set tableHeaders = Nothing
    Set prevYearValues = Nothing
    Set currentYearValues = Nothing
    Set Instances = Nothing
End Sub
