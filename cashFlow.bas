Attribute VB_Name = "cashFlow"
Sub cashFlowIdentifier(twinObj As ClsSheetTwin)
    ' first make sure it is not CBS/CPL
    If IsCPLorCBS(twinObj) = True Then Exit Sub
    ' make sure it is not tax as well
    If InStr(1, CStr(twinObj.source.Name), "tax", vbTextCompare) > 0 Then Exit Sub
    ' make sure it is not FMC_AJE
    If Not twinObj.source.usedRange.Find(What:="Being tax provision for", LookIn:=xlValues, LookAt:=xlPart, MatchCase:=False) Is Nothing Then Exit Sub
    Application.ScreenUpdating = False
    'make sure it is not FS
    If isFS(twinObj) = True Then Exit Sub
    'make sure it is not CSOCE (because my CF identifies SOCE sometimes)
    If IsValidSOCE(twinObj) Then Exit Sub
    ' now make sure it is CF
    If isCF(twinObj) = True Then

        Dim searchRange As Range
        Dim source As Worksheet
        Dim target As Worksheet
        Set source = twinObj.source
        Set target = twinObj.target
        Set searchRange = Application.Intersect(source.usedRange, source.Range("A1:Z300"))
        Dim dataArrFormula As Variant
        Dim dataArrValue As Variant
        
        Application.AskToUpdateLinks = False
        Application.DisplayAlerts = False
        
        dataArrFormula = searchRange.Formula
        dataArrValue = searchRange.Value
        
        Application.AskToUpdateLinks = True
        Application.DisplayAlerts = True
        
        Dim r As Long, c As Long, i As Variant ' iterators
        Dim sourceCurrentYear As Range, targetPriorYear As Range
        
        Dim currentStreak As Long, recordStreak As Long, longestRow As Long, longestColBegin As Long
        For r = 1 To UBound(dataArrFormula, 1)
            currentStreak = 0
            For c = 1 To UBound(dataArrFormula, 2)
                If InStr(1, dataArrFormula(r, c), "=") Or IsNumeric(CStr(dataArrFormula(r, c))) Then
                ' streak only counts formulas or strictly numeric values
                    If r < UBound(dataArrFormula, 1) - 1 Then
                        If InStr(1, dataArrFormula(r + 1, c), "=") > 0 Or IsNumeric(CStr(dataArrFormula(r + 1, c))) Then
                            If InStr(1, dataArrFormula(r + 2, c), "=") > 0 Or IsNumeric(CStr(dataArrFormula(r + 2, c))) Then
                                currentStreak = currentStreak + 1
                                If currentStreak > recordStreak Then
                                    recordStreak = currentStreak
                                    longestRow = r
                                    longestColBegin = c - recordStreak
                                End If
                            End If
                        End If
                    End If
                End If
            Next c
        Next r
        
        'isCF has checked for empty sheet
        If recordStreak <> 0 And longestRow <> 0 And longestColBegin <> 0 Then
            For c = longestColBegin To longestColBegin + recordStreak
                If Left(CStr(dataArrFormula(longestRow, c)), 1) <> "=" Then
                    'only replace the numeric values. remember that the block only contains formulas or numeric, so we have filtered out everything
                    dataArrFormula(longestRow, c) = dataArrValue(longestRow + 1, c)
                End If
            Next c
        End If
        
        If currentStreak = recordStreak Then
            MsgBox "Error. Longest continuous range in '" & twinObj.source.Name & "' could not be found." & _
            " (there are two ranges with identical length, so the range of cash flow values could not be identified."
            GoTo CleanUp
        End If
        'so external links crash excel for some reason. we sanitize and get rid of it, replacing with value
        For r = 1 To UBound(dataArrFormula, 1)
            For c = 1 To UBound(dataArrFormula, 2)
                If InStr(1, dataArrFormula(r, c), "[") > 0 Then
                    dataArrFormula(r, c) = dataArrValue(r, c)
                End If
            Next c
        Next r
        
        ' now we dump everything back
        target.Range(searchRange.Address).Formula = dataArrFormula
        
                    
                            
    
    End If
                
            
CleanUp:
Set priorYearRow = Nothing
Set currentYearRow = Nothing
strI = ""
Set searchRange = Nothing
Set source = Nothing
Set target = Nothing
Set rPrior = Nothing
Set rCurrent = Nothing
If Not IsEmpty(dataArrValues) Then Erase dataArrValues

                
    
End Sub

Public Function IsCPLorCBS(twinObj As ClsSheetTwin) As Boolean
    ' Returns True if the sheet is identified as a BS or PL variant
    On Error GoTo CleanUp
    
    Dim source As Worksheet
    Set source = twinObj.source
    
    ' 1. Check Name via Select Case
    Select Case UCase(Trim(source.Name))
        Case "CBS", "CPL", "FC_BS", "FC_P&L", "FMC_BS", "FMC_P&L", "FMC_BS 2024", "FMC_PL 2024", _
             "FC_BS 2024", "FC_PL 2024", "SBS", "SPL", "PROFIT AND LOSS", "BALANCE SHEET", "BS", "P&L"
            IsCPLorCBS = True
            Exit Function ' Found by name, no need to scan cells
    End Select

    ' 2. Check Content (Year Header Scan)
    Dim searchRange As Range
    Set searchRange = Application.Intersect(source.Range("A1:Z300"), source.usedRange)
    
    If Not searchRange Is Nothing Then
        Dim dataArr As Variant
        dataArr = searchRange.Value ' Load to RAM
        
        If IsArray(dataArr) Then
            Dim r As Long, c As Long, occurrenceCount As Long
            
            ' Loop through rows
            For r = 1 To UBound(dataArr, 1)
                occurrenceCount = 0
                ' Loop through columns
                For c = 1 To UBound(dataArr, 2)
                    ' CStr handles numbers/errors safely
                    If Trim(CStr(dataArr(r, c))) = "2024" Then
                        occurrenceCount = occurrenceCount + 1
                    End If
                Next c
                
                ' Threshold met
                If occurrenceCount > 4 Then
                    IsCPLorCBS = True
                    GoTo CleanUp
                End If
            Next r
        End If
    End If

CleanUp:
    ' Wipe the RAM and kill the pointers
    If IsArray(dataArr) Then Erase dataArr
    Set searchRange = Nothing
    Set source = Nothing
End Function

Public Function isCF(twinObj As ClsSheetTwin) As Boolean
    isCF = False
    Dim source As Worksheet
    Set source = twinObj.source
    If twinObj.source.usedRange.Cells.Count = 1 And IsEmpty(source.usedRange.Cells(1, 1)) Then
    ' guard against empty sheet
        isCF = False
        GoTo CleanUp
    End If
    Dim searchRange As Range
    Dim target As Worksheet
    Set target = twinObj.target
    Set searchRange = Application.Intersect(source.usedRange, source.Range("A1:Z300"))
    Dim matchString As Variant
    matchString = Array("share", "capital", "retained", "SDL", "shareholder", "director", "tax payable", "cash equivalents", "CACE", "amount due")
    Dim currencies As Variant
    currencies = Array("US$", "S$", "$")
    Dim currencyMatch As Long
    Dim i As Variant
    
    'To prevent further bloating we write everything to RAM
    Dim dataArrFormulas As Variant
    dataArrFormulas = searchRange.Formula
    
    Dim r As Long, c As Long, matchCount As Long
    For r = 1 To UBound(dataArrFormulas, 1)
        matchCount = 0
        currencyMatch = 0
        For c = 1 To UBound(dataArrFormulas, 2)
            If InStr(1, dataArrFormulas(r, c), "Prior year bal", vbTextCompare) > 0 Then
                If InStr(1, dataArrFormulas(r + 1, c), "Current year bal", vbTextCompare) > 0 Then
                    isCF = True
                    GoTo CleanUp
                End If
            End If
            
            For Each i In matchString
                If InStr(1, CStr(dataArrFormulas(r, c)), i, vbTextCompare) > 0 Then
                    matchCount = matchCount + 1
                ElseIf r < UBound(dataArrFormulas, 1) Then 'boundary check for array
                    If InStr(1, CStr(dataArrFormulas(r + 1, c)), i, vbTextCompare) > 0 Then
                        matchCount = matchCount + 1
                    End If
                End If
            Next i
            
            If matchCount > 3 Then
                isCF = True
                GoTo CleanUp
            End If
            
            For Each i In currencies
                If UCase(Trim(CStr(dataArrFormulas(r, c)))) = i Then
                    currencyMatch = currencyMatch + 1
                End If
            Next i
            
            If currencyMatch > 4 Then
                isCF = True
                GoTo CleanUp
            End If
            
            If InStr(1, CStr(dataArrFormulas(r, c)), "Cash flows from operating activities", vbTextCompare) > 0 Or _
            InStr(1, CStr(dataArrFormulas(r, c)), "Cash flow from operating activities", vbTextCompare) > 0 Then
                If isFS(twinObj) Then
                    isCF = True
                    GoTo CleanUp
                End If
            End If
                
        Next c
    Next r
    
CleanUp:
    Set searchRange = Nothing
    Set source = Nothing
    Set target = Nothing
    If Not IsEmpty(dataArrFormulas) Then Erase dataArrFormulas
    If Not IsEmpty(matchString) Then Erase matchString
    If Not IsEmpty(currencies) Then Erase currencies
    i = Empty ' since it iterated through strings, it's a pointer to strings
End Function
