Attribute VB_Name = "CPL_orCBS"
Public Sub CPLorCBSIdentifier(twinObj As ClsSheetTwin)
    Set source = twinObj.source
    Dim isIdentified As Boolean
    isIdentified = False
    
    Select Case UCase(Trim(source.Name))
        Case "CBS", "CPL", "FC_BS", "FC_P&L", "FMC_BS", "FMC_P&L", "FMC_BS 2024", "FMC_PL 2024", _
        "FC_BS 2024", "FC_PL 2024", "SBS", "SPL", "PROFIT AND LOSS", "BALANCE SHEET", "BS", "P&L", _
        "BS", "PL"
            isIdentified = True
    End Select
    
    If isFS(twinObj) = False Or isIdentified = True Then ' note: sometimes CBS and CPL can have no gridlines.
        'is not a FS thing
        Dim dataArr As Variant
        Dim formulaArr As Variant
        Dim searchRange As Range
        Dim targetRange As Range
        Dim targetStr As String
        Set searchRange = Application.Intersect(source.Range("A1", "Z300"), source.usedRange)
        If Not searchRange Is Nothing Then dataArr = searchRange.Value
        If Not searchRange Is Nothing Then formulaArr = searchRange.Formula
        If searchRange Is Nothing Then GoTo CleanUp
        ' dataArr is loaded into RAM
        If IsArray(dataArr) Then
            For r = 1 To UBound(dataArr, 1)
                occuranceCount = 0
                
                For c = 1 To UBound(dataArr, 2)
                    If Trim(CStr(dataArr(r, c))) = "2024" Then occuranceCount = occuranceCount + 1
                Next c
        
                If occuranceCount > 4 Then
                    isIdentified = True
                End If
            Next r
        End If
        
        If isIdentified = True Then
            Debug.Print twinObj.source.Name & " identified as CPL/CBS"
            Call ClearNumbersAndFormatCustom(twinObj.target.Range(searchRange.Address))
            For r = 1 To UBound(dataArr, 1)
                For c = 1 To UBound(dataArr, 2)
                    Set targetRange = twinObj.target.Range(searchRange.Address).Cells(r, c)
                    Set sourceRange = twinObj.source.Range(searchRange.Address).Cells(r, c)
                    If Left(CStr(formulaArr(r, c)), 1) <> "=" And UCase(Trim(dataArr(r, c))) = "2024" Then
                        dataArr(r, c) = "2025"
                        targetRange.Value = dataArr(r, c)
                    End If
                    If UCase(Trim(dataArr(r, c))) = "AS AT 31 DECEMBER 2024" Then
                        targetRange.Value = "As at 31 December 2025"
                    End If
                    If UCase(Trim(dataArr(r, c))) = "FOR THE YEAR ENDED 31 DECEMBER 2024" Then
                        targetRange.Value = "For the year ended 31 December 2025"
                    End If
                    If Left(CStr(formulaArr(r, c)), 1) = "=" And Right(CStr(formulaArr(r, c)), 2) = "+1" Then
                        targetStr = CStr(formulaArr(r, c))
                        targetRange.Formula = Left(targetStr, Len(targetStr) - 2)
                        targetRange.ClearComments
                    End If
                    If Left(CStr(formulaArr(r, c)), 1) = "=" And Right(CStr(formulaArr(r, c)), 2) = "-1" Then
                        targetStr = CStr(formulaArr(r, c))
                        targetRange.Formula = Left(targetStr, Len(targetStr) - 2)
                        targetRange.ClearComments
                    End If
                    Set targetRange = Nothing
                Next c
            Next r
        End If
    End If
    
CleanUp:
    Set searchRange = Nothing
    targetStr = ""
    Set targetRange = Nothing
    Set source = Nothing
    If Not IsEmpty(dataArr) Then
        Erase dataArr
    End If
End Sub
