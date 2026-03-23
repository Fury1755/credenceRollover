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
    rowStubsCollection = Array("At 1 January 2023", "At 31 December 2023 and 1 January 2024", "At 31 December 2024", "1 January 2023")
    Dim tableValues As Range
    Dim rowStubIsNothing As Long
    
    If twinObj.usedRange Is Nothing Then GoTo Cleanup
    For Each cell In twinObj.usedRange
        For i = LBound(headerCollection) To UBound(headerCollection)
            If InStr(1, CStr(cell.Value), headerCollection(i)) > 0 And InStr(1, CStr(cell.Offset(1, 0).Value), "S$") > 0 Then
                    If headerCollection(i) = "capital" Then
                        Set tableHeadersBegin = cell
                    ElseIf UCase(Trim(headerCollection(i))) = "TOTAL" Then
                        Set tableHeadersEnd = cell
                    End If
            End If
        Next i
        
        For j = LBound(rowStubsCollection) To UBound(rowStubsCollection)
            If InStr(1, CStr(cell.Value), rowStubsCollection(j)) > 0 Then
                If rowStubsCollection(j) = "At 1 January 2023" Or rowStubsCollection(j) = "1 January 2023" Then
                    Set rowStubsBegin = cell
                ElseIf rowStubsCollection(j) = "At 31 December 2023 and 1 January 2024" Then
                    Set rowStubsMiddle = cell
                ElseIf rowStubsCollection(j) = "At 31 December 2024" Then
                    Set rowStubsEnd = cell
                End If
            End If
        Next j
        
        
        'now that we have both pairs, we create a new instance
        If (Not rowStubsEnd Is Nothing) And (Not tableHeadersEnd Is Nothing) Then
                Set Me.rowStubs = source.Range(rowStubsBegin, rowStubsEnd)
                Set Me.tableHeaders = source.Range(tableHeadersBegin, tableHeadersEnd)
                Set tableValues = Application.Intersect(Me.rowStubs.EntireRow, Me.tableHeaders.EntireColumn)
                If rowStubs.Count Mod 2 = 1 Then
                    Set Me.prevYearValues = Application.Intersect(source.Range(rowStubsBegin, rowStubsMiddle.Offset(-1, 0)).EntireRow, Me.tableHeaders.EntireColumn)
                    Set Me.currentYearValues = Application.Intersect(source.Range(rowStubsMiddle, rowStubsEnd.Offset(-1, 0)).EntireRow, Me.tableHeaders.EntireColumn)
                ElseIf rowStubs.Count Mod 2 = 0 Then
                    MsgBox "SOCE table in " & source.Name & " has even number of rows. Unable to process, check carefully."
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
End If

Cleanup:
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

Private Sub Class_Terminate()
    Set rowStubs = Nothing
    Set tableHeaders = Nothing
    Set prevYearValues = Nothing
    Set currentYearValues = Nothing
    Set Instances = Nothing
End Sub
