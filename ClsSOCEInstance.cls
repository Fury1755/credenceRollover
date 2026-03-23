VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClsSOCEInstance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public rowStubs As Range
Public tableHeaders As Range
Private currentYearValues As Range
Private prevYearValues As Range
Private pParent As ClsSOCETable ' we access this during cleanup

Public Function Init(Parent As ClsSOCETable, currYear As Range, prevYear As Range)
    If Not Parent Is Nothing Then Set Me.rowStubs = Parent.rowStubs
    Set Me.tableHeaders = Parent.tableHeaders
    Set currentYearValues = currYear
    Set prevYearValues = prevYear
    Set pParent = Parent
End Function

Public Function processValues(Parent As ClsSOCETable)
    Dim isCurrentAndPreviousSame As Boolean
    Dim i As Long
    Dim targetPrevYearValues As Range
    If Parent.currentYearValues.Rows.Count = Parent.prevYearValues.Rows.Count And Parent.currentYearValues.Columns.Count = Parent.prevYearValues.Columns.Count Then
        isCurrentAndPreviousSame = True
    Else
        isCurrentAndPreviousSame = False
        MsgBox "SOCE: current and previous years do not have the same number of rows. Cannot run. Location: " & Parent.currentYearValues.Worksheet.Name & " , " & Parent.currentYearValues.Address & " not matched to " & Parent.prevYearValues.Address
        Exit Function
    End If
    
    If isCurrentAndPreviousSame = True Then
        Set targetPrevYearValues = Parent.Parent.target.Range(prevYearValues.Address)
        ' now targetPrevYearValues is pointing to target
        ' loop through with single iterator
        For i = 1 To targetPrevYearValues.Count
            If cellHasHardValues(targetPrevYearValues.Cells(i)) = True Then
                targetPrevYearValues.Cells(i).Value2 = currentYearValues.Cells(i).Value2
            Else
                If cellHasHardValues(targetPrevYearValues.Cells(i)) = False Then
                    ' Debug.Print targetPrevYearValues.Cells(i).Address & " NOT A HARD VALUE"
                End If
            End If
        Next i
    End If
    
    ' now we update years in the row stubs
    Dim rowStubData As Variant
    rowStubData = Me.rowStubs.Value
    Call UpdateYearsInArray(rowStubData)
    Set targetRowStub = Parent.Parent.target.Range(Me.rowStubs.Address)
    targetRowStub.Value2 = rowStubData

End Function

Private Sub Class_Terminate()
    Set rowStubs = Nothing
    Set tableHeaders = Nothing
    Set currentYearValues = Nothing
    Set prevYearValues = Nothing
    Set pParent = Nothing
End Sub
