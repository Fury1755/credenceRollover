Attribute VB_Name = "rolloverMaster"
Sub InitializeRolloverWorkbooks()
Attribute InitializeRolloverWorkbooks.VB_ProcData.VB_Invoke_Func = "I\n14"
    Dim sourceWb As Workbook, targetWb As Workbook
    Dim sourcePath As String, targetPath As String
    Dim twinList As Collection
    Dim twinObj As ClsSheetTwin
    Dim i As Long
    
    On Error GoTo ErrorHandler
    sourcePath = GetFilePath("Select the SOURCE Workbook (Copy From)")
    If sourcePath = "" Then Exit Sub
    
    targetPath = GetFilePath("Select the TARGET Workbook (Paste To)")
    If targetPath = "" Then Exit Sub
    ToggleWaitMode True
    Debug.Print "--- NEW RUN STARTED: " & Now & " ---"

    Set sourceWb = Workbooks.Open(sourcePath, ReadOnly:=True)
    Set targetWb = Workbooks.Open(targetPath, ReadOnly:=False)
    
    Set twinList = New Collection
    Call BuildClassTwins(sourceWb, targetWb, twinList)
    
    Debug.Print "Twins Found: " & twinList.Count

    If twinList.Count = 0 Then
        MsgBox "No matching sheets found!", vbExclamation
        GoTo Cleanup
    End If

    For i = 1 To twinList.Count
    Debug.Print "Index: " & i & " | Sheet: " & twinList(i).target.Name
            ' Debug.Print "MASTER: Entering Iteration " & i & ": " & twinList(i).source.Name
            
            Set twinObj = twinList(i)
                Call shiftColumnsInTwin(twinObj.source, twinObj.target, twinObj)
                'shiftColumnsInTwin is self contained; i.e. runs independently
                Call SOCE_Identifier(twinObj) ' also self contained
                Call CPLorCBSIdentifier(twinObj)
                Call cashFlowIdentifier(twinObj)
                DoEvents
                
            
            If Err.Number <> 0 Then
                Debug.Print "MASTER: Error caught in loop " & twinList(i).source.Name & ": " & Err.Description
                Err.Clear
            Else
                ' Debug.Print "MASTER: Returned from Worker " & twinList(i).source.Name & " with no errors."
            End If
            On Error GoTo 0
            DoEvents
        Next i
    
    Call ForceFullRecalc
Cleanup:
    Debug.Print "Cleaning up memory..."
    
    On Error Resume Next
    ToggleWaitMode False
    If Not twinList Is Nothing Then
        For i = twinList.Count To 1 Step -1
            twinList.Remove i
        Next i
    End If
    Set sourceWb = Nothing
    Set targetWb = Nothing
    Set twinList = Nothing
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Critical Error: " & Err.Description, vbCritical
    Debug.Print "Error " & Err.Number & ": " & Err.Description
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Resume Cleanup
End Sub
Private Sub ToggleWaitMode(ByVal bWait As Boolean)
    With Application
        .ScreenUpdating = Not bWait
        .EnableEvents = Not bWait
        .AskToUpdateLinks = Not bWait
        .DisplayAlerts = Not bWait
        If bWait Then
            .Calculation = xlCalculationManual
        Else
            .Calculation = xlCalculationAutomatic
        End If
    End With
End Sub


' --- Helper Function to handle File Explorer ---
