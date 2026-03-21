VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClsSheetTwin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public source As Worksheet
Public target As Worksheet
Public tgtColOld As Range
Public tgtColNotes As Range
Public usedRange As Range
Public isFinancialStatement As Boolean


' Constructor (what it does when it gets set up)
Public Sub Init(srcWs As Worksheet, tgtWs As Worksheet)

Set Me.source = srcWs
Set Me.target = tgtWs
Set usedRange = Intersect(source.usedRange, source.Range("A1", "Z100"))

End Sub

' Method (what this class can do)
Private Sub Class_Terminate()
Set Me.source = Nothing
Set Me.target = Nothing
Set Me.tgtColOld = Nothing
Set Me.tgtColNotes = Nothing
Set usedRange = Nothing
isFinancialStatement = False
End Sub
