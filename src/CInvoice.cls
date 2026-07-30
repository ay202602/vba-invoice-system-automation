VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CInvoice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mCustomer As Range
Private mCellStart As Range
Private mPasteCells As Range

Public Sub init(Customer As Range, CellStart As Range, PasteCells As Range)
    Set mCustomer = Customer
    Set mCellStart = CellStart
    Set mPasteCells = PasteCells
End Sub

Public Sub CellHide(Value As Boolean)
    mCustomer.Columns.Hidden = Value
End Sub

Public Function CellSearch(Data As String) As Range
    With mCellStart.CurrentRegion
        Set CellSearch = .Offset(1, 0).Resize(.Rows.Count - 1, .Columns.Count).Find( _
            what:=Data, _
            LookIn:=xlValues, _
            lookat:=xlWhole _
            )
    End With
End Function

Public Sub CellClear()
    mPasteCells.ClearContents
End Sub

Public Sub CellCopy()
    With mCellStart.CurrentRegion
        .Offset(1, 0).Resize(.Rows.Count - 1, .Columns.Count).Copy
    End With
End Sub

Public Sub CellPaste()
    mPasteCells.PasteSpecial Paste:=xlPasteValues
End Sub

Public Sub FilterOn(Columns As Long, Value As String)
    mCellStart.AutoFilter FIeld:=Columns, Criteria1:=Value
End Sub

Public Sub FilterOff()
    mCellStart.AutoFilter
End Sub
