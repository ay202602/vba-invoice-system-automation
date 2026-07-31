Attribute VB_Name = "CreateInvoice"
Option Explicit

Sub CreateInvoice()
    Dim ws As Worksheet
    Set ws = Worksheets("売上")
    
    ' 売上シートのA列の最終行を取得
    Dim LastRow As Long
    LastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
    If LastRow < 2 Then
        MsgBox "売上リストが存在しません。", vbCritical
        Exit Sub
    End If
    
    Dim InvWs As Worksheet
    Set InvWs = Worksheets("請求書")
    
    Dim CustomerKey As String
    CustomerKey = InvWs.Range("B2").Value
    
    Dim inv As CInvoice
    Set inv = New CInvoice
    
    inv.init _
        Customer:=ws.Range("C1"), _
        CellStart:=ws.Range("A1"), _
        PasteCells:=InvWs.Range("A5:G17")

    Dim Search As Range
    Set Search = inv.CellSearch(CustomerKey)
    
    If CustomerKey = "" Then
        MsgBox "B2セルに顧客名を入力してください。", vbCritical
        Exit Sub
    ElseIf Search Is Nothing Then
        MsgBox CustomerKey & "は売上リストの顧客名には存在しません。", vbCritical
        Exit Sub
    Else
        Call inv.CellClear
        Call inv.FilterOn(Columns:=3, Value:=CustomerKey)
        Call inv.CellHide(True)
        Call inv.CellCopy
        Call inv.CellPaste
        Call inv.CellHide(False)
        Call inv.FilterOff
    End If
End Sub

