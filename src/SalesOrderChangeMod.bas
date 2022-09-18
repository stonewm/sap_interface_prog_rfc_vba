Attribute VB_Name = "SalesOrderChangeMod"
Option Explicit


Public Function ChangeSalesOrder(OrderNo As String, ItemNo As Integer, NewPricing As String) As String
    Dim functions As New SAPFunctionsOCX.SAPFunctions
    Dim func As SAPFunctionsOCX.Function
    Dim commitFunc As SAPFunctionsOCX.Function
    Dim orderItemIn As SAPTableFactoryCtrl.Table
    Dim orderItemInX As SAPTableFactoryCtrl.Table
    Dim returnTable As SAPTableFactoryCtrl.Table
    
    Dim retVal As String '��������ֻ
    retVal = ""
    
    ' sapConnection is global
    If sapConnection Is Nothing Then
        MsgBox "���¼SAPϵͳ!", vbOKOnly + vbInformation
        Exit Function
    End If
    
    If sapConnection.IsConnected <> tloRfcConnected Then
        MsgBox "���¼SAPϵͳ!", vbOKOnly + vbInformation
        Exit Function
    End If
    
    Set functions.Connection = sapConnection
    Set func = functions.Add("BAPI_SALESORDER_CHANGE")
    
    ' BAPI����-Importing
    func.Exports("SALESDOCUMENT").Value = OrderNo              ' ���۶�����
    func.Exports("ORDER_HEADER_INX").Value("UPDATEFLAG") = "U" ' U��ʾ�޸�
    
    ' BAPI����-Pricing(��LOGIC_SWITCH������)
    func.Exports("LOGIC_SWITCH").Value("PRICING") = NewPricing
    func.Exports("LOGIC_SWITCH").Value("COND_HANDL") = "X"
     
    'BAPI����-ORDER_ITEM_IN / ORDER_ITEM_IN
    Set orderItemIn = func.Tables("ORDER_ITEM_IN")
    Set orderItemInX = func.Tables("ORDER_ITEM_INX")
     
    orderItemIn.AppendRow
    orderItemIn.Value(1, "ITM_NUMBER") = ItemNo
    
    orderItemInX.AppendRow
    orderItemInX.Value(1, "ITM_NUMBER") = ItemNo
    orderItemInX.Value(1, "UPDATEFLAG") = "U"
    
    'BAPI����-����ֵ
    Set returnTable = func.Tables("RETURN")
    'ִ�к���
    If func.Call = False Then
        retVal = DumpReturn(returnTable)
        Exit Function
    Else
        retVal = DumpReturn(returnTable)
        Dim returnOfCommit As SAPTableFactoryCtrl.Table
        Set commitFunc = functions.Add("BAPI_TRANSACTION_COMMIT")
        commitFunc.Exports("WAIT").Value = "X"
        Set returnOfCommit = commitFunc.Tables("RETURN")
        
        If commitFunc.Call = False Then
            MsgBox func.Exception
            Exit Function
        End If
    End If
    
    ChangeSalesOrder = retVal
End Function

'----------------------------
' ��ȡ�����ķ���ֵ
'----------------------------
Private Function DumpReturn(ret As SAPTableFactoryCtrl.Table) As String
    Dim retVal As String
    retVal = ""

    If Not ret Is Nothing Then
        If ret.rowcount > 0 Then
            retVal = "��Ϣ���� " & ret.Value(ret.rowcount, 1) & "," & ret.Value(ret.rowcount, 4)
        End If
    End If
    
    DumpReturn = retVal
End Function

Public Sub RunScript()
    Dim i As Long
    Dim returnVal As String
    
    For i = 4 To Sheet3.UsedRange.rows.Count
        If Sheet3.Range("A" & i).Value = "EOF" Then Exit Sub
        
        Dim leftCell As Range
        Set leftCell = Sheet3.Range("A" & i)
        returnVal = ChangeSalesOrder(leftCell.Value, leftCell.Offset(0, 1).Value, leftCell.Offset(0, 2).Value)
        leftCell.Offset(0, 3).Value = returnVal
    Next
End Sub


'Public Sub TestChangeSalesOrder()
'    Call ChangeSalesOrder("0061000702", 10, "B")
'End Sub
