Attribute VB_Name = "SAPConnectionModule"
Option Explicit

Private logonControl As SAPLogonControl

' 业务组件都需要指定Connection对象，所以设为Public
Public sapConnection As SAPLogonCtrl.Connection

Public Sub Logon()
    If sapConnection Is Nothing Then
        Set logonControl = New SAPLogonCtrl.SAPLogonControl
        Set sapConnection = logonControl.NewConnection()
    End If
    Call sapConnection.Logon(0, False) 'parameter: hwnd, silent logon
End Sub

Public Sub logoff()
    If Not sapConnection Is Nothing Then
        If sapConnection.IsConnected = tloRfcConnected Then
            Call sapConnection.logoff
        End If
    End If
End Sub




