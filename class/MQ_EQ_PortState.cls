VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MQ_EQ_PortState"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements ts_ECAncestor.TaskAncestor

Private Const MODULE_ID = "MQ_EQ_PortState"

Dim Parent As tsEC.clsRunManager
Dim TaskID As String
Dim objBatch As Batch
Dim vParameter As Variant
Dim i As Integer

Private Sub TaskAncestor_Execute(Parameter As Variant)
    On Error GoTo ErrorHandler
    Dim objVfei As Message22
    
    i = 0
    Set objVfei = New vfei.Message22
    
    With objVfei
        .CommandID = "QUERY_STATUS"
        .MachineID = Parent.GetPuObject.EqDrvList(1).EqDrvID
        .MessageType = "C"
        .AddVar "PARAMETER", "A", "PORT_STATUS"
    End With
    Call Parent.SendVfei(objVfei, Me)
    Exit Sub
ErrorHandler:
    Call Parent.LogTrace(0, FormatErrMsg(MODULE_ID & ".Execute", Err, Err.Description))
End Sub

Private Property Set TaskAncestor_Parent(ByVal RHS As Object)
   Set Parent = RHS
End Property

Private Property Get TaskAncestor_Parent() As Object
   Set TaskAncestor_Parent = Parent

End Property



Private Sub TaskAncestor_ReceivedHostMsg(objHostTrans As HostOcx.HostTransaction, isPrimary As Boolean)
    On Error GoTo ErrorHandler
    Exit Sub
ErrorHandler:
    Call Parent.LogTrace(0, FormatErrMsg(MODULE_ID & ".ReceivedHostMsg", Err, Err.Description))
End Sub

Private Sub TaskAncestor_ReceivedVfei(objVfei As vfei.Message22)
    On Error GoTo ErrorHandler
    Dim TempLng As Long
    Dim Exp16 As Long
    Dim LBinaryStr As String
    Dim RBinaryStr As String
    Dim BinaryStr As String
    Dim i As Integer
    Dim sPort1Status As String
    Dim sPort2Status As String
    Dim objTmpVfei As Message22
    Dim objTmpSmif As SmifItem
    Dim sSmif1Status As String
    Dim sSmif2Status As String
    Dim RightStatus As Boolean
    
    If objVfei.CommandID <> "QUERY_STATUS" Then Exit Sub
    If objVfei.ErrorCode <> 0 Then
        Call Parent.LogTrace(0, MODULE_ID & " : Error description=" & objVfei.ErrorText)
        Exit Sub
    End If
    
    Select Case Left$(objVfei.MachineID, 2)
        Case "MI"
            RightStatus = False
            TempLng = CLng(objVfei.Var("PORT_STATUS").Value)
            '--------------------------------------------------------------------------
            'Change Format
            Exp16 = Exp(16 * Log(2))
            
            For i = 1 To 16
               Exp16 = Exp16 / 2
               If ((Exp16) And TempLng) = Exp16 Then
                  BinaryStr = BinaryStr & "1"
               Else
                  BinaryStr = BinaryStr & "0"
               End If
            Next i
            '--------------------------------------------------------------------------
            LBinaryStr = Left(BinaryStr, 8)
            RBinaryStr = Right(BinaryStr, 8)
            If (LBinaryStr And "10100001") = "10100001" Then
                Call Parent.LogTrace(0, MODULE_ID & " Port1 Door open")
                On Error Resume Next
                sPort1Status = Parent.GetPuObject.PortList(1).Attributes("PortStatus").Value
                If Err.Number <> 0 Then
                    Err.Clear
                    Parent.GetPuObject.PortList(1).Attributes.Add("PortStatus").Value = "Full"
                Else
                    Parent.GetPuObject.PortList(1).Attributes.Item("PortStatus").Value = "Full"
                End If
                On Error GoTo ErrorHandler
            Else
                Call Parent.LogTrace(0, MODULE_ID & " Port1 Door close")
                On Error Resume Next
                sPort1Status = Parent.GetPuObject.PortList(1).Attributes("PortStatus").Value
                If Err.Number <> 0 Then
                    Err.Clear
                    Parent.GetPuObject.PortList(1).Attributes.Add("PortStatus").Value = "Busy"
                Else
                    Parent.GetPuObject.PortList(1).Attributes.Item("PortStatus").Value = "Busy"
                End If
                On Error GoTo ErrorHandler
            End If
            If (RBinaryStr And "10100001") = "10100001" Then
               Call Parent.LogTrace(0, MODULE_ID & " Port2 Door open")
                On Error Resume Next
                sPort2Status = Parent.GetPuObject.PortList(2).Attributes("PortStatus").Value
                If Err.Number <> 0 Then
                    Err.Clear
                    Parent.GetPuObject.PortList(2).Attributes.Add("PortStatus").Value = "Full"
                Else
                    Parent.GetPuObject.PortList(2).Attributes.Item("PortStatus").Value = "Full"
                End If
                On Error GoTo ErrorHandler
            Else
               Call Parent.LogTrace(0, MODULE_ID & " Port2 Door close")
                On Error Resume Next
                sPort2Status = Parent.GetPuObject.PortList(2).Attributes("PortStatus").Value
                If Err.Number <> 0 Then
                    Err.Clear
                    Parent.GetPuObject.PortList(2).Attributes.Add("PortStatus").Value = "Busy"
                Else
                    Parent.GetPuObject.PortList(2).Attributes.Item("PortStatus").Value = "Busy"
                End If
                On Error GoTo ErrorHandler
            End If
            
            '
            For Each objTmpSmif In Parent.GetPuObject.SmifList
                Set objTmpVfei = Nothing
                Set objTmpVfei = New vfei.Message22
                With objTmpVfei
                    .CommandID = "QUERY_STATUS"
                    .MachineID = objTmpSmif.SMIFID
                End With
                Call Parent.SendVfei(objTmpVfei, Me)
            Next
        Case "SM"
            Select Case objVfei.MachineID
                Case "SMIF1"
                    i = i + 1
                    If objVfei.Var("POD_PRESENT").Value = False Then
                        Call Parent.LogTrace(0, MODULE_ID & " SMIF1 is not present")
                        On Error Resume Next
                        sSmif1Status = Parent.GetPuObject.SmifList(1).Attributes("SmifStatus").Value
                        If Err.Number <> 0 Then
                            Err.Clear
                            Parent.GetPuObject.SmifList(1).Attributes.Add("SmifStatus").Value = "Normal"
                        Else
                            Parent.GetPuObject.PortList(1).Attributes.Item("SmifStatus").Value = "Normal"
                        End If
                        On Error GoTo ErrorHandler
                    Else
                        Call Parent.LogTrace(0, MODULE_ID & " SMIF1 is present")
                        On Error Resume Next
                        sSmif1Status = Parent.GetPuObject.SmifList(1).Attributes("SmifStatus").Value
                        If Err.Number <> 0 Then
                            Err.Clear
                            Parent.GetPuObject.SmifList(1).Attributes.Add("SmifStatus").Value = "Busy"
                        Else
                            Parent.GetPuObject.SmifList(1).Attributes.Item("SmifStatus").Value = "Busy"
                        End If
                        On Error GoTo ErrorHandler
                    End If
                    
                Case "SMIF2"
                    i = i + 1
                    If objVfei.Var("POD_PRESENT").Value = False Then
                        Call Parent.LogTrace(0, MODULE_ID & " SMIF2 is not present")
                        On Error Resume Next
                        sSmif2Status = Parent.GetPuObject.SmifList(2).Attributes("SmifStatus").Value
                        If Err.Number <> 0 Then
                            Err.Clear
                            Parent.GetPuObject.SmifList(2).Attributes.Add("SmifStatus").Value = "Normal"
                        Else
                            Parent.GetPuObject.PortList(2).Attributes.Item("SmifStatus").Value = "Normal"
                        End If
                        On Error GoTo ErrorHandler
                    Else
                        Call Parent.LogTrace(0, MODULE_ID & " SMIF2 is present")
                        On Error Resume Next
                        sSmif2Status = Parent.GetPuObject.SmifList(2).Attributes("SmifStatus").Value
                        If Err.Number <> 0 Then
                            Err.Clear
                            Parent.GetPuObject.SmifList(2).Attributes.Add("SmifStatus").Value = "Busy"
                        Else
                            Parent.GetPuObject.SmifList(2).Attributes.Item("SmifStatus").Value = "Busy"
                        End If
                        On Error GoTo ErrorHandler
                    End If
                    
            End Select
            
            If i = 2 Then
                Call Parent.TaskComplete(TaskID, "")
                Exit Sub
            End If
    End Select
    Exit Sub
ErrorHandler:
    Call Parent.LogTrace(0, FormatErrMsg(MODULE_ID & ".ReceivedVfei", Err, Err.Description))
End Sub

Private Property Let TaskAncestor_TaskID(ByVal RHS As String)
    TaskID = RHS
End Property

Private Property Get TaskAncestor_TaskID() As String
    TaskAncestor_TaskID = TaskID
End Property


