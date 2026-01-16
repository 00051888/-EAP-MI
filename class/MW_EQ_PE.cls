VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MW_EQ_PE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit


Implements ts_ECAncestor.TaskAncestor

Private Const MODULE_ID = "MW_EQ_PE"

Dim Parent As tsEC.clsRunManager
Dim TaskID As String
Dim objBatch As Batch
Dim iIndex As Integer

Private Sub TaskAncestor_Execute(Parameter As Variant)
    Dim objEqDrv As EqDrv
    
    On Error GoTo ErrorHandler
    
    Set objBatch = Parameter
    
    For Each objEqDrv In Parent.GetPuObject.EqDrvList
       Call Parent.SubscribeVfeiMsg(Me, objEqDrv.EqDrvID, "PE")
    Next
    
    
    
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
    Dim objLot As Lot
    Dim bWait As Boolean
    Dim objnewVfei As New Message22
    Dim Parameter(1 To 4) As String
    Dim sTemp As String
    On Error GoTo ErrorHandler
    
'Yen add for MI-4 dual function
For Each objLot In objBatch.LotList
If (objVfei.CommandID = "PE") Then
If (objVfei.Var("CONTROLLER").Value <> objLot.LotID) Then Exit Sub
End If
Next
'Yen
    
    Select Case objVfei.CommandID
       Case "PE"
          bWait = False
          For Each objLot In objBatch.LotList
             If objLot.Status <> LOT_COMPLETED Then
                objLot.Status = LOT_COMPLETED
                Parent.GetPuObject.PortList(CStr(objLot.Port)).State = PORT_PRESENT
                Call SendLotEventReport(Parent, Me, objLot)
                Call Parent.LogTrace(0, "PE")
                Exit For
             End If
          Next objLot
    
          For Each objLot In objBatch.LotList
             If objLot.Status <> LOT_COMPLETED Then
               bWait = True
               Exit For
             End If
          Next objLot
    
          If bWait = False Then
             Parameter(1) = "LLOT"
             Parameter(2) = "LWAF"
             Parameter(3) = "RLOT"
             Parameter(4) = "RWAF"
             With objnewVfei
                .CommandID = "QUERY_STATUS"
                .MachineID = Parent.GetPuObject.EqDrvList(1).EqDrvID
                .MessageType = "C"
                .AddVar "PARAMETER", "A", Parameter
             End With
             Call Parent.SendVfei(objnewVfei, Me)
          End If
       Case "QUERY_STATUS"
          On Error Resume Next
          For Each objLot In objBatch.LotList
             If objLot.LotID = Trim$(objVfei.Var("LLOT").Value) Then
                sTemp = objLot.Attributes.Item("ProcessedCount").Value
                If Err <> 0 Then
                   Err.Clear
                   objLot.Attributes.Add("ProcessedCount").Value = objVfei.Var("LWAF").Value
                Else
                   objLot.Attributes.Item("ProcessedCount").Value = objVfei.Var("LWAF").Value
                End If
             ElseIf objLot.LotID = Trim$(objVfei.Var("RLOT").Value) Then
                sTemp = objLot.Attributes.Item("ProcessedCount").Value
                If Err <> 0 Then
                   Err.Clear
                   objLot.Attributes.Add("ProcessedCount").Value = objVfei.Var("RWAF").Value
                Else
                   objLot.Attributes.Item("ProcessedCount").Value = objVfei.Var("RWAF").Value
                End If
             End If
             
            Err.Clear
            objLot.Attributes.Add("PROCESS_END_TMST").Value = Format(Now, "yyyy-mm-dd hh:mm:ss.000")
            If Err.Number = 0 Then
                Call Parent.LogTrace(0, "EAP Add PROCESS_END_TMST Attribute Success,")
            Else
                Err.Clear
                Call Parent.LogTrace(0, "EAP Add PROCESS_END_TMST Attribute fail,ErrorText=" & Err.Description)
            End If
             
          Next
          Call Parent.TaskComplete(TaskID, "")
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


