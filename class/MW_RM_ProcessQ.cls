VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MW_RM_ProcessQ"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit


Implements ts_ECAncestor.TaskAncestor

Private Const MODULE_ID = "MW_EQ_ProcessQ"

Dim Parent As tsEC.clsRunManager
Dim TaskID As String
Dim objBatch As Batch

Private Sub TaskAncestor_Execute(Parameter As Variant)
    Dim objVfei As vfei.Message22
    
    On Error GoTo ErrorHandler
    
    Set objBatch = Parameter
    
    If objBatch.LotList.Count = 2 Then
       Set objVfei = New Message22
       With objVfei
          .CommandID = "RUN_CASSETTE"
          .MachineID = Parent.GetPuObject.EqDrvList(1).EqDrvID
          .MessageType = "E"
          .AddVar "CASSETTEID", "A", objBatch.LotList(2).Cassette
       End With
       Call Parent.LogTrace(0, MODULE_ID & " .Send RUN_CASSETTE Vfei Message,and One Batch Two Cassette will SubscribeVfeiMsg('LOT_Q')")
       Call Parent.SubscribeVfeiMsg(Me, Parent.GetPuObject.EqDrvList(1).EqDrvID, "LOT_Q")
       Call Parent.SendVfei(objVfei, Me)
       Exit Sub
    Else
       If Parent.GetPuObject.Attributes("State").Value <> "Run" Then
          Parent.GetPuObject.Attributes("State").Value = "Run"
       End If
       Call Parent.LogTrace(0, MODULE_ID & " .Run Next Task, One Batch One Lot, PU State :" & Parent.GetPuObject.Attributes("State").Value)
       Call Parent.TaskComplete(TaskID, "", Parameter)
       Exit Sub
    End If
    
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
    
    If objVfei.CommandID <> "LOT_Q" Then Exit Sub
    
    If objVfei.ErrorCode = 0 Then
       Call Parent.LogTrace(0, MODULE_ID & " .LOT_Q : " & objBatch.LotList(1).LotID & " and " & objBatch.LotList(2).LotID & " have arrival and ready to load in SMIF")
       If Parent.GetPuObject.Attributes("State").Value <> "Run" Then
          Parent.GetPuObject.Attributes("State").Value = "Run"
       End If
       
       Call Parent.TaskComplete(TaskID, "")
       Exit Sub
    End If
    
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



Private Sub Class_Terminate()
    Set objBatch = Nothing
End Sub

