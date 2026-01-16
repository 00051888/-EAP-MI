VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MW_EQ_MOC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit


Implements ts_ECAncestor.TaskAncestor

Private Const MODULE_ID = "MW_EQ_MOC"

Dim Parent As tsEC.clsRunManager
Dim objBatch As Batch
Dim TaskID As String

Private Sub Class_Terminate()
   Set objBatch = Nothing
End Sub

Private Sub TaskAncestor_Execute(Parameter As Variant)
    Dim objEqDrv As EqDrv
    Dim bWait As Boolean
    

    On Error GoTo ErrorHandler
    
    Set objBatch = Parameter
    On Error Resume Next
    
    For Each objEqDrv In Parent.GetPuObject.EqDrvList
       Call Parent.SubscribeVfeiMsg(Me, objEqDrv.EqDrvID, "MOC")
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
    Dim objPort As PortItem
    Dim objTmpBatch As Batch
    Dim objVfei_1 As vfei.Message22
    Dim bWait As Boolean
    
    On Error GoTo ErrorHandler

'Yen add for MI-4 dual function
For Each objLot In objBatch.LotList
If objVfei.Var("CONTROLLER").Value <> objLot.LotID Then Exit Sub
Next
'Yen

    bWait = False
       
    For Each objLot In objBatch.LotList
       If objLot.Status <> "MOC" Then
          objLot.Status = "MOC"
          Call Parent.LogTrace(0, "Received MOC")
          Parent.GetPuObject.PortList(CStr(objLot.Port)).State = PORT_EMPTY
          Exit For
       End If
    Next objLot
       
    For Each objLot In objBatch.LotList
       If objLot.Status <> "MOC" Then
          bWait = True
          Exit For
       End If
    Next objLot
       
    If bWait = False Then
      ' On Error Resume Next
      ' objBatch.Attributes.Item("Processing") = "N"
     '  For Each objTmpBatch In Parent.GetBatchList
     '     If objTmpBatch.Attributes.Item("Processing") = "W" Then
     '        Set objVfei_1 = New Message22
     '        objVfei_1.CommandID = "PROCESSQ"
     ''        objVfei_1.MachineID = "RM"
     '        objVfei_1.MessageType = "E"
     '        objVfei_1.AddVar "ISPROCESSQ", "A", "FALSE"
     '        Call Parent.SendVfei(objVfei_1, Me)
     '        Exit For
     '     End If
     '  Next objTmpBatch
       'Call Parent.LogTrace(0, "Received MOC")
       Call Parent.TaskComplete(TaskID, "")
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

