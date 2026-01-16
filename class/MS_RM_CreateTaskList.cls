VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MS_RM_CreateTaskList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements ts_ECAncestor.TaskAncestor

Private Const MODULE_ID = "MS_RM_CreateTaskList"

Dim Parent As tsEC.clsRunManager
Dim TaskID As String
Dim objBatch As Batch


Private Sub TaskAncestor_Execute(Parameter As Variant)
    Dim objPort As PortItem
    Dim objSmif As SmifItem
    Dim sTaskList As String
    Dim objLot As Lot
    On Error GoTo ErrorHandler
    Set objBatch = Parameter
    
    sTaskList = Parent.GetTaskAttribute(TaskID, "TASK_LIST")
    
    For Each objLot In objBatch.LotList
       Set objPort = Parent.GetPuObject.PortList(objLot.Port)
       Set objSmif = Parent.GetPuObject.SmifList(objPort.RelatedSMIFID)
       Call Parent.StartTaskList(sTaskList, objSmif)
    Next objLot
    
    Call Parent.TaskComplete(TaskID, "")
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



