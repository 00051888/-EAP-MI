VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MQ_RM_Status"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Option Base 1

Implements ts_ECAncestor.TaskAncestor

Private Const MODULE_ID = "MQ_RM_Status"

Dim parent As tsEC.clsRunManager
Dim TaskID As String

Private Sub TaskAncestor_Execute(Parameter As Variant)
Dim objVfei As Message22
Dim vEqIdle As Boolean
Dim vBatchCount As Integer

On Error Resume Next
On Error GoTo ErrorHandler

vBatchCount = parent.GetBatchList.count
If vBatchCount = 0 Then
    vEqIdle = True
Else
    vEqIdle = False
End If

Set objVfei = New Message22
If vEqIdle Then
    With objVfei
         .CommandID = "EQ_IDLE"
         .MachineID = parent.GetPuObject.EqDrvList(1).EqDrvID
         .MessageType = "E"
    End With
Else
    With objVfei
         .CommandID = "EQ_BUSY"
         .MachineID = parent.GetPuObject.EqDrvList(1).EqDrvID
         .MessageType = "E"
    End With
End If
Call parent.SendVfei(objVfei, Me)
Call parent.TaskComplete(TaskID, "")
Exit Sub

ErrorHandler:
Call parent.LogTrace(0, FormatErrMsg(MODULE_ID & ".Execute", Err, Err.Description))
End Sub

Private Property Set TaskAncestor_Parent(ByVal RHS As Object)
Set parent = RHS
End Property

Private Property Get TaskAncestor_Parent() As Object
Set TaskAncestor_Parent = parent
End Property

Private Sub TaskAncestor_ReceivedHostMsg(objHostTrans As HostOcx.HostTransaction, isPrimary As Boolean)
On Error GoTo ErrorHandler

Exit Sub

ErrorHandler:
Call parent.LogTrace(0, FormatErrMsg(MODULE_ID & ".ReceivedHostMsg", Err, Err.Description))
End Sub

Private Sub TaskAncestor_ReceivedVfei(objVfei As vfei.Message22)

On Error GoTo ErrorHandler

Exit Sub

ErrorHandler:
Call parent.LogTrace(0, FormatErrMsg(MODULE_ID & ".ReceivedVfei", Err, Err.Description))
End Sub

Private Property Let TaskAncestor_TaskID(ByVal RHS As String)
TaskID = RHS
End Property

Private Property Get TaskAncestor_TaskID() As String
TaskAncestor_TaskID = TaskID
End Property



