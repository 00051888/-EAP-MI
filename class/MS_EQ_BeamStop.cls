VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MS_EQ_BeamStop"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements ts_ECAncestor.TaskAncestor

Private Const MODULE_ID = "MS_EQ_BeamStop"

Dim Parent As Object
Dim TaskID As String
Dim pCurrentIdleCount As Long
Dim pTargetIdleCount As Long
Dim p_blnBeamShutdown As Boolean

Private Sub TaskAncestor_Execute(Parameter As Variant)
'於Task Attribute中設定IDLE_COUNT

On Error GoTo ErrorHandler
Call Parent.SubscribeVfeiMsg(Me, "", "EQ_IDLE")
Call Parent.SubscribeVfeiMsg(Me, "", "EQ_BUSY")
pCurrentIdleCount = 0
pTargetIdleCount = CLng(Parent.GetTaskAttribute(TaskID, "IDLE_COUNT"))
p_blnBeamShutdown = False
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
Dim objVfei1 As Message22

On Error GoTo ErrorHandler

If objVfei.ErrorCode <> 0 Then
    Call Parent.LogTrace(0, "Wait EQ Status Failed!!")
    Exit Sub
End If
Select Case objVfei.CommandID
    Case "EQ_IDLE"
        pCurrentIdleCount = pCurrentIdleCount + 1
    Case "EQ_BUSY"
        pCurrentIdleCount = 0
        If p_blnBeamShutdown Then
            p_blnBeamShutdown = False
            Parent.LogEvent "EQ IDLE COUNT", "", "Reset EQ IDLE COUNT"
        End If
        Exit Sub
    Case Else
        Exit Sub
End Select
Parent.LogEvent "EQ IDLE COUNT", "", "EQ IDLE COUNT=" & pCurrentIdleCount
If pCurrentIdleCount >= pTargetIdleCount And Not p_blnBeamShutdown Then
    p_blnBeamShutdown = True
    
    Set objVfei = New Message22
    objVfei.CommandID = "REMOTE_COMMAND"
    objVfei.MachineID = Parent.GetPuObject.EqDrvList(1).EqDrvID
    objVfei.AddVar "COMMAND", "A", "SHUT-DOWN"
    objVfei.AddVar "PARAMETER", "L", ""
    objVfei.MessageType = "C"
    Call Parent.SendVfei(objVfei, Me)
    Parent.LogEvent "EQ BEAM SHUTDOWN", "", "EQ BEAM SHUTDOWN"
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

