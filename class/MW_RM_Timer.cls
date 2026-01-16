VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MW_RM_Timer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements ts_ECAncestor.TaskAncestor

Private Const MODULE_ID = ""

Dim Parent As Object
Dim TaskID As String
Dim Description As String
Dim TimerDelay As String




Private Sub TaskAncestor_Execute(Parameter As Variant)
    Dim objVfei As Message22
    On Error GoTo ErrorHandler
       Set objVfei = New vfei.Message22
        objVfei.CommandID = "TIME_COUNT"
        objVfei.MachineID = "CONSOLE"
        TimerDelay = Parent.GetTaskAttribute(TaskID, "TIMER")
        If IsNumeric(TimerDelay) Then
            objVfei.AddVar "COUNT", "A", TimerDelay
        Else
            objVfei.AddVar "COUNT", "A", "3"
        End If
        Call Parent.SendVfei(objVfei, Me)
        Call Parent.LogTrace(0, "Delay " & IIf(IsNumeric(TimerDelay), TimerDelay, "3") & " Secends, then retry")
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
            Description = Parent.GetTaskAttribute(TaskID, "TEXT")
            If Description <> "" Then
                Call Parent.LogTrace(0, "Time is up,Retry [" & Description & "] again")
            Else
                Call Parent.LogTrace(0, "Time is up.")
            End If
            Call Parent.TaskComplete(TaskID, "")
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


