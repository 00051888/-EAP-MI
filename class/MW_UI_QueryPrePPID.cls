VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MW_UI_QueryPrePPID"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit


Implements ts_ECAncestor.TaskAncestor

Private Const MODULE_ID = "MW_UI_QueryPrePPID"

Dim Parent As tsEC.clsRunManager
Dim TaskID As String
Dim objPU As ProcessingUnit
Dim objTmpHostTrans As Object

Private Sub Class_Terminate()
   Set objPU = Nothing
End Sub

Private Sub TaskAncestor_Execute(Parameter As Variant)
    On Error GoTo ErrorHandler
    
    
    Set objPU = Parameter
    Call Parent.SubscribeHostMsg(Me, "COMMAND_FROM_ECUI")

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
    Dim objnewVfei As New vfei.Message22
    
    On Error GoTo ErrorHandler
    
    If objHostTrans.TransactionName <> "COMMAND_FROM_ECUI" Then Exit Sub
    If objHostTrans.Primary.Item("COMMAND_ID").Value <> "QUERY_STATUS" Then Exit Sub
    
    Set objTmpHostTrans = objHostTrans
    
    'strIniFile = Parent.GetTaskAttribute(TaskID, "INIFILE")
    'strPrePPID = (GetIniSetting(strIniFile, "General", "PrePPID"))
    With objnewVfei
       .CommandID = objHostTrans.Primary.Item("COMMAND_ID").Value
       .MachineID = Parent.GetPuObject.EqDrvList(1).EqDrvID
       .MessageType = "C"
       .AddVar "PARAMETER", "A", objHostTrans.Primary.Item("Parameter").Item(1).Item("ParameterName").Value
    End With
    Call Parent.LogTrace(0, "Query 'PPID'")
    'Call Parent.LogEvent("QUERY_STATUS", GetLotListStr, "PARAMETER='PPID'")
    Call Parent.SendVfei(objnewVfei, Me)
        
    Exit Sub
ErrorHandler:
    Call Parent.LogTrace(0, FormatErrMsg(MODULE_ID & ".ReceivedHostMsg", Err, Err.Description))
End Sub

Private Sub TaskAncestor_ReceivedVfei(objVfei As vfei.Message22)
   '
    
    With objTmpHostTrans.Secondary
        .Item("ErrorCode") = objVfei.ErrorCode
        .Item("ErrorText") = objVfei.ErrorText
        With .Item("ReplyItem")
           .Add.Item("Value").Value = objVfei.Var("PPID").Value
        End With
        Call Parent.LogTrace(0, "PrePPID=" & objVfei.Var("PPID").Value)
        Call Parent.LogEvent("QUERY_STATUS", "", "PrePPID=" & objVfei.Var("PPID").Value)
        
    End With
    
    objTmpHostTrans.Reply
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


