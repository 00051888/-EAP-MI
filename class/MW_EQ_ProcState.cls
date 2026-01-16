VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MW_EQ_ProcState"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit


Implements ts_ECAncestor.TaskAncestor

Private Const MODULE_ID = "MW_EQ_ProcState"

Dim Parent As tsEC.clsRunManager
Dim TaskID As String
Dim objBatch As Batch
Dim objLot As Lot

Private Sub TaskAncestor_Execute(Parameter As Variant)
    Dim objEqDrv As EqDrv
    
    On Error GoTo ErrorHandler
    
    Set objBatch = Parameter
    
    For Each objEqDrv In Parent.GetPuObject.EqDrvList
       Call Parent.SubscribeVfeiMsg(Me, objEqDrv.EqDrvID, "GO_DISPLAY")
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
    Dim objTmpLot       As Lot
    Dim sLotID          As String
    Dim sProcessState   As String
    Dim sProcStateDesc  As String
    On Error GoTo ErrorHandler

    sLotID = objVfei.Var("CONTROLLER").Value
    sProcessState = objVfei.Var("STATE").Value
    
    For Each objTmpLot In objBatch.LotList
        If sLotID = objTmpLot.LotID Then
            Set objLot = objTmpLot
            Exit For
        End If
    Next
    
    Select Case sProcessState
    Case "0"
        sProcStateDesc = "Not_Ready"
    Case "1"
        sProcStateDesc = "Ready"
    Case "2"
        sProcStateDesc = "Unused"
    Case "3"
        sProcStateDesc = "Setting_Up"
    Case "4"
        sProcStateDesc = "Implant_Ready"
    Case "5"
        sProcStateDesc = "Implanting"
    Case "6"
        sProcStateDesc = "Implant_Hold"
    Case "7"
        sProcStateDesc = "Aborting"
    Case Else
        sProcStateDesc = "not_define"
    End Select
    
    If Not (objLot Is Nothing) Then
        Call Parent.LogTrace(0, "Eqp PROCESS_STATE[" & sProcStateDesc & "]")
        
        If sProcessState = "4" Then
            Call Parent.LogEvent("PROCESS_STATE", sLotID, "Ready to Send START-IMPLANT")
            Call Parent.LogTrace(0, "PROCESS_STATE, Eqp is ready to Send START-IMPLANT")
            Call Parent.TaskComplete(TaskID, "")
        End If
    Else
        Call Parent.LogTrace(0, "GO_DISPLAY get objLot failure, at Lot[" & sLotID & "] and batch[" & objBatch.BatchID & "]")
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


