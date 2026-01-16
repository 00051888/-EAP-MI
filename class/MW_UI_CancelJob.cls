VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MW_UI_CancelJob"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements ts_ECAncestor.TaskAncestor

Private Const MODULE_ID = "MW_UI_CancelJob"

Dim Parent As tsEC.clsRunManager
Dim TaskID As String
Dim objBatch As Batch

Private Sub TaskAncestor_Execute(Parameter As Variant)
    On Error GoTo ErrorHandler
    
    Set objBatch = Parameter
    
    Call Parent.SubscribeHostMsg(Me, "CANCEL_JOB_TO_EC")
    
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
    Dim sCancelBatchID As String
    Dim sCancelCassetteID As String
    Dim objLot As Lot
    Dim objPortList As PortList
    Dim TempStr As String
    Dim IsWaitMOC As Boolean
    
    On Error GoTo ErrorHandler
    
    sCancelBatchID = objHostTrans.Primary.Item("BATCH_ID")
    sCancelCassetteID = objHostTrans.Primary.Item("CASSETTE")
    
    If sCancelBatchID <> "NULL" Then
       If sCancelBatchID <> objBatch.BatchID Then Exit Sub
    Else
       For Each objLot In objBatch.LotList
           If objLot.Cassette = sCancelCassetteID Then
              Exit For
           End If
       Next
    
       If objLot.Cassette <> sCancelCassetteID Then Exit Sub
    End If
    
    
    TempStr = UCase$(Trim$(Parent.GetTaskAttribute(TaskID, "IS_WAIT_MOC")))
    IsWaitMOC = Not (TempStr = "" Or TempStr = "N")
        'clean port & stage
    Set objPortList = Parent.GetPuObject.PortList
    
    On Error Resume Next
    
    For Each objLot In objBatch.LotList
        
         objPortList(CStr(objLot.Port)).CassetteID = ""
         If Not IsWaitMOC Then
            objPortList(CStr(objLot.Port)).State = PORT_EMPTY
         End If
       
    Next objLot
    
    
    Call Parent.LogEvent("Cancel Job", objBatch.LotList(1).LotID, "Port:" & objBatch.LotList(1).Port)
    objHostTrans.ErrorCode = 0
    objHostTrans.ErrorText = ""
    objHostTrans.Reply
    '-------------------------------------------------------------------------
    'Update PU.Attribute's porperty("State") = ""    edit by simon 2000/07/04
    Parent.GetPuObject.Attributes("State").Value = ""
    '-------------------------------------------------------------------------
    Call Parent.GetBatchList.Remove(objBatch.BatchID)
    Call Parent.KillTaskList(TaskID)
    
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



