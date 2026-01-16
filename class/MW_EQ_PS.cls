VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MW_EQ_PS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit


Implements ts_ECAncestor.TaskAncestor

Private Const MODULE_ID = "MW_EQ_PS"

Dim Parent As tsEC.clsRunManager
Dim TaskID As String
Dim objBatch As Batch
Dim iIndex As Integer

Private Sub TaskAncestor_Execute(Parameter As Variant)
    Dim objEqDrv As EqDrv
    
    On Error GoTo ErrorHandler
    
    Set objBatch = Parameter
    
    For Each objEqDrv In Parent.GetPuObject.EqDrvList
       Call Parent.SubscribeVfeiMsg(Me, objEqDrv.EqDrvID, "PS")
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
    Dim TempStr As String
    Dim objLot As Lot
    Dim isFind As Boolean
    Dim isBatch As Boolean
    Dim bWait As Boolean
    Dim TempStrA As String
    Dim TempStrB As String
    
    On Error GoTo ErrorHandler
    
'Yen add for MI-4 dual function
For Each objLot In objBatch.LotList
If objVfei.Var("CONTROLLER").Value <> objLot.LotID Then Exit Sub
Next
'Yen
    
    bWait = False
    
    For Each objLot In objBatch.LotList
       If objLot.Status <> LOT_RUNNING Then
          objLot.Status = LOT_RUNNING
          Parent.GetPuObject.PortList(CStr(objLot.Port)).State = PORT_ACTIVE
          Call SendLotEventReport(Parent, Me, objLot)
          Call Parent.LogTrace(0, "PS")
          Exit For
       End If
    Next objLot
    
    For Each objLot In objBatch.LotList
       If objLot.Status <> LOT_RUNNING Then
          bWait = True
          Exit For
       End If
    Next objLot
    
    If bWait = False Then
       On Error Resume Next
       For Each objLot In objBatch.LotList
            Err.Clear
            objLot.Attributes.Add("PROCESS_START_TMST").Value = Format(Now, "yyyy-mm-dd hh:mm:ss.000")
            If Err.Number = 0 Then
                Call Parent.LogTrace(0, "EAP Add PROCESS_START_TMST Attribute Success.")
            Else
                Err.Clear
                Call Parent.LogTrace(0, "EAP Add PROCESS_START_TMST Attribute fail,ErrorText=" & Err.Description)
            End If
       Next objLot
       Call Parent.TaskComplete(TaskID, "")
       Set objBatch = Nothing
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


