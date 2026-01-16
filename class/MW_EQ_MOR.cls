VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MW_EQ_MOR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit


Implements ts_ECAncestor.TaskAncestor

Private Const MODULE_ID = "MW_EQ_MOR"

Dim Parent As tsEC.clsRunManager
Dim TaskID As String
Dim objBatch As Batch
Dim iIndex As Integer

Private Sub TaskAncestor_Execute(Parameter As Variant)
    Dim objEqDrv As EqDrv
    
    On Error GoTo ErrorHandler
    
    Set objBatch = Parameter
    
    For Each objEqDrv In Parent.GetPuObject.EqDrvList
       Call Parent.SubscribeVfeiMsg(Me, objEqDrv.EqDrvID, "MOR")
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
    Dim sPortID As String
    Dim sSmifID As String
    Dim objLot As Lot
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Dim bWait As Boolean
    Dim TempStrA As String
    Dim TempStrB As String
    Dim objnewVfei As vfei.Message22
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    On Error GoTo ErrorHandler
    
'Yen add for MI-4 dual function
For Each objLot In objBatch.LotList
If objVfei.Var("CONTROLLER").Value <> objLot.LotID Then Exit Sub
Next
'Yen

    bWait = False
    ' Albert.L
    Dim MOR_Flag_R As Boolean
    Dim MOR_Flag_L As Boolean
    MOR_Flag_R = True
    MOR_Flag_L = True
    
    ' Define Which Port Need to Check
    For Each objLot In objBatch.LotList
        If objLot.Port = "1" Then MOR_Flag_L = False
        If objLot.Port = "2" Then MOR_Flag_R = False
    Next objLot
       
    
    For Each objLot In objBatch.LotList
       If objLot.LotID = objVfei.Var("LLOT") And objVfei.Var("LPORT") = "1" Then
            objLot.Status = "MOR"
            MOR_Flag_L = True
       End If
    Next objLot
       
    For Each objLot In objBatch.LotList
       If objLot.LotID = objVfei.Var("RLOT") And objVfei.Var("RPORT") = "1" Then
            objLot.Status = "MOR"
            MOR_Flag_R = True
       End If
    Next objLot
          
    If MOR_Flag_L = True And MOR_Flag_R = True Then
        Call Parent.TaskComplete(TaskID, "")
    Else
         Call Parent.LogTrace(0, "Pivot Not Extend Complete")
         Call Parent.LogEvent("Pivot Error", "", "Smif1-" + MOR_Flag_L + ";Smif2-" + MOR_Flag_R)
         'Call Parent.LogEvent("Pivot Error", 111, "Pivot Not Extend Complete")
         'Call Parent.LogAlarm(111, "SET", 0, " Pivot Not Extend Complete")
         'Call AlarmReport(Parent, Me, 111, "Pivot Not Extend Complete", objVfei.ErrorText)
    End If
       'If objLot.Status <> "MOR" Then
       '   objLot.Status = "MOR"
       '   Call Parent.LogTrace(0, "Received MOR")
       '   Exit For
       'End If
    'Next objLot
    
    'For Each objLot In objBatch.LotList
    '   If objLot.Status <> "MOR" Then
    '      bWait = True
    '      Exit For
    '   End If
    'Next objLot
    
    'If bWait = False Then
    '   Call Parent.TaskComplete(TaskID, "")
    '   Set objBatch = Nothing
    'End If
    
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












