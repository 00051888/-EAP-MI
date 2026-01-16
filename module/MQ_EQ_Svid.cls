VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MQ_EQ_Svid"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit


Implements ts_ECAncestor.TaskAncestor

Private Const MODULE_ID = "MQ_EQ_Svid"

Dim Parent As tsEC.clsRunManager
Dim TaskID As String
Dim objBatch As Batch



Private Sub TaskAncestor_Execute(Parameter As Variant)
    Dim objEqDrv As EqDrv
    Dim objLot As Lot
    Dim Svid As String
    Dim i  As Integer
    Dim SvidString As String
    Dim flag As Boolean
    
    Dim objVfei As Message22
       
    On Error GoTo ErrorHandler
    Set objVfei = New Message22
    Svid = Parent.GetTaskAttribute(TaskID, "SVID")
   
    
    
    
   With objVfei
         .CommandID = "QUERY_SVID"
         .MachineID = Parent.GetPuObject.EqDrvList(1).EqDrvID
         .MessageType = "C"
         .AddVar "Svid", "A", Svid
         
         
    End With
    Call Parent.SendVfei(objVfei, Me)
   
    
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
    On Error Resume Next
    Dim objLot As Lot
    Dim LotIdStr As String
    Dim RecipeStr As String
    Dim CommandStr As String
    Dim Description As String
    Set objLot = objBatch.LotList(1)
    
   'Parseing 1101 Implant Information

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


















