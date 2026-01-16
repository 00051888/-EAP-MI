VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MS_EQ_Start_Implant"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements ts_ECAncestor.TaskAncestor

Private Const MODULE_ID = "MS_EQ_Start_Implant"

Dim Parent As tsEC.clsRunManager
Dim TaskID As String
Dim objPU As ProcessingUnit
Dim objBatch As Batch


Private Sub TaskAncestor_Execute(Parameter As Variant)
    Dim objVfei As New Message22
    
    On Error GoTo ErrorHandler

    Set objBatch = Parameter
    Set objPU = Parent.GetPuObject
        
    objVfei.CommandID = "REMOTE_COMMAND"
    objVfei.MachineID = objPU.EqDrvList(1).EqDrvID
    objVfei.AddVar "COMMAND", "A", "START-IMPLANT"
    
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
    Dim objLot As Lot
    Dim sContinue As String
    
    On Error GoTo ErrorHandler
    If objVfei.CommandID <> "REMOTE_COMMAND" Then Exit Sub
    
    If objVfei.ErrorCode = 0 Then
        For Each objLot In objBatch.LotList
           ' Remark For Auto Implant :
           'objLot.Status = LOT_STARTING
           Call SendLotEventReport(Parent, Me, objLot)
        Next objLot
        Call Parent.LogEvent("START-IMPLANT", GetLotListStr, "PortID:" & GetPortListStr)
        Call Parent.LogTrace(5, "Start Implant Successful.")
        
        Call Parent.TaskComplete(TaskID, "")
    Else
        For Each objLot In objBatch.LotList
          ' Remark For Auto Implant :
          'objLot.Status = LOT_STARTERR
           Call SendLotEventReport(Parent, Me, objLot)
        Next objLot
        Call Parent.LogEvent("STARTERR", GetLotListStr, "PortID:" & GetPortListStr)
        Call Parent.LogTrace(5, "Start Implant Failure.")
        Call AlarmReport(Parent, Me, "12", Translate(Parent, "Start Implant «ü¥O¥¢±Ñ: ", "Start Implant Command failure"), objVfei.ErrorText)
        sContinue = Parent.GetTaskAttribute(TaskID, "CONTINUE")
        If sContinue = "Y" Then
           Call Parent.TaskComplete(TaskID, "")
        End If
    End If

    
    Set objBatch = Nothing
    Set objPU = Nothing
    
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


Private Function GetLotListStr() As String
   Dim TempStr As String
   Dim i As Integer
   
   
   For i = 1 To objBatch.LotList.Count
       TempStr = TempStr & objBatch.LotList(i).LotID & " "
   Next
   GetLotListStr = TempStr
   

End Function

Private Function GetPortListStr() As String
   Dim TempStr As String
   Dim i As Integer
   
   
   For i = 1 To objBatch.LotList.Count
       TempStr = TempStr & objBatch.LotList(i).Port & " "
   Next
   GetPortListStr = TempStr
   

End Function

