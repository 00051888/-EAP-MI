VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MS_EQ_Start"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements ts_ECAncestor.TaskAncestor

Private Const MODULE_ID = "MS_EQ_Start"

Dim Parent As tsEC.clsRunManager
Dim TaskID As String
Dim objPU As ProcessingUnit
Dim objBatch As Batch
Dim intStartCnt As Integer

Private Sub TaskAncestor_Execute(Parameter As Variant)
    Dim objLot As Lot
    Dim objVfei As New Message22
    
    On Error GoTo ErrorHandler

    Set objBatch = Parameter
    Set objPU = Parent.GetPuObject
        
    objVfei.CommandID = "REMOTE_COMMAND"
    objVfei.MachineID = objPU.EqDrvList(1).EqDrvID
    objVfei.AddVar "COMMAND", "A", "START"
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
    Dim objnewVfei As vfei.Message22
    Dim sTimerDelay     As String
    Dim sMaxRetryCount  As String
    
    On Error GoTo ErrorHandler
    
    If objVfei.CommandID <> "REMOTE_COMMAND" And _
       objVfei.CommandID <> "TIME_COUNT" Then Exit Sub
    
    sTimerDelay = Parent.GetTaskAttribute(TaskID, "TIMER")
    sMaxRetryCount = Parent.GetTaskAttribute(TaskID, "MAX_RETRY_COUNT")
    
    'DeBug Message
    Call Parent.LogEvent("ReciveVfei-START", GetLotListStr, "sTimerDelay:" & sTimerDelay & ", sMaxRetryCount:" & sMaxRetryCount & ", intStartCnt:" & intStartCnt)
    Call Parent.LogEvent("ReciveVfei-START", GetLotListStr, "VfeiCommand:" & objVfei.CommandID & ", ErrorCode:" & objVfei.ErrorCode & ", EqDrvID:" & objPU.EqDrvList(1).EqDrvID)
    
    'Start Command need to send 3 times (120 sec), if Start-Command failed.
    If objVfei.ErrorCode = 0 Then
        If objVfei.CommandID = "REMOTE_COMMAND" Then
            'Start Command Success, then TaskComplete. And, clear objPU & objBatch.
            For Each objLot In objBatch.LotList
                objLot.Status = "BEAM"
                Call SendLotEventReport(Parent, Me, objLot)
            Next objLot
            Call Parent.LogEvent("START", GetLotListStr, "PortID:" & GetPortListStr)
            Call Parent.LogTrace(5, "Start Successful.")
            
            intStartCnt = 0
            
            Call Parent.TaskComplete(TaskID, "")
            
            Set objPU = Nothing
            Set objBatch = Nothing

        ElseIf objVfei.CommandID = "TIME_COUNT" Then
            'Re-Start-Timer reply, send "START" command.
            Set objnewVfei = New Message22
            
            With objnewVfei
                .CommandID = "REMOTE_COMMAND"
                .MachineID = objPU.EqDrvList(1).EqDrvID
                .MessageType = "C"
                .AddVar "COMMAND", "A", "START"
            End With
            Call Parent.LogEvent("Re-START Sended", GetLotListStr, "PortID:" & GetPortListStr & " Count=" & intStartCnt)
            Call Parent.SendVfei(objnewVfei, Me)
        End If
    Else
        'Start Command failed, Re-Start-Timer sended.
        intStartCnt = intStartCnt + 1
        
        'Send Timer Vfei. Maxinum times = 3.
        If intStartCnt <= CInt(sMaxRetryCount) Then
            Set objnewVfei = New Message22
            With objnewVfei
                .CommandID = "TIME_COUNT"
                .MachineID = "CONSOLE"
                .MessageType = "C"
                'Default TimeDelay is 120 secs.
                If IsNumeric(sTimerDelay) Then
                    .AddVar "COUNT", "A", sTimerDelay
                Else
                    .AddVar "COUNT", "A", "120"
                End If
            
            End With
            Call Parent.LogEvent("Re-START", GetLotListStr, "PortID:" & GetPortListStr & " Count=" & intStartCnt)
            Call Parent.SendVfei(objnewVfei, Me)
        Else
            'Re-Start over 3 times. Close-Task, and clear objPU & objBatch.
            For Each objLot In objBatch.LotList
                objLot.Status = "BEAMERR"
                Call SendLotEventReport(Parent, Me, objLot)
            Next objLot
            Call Parent.LogEvent("START_ERR", GetLotListStr, "PortID:" & GetPortListStr)
            Call Parent.LogTrace(5, "Start Failed.")
            Call AlarmReport(Parent, Me, "12", Translate(Parent, "Start «ü¥O¥¢±Ñ: ", "Start Command failure:"), objVfei.ErrorText)
            sContinue = Parent.GetTaskAttribute(TaskID, "CONTINUE")
            If sContinue = "Y" Then
                Call Parent.TaskComplete(TaskID, "")
            End If
            
            Set objPU = Nothing
            Set objBatch = Nothing
            
        End If
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


