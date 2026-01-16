VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MQ_EQ_PPBody"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements ts_ECAncestor.TaskAncestor

Private Const MODULE_ID = "MQ_EQ_PPBody"

Dim Parent As tsEC.clsRunManager
Dim TaskID As String
Dim objBatch As Batch
Dim objLot As Lot
Dim objPort As PortItem
Dim objSmif As SmifItem
Dim objEqDrv As EqDrv
Dim vParameter As Variant

Private Sub TaskAncestor_Execute(Parameter As Variant)
    Dim objnewVfei As New Message22
    Dim sPortID As String
    Dim sSmifID As String
    Dim sLotID As String
    Dim sRMMode As String
    Dim iIndex As Integer
    Dim objTmpLot As Lot
    
    On Error GoTo ErrorHandler
    If IsObject(Parameter) Then
        Set vParameter = Parameter
    Else
        vParameter = Parameter
    End If
    
    If TypeOf Parameter Is PortItem Then
        Set objPort = Parameter
        sPortID = objPort.PortID
        sSmifID = objPort.RelatedSMIFID
        Set objSmif = Parent.GetPuObject.SmifList(sSmifID)
    ElseIf TypeOf Parameter Is SmifItem Then
        Set objSmif = Parameter
        sSmifID = objSmif.SMIFID
        sPortID = objSmif.RelatedPortID
        Set objPort = Parent.GetPuObject.PortList(sPortID)
        Set objEqDrv = Parent.GetPuObject.EqDrvList(1)
    ElseIf TypeOf Parameter Is Batch Then
        Set objBatch = Parameter
        Set objLot = objBatch.LotList(1)
        sPortID = objLot.Port
        Set objPort = Parent.GetPuObject.PortList(sPortID)
        sSmifID = objPort.RelatedSMIFID
        Set objSmif = Parent.GetPuObject.SmifList(sSmifID)
    ElseIf TypeOf Parameter Is Lot Then
        Set objLot = Parameter
        sPortID = objLot.Port
        Set objPort = Parent.GetPuObject.PortList(sPortID)
        sSmifID = objPort.RelatedSMIFID
        Set objSmif = Parent.GetPuObject.SmifList(sSmifID)
    End If
    
    'Check Recipe Body
    For Each objTmpLot In objBatch.LotList
        objTmpLot.Status = "LOT_CHECKPPBODY"
        Call SendLotEventReport(Parent, Me, objLot)
    Next objTmpLot
 
    If objBatch.LotList.Count > 1 Then
        sLotID = objBatch.LotList(1).LotID & "," & objBatch.LotList(2).LotID
        sPortID = "Port=" & objBatch.LotList(1).Port & "," & objBatch.LotList(2).Port
    Else
        sLotID = objBatch.LotList(1).LotID
        sPortID = "Port=" & objBatch.LotList(1).Port
    End If

   
    'Add by simon 2000/08/15    ' R-的RECIPE 將不做Check
    ' On Error Resume Next
    ' iIndex = InStr(1, objLot.Recipe1, "R-")
    iIndex = InStr(1, objBatch.LotList(1).Recipe1, "R-")
    If iIndex <> 0 Then
        Call Parent.LogTrace(0, "此 Recipe 為 R-開始不Check")
        Call Parent.TaskComplete(TaskID, "")
        Exit Sub
    End If

    ' On Error GoTo ErrorHandler
   
    'Send query PPBody
    With objnewVfei
        .CommandID = "CHECK_RECIPE_BODY"
        .MachineID = Parent.GetPuObject.EqDrvList(1)
        .MessageType = "C"
        If objBatch.LotList.Count = 2 Then
            If objBatch.LotList(1).Recipe1 = objBatch.LotList(2).Recipe1 Then
                .AddVar "PPID", "A", objLot.Recipe1
                .AddVar "LOT1_STEP", "A", objBatch.LotList(1).Step
                .AddVar "LOT2_STEP", "A", objBatch.LotList(2).Step
            Else
                Call AlarmReport(Parent, Me, "013", "The recipe of two Lots is not the same", "Lot=" & sLotID)
                Call Parent.LogTrace(0, "Cassette1's recipe is not equal to Cassette2's recipe")
            End If
        Else
            .AddVar "PPID", "A", objBatch.LotList(1).Recipe1
            .AddVar "LOT1_STEP", "A", objBatch.LotList(1).Step
            .AddVar "LOT2_STEP", "A", ""
        End If
        .AddVar "RUNMODE", "A", objBatch.RunMode
        
    End With
   
    Call Parent.SendVfei(objnewVfei, Me)
   
  
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
    Dim sLotID As String
    Dim sPortID As String
    
    On Error GoTo ErrorHandler
    
    If objVfei.CommandID <> "CHECK_RECIPE_BODY" Then Exit Sub
    
    If objBatch.LotList.Count > 1 Then
       sLotID = objBatch.LotList(1).LotID & "," & objBatch.LotList(2).LotID
       sPortID = "Port=" & objBatch.LotList(1).Port & "," & objBatch.LotList(2).Port
    Else
       sLotID = objBatch.LotList(1).LotID
       sPortID = "Port=" & objBatch.LotList(1).Port
    End If
    
    If objVfei.ErrorCode <> 0 Then
        For Each objLot In objBatch.LotList
            objLot.Status = LOT_MISPPBODY
            Call SendLotEventReport(Parent, Me, objLot)
            Call Parent.LogEvent("Check Recipe Error", sLotID, objVfei.ErrorText)
        Next objLot
        Call AlarmReport(Parent, Me, 1, Translate(Parent, "檢查 Recipe 內容錯誤:", "Check Recipe body error."), objVfei.ErrorText)
    Else
        Call Parent.TaskComplete(TaskID, "")
    End If
    
    Set objBatch = Nothing
    Set objLot = Nothing
    Set objPort = Nothing
    Set objSmif = Nothing
    Set objEqDrv = Nothing
    
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


