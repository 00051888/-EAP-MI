VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MW_SMIF_PodArrive"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements ts_ECAncestor.TaskAncestor

Private Const MODULE_ID = "MW_SMIF_PodArrive"

Dim Parent As tsEC.clsRunManager
Dim TaskID As String
Dim vParameter As Variant
Private Sub TaskAncestor_Execute(Parameter As Variant)
    Dim sPortID As String
    Dim sSmifID As String
    Dim tSubscribeItem As String
    Dim tServerName As String
    Dim objBatch As Batch
    Dim objLot As Lot
    Dim objPort As PortItem
    Dim objSmif As SmifItem
    
    On Error GoTo ErrorHandler
    
    If IsObject(Parameter) Then
        Set vParameter = Parameter
    Else
        vParameter = Parameter
    End If
    
    sPortID = Parent.GetTaskAttribute(TaskID, "PORT")
    sSmifID = Parent.GetTaskAttribute(TaskID, "SMIFID")
    tServerName = Parent.GetTaskAttribute(TaskID, "SERVERNAME")
    
    If sPortID <> "" Then
       Call Parent.SubscribeVfeiMsg(Me, "", "POD_PLACED", "PORT", sPortID)
    ElseIf sSmifID <> "" Then
        Call Parent.SubscribeVfeiMsg(Me, "", "POD_PLACED", "SMIFID", sSmifID)
    
'--- Add for V0501
    Else
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

        tSubscribeItem = Parent.GetTaskAttribute(TaskID, "SUBSCRIBE")
    
        If tSubscribeItem = "" Then
            Call Parent.SubscribeVfeiMsg(Me, "", "POD_PLACED")
        Else
            Select Case Trim(tSubscribeItem)
                Case "PORT"
                   Call Parent.SubscribeVfeiMsg(Me, "", "POD_PLACED", "PORT", sPortID)
                Case "LOT"
                   Call Parent.SubscribeVfeiMsg(Me, "", "POD_PLACED", "LOT_ID", objLot.LotID)
                Case "SMIF"
                   Call Parent.SubscribeVfeiMsg(Me, "", "POD_PLACED", "SMIFID", sSmifID)
                Case "BATCH"
                   Call Parent.SubscribeVfeiMsg(Me, "", "POD_PLACED", "BATCHID", objBatch.BatchID)
            End Select
        End If
        
   End If
    
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
    Dim objPort As PortItem, sPortID As String
    Dim objSmif As SmifItem, sSmifID As String
    Dim TempStr As String
    Dim isFind As Boolean
    Dim sTaskListName As String
    Dim sRMMode As String
    Dim sAcceptableLevel As String
    Dim bolContinue As Boolean
    Dim vObject As Variant
    
    On Error Resume Next
    
    isFind = False
    TempStr = objVfei.Var("SMIFID").Value
    If Err.Number = 0 Then
        isFind = True
        sSmifID = TempStr
'        Set objSMIF = Parent.GetPuObject.SmifList(sSMIFID)
'        sPortID = objSMIF.RelatedPortID
    Else
       Err.Clear
       TempStr = objVfei.Var("PORT").Value
       If Err.Number = 0 Then
          sPortID = TempStr
          Set objPort = Parent.GetPuObject.PortList(sPortID)
          sSmifID = objPort.RelatedSMIFID
          isFind = True
       Else
          isFind = False
           Call Parent.LogTrace(0, "Received a Event " & objVfei.CommandID & "! But can't find associate Object")
       End If
    End If
    
    On Error GoTo ErrorHandler
    If isFind Then
        
         'This action is initial "State" property
         If Parent.GetPuObject.SmifList(1).PODInPlace = False And Parent.GetPuObject.SmifList(2).PODInPlace = False Then
             Parent.GetPuObject.Attributes("State").Value = ""
         End If
       Set objSmif = Parent.GetPuObject.SmifList(sSmifID)
       If objSmif.SMIFType = "SC" Then
         Set vObject = objPort
       Else
         Set vObject = objSmif
       End If
       With objSmif
        .PODInPlace = True
       End With
'--- Modify for IRR047 by hsiulin on 6/29
'    sRMMode = Parent.GetPuObject.ControlMode
    Dim iPos As Integer
    iPos = InStr(Parent.GetPuObject.ControlMode, "/")
    If iPos > 0 Then
        sRMMode = Mid(Parent.GetPuObject.ControlMode, 1, iPos - 1)
    Else
        sRMMode = Parent.GetPuObject.ControlMode
    End If
'---
       sAcceptableLevel = Parent.GetTaskAttribute(TaskID, "ACCEPTABLE_LEVEL")
       bolContinue = False
    
       Select Case (sRMMode)
         Case "ONLINE_REMOTE"
            bolContinue = True
            
         Case "ONLINE_LOCAL"
            If sAcceptableLevel = "LOCAL" Or _
               sAcceptableLevel = "" Then
                bolContinue = True
            End If
       End Select
       If bolContinue = True Then
         sTaskListName = Parent.GetTaskAttribute(TaskID, "TASK_LIST")
         If sTaskListName = "" Then
            sTaskListName = Parent.GetTaskAttribute(TaskID, "TaskList")
         End If
         
         'This action is initial "State" property
        ' If Parent.GetPuObject.SmifList(1).PODInPlace = False And Parent.GetPuObject.SmifList(2).PODInPlace = False Then
        '     Parent.GetPuObject.Attributes("State").Value = ""
        ' End If
         '--------------------------------------------------------
         
         If sTaskListName <> "" Then
            Call Parent.StartTaskList(sTaskListName, vObject)
         Else
            Call Parent.TaskComplete(TaskID, "0", vObject)
         End If
       Else
         Call Parent.LogTrace(0, "Equipment mode is not acceptable to run")
         Call AlarmReport(Parent, Me, "10010", "Equipment control mode is not acceptable to run")
         Call Parent.TaskComplete(TaskID, "1", vObject)
       End If
    Else
       Call Parent.LogTrace(0, "Can't find the related smif")
       Call AlarmReport(Parent, Me, "10011", "Can't find the related smif")
       Call Parent.TaskComplete(TaskID, "1", vObject)
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



