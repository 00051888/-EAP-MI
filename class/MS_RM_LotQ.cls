VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MS_RM_LotQ"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements ts_ECAncestor.TaskAncestor

Private Const MODULE_ID = "MS_RM_LotQ"

Dim parent As tsEC.clsRunManager
Dim TaskID As String
Dim objBatch As Batch
Dim objLot As Lot
Dim objPort As PortItem
Dim objSmif As SmifItem
Dim sSmifID As String
Dim sPortID As String

Dim vParameter As Variant


Private Sub TaskAncestor_Execute(Parameter As Variant)
    Dim objnewVfei As New Message22
    Dim objTmpBatch As Batch
    Dim objTmpLot As Lot
    Dim bSendMsg As Boolean
    Dim sTemp As String
    
    On Error GoTo ErrorHandler
    
    If IsObject(Parameter) Then
        Set vParameter = Parameter
    Else
        vParameter = Parameter
    End If
    
    If TypeOf vParameter Is PortItem Then
        Set objPort = vParameter
        sPortID = objPort.PortID
        sSmifID = objPort.RelatedSMIFID
        Set objSmif = parent.GetPuObject.SmifList(sSmifID)
    ElseIf TypeOf vParameter Is SmifItem Then
        Set objSmif = vParameter
        sSmifID = objSmif.SMIFID
        sPortID = objSmif.RelatedPortID
        Set objPort = parent.GetPuObject.PortList(sPortID)
    ElseIf TypeOf vParameter Is Batch Then
        Set objBatch = vParameter
        Set objLot = objBatch.LotList(1)
        sPortID = objLot.Port
        Set objPort = parent.GetPuObject.PortList(sPortID)
        sSmifID = objPort.RelatedSMIFID
        Set objSmif = parent.GetPuObject.SmifList(sSmifID)
    ElseIf TypeOf vParameter Is Lot Then
        Set objLot = vParameter
        sPortID = objLot.Port
        Set objPort = parent.GetPuObject.PortList(sPortID)
        sSmifID = objPort.RelatedSMIFID
        Set objSmif = parent.GetPuObject.SmifList(sSmifID)
    End If
    Call parent.LogTrace(0, MODULE_ID & " PU State=" & parent.GetPuObject.Attributes("State").Value)
    If parent.GetPuObject.Attributes("State").Value <> "Wait" Then
       'parent.GetPuObject.Attributes("State").Value = "Wait"
       Call parent.LogTrace(0, MODULE_ID & " Execute ,Receive the First(" & objSmif.RelatedCassetteID & ") Cassette, PU State=" & parent.GetPuObject.Attributes("State").Value)
       Call parent.TaskComplete(TaskID, "0", vParameter)
    Else
       'On Error Resume Next
       Call parent.LogTrace(0, MODULE_ID & " .Receive the Second(" & objSmif.RelatedCassetteID & ") Cassette, PU State=" & parent.GetPuObject.Attributes("State").Value)
       For Each objTmpBatch In parent.GetBatchList
             For Each objTmpLot In objTmpBatch.LotList
               ' If Err.Number = 0 Then
                   If objTmpLot.Cassette = objSmif.RelatedCassetteID Then
                      Call parent.LogTrace(0, MODULE_ID & ".Execute ,Batch information has created")
                      parent.GetPuObject.Attributes("State").Value = "Run"
                      Call parent.LogTrace(0, "The second CassetteID =" & objSmif.RelatedCassetteID & ", PU State: " & parent.GetPuObject.Attributes("State").Value)
                      objnewVfei.CommandID = "LOT_Q"
                      objnewVfei.MachineID = parent.GetPuObject.EqDrvList(1).EqDrvID
                      objnewVfei.MessageType = "E"
                      objnewVfei.AddVar "SMIFID", "A", sSmifID
                      objnewVfei.AddVar "PORT", "A", sPortID
                      Call parent.SendVfei(objnewVfei, Me)
                      Call parent.TaskComplete(TaskID, "2", vParameter)
                      Exit Sub
                   End If
             Next objTmpLot
       Next objTmpBatch
    
    
    
       If Not objTmpBatch Is Nothing Then
          Call parent.LogTrace(0, MODULE_ID & ".Execute ,The second CassetteID of Batch isn't same with Pod CassetteID(" & objSmif.RelatedCassetteID & ")")
          Call AlarmReport(parent, Me, "002", Translate(parent, "晶舟號碼的錯誤:", "Cassette is Error:"), Translate(parent, "SMIF 上的晶舟號碼是: ", "Cassette on SMIF  is ") & objSmif.RelatedCassetteID & Translate(parent, " ,請將所有晶舟重新放置 SMIF上, SMIFID=", ",Please put the cassette on the SMIF again,SMIFID=") & objSmif.SMIFID)
          Call parent.TaskComplete(TaskID, "1", vParameter)
          Exit Sub
       Else
          Call parent.LogTrace(0, MODULE_ID & ".Execute , Batch Info hasn't Created, and Wait 'RUN_CASSETTE' Event")
          Call parent.SubscribeVfeiMsg(Me, parent.GetPuObject.EqDrvList(1).EqDrvID, "RUN_CASSETTE")
          Exit Sub
       End If
    End If


    Exit Sub
ErrorHandler:
    Call parent.LogTrace(0, FormatErrMsg(MODULE_ID & ".Execute", Err, Err.Description))
End Sub

Private Property Set TaskAncestor_Parent(ByVal RHS As Object)
   Set parent = RHS
End Property

Private Property Get TaskAncestor_Parent() As Object
   Set TaskAncestor_Parent = parent

End Property



Private Sub TaskAncestor_ReceivedHostMsg(objHostTrans As HostOcx.HostTransaction, isPrimary As Boolean)
    On Error GoTo ErrorHandler

    Exit Sub
ErrorHandler:
    Call parent.LogTrace(0, FormatErrMsg(MODULE_ID & ".ReceivedHostMsg", Err, Err.Description))
End Sub

Private Sub TaskAncestor_ReceivedVfei(objVfei As vfei.Message22)
    Dim objnewVfei As New vfei.Message22
    
    On Error GoTo ErrorHandler
    
    If objVfei.CommandID <> "RUN_CASSETTE" Then Exit Sub
    
    Call parent.LogTrace(0, MODULE_ID & " ReceivedVfei ,Recieve 'RUN_CASSETTE EVENT'")

    If objVfei.Var("CASSETTEID").Value = objSmif.RelatedCassetteID Then
       parent.GetPuObject.Attributes("State").Value = "Run"
       Call parent.LogTrace(0, MODULE_ID & " ReceivedVfei ,Send LOT_Q Vfei Message ,PU State :" & parent.GetPuObject.Attributes("State").Value)
       objnewVfei.CommandID = "LOT_Q"
       objnewVfei.MachineID = parent.GetPuObject.EqDrvList(1).EqDrvID
       objnewVfei.MessageType = "E"
       objnewVfei.AddVar "SMIFID", "A", objSmif.SMIFID
       objnewVfei.AddVar "PORTID", "A", objPort.PortID
       Call parent.SendVfei(objnewVfei, Me)
       Call parent.TaskComplete(TaskID, "2", vParameter)
    Else
       'Parent.GetPuObject.Status = "Ready"
       Call parent.LogTrace(0, MODULE_ID & " ReceivedVfei ,The CassetteID of Lot object is " & objVfei.Var("CASSETTEID").Value & ",But Pod placed CassetteID is " & objSmif.RelatedCassetteID)
       Call parent.LogEvent("SMIF Pod placed CassetteID:", _
                             objSmif.RelatedCassetteID, "MES CassetteID:" & objVfei.Var("CASSETTEID").Value & " SMIFID=" & objSmif.SMIFID)
       Call AlarmReport(parent, Me, "003", Translate(parent, "晶舟號碼的錯誤:", "Cassette is Error:"), Translate(parent, "SMIF 上的晶舟號碼是: ", "Cassette on SMIF  is ") & objSmif.RelatedCassetteID & " , MES 上的晶舟號碼是: " & objVfei.Var("CASSETTEID").Value & Translate(parent, " ,請將所有晶舟重新放置 SMIF上, SMIFID=", ",Please put the cassette on the SMIF again,SMIFID=") & objSmif.SMIFID)
       Call parent.TaskComplete(TaskID, "1", vParameter)

    End If
    Exit Sub
ErrorHandler:
    Call parent.LogTrace(0, FormatErrMsg(MODULE_ID & ".ReceivedVfei", Err, Err.Description))
End Sub

Private Property Let TaskAncestor_TaskID(ByVal RHS As String)
    TaskID = RHS
End Property

Private Property Get TaskAncestor_TaskID() As String
    TaskAncestor_TaskID = TaskID
End Property

Private Sub Class_Terminate()
    Set objBatch = Nothing
    Set objLot = Nothing
    Set objPort = Nothing
    Set objSmif = Nothing
    If IsObject(vParameter) Then
        Set vParameter = Nothing
    End If
End Sub

