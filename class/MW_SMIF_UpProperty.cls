VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MW_SMIF_UpProperty"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements ts_ECAncestor.TaskAncestor

Private Const MODULE_ID = "MW_SMIF_UpProperty"

Dim Parent As tsEC.clsRunManager
Dim TaskID As String
Dim objBatch As Batch
Dim objLot As Lot
Dim objPort As PortItem
Dim objSmif As SmifItem

Dim vParameter As Variant

Private Sub TaskAncestor_Execute(Parameter As Variant)
    Dim sSmifID As String
    Dim sPortID As String
    Dim sINPUT As String
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
    ElseIf TypeOf Parameter Is Batch Then
        Set objBatch = Parameter
    ElseIf TypeOf Parameter Is Lot Then
        Set objLot = Parameter
        sPortID = objLot.Port
        Set objPort = Parent.GetPuObject.PortList(sPortID)
        sSmifID = objPort.RelatedSMIFID
        Set objSmif = Parent.GetPuObject.SmifList(sSmifID)
    End If
    sINPUT = Parent.GetTaskAttribute(TaskID, "INPUT")
    If sINPUT = "Y" Then
       Call Parent.SubscribeVfeiMsg(Me, "", "POD_REMOVE", "PORT", objSmif.RelatedPortID)
    Else
       Call Parent.SubscribeVfeiMsg(Me, "", "POD_REMOVE")
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
    Dim objLot As Lot
    Dim TempStr As String
    Dim isFind As Boolean
    Dim bWait As Boolean
    
    On Error Resume Next
    
    sPortID = objVfei.Var("SMIFPORTID").Value
    Set objPort = Parent.GetPuObject.PortList(sPortID)
    sSmifID = objPort.RelatedSMIFID
    Set objSmif = Parent.GetPuObject.SmifList(sSmifID)
    
    Parent.GetPuObject.PortList(CStr(objLot.Port)).State = PORT_EMPTY
    Call Parent.LogTrace(0, "Receive PORT" & sPortID & " POD_REMOVE")
    Call Parent.LogEvent("POD_REMOVE", objSmif.SMIFID, "Received POD_REMOVE Event")
    With objSmif
       .RelatedCassetteID = ""
       .PODInPlace = False
       .Locked = False
    End With
    
    Call Parent.LogTrace(0, objSmif.SMIFID & " PODInPlace=" & objSmif.PODInPlace)
    Call Parent.LogEvent("PODInPlace", objSmif.SMIFID, objSmif.PODInPlace)
    If (Parent.GetPuObject.SmifList(1).PODInPlace = False) And (Parent.GetPuObject.SmifList(2).PODInPlace = False) Then
       Parent.GetPuObject.Attributes("State").Value = ""
    End If
    
    'Call Parent.TaskComplete(TaskID, "")
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


