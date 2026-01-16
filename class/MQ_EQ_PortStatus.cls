VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MQ_EQ_PreDoorStatus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements ts_ECAncestor.TaskAncestor

Private Const MODULE_ID = "MQ_EQ_PreDoorStatus"

Dim Parent As tsEC.clsRunManager
Dim TaskID As String
Dim objBatch As Batch
Dim objLot As Lot
Dim objPort As PortItem
Dim objSmif As SmifItem
Dim objEqDrv As EqDrv
Dim vParameter As Variant

Private Sub Class_Terminate()
    Set objBatch = Nothing
    Set objLot = Nothing
    Set objPort = Nothing
    Set objSmif = Nothing
    Set objEqDrv = Nothing

End Sub

Private Sub TaskAncestor_Execute(Parameter As Variant)
    Dim objnewVfei As New Message22
    Dim sPortID As String
    Dim sSmifID As String
    Dim sRMMode As String
    Dim Parray(1 To 2) As String
    
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
        Set objSmif = Parent.GetPuObject.SmifList(sSmifID)
    ElseIf TypeOf vParameter Is SmifItem Then
        Set objSmif = vParameter
        sSmifID = objSmif.SMIFID
        sPortID = objSmif.RelatedPortID
        Set objPort = Parent.GetPuObject.PortList(sPortID)
        Set objEqDrv = Parent.GetPuObject.EqDrvList(1)
    ElseIf TypeOf vParameter Is Batch Then
        Set objBatch = vParameter
        Set objEqDrv = Parent.GetPuObject.EqDrvList(1)
    ElseIf TypeOf vParameter Is Lot Then
        Set objLot = vParameter
        sPortID = objLot.Port
        Set objPort = Parent.GetPuObject.PortList(sPortID)
        sSmifID = objPort.RelatedSMIFID
        Set objSmif = Parent.GetPuObject.SmifList(sSmifID)
    End If
    
    Parray(1) = "PROCESS_STATE"
    Parray(2) = "PORT_STATUS"
    With objnewVfei
         .CommandID = "QUERY_STATUS"
         .MachineID = objEqDrv.EqDrvID
         .MessageType = "C"
         .AddVar "PARAMETER", "A", Parray
'
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
    Dim TempLng As Long

    Dim Exp16 As Long
    Dim LBinaryStr As String
    Dim RBinaryStr As String
    Dim BinaryStr As String
    Dim i As Integer
    Dim RightStatus As Boolean
    
    On Error GoTo ErrorHandler
    
    If objVfei.CommandID <> "QUERY_STATUS" Then Exit Sub
    RightStatus = False
    TempLng = CLng(objVfei.Var("PORT_STATUS").Value)
    '--------------------------------------------------------------------------
    'Change Format
    Exp16 = Exp(16 * Log(2))
    
    For i = 1 To 16
       Exp16 = Exp16 / 2
       If ((Exp16) And TempLng) = Exp16 Then
          BinaryStr = BinaryStr & "1"
       Else
          BinaryStr = BinaryStr & "0"
       End If
    Next i
    '--------------------------------------------------------------------------
    LBinaryStr = Left(BinaryStr, 8)
    RBinaryStr = Right(BinaryStr, 8)
    
    If (LBinaryStr And "10100001") = "10100001" Then
       Call Parent.LogTrace(0, "Port1 Door open")
       Parent.GetPuObject.Attributes("Port1State").Value = "Open"
    Else
       Call Parent.LogTrace(0, "Port1 Door close")
       Parent.GetPuObject.Attributes("Port1State").Value = "Close"
    End If
    If (RBinaryStr And "10100001") = "10100001" Then
       Call Parent.LogTrace(0, "Port2 Door open")
       Parent.GetPuObject.Attributes("Port2State").Value = "Open"
    Else
       Call Parent.LogTrace(0, "Port2 Door close")
       Parent.GetPuObject.Attributes("Port2State").Value = "Close"
    End If
    
    If objBatch.LotList.Count = 2 Then
       If (Parent.GetPuObject.Attributes("Port1State").Value = "Close") Or (Parent.GetPuObject.Attributes("Port2State").Value = "Close") Then
          Call Parent.LogEvent("Door Open Failed", _
               GetLotListStr, "PortID:" & GetPortListStr)
          Call AlarmReport(Parent, Me, "011", "The doors open failed", "LotID:" & GetLotListStr)
          Call Parent.LogTrace(0, "The doors open failed!!!")
          Call Parent.TaskComplete(TaskID, "1", vParameter)
       End If
    ElseIf objBatch.LotList.Count = 1 Then
       If Parent.GetPuObject.Attributes("Port" & objBatch.LotList(1).Port & "State").Value = "Close" Then
          Call Parent.LogEvent("Door Open Failed", _
               GetLotListStr, "PortID:" & GetPortListStr)
          Call AlarmReport(Parent, Me, "011", "The doors open failed", "LotID:" & GetLotListStr)
          Call Parent.LogTrace(0, "The doors open failed!!!")
          Call Parent.TaskComplete(TaskID, "1", vParameter)
       End If
    Else
       Call Parent.LogTrace(0, "Information Error")
       Call Parent.TaskComplete(TaskID, "1", vParameter)
    End If
    
    Select Case objVfei.Var("PROCESS_STATE").Value
    '---2007-03-28   Modify MI-4 IMP status check function, allow move in on case 2,4,5,6 by ShangChiun
       Case "1", "2", "3", "4", "5", "6"
          Call Parent.TaskComplete(TaskID, "0", vParameter)
    
    '
    '   Case "2"
    '      Call Parent.LogEvent("System's Process State is Loading/UnLoading(2)", _
    '                            GetLotListStr, "PortID:" & GetPortListStr)
    '      Call AlarmReport(Parent, Me, "021", "╰参矪瞶篈瞷琌 Loading/UnLoading(2)", "LotID:" & GetLotListStr)
    '      Call Parent.LogTrace(0, "System's Process State is Loading/UnLoading(2)!!!")
    '      Call Parent.TaskComplete(TaskID, "1", vParameter)
    '   Case "4"
    '      Call Parent.LogEvent("System's Process State is SetUp(4)", _
    '               GetLotListStr, "Port:" & GetPortListStr)
    '      Call AlarmReport(Parent, Me, "023", "╰参矪瞶篈瞷琌 SetUp(4)", "LotID:" & GetLotListStr)
    '      Call Parent.LogTrace(0, "System's Process State is SetUp(4)!!!")
    '      Call Parent.TaskComplete(TaskID, "1", vParameter)
    '   Case "5"
    '      Call Parent.LogEvent("System's Process State is Implanting(5)", _
    '             GetLotListStr, "PortID:" & GetPortListStr)
    '      Call AlarmReport(Parent, Me, "024", "╰参矪瞶篈瞷琌 Implanting(5)", "LotID:" & GetLotListStr)
    '      Call Parent.LogTrace(0, "System's Process State is Implanting(5)!!!")
    '      Call Parent.TaskComplete(TaskID, "1", vParameter)
    '   Case "6"
    '      Call Parent.LogEvent("System's Process State is Implant Hold(6)", _
    '             GetLotListStr, "PortID:" & GetPortListStr)
    '      Call AlarmReport(Parent, Me, "025", "╰参矪瞶篈瞷琌 Implant Hold(6)", "LotID:" & GetLotListStr)
    '      Call Parent.LogTrace(0, "System's Process State is Implant Hold(6)!!!")
    '     Call Parent.TaskComplete(TaskID, "1", vParameter)
       Case Else
          Call Parent.LogEvent("╰参篈ぃ", _
                   GetLotListStr, "PortID:" & GetPortListStr)
          Call AlarmReport(Parent, Me, "022", Translate(Parent, "╰参篈ぃ", "EQ System Status is Unknow."), "LotID:" & GetLotListStr)
          Call Parent.LogTrace(0, "╰参篈ぃ!!!")
          Call Parent.TaskComplete(TaskID, "1", vParameter)
    End Select
       
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

