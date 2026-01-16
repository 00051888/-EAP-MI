VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MQ_EQ_PostDoorStatus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements ts_ECAncestor.TaskAncestor

Private Const MODULE_ID = "MQ_EQ_PostDoorStatus"

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
    
   With objnewVfei
       .CommandID = "QUERY_STATUS"
       .MachineID = objEqDrv.EqDrvID
       .MessageType = "C"
       .AddVar "PARAMETER", "A", "PORT_STATUS"
   End With
   Call Parent.LogTrace(0, "Query 'PORT_STSTUS'(ProcessEnd)")
   Call Parent.LogEvent("QUERY_STATUS", GetLotListStr, "PARAMETER='PORT_STATUS'")
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
    Dim objLSmif As SmifItem
    Dim objRSmif As SmifItem
    Dim Exp16 As Long
    Dim LBinaryStr As String
    Dim RBinaryStr As String
    Dim BinaryStr As String
    Dim i As Integer
    Dim Check_Flag_R As Boolean
    Dim Check_Flag_L As Boolean
    Dim Status_Flag_R As Boolean
    Dim Status_Flag_L As Boolean
    
    On Error GoTo ErrorHandler
    
    
    If objVfei.CommandID <> "QUERY_STATUS" Then Exit Sub
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
    '___________________________________________________
    ' Define Which Port Need to Check
    Check_Flag_R = True
    Check_Flag_L = True
    
      
    For Each objLot In objBatch.LotList
        If objLot.Port = "1" Then Check_Flag_L = False
        If objLot.Port = "2" Then Check_Flag_R = False
    Next objLot
      
        
        If (LBinaryStr And "10100101") = "10100101" Then
            Set objLSmif = Parent.GetPuObject.SmifList("SMIF1")
            Call Parent.LogTrace(0, "PORT1 Door open ,and Cassette present")
            Call Parent.LogEvent("PORT_STSTUS", "PORT" & objLSmif.RelatedPortID, "PORT1 Door open ,and Cassette present")
            Check_Flag_L = True
         
        Else
            Set objLSmif = Parent.GetPuObject.SmifList("SMIF1")
            Call Parent.LogEvent("PORT_STSTUS", "PORT" & objLSmif.RelatedPortID, "Port1 Status isn't Correct")
            'Call Parent.LogTrace(0, "Port1 Status isn't Correct")
        End If
               
        If (RBinaryStr And "10100101") = "10100101" Then
            Set objRSmif = Parent.GetPuObject.SmifList("SMIF2")
            Call Parent.LogTrace(0, "PORT2 Door open ,and Cassette present")
            Call Parent.LogEvent("PORT_STSTUS", "PORT" & objRSmif.RelatedPortID, "PORT2 Door open ,and Cassette present")
            Check_Flag_R = True
          
        Else
            Set objRSmif = Parent.GetPuObject.SmifList("SMIF2")
            Call Parent.LogEvent("PORT_STSTUS", "PORT" & objLSmif.RelatedPortID, "Port2 Status isn't Correct")
            'Call Parent.LogTrace(0, "Port2 Status isn't Correct")
        End If
    
        If Check_Flag_R = True And Check_Flag_L = True Then
            Call Parent.TaskComplete(TaskID, "OK")
        Else
            Call Parent.LogEvent("PORT_STSTUS", "PORT1 or PORT2", "Port1 or Port2 Status isn't Correct")
            Call Parent.LogTrace(0, "EQ Door Status Error")
            'Call AlarmReport(Parent, Me, 1, "檢查機門狀態錯誤", "檢查機門狀態錯誤")
            Call Parent.TaskComplete(TaskID, "FAIL")
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



