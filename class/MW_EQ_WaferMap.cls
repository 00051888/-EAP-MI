VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MW_EQ_WaferMap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit


Implements ts_ECAncestor.TaskAncestor

Private Const MODULE_ID = "MW_EQ_WaferMap"

Dim Parent As tsEC.clsRunManager
Dim TaskID As String
Dim objBatch As Batch
Dim objSmif As SmifItem
Dim objPort As PortItem
Dim objlot As Lot
Dim ReplyStr As String
Dim EC_Key_String As String
Dim OutputStatus As String


Private Sub Class_Terminate()
    Set objBatch = Nothing
    Set objPort = Nothing
    Set objSmif = Nothing
End Sub

Private Sub TaskAncestor_Execute(Parameter As Variant)
    Dim objEqDrv As EqDrv
    
    On Error GoTo ErrorHandler
    
    Set objBatch = Parameter
    
    For Each objEqDrv In Parent.GetPuObject.EqDrvList
       Call Parent.SubscribeVfeiMsg(Me, objEqDrv.EqDrvID, "WAFER_MAPPING")
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
    'ALARM_ACK_TO_EC <%Subject> <%TID> <%ID>  <ALARM_ID> <CHOICE> <EC_KEY_STRING>
     If objHostTrans.TransactionName = "ALARM_ACK_TO_EC" Then
        If objHostTrans.Primary.Item("EC_KEY_STRING") = EC_Key_String Then
            objHostTrans.Reply
            Select Case objHostTrans.Primary.Item("CHOICE")
                Case "HOLDLOT"
                    Call Parent.TaskComplete(TaskID, ReplyStr)
                Case "CONTINUE"
                    Call Parent.TaskComplete(TaskID, "=")
            End Select
        End If
    End If
    
    Exit Sub
ErrorHandler:
    Call Parent.LogTrace(0, FormatErrMsg(MODULE_ID & ".ReceivedHostMsg", Err, Err.Description))
End Sub

Private Sub TaskAncestor_ReceivedVfei(objVfei As vfei.Message22)
    Dim objlot As Lot
    Dim bWait As Boolean
    Dim vTemp As Variant
    Dim TempStr As String
    Dim objAttribute As tsBatchLot.ExtAttribute
    Dim sHoldComment As String
    Dim sAlarmComment As String
    
    On Error GoTo ErrorHandler
    
'Yen add for MI-4 dual function
For Each objlot In objBatch.LotList
If objVfei.Var("CONTROLLER").Value <> objlot.LotID Then Exit Sub
Next
'Yen
    
    If objVfei.CommandID <> "WAFER_MAPPING" Then Exit Sub
    bWait = False
    
    For Each objlot In objBatch.LotList
       If objlot.Status <> "MAPPING" Then
          objlot.Status = "MAPPING"
          Exit For
       End If
    Next
    
    For Each objlot In objBatch.LotList
       If objlot.Status <> "MAPPING" Then
          bWait = True
          Exit For
       End If
    Next
    
    On Error Resume Next
    If bWait = False Then
       For Each objlot In objBatch.LotList
          If objlot.LotID = Trim$(objVfei.Var("LLOT").Value) Then
             objlot.Attributes.Add("MappingCount").Value = objVfei.Var("LWAF").Value
             
             TempStr = "0000000000000000000000000"
             If Err.Number = 0 Then
                 Set objAttribute = objlot.Attributes("MappingString")
                 If Err.Number = 0 Then
                    objAttribute.Value = TempStr
                 Else
                    objlot.Attributes.Add("MappingString").Value = TempStr
                 End If
             End If
          ElseIf objlot.LotID = Trim$(objVfei.Var("RLOT").Value) Then
             objlot.Attributes.Add("MappingCount").Value = objVfei.Var("RWAF").Value
             TempStr = "0000000000000000000000000"
             If Err.Number = 0 Then
                 Set objAttribute = objlot.Attributes("MappingString")
                 If Err.Number = 0 Then
                    objAttribute.Value = TempStr
                 Else
                    objlot.Attributes.Add("MappingString").Value = TempStr
                 End If
             End If
          End If
       Next
       Err.Clear

       If objBatch.RunMode = "PR" Then
          For Each objlot In objBatch.LotList
              vTemp = objlot.Attributes.Item("MappingCount").Value
              If vTemp <> objlot.Qty Then
                  ReplyStr = "<>"
                  Exit For
              Else
                  ReplyStr = "="
              End If
          Next objlot
             
          If ReplyStr <> "=" Then
             TempStr = UCase$(Trim$(Parent.GetTaskAttribute(TaskID, "ACTION")))
             If Len(TempStr) = 0 Then
                TempStr = "OK"
             End If
             Select Case TempStr
                Case "OK"
                     sAlarmComment = "LotID=" & objlot.LotID & Translate(Parent, ",Mes 计秖=", " ,MES quantity=") & CStr(objlot.Qty) & Translate(Parent, ",诀厨计=", " , Equipment quantity=") & objlot.Attributes("MappingCount").Value & "," & "WaferMap=" & TempStr
                     Call AlarmReport(Parent, Me, "107", Translate(Parent, "秈计秖岿粇", "MoveIn quantity error."), sAlarmComment, OK)
                     Call Parent.TaskComplete(TaskID, ReplyStr)

                Case "HOLD"
                      On Error Resume Next
                      TempStr = objlot.Attributes("MappingString")
                      If Err.Number <> 0 Then
                        Err.Clear
                        TempStr = ""
                      End If
                      On Error GoTo ErrorHandler
                      sHoldComment = "LotID=" & objlot.LotID & ",Mes 计秖=" & CStr(objlot.Qty) & ",诀厨计=" & objlot.Attributes("MappingCount").Value & "," & "WaferMap=" & TempStr
                      sAlarmComment = "LotID=" & objlot.LotID & Translate(Parent, ",Mes 计秖=", " ,MES quantity=") & CStr(objlot.Qty) & Translate(Parent, ",诀厨计=", " , Equipment quantity=") & objlot.Attributes("MappingCount").Value & "," & "WaferMap=" & TempStr
                      
                      Call AlarmReport(Parent, Me, "107", Translate(Parent, "秈计秖岿粇", "MoveIn quantity error."), sAlarmComment, AckHold)
                      On Error Resume Next
                      Set objAttribute = objlot.Attributes("HoldComment")
                      If Err.Number = 0 Then
                        objAttribute.Value = sHoldComment
                      Else
                        objlot.Attributes.Add("HoldComment").Value = sHoldComment
                      End If
                      On Error GoTo ErrorHandler
                     Call Parent.TaskComplete(TaskID, ReplyStr)
                Case "HOLD_CONTINUE", "CONTINUE_HOLD"
                      On Error Resume Next
                      TempStr = objlot.Attributes("MappingString")
                      If Err.Number <> 0 Then
                        Err.Clear
                        TempStr = ""
                      End If
                      On Error GoTo ErrorHandler
                      sHoldComment = "LotID=" & objlot.LotID & ",Mes 计秖=" & CStr(objlot.Qty) & ",诀厨计=" & objlot.Attributes("MappingCount").Value & "," & "WaferMap=" & TempStr
                      sAlarmComment = "LotID=" & objlot.LotID & Translate(Parent, ",Mes 计秖=", " ,MES quantity=") & CStr(objlot.Qty) & Translate(Parent, ",诀厨计=", " , Equipment quantity=") & objlot.Attributes("MappingCount").Value & "," & "WaferMap=" & TempStr
                     
                      EC_Key_String = AlarmReport(Parent, Me, "107", Translate(Parent, "秈计秖岿粇", "MoveIn quantity error."), sAlarmComment, HoldLot_Continue)

                      On Error Resume Next
                      Set objAttribute = objlot.Attributes("HoldComment")
                      If Err.Number = 0 Then
                        objAttribute.Value = sHoldComment
                      Else
                        objlot.Attributes.Add("HoldComment").Value = sHoldComment
                      End If
                      On Error GoTo ErrorHandler
                     Call Parent.SubscribeHostMsg(Me, "ALARM_ACK_TO_EC")

             End Select
          Else
             Call Parent.TaskComplete(TaskID, "=")
          End If
       Else
           Call Parent.TaskComplete(TaskID, "=")
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


