VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MS_RM_ChkMoveOutQty"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit


Implements ts_ECAncestor.TaskAncestor

Private Const MODULE_ID = "MS_RM_ChkMoveInQty"

Dim Parent As tsEC.clsRunManager
Dim TaskID As String
Dim objBatch As Batch
Dim EC_Key_String As String
Dim OutputStatus As String
Dim BatchLots As Integer


Private Sub TaskAncestor_Execute(Parameter As Variant)
    Dim objLot As Lot
    Dim vTemp As Variant
    Dim TempStr As String
    Dim isFind As Boolean
    Dim sLotID As String
    
    On Error GoTo ErrorHandler

    Set objBatch = Parameter
    
    On Error Resume Next
    
    For Each objLot In objBatch.LotList
       vTemp = objLot.Attributes.Item("ProcessedCount").Value
       If vTemp <> objLot.Qty Then
             OutputStatus = "<>"
             Exit For
       Else
         OutputStatus = "="
       End If
    Next
    
    If objBatch.RunMode = "PR" Then

       TempStr = UCase$(Parent.GetTaskAttribute(TaskID, "USER_DECIDE"))

       If OutputStatus <> "=" Then
          If TempStr = "Y" Then
             EC_Key_String = AlarmReport(Parent, Me, "2100", Translate(Parent, "矪瞶Ч计ぃ才", "Processed quantity is not match"), _
                  "Lot= " & objLot.LotID & Translate(Parent, " , MES 计秖=", " ,MES quantity=") & CStr(objLot.Qty) & Translate(Parent, " 矪瞶Ч计 =", " ,Processed quantity=") & objLot.Attributes.Item("ProcessedCount").Value _
                  , HoldLot_Continue)

             Call Parent.SubscribeHostMsg(Me, "ALARM_ACK_TO_EC")
          Else
             EC_Key_String = AlarmReport(Parent, Me, "2100", Translate(Parent, "矪瞶Ч计ぃ才", "Processed quantity is not match"), _
                 "Lot= " & objLot.LotID & Translate(Parent, " , MES 计秖=", " ,MES quantity=") & CStr(objLot.Qty) & Translate(Parent, " 矪瞶Ч计 =", " ,Processed quantity=") & objLot.Attributes.Item("ProcessedCount").Value _
                 , AckHold)
             Call Parent.TaskComplete(TaskID, OutputStatus)
          End If
       Else
          Call Parent.TaskComplete(TaskID, OutputStatus)
       End If
    Else
    'Pilot 1999,8,3
        If OutputStatus <> "=" Then

           EC_Key_String = AlarmReport(Parent, Me, "2100", Translate(Parent, "矪瞶Ч计ぃ才", "Processed quantity is not match"), _
                     "Lot= " & objLot.LotID & Translate(Parent, " , MES 计秖=", " ,MES quantity=") & CStr(objLot.Qty) & Translate(Parent, " 矪瞶Ч计 =", " ,Processed quantity=") & objLot.Attributes.Item("ProcessedCount").Value _
                     , OK)

        End If
            Call Parent.TaskComplete(TaskID, OutputStatus)
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

    If objHostTrans.TransactionName = "ALARM_ACK_TO_EC" Then
        With objHostTrans.Primary
            If .Item("EC_KEY_STRING") = EC_Key_String Then
                objHostTrans.Reply
                Select Case .Item("CHOICE")
                    Case "HOLDLOT"
                        Call Parent.TaskComplete(TaskID, OutputStatus)
                    Case "CONTINUE"
                        Call Parent.TaskComplete(TaskID, "=")
                End Select
            End If
        End With
    End If

    Exit Sub
ErrorHandler:
    Call Parent.LogTrace(0, FormatErrMsg(MODULE_ID & ".ReceivedHostMsg", Err, Err.Description))
End Sub

Private Sub TaskAncestor_ReceivedVfei(objVfei As vfei.Message22)
    On Error GoTo ErrorHandler

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



