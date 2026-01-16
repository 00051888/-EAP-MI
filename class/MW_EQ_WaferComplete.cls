VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MW_EQ_WAFERCOMPLETE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit


Implements ts_ECAncestor.TaskAncestor

'For Check Double Implanted
Private Const MODULE_ID = "MW_EQ_WaferComplete"

Dim Parent As tsEC.clsRunManager
Dim TaskID As String
Dim objBatch As Batch
Dim iIndex As Integer
Dim sRotate As Integer

Private Sub TaskAncestor_Execute(Parameter As Variant)
    Dim objEqDrv As EqDrv
    Dim objlot As Lot
    
    On Error GoTo ErrorHandler
    
    Set objBatch = Parameter
    
    For Each objlot In objBatch.LotList
        objlot.Attributes.Add("ImplantedString") = "00000000000000000000000000000000000000000000000000"
    Next
    
    For Each objEqDrv In Parent.GetPuObject.EqDrvList
       Call Parent.SubscribeVfeiMsg(Me, objEqDrv.EqDrvID, "Wafer_complete")
       Call Parent.SubscribeVfeiMsg(Me, objEqDrv.EqDrvID, "PS")
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
    
    Exit Sub
ErrorHandler:
    Call Parent.LogTrace(0, FormatErrMsg(MODULE_ID & ".ReceivedHostMsg", Err, Err.Description))
End Sub

Private Sub TaskAncestor_ReceivedVfei(objVfei As vfei.Message22)
    Dim TempStr As String
    Dim objlot As Lot
    Dim isFind As Boolean
    Dim isBatch As Boolean
    Dim objAttribute As tsBatchLot.ExtAttribute
    Dim sHoldComment As String
    Dim iSlotID As Integer
    Dim strMappingstring As String
    Dim objMyVfei As Message22
        
    On Error GoTo ErrorHandler
    
'Yen add for MI-4 dual function
For Each objlot In objBatch.LotList
If objVfei.Var("CONTROLLER").Value <> objlot.LotID Then Exit Sub
Next
'Yen
    
    'Rotate (Wafer Complete Information )
    If objVfei.CommandID = "PS" Then
        sRotate = objVfei.Var("ROTATE").Value
    Else

    On Error Resume Next
    TempStr = objVfei.Var("LOT_ID").Value
    iSlotID = objVfei.Var("SLOT").Value
    
    'If iSlotID > 25 Then iSlotID = iSlotID - 25
       
    'For Check Bouble ImplantedString
    For Each objlot In objBatch.LotList
          If objlot.LotID = TempStr Then
                Set objAttribute = objlot.Attributes("ImplantedString")
                strMappingstring = objAttribute.Value
                          
                    'Implanted String Compare
                    If Mid(strMappingstring, iSlotID, 1) > sRotate * objBatch.LotList.Count - 1 Then
                            '~~Stop Eq when abnormal implant. Shang-Chun Add on 2008/09/03
                            Set objMyVfei = New Message22
                            objMyVfei.CommandID = "REMOTE_COMMAND"
                            objMyVfei.MachineID = Parent.GetPuObject.EqDrvList(1).EqDrvID
                            objMyVfei.AddVar "COMMAND", "A", "STOP"
                            objMyVfei.MessageType = "C"
                            
                            Call Parent.SendVfei(objMyVfei, Me)
                            Call Parent.LogTrace(0, "Send remote command <STOP> to eq.")
                            '~~
                    
                    
                    
                    sHoldComment = "LotID=" & objlot.LotID & " 疑重複稙入,於第" & CStr(iSlotID) & "開始"
                    objlot.Attributes.Add("HoldComment").Value = sHoldComment
                     
                    
                    Call AlarmReport(Parent, Me, "107", Translate(Parent, " 疑重複稙入,於第", "Suspected repeated implantation.Starting from Slot: ") & CStr(iSlotID) & Translate(Parent, "開始,請迅速通知製程和設備工程師", " .Please inform PE and EQ Engineer as soon as possible "), sHoldComment, AckHold)
                    Call Parent.TaskComplete(TaskID, "FAIL")
                            
                            Else
                            Mid(strMappingstring, iSlotID, 1) = Mid(strMappingstring, iSlotID, 1) + 1
                            objlot.Attributes("ImplantedString").Value = strMappingstring
                    End If
          End If
    Next
    
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


