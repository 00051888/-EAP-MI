VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MW_SMIF_UnLock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements ts_ECAncestor.TaskAncestor

Private Const MODULE_ID = "MW_SMIF_UnLock"

Dim Parent As tsEC.clsRunManager
Dim TaskID As String

Dim sPort As String

Dim objBatch As Batch
Dim objPort As PortItem
Dim objSmif As SmifItem

Dim vParameter As Variant

Private Sub TaskAncestor_Execute(Parameter As Variant)
    Dim sSmifID As String
    Dim sPortID As String
    Dim sIsToEq As String
    Dim objLot As Lot
    Dim objTmpPort As PortItem
    
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
    End If
    On Error GoTo ErrorHandler
    
    If IsObject(Parameter) Then
        Set vParameter = Parameter
    Else
        vParameter = Parameter
    End If
   Set objBatch = vParameter
                    
     sIsToEq = Parent.GetTaskAttribute(TaskID, "TO_EQ")
    If sIsToEq <> "" And (UCase(sIsToEq) = "TRUE" Or UCase(sIsToEq) = "Y") Then
        For Each objLot In objBatch.LotList
            Call Parent.SubscribeVfeiMsg(Me, "", "UNLOCK_PORT_COMPLETED", "PORT", objLot.Port)
        Next
        
    Else
        For Each objLot In objBatch.LotList
            Set objTmpPort = Parent.GetPuObject.PortList(objLot.Port)
            Call Parent.SubscribeVfeiMsg(Me, "", "UNLOCK_PORT_COMPLETED", "SMIFID", objTmpPort.RelatedSMIFID)
        Next
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
    bWait = False
    sSmifID = objVfei.Var("SMIFID").Value
    sPortID = objVfei.Var("SMIFPORTID").Value
    Set objSmif = Parent.GetPuObject.SmifList(sSmifID)
    Set objPort = Parent.GetPuObject.PortList(sPortID)

    For Each objLot In objBatch.LotList
       If objLot.Status <> "UNLOCK" Then
          objLot.Status = "UNLOCK"
          Call Parent.LogTrace(0, "Recived POD_UNLOCK")
          With objSmif
             .RelatedCassetteID = ""
             .PODInPlace = False
             .Locked = False
          End With
          Exit For
       End If
    Next objLot
    
    For Each objLot In objBatch.LotList
       If objLot.Status <> "UNLOCK" Then
          bWait = True
          Exit For
       End If
    Next objLot
    
    If bWait = False Then
       Call Parent.TaskComplete(TaskID, "")
       Set objBatch = Nothing
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



