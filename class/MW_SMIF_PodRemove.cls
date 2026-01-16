VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MW_SMIF_PodRemove"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements ts_ECAncestor.TaskAncestor

Private Const MODULE_ID = "MW_SMIF_PodRemove"

Dim Parent As tsEC.clsRunManager
Dim TaskID As String
Dim objBatch As Batch
Dim vParameter As Variant

Private Sub TaskAncestor_Execute(Parameter As Variant)
    
    On Error GoTo ErrorHandler
    
    If IsObject(Parameter) Then
        Set vParameter = Parameter
    Else
        vParameter = Parameter
    End If
   Set objBatch = vParameter
                    
   Call Parent.SubscribeVfeiMsg(Me, "", "POD_REMOVE")
   
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
    
'Yen add for MI-4 dual function
For Each objLot In objBatch.LotList
If objVfei.Var("SMIFPORTID").Value <> objLot.Port Then Exit Sub
Next
'Yen
    
    bWait = False
    sSmifID = objVfei.Var("SMIFID").Value
    sPortID = objVfei.Var("SMIFPORTID").Value
    Set objSmif = Parent.GetPuObject.SmifList(sSmifID)
    Set objPort = Parent.GetPuObject.PortList(sPortID)

    For Each objLot In objBatch.LotList
       If objLot.Status <> LOT_REMOVE Then
          objLot.Status = LOT_REMOVE
          Parent.GetPuObject.PortList(CStr(objLot.Port)).State = PORT_EMPTY
          Call SendLotEventReport(Parent, Me, objLot)
          Call Parent.LogTrace(0, "POD_REMOVE")
          With objSmif
             .RelatedCassetteID = ""
             .PODInPlace = False
             .Locked = False
          End With
          Exit For
       End If
    Next objLot
    
    For Each objLot In objBatch.LotList
       If objLot.Status <> LOT_REMOVE Then
          bWait = True
          Exit For
       End If
    Next objLot
    
    If bWait = False Then
       Call Parent.TaskComplete(TaskID, "")
       Set objBatch = Nothing
    End If
    
    
    'isFind = False
    'sSmifID = objVfei.Var("SMIFID").Value
    'sPortID = objVfei.Var("SMIFPORTID").Value
    
    'If Err.Number = 0 Then
    '    isFind = True
    '    Set objSmif = Parent.GetPuObject.SmifList(sSmifID)
    '    Set objPort = Parent.GetPuObject.PortList(sPortID)
    '
    'Else
    '   Err.Clear
    '   TempStr = objVfei.Var("PORT").Value
    '   If Err.Number = 0 Then
    '      sPortID = TempStr
    '      Set objPort = Parent.GetPuObject.PortList(sPortID)
    '      sSmifID = objPort.RelatedSMIFID
    '      isFind = True
    '   Else
    '      isFind = False
    '       Call Parent.LogTrace(0, "Received a Event " & objVfei.CommandID & "! But can't find associate Object")
   '    End If
    ''End If
    
    
    'If isFind Then
    '   bWait = False
    ''      On Error Resume Next
    '   For Each objLot In objBatch.LotList
    '      If objLot.Cassette = objPort.CassetteID Then
    '          objLot.Chamber = "PODREMOVE"
    '          Exit For
    '      End If
    '   Next
    '
    '   For Each objLot In objBatch.LotList
    '      If objLot.Chamber <> "PODREMOVE" Then
    '         bWait = True
    '         Exit For
    '      End If
    '      With objSmif
    '         .RelatedCassetteID = ""
    '         .PODInPlace = False
    '         .Locked = False
    '      End With
    '   Next
     '  If bWait = False Then
     '     Call Parent.TaskComplete(TaskID, "")
     '     Set objBatch = Nothing
     ''  End If
    'End If

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


