VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MW_EQ_MIC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit


Implements ts_ECAncestor.TaskAncestor

Private Const MODULE_ID = "MW_EQ_MIC"

Dim Parent As tsEC.clsRunManager
Dim TaskID As String
Dim objBatch As Batch
Dim iIndex As Integer

Private Sub TaskAncestor_Execute(Parameter As Variant)
    Dim objEqDrv As EqDrv
    Dim objlot As Lot
    Dim bWait As Boolean
    On Error GoTo ErrorHandler
    
    Set objBatch = Parameter
    On Error Resume Next
    
'For Each objEqDrv In Parent.GetPuObject.EqDrvList
'Call Parent.SubscribeVfeiMsg(Me, objEqDrv.EqDrvID, "MIC")
Call Parent.SubscribeVfeiMsg(Me, "", "LOAD_CYCLE_COMPLETED")
'Next

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
    Dim objTmpLot As Lot
    Dim objTempLot As Lot
    Dim bWait As Boolean
    Dim sSmifID As String
    Dim sPortID As String
    On Error GoTo ErrorHandler

'Yen add for MI-4 dual function
For Each objTmpLot In objBatch.LotList
If objVfei.Var("SMIFPORTID").Value <> objTmpLot.Port Then Exit Sub
Next
'Yen
           
    bWait = False
    sSmifID = objVfei.Var("SMIFID").Value
    sPortID = objVfei.Var("SMIFPORTID").Value
        
    For Each objTmpLot In objBatch.LotList
       If objTmpLot.Cassette = Parent.GetPuObject.SmifList(sSmifID).RelatedCassetteID Then
          objTmpLot.Status = "MIC"
          Parent.GetPuObject.PortList(sPortID).State = PORT_PRESENT
          Call Parent.LogTrace(0, objTmpLot.LotID & "Received MIC")
          Call Parent.LogEvent("MIC", objTmpLot.LotID, "Received MIC")
          
          If chk_LotAtt_exist(objTmpLot, "LOAD_COMPLETE_TMST") Then
               Call Parent.LogTrace(0, sSmifID & " Add LOAD_COMPLETE_TMST Attribute fail,ErrorText=" & Err.Description)
          Else
             objTmpLot.Attributes.Add("LOAD_COMPLETE_TMST").Value = Format(Now, "yyyy-mm-dd hh:mm:ss.000")
             Call Parent.LogTrace(0, sSmifID & "  Add LOAD_COMPLETE_TMST Attribute Success.")
          End If
          
          
          Exit For
       End If
    Next objTmpLot
    
    For Each objTempLot In objBatch.LotList
       If objTempLot.Status <> "MIC" Then
          bWait = True
          Exit For
       End If
    Next objTempLot
    
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
Private Function chk_LotAtt_exist(objTmpLot As Lot, sLotAtt As String) As Boolean
    Dim strtmp As String
    
    On Error Resume Next
    
    strtmp = objTmpLot.Attributes.Item(sLotAtt).Value
    If Err.Number = 0 Then
        chk_LotAtt_exist = True
        Call Parent.LogTrace(0, "Lot attribute " & sLotAtt & " is exist.")
    Else
        chk_LotAtt_exist = False
    End If
    
    Err.Clear
    Exit Function
    
End Function

