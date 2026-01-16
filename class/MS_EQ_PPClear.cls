VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MS_EQ_PPCLEAR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements ts_ECAncestor.TaskAncestor

Private Const MODULE_ID = "MS_EQ_PPCLEAR"

Dim Parent As tsEC.clsRunManager
Dim TaskID As String
Dim objPU As ProcessingUnit
Dim objBatch As Batch


Private Sub TaskAncestor_Execute(Parameter As Variant)
    Dim objVfei As Message22
    Dim objLot As Lot
    Dim objPort As PortItem
    Dim i As Integer
    Dim Msg As String
    
    On Error GoTo ErrorHandler

    Set objBatch = Parameter
    Set objPU = Parent.GetPuObject

  '  For i = 2 To Parent.GetPuObject.Attributes.Count
  '     Parent.GetPuObject.Attributes.Remove (2)
  '  Next i
  
  
'    Set objVfei = New Message22
'    objVfei.CommandID = "REMOTE_COMMAND"
'    objVfei.MachineID = objPU.EqDrvList(1).EqDrvID
'    If objBatch.LotList.count = 2 Then
'        objVfei.AddVar "COMMAND", "A", "PPCLEAR-DUAL"
'    Else
'        For Each objLot In objBatch.LotList
'           If objLot.Port = 1 Then
'              objVfei.AddVar "COMMAND", "A", "PPCLEAR-LEFT"
'           Else
'              objVfei.AddVar "COMMAND", "A", "PPCLEAR-RIGHT"
'           End If
'        Next
'    End If
    
        For Each objLot In objBatch.LotList
           If objLot.Port = 1 Then
                Set objVfei = New Message22
                objVfei.CommandID = "REMOTE_COMMAND"
                objVfei.MachineID = objPU.EqDrvList(1).EqDrvID
                objVfei.AddVar "COMMAND", "A", "PPCLEAR-LEFT"
     
                Call Parent.SendVfei(objVfei, Me)
          Else
                Set objVfei = New Message22
                objVfei.CommandID = "REMOTE_COMMAND"
                objVfei.MachineID = objPU.EqDrvList(1).EqDrvID
                objVfei.AddVar "COMMAND", "A", "PPCLEAR-RIGHT"
    
                Call Parent.SendVfei(objVfei, Me)
           End If
        Next
    
  
'    Set objVfei = New Message22
'    objVfei.CommandID = "REMOTE_COMMAND"
'    objVfei.MachineID = objPU.EqDrvList(1).EqDrvID
'    objVfei.AddVar "COMMAND", "A", "PPSELECT"
'    If objBatch.LotList.count = 2 Then
'        If objBatch.LotList(1).Recipe1 = objBatch.LotList(2).Recipe1 Then
'            objVfei.AddVar "PPID", "A", objBatch.LotList(1).Recipe1
'            objVfei.AddVar "LOC", "B", 3
'            For Each objPort In Parent.GetPuObject.PortList
'               If objPort.PortID = 1 Then
'                   For Each objLot In objBatch.LotList
'                      If objPort.CassetteID = objLot.Cassette Then
'                         objVfei.AddVar "MID1", "A", objLot.LotID
'                      Else
'                         objVfei.AddVar "MID2", "A", objLot.LotID
'                      End If
'                   Next objLot
'               Else
'                   For Each objLot In objBatch.LotList
'                      If objPort.CassetteID = objLot.Cassette Then
'                         objVfei.AddVar "MID2", "A", objLot.LotID
'                      Else
'                         objVfei.AddVar "MID1", "A", objLot.LotID
'                      End If
'                   Next objLot
'               End If
'
'               Msg = objVfei.Build
'               Dim objTmpVfei As New Message22
'               objTmpVfei.Parse Msg
'
'               objVfei.Clear
'               With objVfei
'                  .CommandID = objTmpVfei.CommandID
'                  .MachineID = objTmpVfei.MachineID
'                  .AddVar "COMMAND", "A", "PPSELECT"
'                  .AddVar "PPID", "A", objTmpVfei.Var("PPID").Value
'                  .AddVar "LOC", "B", objTmpVfei.Var("LOC").Value
'                  .AddVar "MID1", "A", objTmpVfei.Var("MID1").Value
'                  .AddVar "MID2", "A", objTmpVfei.Var("MID2").Value
'               End With
'               Exit For
'           Next objPort
'
'        Else
'            Call Parent.LogTrace(0, MODULE_ID & "=" & "Lot_ID:" & GetLotListStr & " have differnt recipe name")
'            Call Parent.LogEvent("Recipe Error", GetLotListStr, "Port:" & GetPortListStr)
'            Call AlarmReport(Parent, Me, "101", "程式錯誤", GetLotListStr & " 此 Batch 的兩個 LOT 的程式名稱不相同")
'            Exit Sub
'        End If
'    Else
'        For Each objLot In objBatch.LotList
'           For Each objPort In Parent.GetPuObject.PortList
'              If objPort.CassetteID = objLot.Cassette Then
'                  objVfei.AddVar "PPID", "A", objLot.Recipe1
'                  objVfei.AddVar "LOC", "B", objPort.PortID
'                  If objPort.PortID = 1 Then
'                     objVfei.AddVar "MID1", "A", objLot.LotID
'                     objVfei.AddVar "MID2", "A", " "
'                  Else
'                     objVfei.AddVar "MID1", "A", " "
'                     objVfei.AddVar "MID2", "A", objLot.LotID
'                  End If
'                  Exit For
'              End If
'           Next objPort
'           Exit For
'        Next objLot
'    End If
    
'    Debug.Print objVfei.Var("Command").Value
    
'    Call Parent.SendVfei(objVfei, Me)
    
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
    Dim objLot As Lot
    Dim tmpRecipe As String
    On Error GoTo ErrorHandler
    
    
    If objVfei.CommandID <> "REMOTE_COMMAND" Then Exit Sub
    
    If objVfei.ErrorCode <> 0 Then
   
           Call Parent.LogEvent("PPCLEARERR", GetLotListStr, " PortID:" & GetPortListStr)
           Call Parent.LogTrace(5, "Recipe Select Failure.")
           Call AlarmReport(Parent, Me, "12", Translate(Parent, "程式清除失敗: ", "Clean Recipe failure. "), objVfei.ErrorText)
    
    Else
                   
           Call Parent.LogEvent("PPCLEAR", GetLotListStr, " PortID:" & GetPortListStr)
           Call Parent.LogTrace(5, "Recipe Clear Successful.")
              
           Call Parent.TaskComplete(TaskID, "")
    End If
 
    Set objBatch = Nothing
    Set objPU = Nothing
    
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


