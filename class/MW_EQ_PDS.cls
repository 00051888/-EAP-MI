VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MW_EQ_PDS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit


Implements ts_ECAncestor.TaskAncestor

Private Const MODULE_ID = "MW_EQ_PDS"

Dim Parent As tsEC.clsRunManager
Dim TaskID As String
Dim objBatch As Batch

Private Sub TaskAncestor_Execute(Parameter As Variant)
    Dim objEqDrv As EqDrv
    
    On Error GoTo ErrorHandler
    
    Set objBatch = Parameter
    
    For Each objEqDrv In Parent.GetPuObject.EqDrvList
       Call Parent.SubscribeVfeiMsg(Me, objEqDrv.EqDrvID, "PDS")
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
    Dim objLot As Lot
    Dim isFind As Boolean
    Dim isBatch As Boolean
    Dim sSlot As String
    Dim sSite As String
    Dim sChamber As String
    Dim sRun As String
    Dim sRecipeStep As String
    Dim objAttribute As tsBatchLot.ExtAttribute

    
    
    On Error GoTo ErrorHandler
    
'Yen add for MI-4 dual function
For Each objLot In objBatch.LotList
If objVfei.Var("CONTROLLER").Value <> objLot.LotID Then Exit Sub
Next
'Yen
    
    On Error Resume Next
    TempStr = Trim$(objVfei.Var("LOT_ID").Value)
    If Err.Number = 0 Then
       For Each objLot In objBatch.LotList
           If objLot.LotID = TempStr Then
              isFind = True
               Exit For
           End If
       Next
       
    Else
       Err.Clear
       TempStr = objVfei.Var("CASSETTE").Value
       If Err.Number = 0 Then
          For Each objLot In objBatch.LotList
              If objLot.Cassette = TempStr Then
                 isFind = True
                 Exit For
              End If
          Next
          
       Else
          Err.Clear
          TempStr = objVfei.Var("BATCH_ID").Value
          If Err.Number = 0 Then
             If objBatch.BatchID = TempStr Then
                isFind = True
                isBatch = True
             End If
          Else
             Err.Clear
             TempStr = objVfei.Var("PORT").Value
             If Err.Number = 0 Then
                For Each objLot In objBatch.LotList
                    If objLot.Port = TempStr Then
                       isFind = True
                       Exit For
                    End If
                Next
             Else
                Call Parent.LogTrace(0, "Received a Event " & objVfei.CommandID & "! But can't find associate Object")
             End If
          End If
       End If
       
    End If
    
    
    If isFind Then
       If isBatch Then
          Set objLot = objBatch.LotList(1)
       End If
       
       
       On Error Resume Next
       TempStr = objVfei.Var("QTY").Value
       If Err.Number = 0 Then
          Set objAttribute = objLot.Attributes("ProcessedCount")
          If Err.Number = 0 Then
             objAttribute.Value = TempStr
          Else
             objLot.Attributes.Add("ProcessedCount").Value = TempStr
          End If
       End If
       
        On Error Resume Next
        sSlot = objVfei.Var("SLOT").Value
        sSite = objVfei.Var("SITE").Value
        sChamber = objVfei.Var("CHAMBER").Value
        sRecipeStep = objVfei.Var("RECIPE_STEP").Value
        sRun = objVfei.Var("RUN").Value
        On Error GoTo ErrorHandler
        If sSlot = "" Then sSlot = "1"
        If sSite = "" Then sSite = "1"
        If sChamber = "" Then sChamber = "1"
        If sRecipeStep = "" Then sRecipeStep = "1"
        If sRun = "" Then sRun = "1"
        
        objVfei.Var("PDS").Var("OPERATOR_ID").Value = objLot.OperatorID
        
       If objBatch.RunMode = "PR" Then
          Call FillEqpData(objLot, objVfei, CInt(sSlot), CInt(sSite), sChamber, sRecipeStep, sRun)
          Call FillMesData(objLot, objVfei, CInt(sSlot), CInt(sSite), sChamber, sRecipeStep, sRun)
       ElseIf objBatch.RunMode = "TEST" Or objBatch.RunMode = "DMMTEST" Then
          Call FillMonitorData(objBatch, objVfei, CInt(sSlot), CInt(sSite), sChamber, sRecipeStep, sRun)
'For CFS RUN and NS Run for v0504
       ElseIf objBatch.RunMode = "CFS" Then
          Call FillEqpData(objLot, objVfei, CInt(sSlot), CInt(sSite), sChamber, sRecipeStep, sRun)
       ElseIf objBatch.RunMode = "NS" Or objBatch.RunMode = "DMMNS" Then
          Call FillEqpData(objLot, objVfei, CInt(sSlot), CInt(sSite), sChamber, sRecipeStep, sRun)
       End If
          
'       Call Parent.TaskComplete(TaskID, "")
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








          
Private Sub FillEqpData(objLot As Lot, objVfei As vfei.Message22, Optional SlotNo As Integer, Optional SiteNo As Integer, Optional sChamber As String, Optional sRecipeStep As String, Optional sRun As String)
   Dim objEqpData As EQP_DataItem
   Dim objTriplet As Triplet
   
   
   On Error Resume Next
   For Each objEqpData In objLot.EQPData
       Set objTriplet = objVfei.Var("PDS").Var(objEqpData.EQParameter)
       If Err.Number = 0 Then
          With objEqpData.DataList.Add(SlotNo, SiteNo)
            .Value = objTriplet.Value
            .Chamber = sChamber
            .RecipeStep = sRecipeStep
            .Run = sRun
          End With
          Call Parent.LogTrace(5, "Add a value into EqpData Name=" & objTriplet.Name & " Value=" & objTriplet.Value)
       Else
          Err.Clear
       End If
   
   Next
   
 

End Sub

Private Sub FillMonitorData(objBatch As Batch, objVfei As vfei.Message22, Optional SlotNo As Integer, Optional SiteNo As Integer, Optional sChamber As String, Optional sRecipeStep As String, Optional sRun As String)
   Dim objMonData As Monitor_DataItem
   Dim objTriplet As Triplet
   
   
   On Error Resume Next
   For Each objMonData In objBatch.MonitorData
       Set objTriplet = objVfei.Var("PDS").Var(objMonData.EQParameter)
       If Err.Number = 0 Then
          With objMonData.DataList.Add(SlotNo, SiteNo)
            .Value = objTriplet.Value
            .Chamber = sChamber
            .RecipeStep = sRecipeStep
            .Run = sRun
          End With
          Call Parent.LogTrace(5, "Add a value into Monitor Data Name=" & objTriplet.Name & " Value=" & objTriplet.Value)
       Else
          Err.Clear
       End If
   
   Next
   
 

End Sub

Private Sub FillMesData(objLot As Lot, objVfei As vfei.Message22, Optional SlotNo As Integer, Optional SiteNo As Integer, Optional sChamber As String, Optional sRecipeStep As String, Optional sRun As String)
   Dim objMesData As MES_DataItem
   Dim objTriplet As Triplet
   
   On Error Resume Next
   
   For Each objMesData In objLot.MESData
       Set objTriplet = objVfei.Var("PDS").Var(objMesData.EQParameter)
       If Err.Number = 0 Then
          With objMesData.DataList.Add(SlotNo, SiteNo)
            .Value = objTriplet.Value
            .Chamber = sChamber
            .RecipeStep = sRecipeStep
            .Run = sRun
          End With
          Call Parent.LogTrace(5, "Add a value into MesData Name=" & objTriplet.Name & " Value=" & objTriplet.Value)
       Else
          Err.Clear
       End If
   Next

End Sub


