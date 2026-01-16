VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MW_UI_BForm_Mi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements ts_ECAncestor.TaskAncestor

Private Const MODULE_ID = "MW_UI_BForm_Mi"

Dim parent As tsEC.clsRunManager
Dim TaskID As String

Dim objPU As ProcessingUnit
Dim objBatch As Batch
Dim LotNumber As Integer
Dim objHostTransNew As HostTransaction



Private Function CheckEqCapacity() As Boolean
    Dim nMaxBatchSize As Integer
    Dim nCurrentBatchSize As Integer
    Dim i As Integer
    
    
    If parent.GetBatchList.Count < objPU.MaxBatch Then
        CheckEqCapacity = True
    Else
        CheckEqCapacity = False
    End If
    
    
  
End Function



Private Function CreateBatch(BatchItem As HostItem) As Batch
    Dim objBatch As Batch
    Dim sCassetteID  As String
    
    Dim LotItem As HostItem
    Dim SlotItem As HostItem
    Dim AttrbuteItem As HostItem
    On Error GoTo ErrorHandler
    
    Set objBatch = New Batch
    
    objBatch.BatchID = BatchItem.Item("BATCH_ID")
    For Each LotItem In BatchItem.Item("LotList")
        sCassetteID = LotItem.Item("Cassette")
        With objBatch.LotList.Add(sCassetteID)
            .LotID = LotItem.Item("LOT_ID")
            .Qty = LotItem.Item("Qty")
            .BatchID = LotItem.Item("BATCH_ID")
            .Recipe1 = LotItem.Item("RECIPE1")
            .Recipe2 = LotItem.Item("RECIPE2")
            .ReticleSlot = LotItem.Item("RETICLE_SLOT")
            .Port = LotItem.Item("PORT")
            .Chamber = LotItem.Item("CHAMBER")
            .LoadSequence = LotItem.Item("LOADSEQ")
            .BoatPosition = LotItem.Item("BOAT_POSITION")
            .Status = LotItem.Item("STATUS")
            .LotType = LotItem.Item("LotType")
            objBatch.Operator = LotItem.Item("OPERATOR")
            .OperatorID = LotItem.Item("OPERATOR")
            objBatch.FMCassette = LotItem.Item("FMCASSETTE")
            .FMCassette = LotItem.Item("FMCASSETTE")
            objBatch.RunMode = LotItem.Item("RunMode")
            '.RunType = LotValue(BATCH_INFO_RUNMODE)
            .Slot = LotItem.Item("SLOTS")
            .PROCESS = LotItem.Item("PROCESS")
            .ProcessVersion = LotItem.Item("ProcessVersion")
            .Route = LotItem.Item("Route")
            .RouteName = LotItem.Item("RouteName")
            .Step = LotItem.Item("Step")
            .StepName = LotItem.Item("StepName")
            .Product = LotItem.Item("Product")
            .Reticle = LotItem.Item("Reticle")
            .Layer = LotItem.Item("Layer")
            .ReticleLayer = LotItem.Item("ReticleLayer")
            
            ' check slot recipe, if yes, build it
            For Each SlotItem In LotItem.Item("SlotList")
                .SlotRecipeList.Add(SlotItem.Item("Slot")) = SlotItem.Item("Recipe")
            Next
           
            '**************************************
            For Each AttrbuteItem In LotItem.Item("LotAttributeList")
                .Attributes.Add("_UI_" & AttrbuteItem.Item("AttrName")) = AttrbuteItem.Item("AttrValue")
            Next
            '**************************************
           
        End With
        
    Next
    
    ' Add by James on 2004/12/14
    For Each AttrbuteItem In BatchItem.Item("BatchAttributeList")
        objBatch.Attributes.Add("_UI_" & AttrbuteItem.Item("AttrName")) = AttrbuteItem.Item("AttrValue")
    Next
    
    Set CreateBatch = objBatch
    Exit Function
    
ErrorHandler:
    Call parent.LogTrace(0, FormatErrMsg(MODULE_ID & ".CreateBatch", Err, Err.Description))
    Set CreateBatch = Nothing
End Function

Private Sub TaskAncestor_Execute(Parameter As Variant)
    On Error GoTo ErrorHandler
    
    Set objPU = Parameter
   Call parent.SubscribeHostMsg(Me, "BATCH_INFO_FROM_ECUI")

    Exit Sub
ErrorHandler:
    Call parent.LogTrace(0, FormatErrMsg(MODULE_ID & ".Execute", Err, Err.Description))
End Sub

Private Property Set TaskAncestor_Parent(ByVal RHS As Object)
   Set parent = RHS
End Property

Private Property Get TaskAncestor_Parent() As Object
   Set TaskAncestor_Parent = parent

End Property

Private Sub TaskAncestor_ReceivedHostMsg(objHostTrans As HostOcx.HostTransaction, isPrimary As Boolean)
   
    Dim objLot As Lot
    Dim sTaskListName As String
    Dim objVfei As Message22
    Dim TempStr As String
    
    
    On Error GoTo ErrorHandler
    
    If objHostTrans.TransactionName <> "BATCH_INFO_FROM_ECUI" Then Exit Sub
    
    ' 20150608 Judy mark & add
    Set objHostTransNew = Nothing
    Set objHostTransNew = objHostTrans
    
    
    If CheckEqCapacity Then
        '---20150608 Judy mark
        'objHostTrans.Reply
               
        If CheckLots(objHostTrans.Primary) Then
            Set objBatch = CreateBatch(objHostTrans.Primary)

            If CheckRecipe Then
               '---20150608 Judy mark & add
               '          Call parent.GetBatchList.AddBatchObject(objBatch)
               '          sTaskListName = parent.GetTaskAttribute(TaskID, objBatch.RunMode)
               '          Call parent.StartTaskList(sTaskListName, objBatch)
               Call CheckBatchList
            Else
                TempStr = Trim$(parent.GetTaskAttribute(TaskID, "AUTO_REQUERY"))
                If TempStr = "" Or UCase$(TempStr) = "Y" Then
              ' Check Recipe Failure , Do a Recipe Requery
                   Set objVfei = New Message22
                   With objVfei
                      .CommandID = "QUERY_RECIPE_LIST"
                      .MachineID = objPU.EqDrvList(1).EqDrvID
                      .MessageType = "C"
                   End With
                   Call parent.SendVfei(objVfei, Me)
'--- Modify for AutomationUI 2.3.1 by hsiulin on 2000/7/5
                   For Each objLot In objBatch.LotList
                      objLot.Status = "WAIT_QUERY_RECIPES"
                      Call SendLotEventReport(parent, Me, objLot)
                   Next
'---
                Else
                   For Each objLot In objBatch.LotList
                      objLot.Status = "MISPPID"
                      Call SendLotEventReport(parent, Me, objLot)
                   Next
                End If
           End If
        Else
           Call parent.TaskComplete(TaskID, "")
           Call parent.LogTrace(0, "Wait another Lot")
        End If
    Else
       ' Check Equipment Capacity Failure
       objHostTrans.ErrorCode = 1004
       objHostTrans.ErrorText = ParseingString("Over Equipment Capacity", "~")
       objHostTrans.Reply
    End If
    
    Exit Sub
ErrorHandler:
    Call parent.LogTrace(0, FormatErrMsg(MODULE_ID & ".ReceivedHostMsg", Err, Err.Description))
End Sub

Private Function CheckRecipe() As Boolean
    Dim sRecipe As String
    Dim objRecipe As RecipeItem
    Dim i As Integer
    Dim NeedCheckRecipe As Boolean
    
    NeedCheckRecipe = parent.GetTaskAttribute(TaskID, "CHECK_RECIPE") <> "N"
    
    If Not NeedCheckRecipe Then
       ' It doesn't need to check recipe
       CheckRecipe = True
       Exit Function
    End If
    
    
    On Error Resume Next
    If objBatch.LotList(1).SlotRecipeList.Count > 0 Then
       CheckRecipe = True
       For i = 1 To objBatch.LotList(1).SlotRecipeList.Count
           sRecipe = objBatch.LotList(1).SlotRecipeList(i).Recipe
           Set objRecipe = parent.GetPuObject.RecipeList(sRecipe)
           If Err.Number <> 0 Then
              CheckRecipe = False
              Exit Function
           Else
              If Not objRecipe.Enable Then
                 CheckRecipe = False
                 Exit Function
              End If
           End If
       Next
    Else
       sRecipe = objBatch.LotList(1).Recipe1
       Set objRecipe = objPU.RecipeList(sRecipe)
       If Err.Number <> 0 Then
           CheckRecipe = False
           Exit Function
       Else
           CheckRecipe = objRecipe.Enable
       End If
    End If
    
End Function

Private Sub TaskAncestor_ReceivedVfei(objVfei As vfei.Message22)
    Dim Recipes As Variant
    Dim sTaskListName  As String
    Dim objLot As Lot
    Dim i As Integer
    
    On Error GoTo ErrorHandler
    
    If objVfei.ErrorCode = 0 Then
       For i = 1 To objPU.RecipeList.Count
           objPU.RecipeList.Remove 1
       Next
       Recipes = objVfei.Var("RECIPE_LIST").Value
       On Error Resume Next 'Pilot aviod Duplicate RecipeID , EX : APCVD
       For i = LBound(Recipes) To UBound(Recipes)
           objPU.RecipeList.Add(CStr(Recipes(i))).Enable = True
       Next
       On Error GoTo ErrorHandler   'Pilot
       If CheckRecipe Then
            '---20150608 Judy mark & add
            ' Call parent.GetBatchList.AddBatchObject(objBatch)
            ' sTaskListName = parent.GetTaskAttribute(TaskID, objBatch.RunMode)
            ' Call parent.StartTaskList(sTaskListName, objBatch)
              
              Call CheckBatchList
       
       Else
          For Each objLot In objBatch.LotList
              objLot.Status = "MISPPID"
              Call SendLotEventReport(parent, Me, objLot)
          Next
       End If
       
       
    Else
       Call parent.LogTrace(0, "Query Eq Recipe Failure. " & objVfei.ErrorText)
       For Each objLot In objBatch.LotList
           objLot.Status = "MISPPID"
           Call SendLotEventReport(parent, Me, objLot)
       Next
    End If
    
    
    Exit Sub
ErrorHandler:
    Call parent.LogTrace(0, FormatErrMsg(MODULE_ID & ".ReceivedVfei", Err, Err.Description))
End Sub

Private Property Let TaskAncestor_TaskID(ByVal RHS As String)
    TaskID = RHS
End Property

Private Property Get TaskAncestor_TaskID() As String
    TaskAncestor_TaskID = TaskID
End Property

Private Function CheckLots(BatchItem As HostItem) As Boolean
 '  Dim i As Integer
   Dim sCassetteID As String
   Dim objSmif As SmifItem
   
   On Error Resume Next
   Select Case BatchItem.Item("LotList").Count
      Case 1
          CheckLots = True
      Case 2
          For Each objSmif In objPU.SmifList
             'For i = 1 To 2
                sCassetteID = BatchItem.Item("LotList").Item(1).Item("Cassette").Value
                If sCassetteID = objSmif.RelatedCassetteID Then
                    BatchItem.Item("LotList").Item(1).Item("PORT").Value = objSmif.RelatedPortID
                    If objSmif.RelatedPortID = 1 Then
                       BatchItem.Item("LotList").Item(2).Item("PORT").Value = 2
                       CheckLots = True
                    ElseIf objSmif.RelatedPortID = 2 Then
                       BatchItem.Item("LotList").Item(2).Item("PORT").Value = 1
                       CheckLots = True
                    End If
                    Exit For
                End If
             'Next i
             'Exit For
          Next objSmif
          
         ' If objSmif.RelatedPortID = 1 Then
         '    BatchItem.Item("LotList").Item(2).Item("PORT").Value = 2
         '    CheckLots = True
         ' ElseIf objSmif.RelatedPortID = 2 Then
         '    BatchItem.Item("LotList").Item(2).Item("PORT").Value = 1
         '    CheckLots = True
         ' End If
          'ElseIf i = 2 And objSmif.RelatedPortID = 1 Then
          '   BatchItem.Item("LotList").Item(1).Item("PORT").Value = 2
          '   CheckLots = True
          'ElseIf i = 2 And objSmif.RelatedPortID = 2 Then
          '   BatchItem.Item("LotList").Item(1).Item("PORT").Value = 1
          '   CheckLots = True
          'End If
          
   End Select
End Function
Private Function CheckBatchList() As Boolean
         Dim oRM As tsEC.clsRunManager
         Dim oMainPUAttr As ts_PU.ExtAttribute
         Dim colPU As clsRunManager
         Dim objRmBatch As Batch
         Dim objRmLot As Lot
         Dim objLot As Lot
         Dim iLotListCount As Integer
         Dim blnFindTheSameCassette, OutOfPortMaxLot As Boolean
         Dim CfgPortMaxLotCount As String
         Dim sTaskListName  As String

12       CfgPortMaxLotCount = GetIniSetting(parent.Path & "\Rm.ini", "PortMaxLot", "MaxLotCount")
13       CfgPortMaxLotCount = IIf(CfgPortMaxLotCount = "", "1", CfgPortMaxLotCount)

15       Set oRM = parent.GetRMList.Item(parent.GetRMList.Count)
    
17       If Not oRM Is Nothing Then
18           On Error Resume Next
19           Set oMainPUAttr = oRM.GetPuObject.Attributes.Item("RmList")
20           If Err Then Err.Clear
21           If Not (oMainPUAttr Is Nothing) Then ' for Muiti PU
22               For Each objLot In objBatch.LotList 'Batch_info from UI
23                   For Each colPU In oRM.GetPuObject.Attributes.Item("RmList").Value
24                       For Each objRmBatch In colPU.GetBatchList
25                           For Each objRmLot In objRmBatch.LotList 'Current lot in EC
26                               If objRmLot.Port = objLot.Port Then 'Count the same Port
27                                   iLotListCount = iLotListCount + 1
28                               End If
29                               If objRmLot.Cassette = objLot.Cassette Then 'Check the same Cassette
30                                   blnFindTheSameCassette = True
31                               End If
32                           Next
33                       Next
34                   Next
35                   If iLotListCount >= CInt(CfgPortMaxLotCount) Then 'Check the same Port
36                       OutOfPortMaxLot = True
37                   End If
38                   iLotListCount = 0
39               Next
40           Else 'for Single PU
41               For Each objLot In objBatch.LotList 'Batch_info from UI
42                   For Each objRmBatch In parent.GetBatchList
43                       For Each objRmLot In objRmBatch.LotList 'Current lot in EC
44                           If objRmLot.Port = objLot.Port Then 'Count the same Port
45                               iLotListCount = iLotListCount + 1
46                           End If
47                           If objRmLot.Cassette = objLot.Cassette Then 'Check the same Cassette
48                               blnFindTheSameCassette = True
49                           End If
50                       Next
51                   Next
52                   If iLotListCount >= CInt(CfgPortMaxLotCount) Then 'Check the same Port
53                       OutOfPortMaxLot = True
54                   End If
55                   iLotListCount = 0
56               Next
57           End If
        
59           If Not (blnFindTheSameCassette = True Or OutOfPortMaxLot = True) Then
60               objHostTransNew.Reply
61               Debug.Print Format(Time(), "HH:MM:SS") & "  " & parent.GetPuObject.PUID & " Start Task"
        
63               Call parent.GetBatchList.AddBatchObject(objBatch)
64               sTaskListName = parent.GetTaskAttribute(TaskID, objBatch.RunMode)
65               Call parent.StartTaskList(sTaskListName, objBatch)
            
67           ElseIf blnFindTheSameCassette = True Then
68               For Each objLot In objBatch.LotList
69                   objLot.Status = Translate(parent, "重覆進站", "Repeated MoveIn")
70                   Call SendLotEventReport(parent, Me, objLot)
71               Next
72               objHostTransNew.ErrorCode = 1004
73               objHostTransNew.ErrorText = Translate(parent, "重覆進站", ParseingString("This LOT has been processed!", "~"))
74               objHostTransNew.Reply
75               Call parent.LogTrace("0", "Cassette " & objBatch.LotList(1).LotID & "重覆進站")
76           ElseIf OutOfPortMaxLot = True Then
77               For Each objLot In objBatch.LotList
78                   objLot.Status = Translate(parent, "超出進站上限", "Over Equipment Capacity")
79                   Call SendLotEventReport(parent, Me, objLot)
80               Next
81               objHostTransNew.ErrorCode = 1004
82               objHostTransNew.ErrorText = Translate(parent, "該PORT已達材料進貨上限,不允許進站!", ParseingString("The Port is Over Equipment Capacity to process, not allowed to move in.", "~"))
83               objHostTransNew.Reply
84               Call parent.LogTrace("0", "Cassette " & objBatch.LotList(1).LotID & "超出進站上限")
85           End If
        
87       End If
    
89       Exit Function
90 ErrorHandler:

End Function

