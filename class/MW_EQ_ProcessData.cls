VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MW_EQ_ProcessData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements ts_ECAncestor.TaskAncestor

Private Const MODULE_ID = "MW_EQ_ProcessData"


Dim pBatchProcessData As BatchProcessData

Dim Parent As tsEC.clsRunManager
Dim TaskID As String
Dim objBatch As Batch
Dim objLot As Lot
Dim colWTW As Collection
Dim aTID_AUTOS() As String
Dim aTID_AUTOR() As String
Dim aEventTime_AUTOS() As String
Dim aEventTime_AUTOR() As String
Dim aWtwEventName_AUTOS() As String
Dim aWtwEventName_AUTOR() As String

Private Sub TaskAncestor_Execute(Parameter As Variant)
    Dim objEqDrv As EqDrv
    Dim vLotID As String
    
    Dim vCassetteID As String
    Dim vRecipe As String
    Dim vPortID As String

    On Error GoTo ErrorHandler
    Set objBatch = Parameter
    
    Set pBatchProcessData = New BatchProcessData
    pBatchProcessData.BatchID = objBatch.BatchID
    
    Set objLot = objBatch.LotList(1)
    
    ReDim Preserve aTID_AUTOS(1) As String
    ReDim Preserve aTID_AUTOR(1) As String
    ReDim Preserve aEventTime_AUTOS(1) As String
    ReDim Preserve aEventTime_AUTOR(1) As String
    ReDim Preserve aWtwEventName_AUTOS(1) As String
    ReDim Preserve aWtwEventName_AUTOR(1) As String
    
    vLotID = Trim(objLot.LotID)
    vCassetteID = Trim(objLot.Cassette)
    vRecipe = Trim(objLot.Recipe1)
    vPortID = Trim(objLot.Port)
    pBatchProcessData.AddLotList vLotID, vCassetteID, vRecipe, vPortID
    
    Call Parent.SubscribeVfeiMsg(Me, , "PS") 'CEID 11 ; by Lot "PS"
    Call Parent.SubscribeVfeiMsg(Me, , "PE") 'CEID 12 ; by Lot "PE"
    Call Parent.SubscribeVfeiMsg(Me, , "Beamscan_started") 'CEID 5, by Lot "AUTOBEAM_STARTUP"
    Call Parent.SubscribeVfeiMsg(Me, , "Beamscan_Complete") 'CEID 10, by Lot "AUTOBEAM_READY"

    Call Parent.SubscribeVfeiMsg(Me, , "Wafer_started") 'CEID 13 ; by Wafer "IMPLANT_W_STARTUP"
    Call Parent.SubscribeVfeiMsg(Me, , "Wafer_complete") 'CEID 14 ; by Wafer "IMPLANT_W_COMPLETE"
    Call Parent.SubscribeVfeiMsg(Me, , "MIC") 'CEID 61
    Call Parent.SubscribeVfeiMsg(Me, , "MOR") 'CEID 65
    
    Set colWTW = New Collection
    colWTW.Add "AUTOBEAM_STARTUP", "Beamscan_started"
    colWTW.Add "AUTOBEAM_READY", "Beamscan_Complete"
    colWTW.Add "IMPLANT_W_STARTUP", "Wafer_started"
    colWTW.Add "IMPLANT_W_COMPLETE", "Wafer_complete"
    
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
         Dim vTID As String
         Dim vEventName As String
         Dim vWtwEventName As String
         Dim vEventTime As String
         Dim vLotID As String
         Dim vPortID As String
         Dim vSlotID As String
         Dim iSlotID As Integer
         Dim tempEventTime As String
         Dim tmpLotdata As LotListData
         Dim objLotData As LotListData

13       On Error GoTo ErrorHandler
    
         ' Record EventName
16       vEventName = objVfei.CommandID
17       If CheckColItemIsExist(vEventName) = True Then
18           vWtwEventName = colWTW.Item(vEventName)
19       Else
20           vWtwEventName = vEventName
21       End If
22       vEventTime = Format(Now, "yyyy-mm-dd hh:mm:ss")
23       vTID = CStr(objVfei.TransactionID)

25       Select Case vWtwEventName
             Case "MIC", "PE"
27               vLotID = Trim(objVfei.Var("CONTROLLER").Value)
28               Set objLotData = pBatchProcessData.LotListData(1)
29               If vLotID = objLotData.LotID Then
30                   objLotData.AddLotEvent vTID, vWtwEventName, vEventTime
31               End If
32               Set objLotData = Nothing
33           Case "PS" '收到PS檢查Buffer內的Autobeam Event,成對就送出,然後清空Buffer
34               vLotID = Trim(objVfei.Var("CONTROLLER").Value)
35               Set objLotData = pBatchProcessData.LotListData(1)
36               If vLotID = objLotData.LotID Then
37                   objLotData.AddLotEvent vTID, vWtwEventName, vEventTime
38                   If aTID_AUTOS(1) <> "" And aTID_AUTOR(1) <> "" Then
39                       objLotData.AddLotEvent aTID_AUTOS(1), aWtwEventName_AUTOS(1), aEventTime_AUTOS(1)
40                       objLotData.AddLotEvent aTID_AUTOR(1), aWtwEventName_AUTOR(1), aEventTime_AUTOR(1)
41                   End If
42                   ClearBuffer
43               End If
44               Set objLotData = Nothing
45           Case "AUTOBEAM_STARTUP"
46               vLotID = Trim(objVfei.Var("CONTROLLER").Value)
47               Set objLotData = pBatchProcessData.LotListData(1)
48               If vLotID = objLotData.LotID Then
                     '放到buffer內
50                   Call Handle_AUTOS(vTID, vEventTime, vWtwEventName)
52               End If
53               Set objLotData = Nothing
            
55           Case "AUTOBEAM_READY"
56               vLotID = Trim(objVfei.Var("CONTROLLER").Value)
57               Set objLotData = pBatchProcessData.LotListData(1)
58               If vLotID = objLotData.LotID Then
                     '放到buffer內
60                   Call Handle_AUTOR(vTID, vEventTime, vWtwEventName)
62               End If
63               Set objLotData = Nothing
            
65           Case "IMPLANT_W_STARTUP"
66               vLotID = Trim(objVfei.Var("CONTROLLER").Value)
67               vSlotID = Trim(objVfei.Var("SLOT").Value)
68               Set objLotData = pBatchProcessData.LotListData(1)
69               If vLotID = objLotData.LotID Then
70                   If CInt(vSlotID) > 25 Then
71                       iSlotID = CInt(vSlotID) - 25
72                       objLotData.AddWaferEvent vTID, CStr(iSlotID), "1", objLotData.Recipe, vWtwEventName, vEventTime
73                   Else
74                       objLotData.AddWaferEvent vTID, vSlotID, "1", objLotData.Recipe, vWtwEventName, vEventTime
75                   End If
                
77               End If
78               Set objLotData = Nothing
            
80           Case "IMPLANT_W_COMPLETE"
81               vLotID = Trim(objVfei.Var("CONTROLLER").Value)
82               vSlotID = Trim(objVfei.Var("SLOT").Value)
83               Set objLotData = pBatchProcessData.LotListData(1)
84               If vLotID = objLotData.LotID Then
85                   If CInt(vSlotID) > 25 Then
86                       iSlotID = CInt(vSlotID) - 25
87                       objLotData.AddWaferEvent vTID, CStr(iSlotID), "1", objLotData.Recipe, vWtwEventName, vEventTime
88                   Else
89                       objLotData.AddWaferEvent vTID, vSlotID, "1", objLotData.Recipe, vWtwEventName, vEventTime
90                   End If
91                   If aTID_AUTOS(1) <> "" And aTID_AUTOR(1) <> "" Then
92                       objLotData.AddLotEvent aTID_AUTOS(1), aWtwEventName_AUTOS(1), aEventTime_AUTOS(1)
930                      objLotData.AddLotEvent aTID_AUTOR(1), aWtwEventName_AUTOR(1), aEventTime_AUTOR(1)
94                   End If
95                   ClearBuffer
96               End If
97               Set objLotData = Nothing
            
99           Case "MOR"
100              vLotID = Trim(objVfei.Var("CONTROLLER").Value)
101              Set objLotData = pBatchProcessData.LotListData(1)
102              If vLotID = objLotData.LotID Then
103                  objLotData.AddLotEvent vTID, vWtwEventName, vEventTime
104                  Set tmpLotdata = objLotData
105              End If
106              Set objLotData = Nothing
        
108      End Select
    
110      If vWtwEventName = "MOR" Then
111          If Not tmpLotdata Is Nothing Then
112              Call LogData(tmpLotdata)
113              Call SendLotEventData(tmpLotdata, vEventTime)
114              Call SendWaferEventData(tmpLotdata, vEventTime)
115          End If
116      End If
    
118      Exit Sub

120 ErrorHandler:
121      Call Parent.LogTrace(0, FormatErrMsg(MODULE_ID & ".ReceivedVfei", Err, "[" & Erl & "]" & Err.Description))
End Sub
Private Sub Handle_AUTOS(vTID As String, vEventTime As String, vWtwEventName As String)
         Dim kIndex As Integer
    
4        On Error GoTo ErrorHandler

6        Call Parent.LogTrace(0, "Handle_AUTOS")
7        kIndex = 1
8        aTID_AUTOS(kIndex) = vTID
9        aEventTime_AUTOS(kIndex) = vEventTime
10       aWtwEventName_AUTOS(kIndex) = vWtwEventName
11       Exit Sub
12 ErrorHandler:
13       Call Parent.LogTrace(0, FormatErrMsg(MODULE_ID & ".Handle_AUTOS", Err, "[" & Erl & "]" & Err.Description))
End Sub
Private Sub Handle_AUTOR(vTID As String, vEventTime As String, vWtwEventName As String)
         Dim kIndex As Integer

4        On Error GoTo ErrorHandler

6        Call Parent.LogTrace(0, "Handle_AUTOR")
7        kIndex = 1
8        aTID_AUTOR(kIndex) = vTID
9        aEventTime_AUTOR(kIndex) = vEventTime
10       aWtwEventName_AUTOR(kIndex) = vWtwEventName
11       Exit Sub
12 ErrorHandler:
13       Call Parent.LogTrace(0, FormatErrMsg(MODULE_ID & ".Handle_AUTOR", Err, "[" & Erl & "]" & Err.Description))
End Sub
Private Sub ClearBuffer()
    Dim j As Integer

    For j = 1 To 1
        aTID_AUTOS(j) = ""
        aEventTime_AUTOS(j) = ""
        aWtwEventName_AUTOS(j) = ""
        aTID_AUTOR(j) = ""
        aEventTime_AUTOR(j) = ""
        aWtwEventName_AUTOR(j) = ""
    Next
 
End Sub

Private Sub LogData(vLotData As LotListData)
    Dim i As Integer
    Dim vLotEventCol As Collection
    Dim vWaferEventCol As Collection
    Dim vLotEvent As LotEventData
    Dim vWaferEvent As SlotChamberData

    On Error GoTo ErrorHandler
   
    Set vLotEventCol = vLotData.GetLotEventCol
    Set vWaferEventCol = vLotData.GetWaferEventCol
    Parent.LogTrace 0, "************************************************************"
    Parent.LogTrace 0, "LOT EVENT"
    Parent.LogTrace 0, "PU,TID,BatchID,LotID,LotRecipe,EventName,EventTMST"
    Parent.LogTrace 0, "************************************************************"
    For i = 1 To vLotEventCol.Count
        Set vLotEvent = vLotEventCol.Item(i)
        Parent.LogTrace 0, Parent.PUID & "," & vLotEvent.TID & "," & pBatchProcessData.BatchID & "," & vLotData.LotID & "," & vLotData.Recipe & "," & vLotEvent.EventName & "," & vLotEvent.EventTMST
    Next i
    Parent.LogTrace 0, "************************************************************"
    Parent.LogTrace 0, "WAFER EVENT"
    Parent.LogTrace 0, "PU,TID,BatchID,LotID,LotRecipe,SlotID,Chamber,Recipe,EventName,EventTMST"
    Parent.LogTrace 0, "************************************************************"
    For i = 1 To vWaferEventCol.Count
        Set vWaferEvent = vWaferEventCol.Item(i)
        Parent.LogTrace 0, Parent.PUID & "," & vWaferEvent.TID & "," & pBatchProcessData.BatchID & "," & vLotData.LotID & "," & vLotData.Recipe & "," & vWaferEvent.sLotID & "," & vWaferEvent.Chamber & "," & vWaferEvent.Recipe & "," & vWaferEvent.EventName & "," & vWaferEvent.EventTMST
    Next i
    Parent.LogTrace 0, "************************************************************"
    Parent.LogTrace 0, "END"
    Parent.LogTrace 0, "************************************************************"
        
    
    
    Exit Sub

ErrorHandler:
    Call Parent.LogTrace(0, FormatErrMsg(MODULE_ID & ".LogData", Err, Err.Description))
End Sub

'[Begin]
'Destination=AS
'Primary=WTW_LOT_EVENT <%Subject> <%TID> <Lot> <Batch> <PU> <Port> <LotRecipe> <UpdateTime> <#EventList>
'    [ <Event> <EventTime> <TID> ]
'Secondary=//WTW_LOT_EVENT <%Status> <%TID> <Lot>
'End
Private Sub SendLotEventData(vLotData As LotListData, iTime As String)
         Dim i As Integer
         Dim vLotEventCol As Collection
         Dim vWaferEventCol As Collection
         Dim vLotEvent As LotEventData
         Dim objHostTrans As HostTransaction

7        On Error GoTo ErrorHandler
    

10       Set vLotEventCol = vLotData.GetLotEventCol
11       Set objHostTrans = Parent.NewHostTransaction(Me, "WTW_LOT_EVENT")
12       With objHostTrans.Primary
13           .Item("Lot") = vLotData.LotID
14           .Item("Batch") = pBatchProcessData.BatchID
15           .Item("PU") = Parent.PUID
16           .Item("Port") = vLotData.PortID
17           .Item("LotRecipe") = vLotData.Recipe
18           .Item("UpdateTime") = iTime
19           For i = 1 To vLotEventCol.Count
20               With .Item("EventList").Add
21                   Set vLotEvent = vLotEventCol.Item(i)
22                   .Item("Event") = vLotEvent.EventName
23                   .Item("EventTime") = vLotEvent.EventTMST
24                   .Item("TID") = vLotEvent.TID
25               End With
26           Next i
27       End With
28       objHostTrans.Send
29       Call Parent.LogTrace(0, "Upload WTW Lot(" & vLotData.LotID & ") Event Data to AS")
        
31       Set objHostTrans = Nothing

33       Exit Sub

35 ErrorHandler:
36       Call Parent.LogTrace(0, FormatErrMsg(MODULE_ID & ".SendLotEventData", Err, "[" & Erl & "]" & Err.Description))
End Sub
Private Function CheckColItemIsExist(vKey As String) As Boolean
1        On Error Resume Next
         Dim stemp As String
3        CheckColItemIsExist = False
4        stemp = colWTW.Item(vKey)
5        If Err.Number = 0 Then CheckColItemIsExist = True
6        Err.Clear
End Function

'[Begin]
'Destination=AS
'Primary=WTW_WAFER_EVENT <%Subject> <%TID> <Lot> <Batch> <PU> <Port> <LotRecipe> <UpdateTime> <#EventList>
'    [<Slot> <chamber> <Recipe> <Event> <EventTime> <TID>]
'Secondary=//WTW_WAFER_EVENT <%Status> <%TID> <Lot>
'End
Private Sub SendWaferEventData(vLotData As LotListData, iTime As String)
         Dim i As Integer
         Dim vWaferEventCol As Collection
         Dim vWaferEvent As SlotChamberData
         Dim objHostTrans As HostTransaction


7        On Error GoTo ErrorHandler
 

10       Set vWaferEventCol = vLotData.GetWaferEventCol
11       Set objHostTrans = Parent.NewHostTransaction(Me, "WTW_WAFER_EVENT")
12       With objHostTrans.Primary
13           .Item("Lot") = vLotData.LotID
14           .Item("Batch") = pBatchProcessData.BatchID
15           .Item("PU") = Parent.PUID
16           .Item("Port") = vLotData.PortID
17           .Item("LotRecipe") = vLotData.Recipe
18           .Item("UpdateTime") = iTime
19           For i = 1 To vWaferEventCol.Count
20               With .Item("EventList").Add
21                   Set vWaferEvent = vWaferEventCol.Item(i)
22                   .Item("Slot") = vWaferEvent.sLotID
23                   .Item("chamber") = vWaferEvent.Chamber
24                   .Item("Recipe") = vWaferEvent.Recipe
25                   .Item("Event") = vWaferEvent.EventName
26                   .Item("EventTime") = vWaferEvent.EventTMST
27                   .Item("TID") = vWaferEvent.TID
28               End With
29           Next i
30       End With
31       objHostTrans.Send
32       Call Parent.LogTrace(0, "Upload WTW Wafer(" & vLotData.LotID & ") Event Data to AS")
        
34       Set objHostTrans = Nothing
35       Exit Sub

37 ErrorHandler:
38       Call Parent.LogTrace(0, FormatErrMsg(MODULE_ID & ".SendWaferEventData", Err, "[" & Erl & "]" & Err.Description))
End Sub


Private Property Let TaskAncestor_TaskID(ByVal RHS As String)
    TaskID = RHS
End Property

Private Property Get TaskAncestor_TaskID() As String
    TaskAncestor_TaskID = TaskID
End Property

