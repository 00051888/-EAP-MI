VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MW_LMC_Event"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements ts_ECAncestor.TaskAncestor

Private Const MODULE_ID = "MW_LMC_Event"

Dim Parent As tsEC.clsRunManager
Dim TaskID As String
Dim objBatch As Batch
Dim objPort As PortItem
Dim objEqDrv As EqDrv
Dim iLMCEventCount As Integer
Dim sLMCEvent() As String

Private Sub TaskAncestor_Execute(Parameter As Variant)
         Dim objEqDrv As EqDrv
         Dim sLMC_EventName As String
         Dim sLMC_VefiMsg As String
         Dim tempLMC_Criteria As String
         Dim iLMC_CriteriaCount As Integer
         Dim sLMC_CriteriaList As String
         Dim sLMC_Criteria As String
         Dim sTestLMCEvent As String
         Dim objLot As Lot
         Dim i, j As Integer
    
12       On Error GoTo ErrorHandler
    
14       Set objBatch = Parameter

16       iLMC_CriteriaCount = 0
17       iLMCEventCount = Parent.GetTaskAttribute(TaskID, "LMC_Count")
         
19       ReDim sLMCEvent(iLMCEventCount) As String
         
21       For i = 1 To iLMCEventCount
22           sLMCEvent(i) = Parent.GetTaskAttribute(TaskID, "LMCEvent" & i)
23           If Err.Number = 0 And sLMCEvent(i) <> "" Then
24               sLMC_EventName = GetToken2(sLMCEvent(i), ",", 1)
25               If Err.Number <> 0 Or sLMC_EventName = "" Then
26                   Err.Clear
27                   Call Parent.LogTrace(0, "[MW_LMC_EVENT] GetTaskAttribute Failed, Not Set =>LMC EventName")
28               End If
29               sLMC_VefiMsg = GetToken2(sLMCEvent(i), ",", 2)
30               If Err.Number <> 0 Or sLMC_VefiMsg = "" Then
31                   Err.Clear
32                   Call Parent.LogTrace(0, "[MW_LMC_EVENT] GetTaskAttribute Failed, Not Set => Subscribe VFEI Name")
33               End If
34               tempLMC_Criteria = GetToken2(sLMCEvent(i), ",", 3)
35               If Err.Number <> 0 Or tempLMC_Criteria = "" Then
36                   Err.Clear
37                   Call Parent.LogTrace(0, "[MW_LMC_EVENT] GetTaskAttribute Failed, Not Set => Subscribe VFEI Criteria")
38               End If
39               sLMC_CriteriaList = Set_LMC_VEFI_Criteria(UCase(tempLMC_Criteria))
40               If sLMC_VefiMsg <> "N" Then
41                   If sLMC_CriteriaList = "" Then
42                       Call Parent.SubscribeVfeiMsg(Me, objEqDrv.EqDrvID, sLMC_VefiMsg)
43                   ElseIf sLMC_CriteriaList = "N/A" Then
44                       Call Parent.LogTrace(0, "VEFI Criteria not on SPEC = " & sLMC_CriteriaList)
45                   Else
46                       iLMC_CriteriaCount = GetTokenCount2(sLMC_CriteriaList, ",")
47                       For Each objEqDrv In Parent.GetPuObject.EqDrvList
48                           For j = 1 To iLMC_CriteriaCount
49                               sLMC_Criteria = GetToken2(sLMC_CriteriaList, ",", j)
50                               Call Parent.SubscribeVfeiMsg(Me, objEqDrv.EqDrvID, sLMC_VefiMsg, tempLMC_Criteria, sLMC_Criteria)
51                           Next j
52                       Next
53                   End If
54               End If
55           Else
56               Err.Clear
57               Call Parent.LogTrace(0, "[MW_LMC_EVENT] GetTaskAttribute Failed, Not Set =>" & "LMCEvent" & i)
58           End If
59       Next i

61       sTestLMCEvent = Parent.GetTaskAttribute(TaskID, "LMCEvent" & iLMCEventCount + 1)
62       If sTestLMCEvent <> "" Then
63           Call Parent.LogTrace(0, "[MW_LMC_EVENT] GetTaskAttribute Failed, Count Error =>" & "LMCEvent" & iLMCEventCount + 1 & "," & sTestLMCEvent)
64       End If
        
66       Call Parent.LogTrace(0, "[MW_LMC_EVENT] Execute to SubscribeVfeiMsg is OK.")
    
68       Exit Sub
69 ErrorHandler:
70       Call Parent.LogTrace(0, FormatErrMsg(MODULE_ID & ".Execute", Err, "[" & Erl & "]" & Err.Description))
End Sub

Private Property Set TaskAncestor_Parent(ByVal RHS As Object)
1        Set Parent = RHS
End Property

Private Property Get TaskAncestor_Parent() As Object
1        Set TaskAncestor_Parent = Parent
End Property

Private Sub TaskAncestor_ReceivedHostMsg(objHostTrans As HostOcx.HostTransaction, isPrimary As Boolean)
1        On Error GoTo ErrorHandler

3        Exit Sub
4 ErrorHandler:
5        Call Parent.LogTrace(0, FormatErrMsg(MODULE_ID & ".ReceivedHostMsg", Err, "[" & Erl & "]" & Err.Description))
End Sub

Private Sub TaskAncestor_ReceivedVfei(objVfei As vfei.Message22)
         Dim objLot As Lot
         Dim i As Integer
         
         Dim sLMC_EventName As String
         Dim sLMC_VefiMsg As String
         Dim tempLMC_Criteria As String
         Dim sVefi_Criteria As String
         
         Dim sLMC_EqpID As String
         Dim sLMC_BatchID As String
         Dim sLMC_CassetteID As String
         Dim sLMC_LotID As String
         Dim sLMC_LotType As String
         Dim sLMC_PortID As String
         Dim sLMC_Recipe1 As String
         Dim sLMC_ProductID As String
         Dim sLMC_ChamberID As String
         Dim sLMC_SlotID As String
         Dim slotValue As Integer
         
    
22       On Error GoTo ErrorHandler
23       sLMC_EqpID = ""
24       sLMC_BatchID = ""
25       sLMC_CassetteID = ""
26       sLMC_LotID = ""
27       sLMC_LotType = ""
28       sLMC_PortID = ""
29       sLMC_Recipe1 = ""
30       sLMC_ProductID = ""
31       sLMC_ChamberID = ""
32       sLMC_SlotID = ""
         
34       Call Parent.LogTrace(0, "[MW_LMC_EVENT] Received an VEFI Event " & objVfei.CommandID)

36       For i = 1 To iLMCEventCount
37           sLMC_EventName = GetToken2(sLMCEvent(i), ",", 1)
38           sLMC_VefiMsg = GetToken2(sLMCEvent(i), ",", 2)
39           tempLMC_Criteria = GetToken2(sLMCEvent(i), ",", 3)

41           If tempLMC_Criteria <> "N" And tempLMC_Criteria <> "" Then
42               On Error Resume Next
43               sVefi_Criteria = objVfei.Var(UCase(Trim(tempLMC_Criteria))).Value
44               If Err.Number <> 0 Then
45                   Err.Clear
46                   sVefi_Criteria = "NULL"
47               End If
48               On Error GoTo ErrorHandler
49           End If
             
51           If objVfei.CommandID = sLMC_VefiMsg And sLMC_VefiMsg <> "N" And sVefi_Criteria <> "NULL" Then
52               For Each objLot In objBatch.LotList
53                   If tempLMC_Criteria <> "N" Then
54                       If InStr(UCase(tempLMC_Criteria), "CONTROLLER") > 0 And sVefi_Criteria <> "" And sVefi_Criteria <> "NULL" Then
55                           If sVefi_Criteria = objLot.LotID Then
56                               sLMC_LotID = objLot.LotID
57                               sLMC_PortID = objLot.Port
58                               sLMC_ChamberID = "NULL"
59                               If sLMC_EventName = "WAFER_START" Or sLMC_EventName = "WAFER_END" Then
60                                   slotValue = val(objVfei.Var("SLOT").Value)
61                                   If slotValue > 25 Then
62                                       slotValue = slotValue - 25
63                                   End If
64                                   sLMC_SlotID = CStr(slotValue)
65                               Else
66                                   sLMC_SlotID = "NULL"
67                               End If
68                               sLMC_BatchID = objBatch.BatchID
69                           End If
70                       Else
71                           Call Parent.LogTrace(0, "[MW_LMC_EVENT] TaskAttribute ItemName don't set PORT LOT CHAMBER SLOT or VEFI Msg Value has error, please check it. =>TaskAttribute:" & tempLMC_Criteria & ",VEFI Msg:" & sVefi_Criteria)
72                       End If
73                   ElseIf tempLMC_Criteria = "N" Then
74                       sLMC_PortID = objLot.Port
75                       sLMC_LotID = objLot.LotID
76                       sLMC_ChamberID = "NULL"
77                       sLMC_SlotID = "NULL"
78                       sLMC_BatchID = objBatch.BatchID
79                   Else
80                       Call Parent.LogTrace(0, "[MW_LMC_EVENT] TaskAttribute ItemName is empty, please check it. =>" & tempLMC_Criteria)
81                   End If
82                   sLMC_EqpID = Parent.PUID
83                   sLMC_CassetteID = objLot.Cassette
84                   sLMC_LotType = objLot.LotType
85                   sLMC_Recipe1 = objLot.Recipe1
86                   sLMC_ProductID = objLot.Product
87                   Call Send_LMC_Evnet(Parent, Me, sLMC_EventName, sLMC_EqpID, sLMC_CassetteID, sLMC_LotID, sLMC_LotType, sLMC_BatchID, sLMC_PortID, sLMC_Recipe1, sLMC_ProductID, sLMC_ChamberID, sLMC_SlotID)
88                   Call Parent.LogTrace(0, "[MW_LMC_EVENT] (LOT=>" & objLot.LotID & ") Send an LMC_Event=>" & sLMC_EventName & ",for VEFIMsg:" & objVfei.CommandID)
89               Next
90           End If
91       Next i
    
93       Exit Sub
94 ErrorHandler:
95       Call Parent.LogTrace(0, FormatErrMsg(MODULE_ID & ".ReceivedVfei", Err, "[" & Erl & "]" & Err.Description))
End Sub

Private Property Let TaskAncestor_TaskID(ByVal RHS As String)
1        TaskID = RHS
End Property

Private Property Get TaskAncestor_TaskID() As String
1        TaskAncestor_TaskID = TaskID
End Property

Public Function Set_LMC_VEFI_Criteria(sCriteriaType As String) As String
         Dim objHostTrans As HostTransaction
         Dim objLot As Lot
         Dim stempCriteria As String
         Dim i As Integer
         
6        On Error GoTo ErrorHandler
7        stempCriteria = ""
8        Select Case (sCriteriaType)
             Case "CONTROLLER"
10               For Each objLot In objBatch.LotList
11                   stempCriteria = stempCriteria & objLot.LotID & ","
12               Next objLot
13               Call Parent.LogTrace(0, "VEFI Criteria(LOT)=" & stempCriteria)
14           Case "N"
15               stempCriteria = ""
16               Call Parent.LogTrace(0, "VEFI Criteria(Nothing)=" & stempCriteria)
17           Case Else
18               stempCriteria = "N/A"
19               Call Parent.LogTrace(0, "VEFI Criteria(Error)=" & stempCriteria)
20       End Select

22       Set_LMC_VEFI_Criteria = stempCriteria
         
24       Call Parent.LogTrace(0, "Set_LMC_VEFI_Criteria is over=>" & sCriteriaType)
25       Exit Function
26 ErrorHandler:
27       Call Parent.LogTrace(0, FormatErrMsg(MODULE_ID & ".Set_LMC_VEFI_Criteria", Err, "[" & Erl & "]" & Err.Description))
End Function

Public Sub Send_LMC_Evnet(Parent As tsEC.clsRunManager, objTask As Object, sEventName As String, sPU As String, sCassette As String, sLot As String, sLotType As String, sBatchID As String, sPort As String, sPPID As String, sProduct As String, sChamber As String, sSlot As String)
         Dim objHostTrans As HostTransaction
         
3        On Error GoTo ErrorHandler
4        Set objHostTrans = Parent.NewHostTransaction(objTask, "LMC_EVENT")
5        With objHostTrans.Primary
6            .Item("EQPID") = sPU
7            .Item("EVENT_NAME") = sEventName
8            .Item("CARRIER_ID") = sCassette
9            .Item("LOT_ID") = sLot
10           .Item("LOT_TYPE") = sLotType
11           .Item("PORT_ID") = sPort
12           .Item("WAFER_ID") = sSlot
13           .Item("CHAMBER_ID") = sChamber
14           .Item("BATCHID") = sBatchID
15           .Item("PPID") = sPPID
16           .Item("PRODUCTID") = sProduct
17           .Item("PLANID") = "NULL"
18           .Item("SUBPLANID") = "NULL"
19           .Item("STEPSEQ") = "NULL"
20           .Item("STEPID") = "NULL"
21           .Item("PRIORITY") = "NULL"
22           .Item("CHAMBER_PPID") = "NULL"
23           .Item("SLOT_NO") = "NULL"
24           .Item("CHAMBER_SITEID") = "NULL"
25           .Item("TOOLTIME") = Format(Now, "YYYYMMDDHHMMSS00")
26       End With
27       objHostTrans.Send
28       Call Parent.LogTrace(0, "Send LMC_Evnet,Cassette:" & sCassette & ";; EVENT_NAME:" & sEventName)
29       Exit Sub
30 ErrorHandler:
31       Call Parent.LogTrace(0, FormatErrMsg(MODULE_ID & ".Send_LMC_Evnet", Err, "[" & Erl & "]" & Err.Description))

End Sub
Function GetTokenCount2(ByVal str As String, dm As String) As Integer
         Dim tmp_str As String
         Dim i As Integer
         Dim no_dm As Boolean
    
5        no_dm = True
6        GetTokenCount2 = 0
7        tmp_str = str
8        While Len(tmp_str) <> 0
9            i = InStr(tmp_str, dm)
10           If i = 0 Then
11               If no_dm = True Then
12                   GetTokenCount2 = 0
13                   Exit Function
14               Else
15                   GetTokenCount2 = GetTokenCount2 + 1
16                   Exit Function
17               End If
18           End If
19           no_dm = False
20           tmp_str = Right$(tmp_str, Len(tmp_str) - i)
21           GetTokenCount2 = GetTokenCount2 + 1
22       Wend

End Function

