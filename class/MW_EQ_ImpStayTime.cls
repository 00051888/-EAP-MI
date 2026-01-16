VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MW_EQ_ImpStayTime"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements ts_ECAncestor.TaskAncestor

Private Const MODULE_ID = "MW_EQ_ImpStayTime"

Dim Parent      As Object
Dim TaskID      As String
Dim objBatch    As Batch
Dim objLot      As Lot
Dim ImpCompTime As Date

Private Sub TaskAncestor_Execute(Parameter As Variant)
    On Error GoTo ErrorHandler
    
    Set objBatch = Parameter
    
    Call Parent.SubscribeVfeiMsg(Me, Parent.GetPuObject.EqDrvList(1).EqDrvID, "Wafer_complete")
    
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
    Dim iStayTime   As Integer
    Dim sTemp       As String
    Dim sLotID      As String
    Dim sSlotID     As String
    Dim bSendSPC    As Boolean
    Dim tmpObjLot   As Lot
    Dim loLot       As Lot
    
    On Error GoTo ErrorHandler
    
    If objVfei.ErrorCode <> 0 Then
        Call Parent.LogTrace(0, MODULE_ID & ".ReceivedVfei CommandID=" & objVfei.CommandID & _
                                            " ErrCode=" & objVfei.ErrorCode & _
                                            " ErrText=" & objVfei.ErrorText)
        Exit Sub
    End If
    
    If objVfei.CommandID <> "Wafer_complete" Then Exit Sub
    
    If objVfei.CommandID = "Wafer_complete" Then
        sLotID = objVfei.Var("LOT_ID").Value
        sSlotID = objVfei.Var("SLOT").Value
        
        'Call Parent.LogTrace(0, MODULE_ID & ": LotId=" & sLotID & " Slot=" & sSlotID & " Wafer_complete")
        
        If VerifyBatchAndLot(objVfei) Then   'Check objBatch and objLot
            'Calculate Imp_StayTime
            On Error Resume Next
            sTemp = objBatch.Attributes.Item("ImpStayTime").Value
            If Err <> 0 Then
                '該objBatch第一次implant
                Call Parent.LogTrace(0, MODULE_ID & ": LotId=" & sLotID & " Slot=" & sSlotID & " first Wafer_complete")
                bSendSPC = False
                sTemp = "0"
                ImpCompTime = Now
                objBatch.Attributes.Add("ImpStayTime").Value = sTemp
            Else
                bSendSPC = True
                sTemp = DateDiff("S", ImpCompTime, Now)
                objBatch.Attributes.Item("ImpStayTime").Value = sTemp
                Call Parent.LogTrace(0, MODULE_ID & ": LotId=" & sLotID & " Slot=" & sSlotID & " ImpStayTime =" & sTemp)
                
                '算完StayTime, 將現在時間重新填入ImpCompTime, 用於下次計算
                ImpCompTime = Now
            End If
            Err.Clear
            On Error GoTo ErrorHandler
            
            'Send Imp_StayTime to MoniSPC
            If bSendSPC Then
                For Each loLot In objBatch.LotList
                    If loLot.LotID = sLotID Then
                        Set tmpObjLot = loLot
                        Exit For
                    End If
                Next
                
                If tmpObjLot Is Nothing Then
                    Call Parent.LogTrace(0, MODULE_ID & ": LotId=" & sLotID & " cannot fine tmpObjLot to send SPC")
                Else
                    Call SENDSPCDATA(objBatch, tmpObjLot, "STAYTIME", sTemp, CInt(sSlotID))
                End If
            End If
        Else
            Call Parent.LogTrace(0, MODULE_ID & ": LotId=" & objVfei.Var("LOT_ID").Value & " not at Batch[" & objBatch.BatchID & "]")
        End If
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

Private Function VerifyBatchAndLot(objVfei As vfei.Message22) As Boolean
    Dim loLot As Lot
    Dim isFind As Boolean
    Dim TempStr As String
    
    isFind = False
    
    On Error Resume Next
    Call Parent.LogTrace(0, "VerifyBatchAndLot: LotId=" & objVfei.Var("LOT_ID").Value & " BatchID=" & objBatch.BatchID)
        
    TempStr = objVfei.Var("LOT_ID").Value
    If Err.Number = 0 Then
        For Each loLot In objBatch.LotList
            If loLot.LotID = TempStr Then
                isFind = True
                Exit For
            End If
        Next
    End If
    Err.Clear

    TempStr = objVfei.Var("CASSETTE").Value
    If Err.Number = 0 Then
        Err.Clear
        For Each loLot In objBatch.LotList
            If loLot.Cassette = TempStr Then
                isFind = True
                Exit For
            End If
        Next
    End If
    Err.Clear
    
    VerifyBatchAndLot = isFind
    
    If isFind Then
        Call Parent.LogTrace(0, "VerifyBatchAndLot: LotId=" & objVfei.Var("LOT_ID").Value & " is in BatchID=" & objBatch.BatchID)
    Else
        Call Parent.LogTrace(0, "Received an Event " & objVfei.CommandID & ". But can't find associate Object!")
    End If
    
    Exit Function

ErrorHandler:
    Call Parent.LogTrace(0, FormatErrMsg(MODULE_ID & ".FindLot", Err, Err.Description))
End Function

Private Sub SENDSPCDATA(objBatch As Batch, objLot As Lot, sDataName As String, sDataValue As String, Optional slotNo As Integer)
'Primary=SPC_MONITORDATA <%Subject> <%TID> <FMCassette> <processPU>
'        <P_Recipe> <Parameter> <measurePU> <M_Recipe> <Operator>
'        <#Slots>
'        [<Slot> <Site> <Value> <ASCII_value> <Chamber>
'        ]
'Secondary=//SPC_MONITORDATA <%Status> <%TID>

    Dim objHostTrans As HostTransaction

    On Error Resume Next
    'Call Parent.LogTrace(0, "Start - Upload Delta Data to MoniSPC:" & objLot.Recipe1 & "/" & sDataName & "/" & sDataValue)
    
    Set objHostTrans = Parent.NewHostTransaction(Me, "SPC_MONITORDATA")
    With objHostTrans.Primary
        .Item("FMCassette").Value = objLot.Cassette
        .Item("processPU").Value = Parent.PUID
        .Item("measurePU").Value = "NULL"
        .Item("Operator").Value = objLot.OperatorID
        .Item("P_Recipe").Value = objLot.Recipe1
        .Item("Parameter").Value = sDataName
        .Item("M_Recipe").Value = "NULL"
    End With

    With objHostTrans.Primary.Item("Slots").Add
        .Item("Slot").Value = slotNo
        .Item("Site").Value = 1
        .Item("Value").Value = sDataValue
        .Item("ASCII_value").Value = sDataValue
        .Item("Chamber").Value = 1
    End With
    objHostTrans.Send
    Call Parent.LogTrace(0, "Upload Delta Data to MoniSPC:" & objLot.Recipe1 & "/" & sDataName)
    Set objHostTrans = Nothing
    
End Sub

'''Private Function FindLot(objVfei As vfei.Message22) As Boolean
'''    Dim loLot As Lot
'''    Dim isFind As Boolean
'''    Dim isBatch As Boolean
'''    Dim TempStr As String
'''
'''    isFind = False
'''    isBatch = False
'''
'''    On Error Resume Next
'''
'''    Err.Clear
'''    TempStr = objVfei.Var("LOT_ID").Value
'''    If Err.Number = 0 Then
'''        For Each loLot In objBatch.LotList
'''            If loLot.LotID = TempStr Then
'''                isFind = True
'''                Exit For
'''            End If
'''        Next
'''    End If
'''    If isFind = True Then GoTo EndOfFinding
'''
'''    Err.Clear
'''    TempStr = objVfei.Var("CASSETTE").Value
'''    If Err.Number = 0 Then
'''        Err.Clear
'''        For Each loLot In objBatch.LotList
'''            If loLot.Cassette = TempStr Then
'''                isFind = True
'''                Exit For
'''            End If
'''        Next
'''    End If
'''    If isFind = True Then GoTo EndOfFinding
'''
'''    Err.Clear
'''    TempStr = objVfei.Var("BATCH_ID").Value
'''    If Err.Number = 0 Then
'''        If objBatch.BatchID = TempStr Then
'''            isFind = True
'''            isBatch = True
'''        End If
'''    End If
'''    If isFind = True Then GoTo EndOfFinding
'''
'''    Err.Clear
'''    TempStr = objVfei.Var("PORT").Value
'''    If Err.Number = 0 Then
'''        For Each loLot In objBatch.LotList
'''            If loLot.Port = TempStr Then
'''                isFind = True
'''                Exit For
'''            End If
'''        Next
'''    End If
'''    If isFind = True Then GoTo EndOfFinding
'''
'''    Err.Clear
'''    TempStr = objVfei.Var("SMIFPORTID").Value
'''    If Err.Number = 0 Then
'''        For Each loLot In objBatch.LotList
'''            If loLot.Port = TempStr Then
'''                isFind = True
'''                Exit For
'''            End If
'''        Next
'''    End If
'''    If isFind = True Then GoTo EndOfFinding
'''
'''    Call Parent.LogTrace(0, "Received an Event " & objVfei.CommandID & ". But can't find associate Object!")
'''
'''EndOfFinding:
'''    FindLot = isFind
'''
'''Exit Function
'''ErrorHandler:
'''    Call Parent.LogTrace(0, FormatErrMsg(MODULE_ID & ".FindLot", Err, Err.Description))
'''End Function

'''Private Sub Send_PDS(ByVal psParaName As String, _
'''                     ByVal psSlotId As String, _
'''                     ByVal psParaValue As String)
'''    Dim objVfei As vfei.Message22
'''
'''    On Error GoTo ErrorHandler
'''
'''    Set objVfei = New Message22
'''    objVfei.CommandID = "PDS"
'''    objVfei.MachineID = Parent.GetPuObject.EqDrvList(1).EqDrvID
'''    objVfei.MessageType = "E"
'''
'''    objVfei.AddVar "LOT_ID", "A", objLot.LotID
'''    objVfei.AddVar "SLOT", "U4", psSlotId
'''    objVfei.AddVar "CHAMBER", "A", ""
'''    objVfei.AddVar "RECIPESTEP", "A", "ALL " & objLot.Recipe1
'''    objVfei.AddVar "SITE", "I1", "1"
'''
'''    objVfei.AddVar "PDS", "L", ""
'''    objVfei.Var("PDS").AddVar psParaName, "A", psParaValue
'''
'''    Call Parent.SendVfei(objVfei, Me)
'''
'''Exit Sub
'''ErrorHandler:
'''    Call Parent.LogTrace(0, FormatErrMsg(MODULE_ID & ".Send_PDS", Err, Err.Description))
'''
'''End Sub
