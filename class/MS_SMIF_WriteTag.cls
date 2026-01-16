VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MS_SMIF_WriteTag"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements ts_ECAncestor.TaskAncestor

Private Const MODULE_ID = "GS_SMIF_WriteTag"

Dim parent As tsEC.clsRunManager
Dim TaskID As String
Dim objBatch As Batch
Dim objLot As Lot
Dim objPort As PortItem
Dim objSmif As SmifItem
Dim sFAB As String
Dim sPortID As String
Dim sSmifID As String
Dim m_TagFileName As String
Dim vParameter As Variant
Dim HoldLotCount As Integer
Private Function GetPosFromIniSetting(tKey As String) As Integer

    Dim iPos As Integer
    Dim iItemCount As Integer
    Dim sItemValue As String
    Dim sItemName As String
    Dim iLength As Integer
    Dim iIndex As Integer
    Dim sTAGInfo As String
    Dim m_TagFileName As String
        
    GetPosFromIniSetting = 0
    m_TagFileName = parent.GetTaskAttribute(TaskID, "TagFile")
   
    If m_TagFileName = "" Then
        m_TagFileName = App.Path & "\TagCfg.ini"
    End If
    
    iItemCount = Val(GetIniSetting(m_TagFileName, "TAG_ITEM_FORMAT", "ItemCount"))
        
    For iIndex = 1 To iItemCount
        sItemValue = Trim$(GetIniSetting(m_TagFileName, "TAG_ITEM_FORMAT", "Item" & CStr(iIndex)))
        sItemName = Trim$(Gettoken(sItemValue, ",", 1))
        If sItemName = tKey Then
            GetPosFromIniSetting = GetPosFromIniSetting + 1
            Exit For
        End If
        GetPosFromIniSetting = GetPosFromIniSetting + Val(Gettoken(sItemValue, ",", 3))
    Next
        

End Function


'Destination = VHI
'Primary=QUERY_TAG_INFO <%Subject> <%TID> <TAG_ID> <LOT_ID> <Cassette_ID>
'Secondary=//QUERY_TAG_INFO <%Status> <%TID> <LOT_ID> <LotPriority> <PU> <Cassette_ID> <LotQuantity> <LotStatus> <LotType> <Recipe> <ReticleFamily> <ProductNumber> <ProcessNumber> <BatteryStartDate> <WaferFlag> <CleanDueDate> <DirtyBit>
Private Sub TaskAncestor_Execute(Parameter As Variant)
    Dim objHostTrans As HostTransaction
    Dim objnewVfei As New Message22
    Dim sRMMode As String
    
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
        Set objSmif = parent.GetPuObject.SmifList(sSmifID)
    ElseIf TypeOf Parameter Is SmifItem Then
        Set objSmif = Parameter
        sSmifID = objSmif.SMIFID
        sPortID = objSmif.RelatedPortID
        Set objPort = parent.GetPuObject.PortList(sPortID)
    ElseIf TypeOf Parameter Is Batch Then
        Set objBatch = Parameter
        Set objLot = objBatch.LotList(1)
        sPortID = objLot.Port
        Set objPort = parent.GetPuObject.PortList(sPortID)
        sSmifID = objPort.RelatedSMIFID
        Set objSmif = parent.GetPuObject.SmifList(sSmifID)
    ElseIf TypeOf Parameter Is Lot Then
        'modified by weixi for furnace 3/22
        Dim objBatchNdx As Batch
        Dim objLotNdx As Lot
        Set objLot = Parameter
        sPortID = objLot.Port
        Set objPort = parent.GetPuObject.PortList(sPortID)
        sSmifID = objPort.RelatedSMIFID
        Set objSmif = parent.GetPuObject.SmifList(sSmifID)
        For Each objBatchNdx In parent.GetBatchList
            For Each objLotNdx In objBatchNdx.LotList
                If objLotNdx.LotID = objLot.LotID Then
                    Set objBatch = objBatchNdx
                    Exit For
                End If
            Next objLotNdx
            If Not objBatch Is Nothing Then
                Exit For
            End If
        Next objBatchNdx
    End If


    Call parent.SubscribeVfeiMsg(Me, "", "READ_TAG_STATUS", "SMIFID", sSmifID)
   
    Set objSmif = parent.GetPuObject.SmifList(sSmifID)
   
    m_TagFileName = parent.GetTaskAttribute(TaskID, "TagFile")
   
    If m_TagFileName = "" Then
        m_TagFileName = App.Path & "\TagCfg.ini"
    End If
   
    '   sPortID = objSmif.RelatedPortID
   
    objnewVfei.CommandID = "READ_TAG"
    objnewVfei.MachineID = sSmifID
    If objSmif.SMIFType = "SC" Then
        objnewVfei.AddVar "PORT", "U1", sPortID
    End If
    
    '--- Modify for IRR047 by hsiulin on 6/29
    '    sRMMode = Parent.GetPuObject.ControlMode
    Dim iPos As Integer
    iPos = InStr(parent.GetPuObject.ControlMode, "/")
    If iPos > 0 Then
        sRMMode = Mid(parent.GetPuObject.ControlMode, 1, iPos - 1)
    Else
        sRMMode = parent.GetPuObject.ControlMode
    End If
    '---
    
    Select Case (sRMMode)
        Case "ONLINE_REMOTE"
            Call parent.SendVfei(objnewVfei, Me)
        Case "ONLINE_LOCAL"
            Call parent.TaskComplete(TaskID, "0", vParameter)
        Case Else
            Call parent.TaskComplete(TaskID, "1", vParameter)
    End Select
    
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
    Dim objnewVfei As New Message22
    Dim sRawData As String
    Dim ItemCount As Integer
    Dim ItemValue As String
    Dim ItemName As String
    Dim HostValue As String
    Dim Length As Integer
    Dim iIndex As Integer
    Dim sTaskAttHoldCode As String
    Dim sTaskAttHoldComment As String
    Dim sTmpLot As String
    
    On Error GoTo ErrorHandler
    
    If objHostTrans.TransactionName = "QUERY_TAG_INFO" Then
        If InStr(objHostTrans.Tag, "LotMix") <> 0 Then
           sTmpLot = Trim(objHostTrans.Secondary.Item("LOT_ID"))
           
           sTaskAttHoldCode = parent.GetTaskAttribute(TaskID, "HoldCode")
          '若Task Attribute沒有設定HoldCode => default = H3232
           If sTaskAttHoldCode = "" Then sTaskAttHoldCode = "H3232"
           
           sTaskAttHoldComment = parent.GetTaskAttribute(TaskID, "HoldComment")
           '若Task Attribute沒有設定HoldComment => default = objHostTrans.Tag
           If sTaskAttHoldComment = "" Then sTaskAttHoldComment = objHostTrans.Tag
           
           Call SendHoldLot(sTmpLot, sTaskAttHoldCode, sTaskAttHoldComment)
        Else
           ItemCount = Val(GetIniSetting(m_TagFileName, "TAG_ITEM_FORMAT", "ItemCount"))
           For iIndex = 1 To ItemCount
               ItemValue = Trim$(GetIniSetting(m_TagFileName, "TAG_ITEM_FORMAT", "Item" & CStr(iIndex)))
               ItemName = Trim$(Gettoken(ItemValue, ",", 1))
               If objHostTrans.Secondary.Item("WaferFlag") <> "Y" And ItemName = "ID:" Then
                   ItemName = "NW:"
               End If
               If ItemName = "FB:" Then
                   If sFAB <> "" Then
                       HostValue = sFAB
                   Else
                       HostValue = "8F"
                   End If
               Else
                   HostValue = Trim(objHostTrans.Secondary.Item(Trim$(Gettoken(ItemValue, ",", 2))))
               End If
               If ItemName = "PU:" Then
                  If objHostTrans.Secondary.Item("PU") = "NULL" Then
                     If objLot.Attributes("MoveOutPU") = "NULL" Then
                         If objLot.Attributes("MoveOutPUF") = "NULL" Then
                             HostValue = ""
                         Else
                             HostValue = objLot.Attributes("MoveOutPUF")
                         End If
                     Else
                         HostValue = objLot.Attributes("MoveOutPU")
                     End If
                  End If
                End If
                        
                Length = Val(Gettoken(ItemValue, ",", 3))
                If HostValue <> "NULL" Then
                   sRawData = sRawData & Mid$(Format(ItemName + HostValue, "!" & String(Length, "@")), 1, Length)
                Else
                   sRawData = sRawData & Mid$(Format(ItemName, "!" & String(Length, "@")), 1, Length)
                End If
                Debug.Print sRawData
             Next
             sRawData = Format(sRawData, "!" & String(14 * 16, "@"))
                
             objnewVfei.CommandID = "WRITE_TAG"
             objnewVfei.MachineID = sSmifID
             If objSmif.SMIFType = "SC" Then
                 objnewVfei.AddVar "PORTID", "A", sPortID
             End If
             objnewVfei.AddVar "RAW_TAG_INFO", "A", sRawData
             Call parent.SendVfei(objnewVfei, Me)
         End If
    ElseIf objHostTrans.TransactionName = "EQALARM_EAS" Then
    
    
    ElseIf objHostTrans.TransactionName = "HOLD_LOT" Then
           HoldLotCount = HoldLotCount - 1
           If objHostTrans.ErrorCode <> 0 Then
               If objHostTrans.ErrorCode = vbObjectError + 2000 + 5 Then
                   Call AlarmReport(parent, Me, "0110", "Hold " & objHostTrans.Tag & " Failure.", objHostTrans.ErrorText)
               Else
                   Call AlarmReport(parent, Me, "0210", "Hold " & objHostTrans.Tag & " Failure.", objHostTrans.ErrorText)
               End If
           Else
               Call parent.LogTrace(0, MODULE_ID & "HoldLot " & objHostTrans.Tag & " Success.")
           End If

           If HoldLotCount = 0 Then
               Call parent.TaskComplete(TaskID, "1", vParameter)
           End If
        
    End If
         
    Exit Sub
ErrorHandler:
    Call parent.LogTrace(0, FormatErrMsg(MODULE_ID & ".ReceivedHostMsg", Err, Err.Description))
End Sub

Private Sub TaskAncestor_ReceivedVfei(objVfei As vfei.Message22)
    'Mark on 6/9 by hsiulin
    '    Dim sSmifID As String
    '----------
    Dim sTAGInfo  As String
    Dim sCassetteID As String
    'Dim objBatch As Batch, objLot As Lot
    Dim objHostTrans As HostTransaction
    Dim sSERIAL_ID As String
    Dim iPlaceIndex As Integer
    Dim objnewVfei As New Message22
    Dim objTmpBatch As Batch
    Dim objTmpLot As Lot
    Dim sIsOEMLot As String
    Dim IsOEMLot As Boolean
    Dim iCheckPos As Integer
        
    iCheckPos = GetPosFromIniSetting("LOT:")
    '---
    Select Case objVfei.CommandID
        Case "READ_TAG"
            sSmifID = objVfei.MachineID
            Set objSmif = parent.GetPuObject.SmifList(sSmifID)
            If objVfei.ErrorCode <> 0 Then
                '20150605 Judy add,Read Tag Failed Before Writing Tag,send alarm to EQ Alarm
                Call SendRecordEQAlarm("SORTER01", "-10001", "Read Tag Failed before Writing Tag.")
                Call AlarmReport(parent, Me, "0102", "Read Tag Failed before Write Tag", "Please update tag using Sorter or ALU, Port:" & objPort.PortID & ", Smif:" & objSmif.SMIFID)
                Call parent.LogTrace(0, "Read TAG Failed before Write Tag: " & objVfei.ErrorText)
                Call parent.TaskComplete(TaskID, "1", vParameter)
                Exit Sub
            End If

            sTAGInfo = objVfei.Var("RAW_TAG_INFO").Value
            
            On Error Resume Next
            sCassetteID = Trim$(Mid$(sTAGInfo, 4, 13))
            IsOEMLot = False
            For Each objTmpBatch In parent.GetBatchList
                For Each objTmpLot In objTmpBatch.LotList
                    If objTmpLot.Cassette = Trim(sCassetteID) Then
                        Set objBatch = objTmpBatch
                        Set objLot = objTmpLot
                    End If
                Next objTmpLot
            Next objTmpBatch
            
            sIsOEMLot = ""
            sIsOEMLot = objLot.Attributes("_UI_IsOEMLOT").Value
            If Err.Number = 0 Then
                If sIsOEMLot = "Y" Then
                    IsOEMLot = True
                End If
            End If
            Err.Clear
            On Error GoTo ErrorHandler
            
            If (parent.GetBatchList(1).RunMode = RUNMODE_PR Or parent.GetBatchList(1).RunMode = RUNMODE_MEPR) And _
                (Not IsOEMLot) Then
                ' modified by weixi 03/28
                sCassetteID = Trim$(Mid$(sTAGInfo, 4, 13))
                sSERIAL_ID = objVfei.Var("SERIAL_ID").Value
                sFAB = Trim$(Mid$(sTAGInfo, 191, 2))
                If sCassetteID <> objSmif.RelatedCassetteID Then                    'Modify by simon 2000/09/20
                    '20150605 Judy Modify,for Lot Mix
                    'Call AlarmReport(Parent, Me, "0101", "Error occur while Write Tag", "Please update tag using Sorter or ALU, Port:" & objPort.PortID & ", Smif:" & objSmif.SMIFID)
                    Call AlarmReport(parent, Me, "0101", sCassetteID & "<>" & objSmif.RelatedCassetteID & " (疑似混料)Lot maybe mix !! Cassette is not right. Error occur while Write Tag", "Please double confirm if Lot mix!!  Port:" & objSmif.RelatedCassetteID & ", Smif:" & objSmif.SMIFID)
                    Call SendQueryTagInfo(sCassetteID, "LotMix!" & sCassetteID & "<>" & objSmif.RelatedCassetteID)  '查詢LotId 進行Hold Lot
                    Call SendQueryTagInfo(objSmif.RelatedCassetteID, "LotMix!" & sCassetteID & "<>" & objSmif.RelatedCassetteID) '查詢LotId 進行Hold Lot
                    Call parent.LogTrace(0, "Lot maybe mix . This Cassette is not right")
                    'Call Parent.TaskComplete(TaskID, "1", vParameter)
                    Exit Sub
                End If
                '-- Add for check tag methed from Instr to get from ini file FOR V0501
                '            iPlaceIndex = InStr(1, sTAGInfo, "LOT", vbTextCompare)
                '            If iPlaceIndex <> 0 Then
                If Mid(sTAGInfo, iCheckPos, 3) = "LOT" Then
                    Call AlarmReport(parent, Me, "0101", "Error occur while Write Tag", "Please update tag using Sorter or ALU, Port:" & objPort.PortID & ", Smif:" & objSmif.SMIFID)
                    Call parent.LogTrace(0, "This POD is not right, Please Check")
                    Call parent.TaskComplete(TaskID, "1", vParameter)
                    Exit Sub
                End If
                Set objHostTrans = parent.NewHostTransaction(Me, "QUERY_TAG_INFO")
                With objHostTrans.Primary
                    .Item("TAG_ID") = Trim(sSERIAL_ID)
                End With
                Call objHostTrans.Send
            Else        'add 12/28 for other modes
                '-- Add for check tag methed from Instr to get from ini file FOR V0501
                '               If Mid$(sTAGInfo, 17, 3) = "EMP" Then
                '                  iPlaceIndex = InStr(1, sTAGInfo, "EMP", vbTextCompare)
                If Mid(sTAGInfo, iCheckPos, 3) = "EMP" Then
                    iPlaceIndex = iCheckPos
                    '                  sTAGInfo = Left(sTAGInfo, iPlaceIndex - 1) & "LOT" & Mid$(sTAGInfo, 20, Len(sTAGInfo) - 19)
                    Mid(sTAGInfo, iPlaceIndex, 4) = "LOT:"
                    '----
                    objnewVfei.CommandID = "WRITE_TAG"
                    objnewVfei.MachineID = sSmifID
                    If objSmif.SMIFType = "SC" Then
                        objnewVfei.AddVar "PORTID", "A", sPortID
                    End If
                    objnewVfei.AddVar "RAW_TAG_INFO", "A", sTAGInfo
                    Call parent.SendVfei(objnewVfei, Me)
                Else
                    'Modify by James Lin 00/03/21
                    '-- Add for check tag methed from Instr to get from ini file FOR V0501  by hsiulin
                    '                   If Mid$(sTAGInfo, 17, 3) = "LOT" Then
                    If Mid(sTAGInfo, iCheckPos, 3) = "LOT" Then
                        '-------
                        Call AlarmReport(parent, Me, "0101", Translate(parent, "Tag 為 LOT，無法 write Tag", "Tag is LOT, can not write Tag."), Translate(parent, "請出站後用 Sorter 或 ALU update, Port:", " Please use Sorter or ALU update Tag when the LOT move out. Port:") & objPort.PortID & ", Smif:" & objSmif.SMIFID)
                        Call parent.LogTrace(0, "Tag 為 LOT，無法 write Tag, Port:" & objPort.PortID & ", Smif:" & objSmif.SMIFID)
                    Else
                        Call AlarmReport(parent, Me, "0101", Translate(parent, "Tag 格式不對，無法 write Tag", "Tag's format is error,can not write Tag"), "Port:" & objPort.PortID & ", Smif:" & objSmif.SMIFID)
                        Call parent.LogTrace(0, "Tag 格式不對，無法 write Tag, Port:" & objPort.PortID & ", Smif:" & objSmif.SMIFID)
                    End If
                   
                    Call parent.TaskComplete(TaskID, "1", vParameter)
                    Exit Sub
                End If
            End If
        
        Case "WRITE_TAG"
            If objVfei.ErrorCode <> 0 Then
                Call AlarmReport(parent, Me, "0101", "Write Tag Failed", "Please update tag using Sorter or ALU, Port:" & objPort.PortID & ", Smif:" & objSmif.SMIFID)
                Call SendRecordEQAlarm("SORTER01", "-10001", "Write Tag Failed after Writing Tag.")
                Call parent.LogTrace(0, "Write Tag Failed: " & objVfei.ErrorText)
                Call parent.TaskComplete(TaskID, "1", vParameter)
            Else
                Call parent.LogTrace(0, "Write Tag successful")
                Call parent.TaskComplete(TaskID, "0", vParameter)
            End If

           
    End Select
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


Private Sub Class_Terminate()
    Set objBatch = Nothing
    Set objLot = Nothing
    Set objPort = Nothing
    Set objSmif = Nothing
    If IsObject(vParameter) Then
        Set vParameter = Nothing
    End If
End Sub
Private Sub SendRecordEQAlarm(sUnRealPU As String, sALID As String, sALTX As String)
         Dim sAlarmText As String
         Dim objHostTrans As HostTransaction

4        Set objHostTrans = parent.NewHostTransaction(Me, "EQALARM_EAS")
5        With objHostTrans.Primary
6            .Item("PU") = sUnRealPU
7            .Item("AlarmType") = "EQAlarm"
8            .Item("AlarmID") = sALID
9            .Item("AlarmText") = parent.PUID & "/" & sALTX
10       End With
11       objHostTrans.Send
12       Exit Sub
13 ErrorHandler:
14       Call parent.LogTrace(0, FormatErrMsg(MODULE_ID & ".SendRecordEQAlarm", Err, Err.Description))
End Sub

Private Sub SendHoldLot(sLotID As String, sHoldLotReason As String, sHoldLotComment As String)
1        On Error GoTo ErrorHandler
         Dim objHostTrans As HostTransaction
    
4        Call parent.LogEvent("Hold Lot", sLotID, "")
    
6        Set objHostTrans = parent.NewHostTransaction(Me, "HOLD_LOT")
7        With objHostTrans.Primary
8            .Item("LOT_ID") = sLotID
9            .Item("Engineer") = "SEMIAUTO"
10           .Item("Operator") = "SEMIAUTO"
11           .Item("HoldReason") = sHoldLotReason
12           .Item("Comment") = sHoldLotComment
13       End With
14       objHostTrans.Tag = sLotID
15       objHostTrans.Send
16       HoldLotCount = HoldLotCount + 1
17       Exit Sub
18 ErrorHandler:
19       Call parent.LogTrace(0, FormatErrMsg(MODULE_ID & ".SendHoldLot", Err, Err.Description))
End Sub

Private Sub SendQueryTagInfo(sCassetteID As String, sReason As String)
1        On Error GoTo ErrorHandler
         Dim objHostTransNew As HostTransaction
    
4        Set objHostTransNew = parent.NewHostTransaction(Me, "QUERY_TAG_INFO")
5        With objHostTransNew.Primary
6            .Item("Cassette_ID") = Trim(sCassetteID)
7        End With
8        objHostTransNew.Tag = sReason
9        Call objHostTransNew.Send
10       Exit Sub
11 ErrorHandler:
12       Call parent.LogTrace(0, FormatErrMsg(MODULE_ID & ".SendQueryTagInfo", Err, Err.Description))
End Sub


