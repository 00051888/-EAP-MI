VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MS_IMP_RCMS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements ts_ECAncestor.TaskAncestor

Private Const MODULE_ID = "MS_IMP_RCMS"

Dim Parent As tsEC.clsRunManager
Dim TaskID As String
Dim objBatch As Batch
Dim objLot As Lot
Dim objPort As PortItem
Dim objSmif As SmifItem
Dim objEqDrv As EqDrv
Dim vParameter As Variant

Private Sub TaskAncestor_Execute(Parameter As Variant)
    Dim objNewVfei As New Message22
    Dim sPortID As String
    Dim sSmifID As String
    Dim sLotID As String
    Dim sRMMode As String
    Dim iIndex As Integer
    Dim objTmpLot As Lot
    
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
        Set objEqDrv = Parent.GetPuObject.EqDrvList(1)
    ElseIf TypeOf Parameter Is Batch Then
        Set objBatch = Parameter
        Set objLot = objBatch.LotList(1)
        sPortID = objLot.Port
        Set objPort = Parent.GetPuObject.PortList(sPortID)
        sSmifID = objPort.RelatedSMIFID
        Set objSmif = Parent.GetPuObject.SmifList(sSmifID)
    ElseIf TypeOf Parameter Is Lot Then
        Set objLot = Parameter
        sPortID = objLot.Port
        Set objPort = Parent.GetPuObject.PortList(sPortID)
        sSmifID = objPort.RelatedSMIFID
        Set objSmif = Parent.GetPuObject.SmifList(sSmifID)
    End If
    
    'Check Recipe Body
    For Each objTmpLot In objBatch.LotList
        objTmpLot.Status = "LOT_CHECKPPBODY"
        Call SendLotEventReport(Parent, Me, objLot)
    Next objTmpLot
 
    If objBatch.LotList.Count > 1 Then
        sLotID = objBatch.LotList(1).LotID & "," & objBatch.LotList(2).LotID
        sPortID = "Port=" & objBatch.LotList(1).Port & "," & objBatch.LotList(2).Port
    Else
        sLotID = objBatch.LotList(1).LotID
        sPortID = "Port=" & objBatch.LotList(1).Port
    End If

   
    ' R-的RECIPE 將不做Check
    iIndex = InStr(1, objBatch.LotList(1).Recipe1, "R-")
    If iIndex <> 0 Then
        Call Parent.LogTrace(0, MODULE_ID & " : 此 Recipe 為 R-開始不Check")
        Call Parent.TaskComplete(TaskID, "")
        Exit Sub
    End If

    'Send query PPBody
    With objNewVfei
        .CommandID = "CHECK_BODY_FORAPC"
        .MachineID = Parent.GetPuObject.EqDrvList(1)
        .MessageType = "C"
        If objBatch.LotList.Count = 2 Then
            If objBatch.LotList(1).Recipe1 = objBatch.LotList(2).Recipe1 Then
                .AddVar "PPID", "A", objLot.Recipe1
                .AddVar "LOT1_STEP", "A", objBatch.LotList(1).Step
                .AddVar "LOT2_STEP", "A", objBatch.LotList(2).Step
            Else
                Call AlarmReport(Parent, Me, "013", "The recipe of two Lots is not the same", "Lot=" & sLotID)
                Call Parent.LogTrace(0, "Cassette1's recipe is not equal to Cassette2's recipe")
            End If
        Else
            .AddVar "PPID", "A", objBatch.LotList(1).Recipe1
            .AddVar "LOT1_STEP", "A", objBatch.LotList(1).Step
            .AddVar "LOT2_STEP", "A", ""
        End If
        .AddVar "RUNMODE", "A", objBatch.RunMode
        
    End With
   
    Call Parent.SendVfei(objNewVfei, Me)
   
  
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
    Dim i As Integer, j As Integer
    Dim sAPC_Result As String, ErrorText As String
    Dim compList As Object, dataList As Object

    If objHostTrans.TransactionName <> "STARTCONTROLJOB" Then Exit Sub

    If objHostTrans.ErrorCode <> 0 Then
        Call AlarmReport(Parent, Me, objHostTrans.ErrorCode, "IMPRCMS(APC) Check Failed", objHostTrans.ErrorText)
        Call Parent.LogEvent("IMPRCMS FAILED", "", "")
        Exit Sub
    End If

    Set compList = objHostTrans.Secondary.Item("ComponentNumberList")
    For i = 1 To compList.Count
        With compList.Item(i)
            If UCase(.Item("ComponentID").Value) = "CHECK_RESULT" Then
                Set dataList = .Item("[DataNumberList]")
                For j = 1 To dataList.Count
                    If dataList.Item(j).Item(1).Value = "RESULT" Then
                        sAPC_Result = dataList.Item(j).Item(2).Value
                        Exit For
                    End If
                Next
            End If
        End With
    Next

    If sAPC_Result = "PASS" Then
        Call Parent.LogEvent("IMPRCMS PASS", "", "")
        Call Parent.TaskComplete(TaskID, "")
    Else
        Call Parent.LogEvent("IMPRCMS FAILED", "", "")
        Call AlarmReport(Parent, Me, "", "IMPRCMS(APC) Check Failed", "")
    End If
    
    Set objBatch = Nothing
    Set objLot = Nothing
    Set objPort = Nothing
    Set objSmif = Nothing
    Set objEqDrv = Nothing

    Exit Sub
ErrorHandler:
    Call Parent.LogTrace(0, FormatErrMsg(MODULE_ID & ".ReceivedHostMsg", Err, Err.Description))
End Sub

Private Sub TaskAncestor_ReceivedVfei(objVfei As vfei.Message22)
    Dim objLot As Lot
    Dim sLotID As String
    Dim sPortID As String
    
    On Error GoTo ErrorHandler
    
    If objVfei.CommandID <> "CHECK_BODY_FORAPC" Then Exit Sub
    
    If objBatch.LotList.Count > 1 Then
       sLotID = objBatch.LotList(1).LotID & "," & objBatch.LotList(2).LotID
       sPortID = "Port=" & objBatch.LotList(1).Port & "," & objBatch.LotList(2).Port
    Else
       sLotID = objBatch.LotList(1).LotID
       sPortID = "Port=" & objBatch.LotList(1).Port
    End If
    
    If objVfei.ErrorCode <> 0 Then
        For Each objLot In objBatch.LotList
            objLot.Status = LOT_MISPPBODY
            Call SendLotEventReport(Parent, Me, objLot)
            Call Parent.LogEvent("Check Recipe Error", sLotID, objVfei.ErrorText)
        Next objLot
        Call AlarmReport(Parent, Me, 1, Translate(Parent, "檢查 Recipe 內容錯誤:", "Check Recipe body error."), objVfei.ErrorText)
    Else
        Call Parent.LogTrace(0, "Recipe : " & objBatch.LotList(1).Recipe1 & " , IMPRCMS COMPARE : " & objVfei.ErrorText)
        Call Send_StartControlJob(objVfei.ErrorText)
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

Private Sub Send_StartControlJob(APCResult As String)
    Dim objHostTrans As HostTransaction
    Dim arrRecords As Variant
    Dim idx As Integer
    Dim key As String, val As String, keyType As String
    Dim lparen As Long, rparen As Long, colon As Long
    On Error GoTo ErrorHandler

    Set objHostTrans = Parent.NewHostTransaction(Me, "STARTCONTROLJOB")
    With objHostTrans.Primary
        .Item("ToolName") = Parent.PUID
        .Item("ApcActionType") = "STARTCONTROLJOB"
        .Item("BatchID") = objBatch.BatchID
    
        With .Item("LotNumberList").Add
                        
            .Item("LotID") = objBatch.LotList(1).Cassette
                                                
            'For assign context_matching
            With .Item("ItemNumberList").Add
                .Item("ItemName") = "PRODUCT"
                .Item("ItemValue") = objBatch.LotList(1).Product
            End With
            With .Item("ItemNumberList").Add
                .Item("ItemName") = "TOOLNAME"
                .Item("ItemValue") = Parent.PUID
            End With
            With .Item("ItemNumberList").Add
                .Item("ItemName") = "PPID"
                .Item("ItemValue") = objBatch.LotList(1).Recipe1
            End With
            With .Item("ItemNumberList").Add
                .Item("ItemName") = "PROCESS"
                .Item("ItemValue") = objBatch.LotList(1).PROCESS
            End With
            With .Item("ItemNumberList").Add
                .Item("ItemName") = "ROUTE"
                .Item("ItemValue") = objBatch.LotList(1).Route
            End With
            With .Item("ItemNumberList").Add
                .Item("ItemName") = "STEP"
                .Item("ItemValue") = objBatch.LotList(1).Step
            End With
            
            '8F APC Server Input
            '------------------------------------------------
            arrRecords = APCSplit(APCResult, ";")
            For idx = LBound(arrRecords) To UBound(arrRecords)
                lparen = InStr(arrRecords(idx), "(")
                rparen = InStr(arrRecords(idx), ")")
                colon = InStr(arrRecords(idx), ":")
                If lparen > 0 And rparen > lparen And colon > lparen And colon < rparen Then
                    key = Trim(Left(arrRecords(idx), lparen - 1))
                    keyType = Mid(arrRecords(idx), lparen + 1, colon - lparen - 1)
                    val = Mid(arrRecords(idx), colon + 1, rparen - colon - 1)
                    If UCase$(keyType) = "EAP" Then
                        With .Item("ItemNumberList").Add
                            .Item("ItemName") = key & "_EAP"
                            .Item("ItemValue") = val
                        End With
                    ElseIf UCase$(keyType) = "BODY" Then
                        With .Item("ItemNumberList").Add
                            .Item("ItemName") = key
                            .Item("ItemValue") = val
                        End With
                    End If
                End If
            Next idx
            '------------------------------------------------

        End With
    End With

    objHostTrans.Send
    
    Set objHostTrans = Nothing
    
    Exit Sub
                    
ErrorHandler:
    Call Parent.LogTrace(0, FormatErrMsg(MODULE_ID & "Send_START_CONTROL_JOB", Err.Number, "[" & Erl & "]" & Err.Description & Erl))
End Sub
Function APCSplit(ByVal strInput As String, ByVal strSep As String) As Variant
    Dim result() As String
    Dim sepCount As Long, pos As Long, startPos As Long, n As Long, arrLen As Long

    ' 防呆：分隔符不可為空
    If strSep = "" Then
        APCSplit = Array(strInput)
        Exit Function
    End If

    ' ? 預先計算分隔符數量
    pos = 1
    Do While InStr(pos, strInput, strSep) > 0
        sepCount = sepCount + 1
        pos = InStr(pos, strInput, strSep) + Len(strSep)
    Loop

    arrLen = sepCount + 1
    ReDim result(arrLen - 1)

    ' ? 再分割內容
    startPos = 1
    For n = 0 To arrLen - 2
        pos = InStr(startPos, strInput, strSep)
        result(n) = Mid(strInput, startPos, pos - startPos)
        startPos = pos + Len(strSep)
    Next n
    result(arrLen - 1) = Mid(strInput, startPos)

    APCSplit = result
End Function
