VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MW_EQ_PDS2"
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
Dim vShowAlarm As Boolean

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
If objHostTrans.TransactionName <> "EQP_DATA" Then Exit Sub
If objHostTrans.ErrorCode <> 0 Then
    If vShowAlarm = True Then   '因PDS中每個參數均會送一次Txn到AS，若AS發生異常，為避免產生大量Alarm到UI，因此限制一次的PDS最多只送一次Alarm到UI
        Call AlarmReport(Parent, Me, objHostTrans.ErrorCode, Translate(Parent, "上傳工單資料到自動化伺服器時發生異常.", "Abnormality occurred when uploading work form data to Automation Server"), objHostTrans.ErrorText)
        vShowAlarm = False
    End If
    Call Parent.LogTrace(0, "Upload EQP Data to AS error")
End If

'--- add for check if Out of Spec for V0504 (IRR067)
'Secondary=//EQP_DATA <%Status> <%TID> <#Parameters>
'        [ <Parameter> <Out_of_Spec_Flag> ]
Dim objSourceItem As HostItem
For Each objSourceItem In objHostTrans.Secondary.Item("Parameters")
    If objSourceItem.Item("Out_of_Spec_Flag").Value <> "0" Then
        Call Parent.LogTrace(0, "EQP_DATA: " & objSourceItem.Item("Parameter").Value & " Out of Spec")
    End If
Next
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
Dim MonitorType As String


On Error GoTo ErrorHandler

'Yen add for MI-4 dual function
For Each objLot In objBatch.LotList
If objVfei.Var("CONTROLLER").Value <> objLot.LotID Then Exit Sub
Next
'Yen

On Error Resume Next
TempStr = objVfei.Var("LOT_ID").Value
If Err.Number = 0 Then
    For Each objLot In objBatch.LotList
        If UCase(objLot.LotID) = UCase(Trim(TempStr)) Then
            isFind = True
            Exit For
        End If
    Next
    
Else
    Err.Clear
    TempStr = objVfei.Var("CASSETTE").Value
    If Err.Number = 0 Then
        For Each objLot In objBatch.LotList
            If UCase(objLot.Cassette) = UCase(Trim(TempStr)) Then
                isFind = True
                Exit For
            End If
        Next
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
            Err.Clear
            TempStr = objVfei.Var("BATCH_ID").Value
            If Err.Number = 0 Then
                If TempStr = "OK" Then
                    isFind = True
                    isBatch = True
                End If
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
    'On Error GoTo ErrorHandler
    If sSlot = "" Then sSlot = "1"
    If sSite = "" Then sSite = "1"
    If sChamber = "" Then sChamber = "1"
    If sRecipeStep = "" Then sRecipeStep = "1"
    If sRun = "" Then sRun = "1"
    
    MonitorType = objBatch.LotList(1).Attributes("_UI_MONITORTYPE").Value
    If objBatch.RunMode = "PR" Then
        Call FillEqpData(objLot, objVfei, CInt(sSlot), CInt(sSite), sChamber, sRecipeStep, sRun)
        Call FillMesData(objLot, objVfei, CInt(sSlot), CInt(sSite), sChamber, sRecipeStep, sRun)
    ElseIf objBatch.RunMode = "TEST" Or objBatch.RunMode = "DMMTEST" Then
        Call FillMonitorData(objBatch, objVfei, CInt(sSlot), CInt(sSite), sChamber, sRecipeStep, sRun)
    ElseIf objBatch.RunMode = "DUMMYTEST" And MonitorType = "PMTest" Then
        Call FillMonitorData(objBatch, objVfei, CInt(sSlot), CInt(sSite), sChamber, sRecipeStep, sRun)
    ElseIf objBatch.RunMode = "CFS" Then
       Call FillEqpData(objLot, objVfei, CInt(sSlot), CInt(sSite), sChamber, sRecipeStep, sRun)
    ElseIf objBatch.RunMode = "NS" Or objBatch.RunMode = "DMMNS" Then
       Call FillEqpData(objLot, objVfei, CInt(sSlot), CInt(sSite), sChamber, sRecipeStep, sRun)
    End If

End If
Exit Sub

ErrorHandler:
Call Parent.LogTrace(0, FormatErrMsg(MODULE_ID & ".ReceivedVfei", Err, Err.Description))
End Sub

Private Function ReplaceChar(orgStr As String, oldChar As String, newChar As String) As String
ReplaceChar = orgStr
If Len(oldChar) <> 1 Or Len(newChar) <> 1 Then Exit Function
Do Until InStr(1, ReplaceChar, oldChar) = 0
    ReplaceChar = Mid(ReplaceChar, 1, InStr(1, ReplaceChar, oldChar) - 1) & newChar & Mid(ReplaceChar, InStr(1, ReplaceChar, oldChar) + 1)
Loop
End Function

Private Property Let TaskAncestor_TaskID(ByVal RHS As String)
TaskID = RHS
End Property

Private Property Get TaskAncestor_TaskID() As String
TaskAncestor_TaskID = TaskID
End Property
          
Private Sub FillEqpData(objLot As Lot, objVfei As vfei.Message22, Optional SlotNo As Integer, Optional SiteNo As Integer, Optional sChamber As String, Optional sRecipeStep As String, Optional sRun As String)
Dim objEqpData As EQP_DataItem
Dim objTriplet As Triplet
Dim i As Integer
Dim isNewSlot As Boolean
Dim vSiteCount As Integer
Dim k As Integer
'Dim vEQParameter As String
'Dim vEQParameterCount As Integer

On Error Resume Next
vSiteCount = 1
isNewSlot = True
vShowAlarm = True
For Each objEqpData In objLot.EQPData
        Set objTriplet = objVfei.Var("PDS").Var(objEqpData.EQParameter)
        If Err.Number = 0 Then
            For i = 1 To objEqpData.DataList.Count
                If objEqpData.DataList.Item(i).Slot = SlotNo Then
                    isNewSlot = False
                    vSiteCount = vSiteCount + 1
                End If
            Next i
            If isNewSlot = True Then
                With objEqpData.DataList.Add(SlotNo, SiteNo)
                    .Value = objTriplet.Value
                    .Chamber = sChamber
                    .RecipeStep = sRecipeStep
                    .Run = sRun
                End With
                If objEqpData.SetupFlag <> "F" Then
                    SendEQPDataToAS objEqpData.Parameter, objTriplet.Value, objEqpData.DataType, objLot, SlotNo, SiteNo, sChamber, sRecipeStep, sRun
                End If
            Else
                With objEqpData.DataList.Add(SlotNo, vSiteCount)
                    .Value = objTriplet.Value
                    .Chamber = sChamber
                    .RecipeStep = sRecipeStep
                    .Run = sRun
                End With
                If objEqpData.SetupFlag <> "F" Then
                    SendEQPDataToAS objEqpData.Parameter, objTriplet.Value, objEqpData.DataType, objLot, SlotNo, vSiteCount, sChamber, sRecipeStep, sRun
                End If
                vSiteCount = 1
            End If
            Call Parent.LogTrace(5, "Add a value into EqpData Name=" & objTriplet.Name & " Value=" & objTriplet.Value)
        Else
            Err.Clear
        End If
Next
End Sub

Private Sub SendEQPDataToAS(Parameter As String, ParameterValue As String, DataType As String, objLot As Lot, sLotID As Integer, SiteID As Integer, Chamber As String, RecipeStep As String, Run As String)
Dim objHostTrans As HostTransaction

On Error GoTo ErrorHandler
Set objHostTrans = Parent.NewHostTransaction(Me, "EQP_DATA")
objHostTrans.Primary.Item("processPU") = Parent.PUID
objHostTrans.Primary.Item("P_Recipe") = objLot.Recipe1
objHostTrans.Primary.Item("Operator") = Parent.GetBatchList(objLot.BatchID).Operator
objHostTrans.Primary.Item("Batch") = objLot.BatchID
objHostTrans.Primary.Item("Lot") = objLot.LotID

With objHostTrans.Primary.Item("Data").Add
    .Item("SlotID") = CStr(SiteID) & Format(sLotID, "00")
    .Item("Parameter") = Parameter
    If DataType = "0" Then 'DataType is Numberic
        .Item("Numeric_value") = ParameterValue
        .Item("ASCII_value") = "NULL"
    Else
        .Item("Numeric_value") = "NULL"
        .Item("ASCII_value") = ParameterValue
    End If
    .Item("Chamber") = Chamber
    .Item("recipe_step") = RecipeStep
    .Item("run") = Run
End With

If objHostTrans.Primary.Item("Data").Count > 0 Then
    objHostTrans.Send
End If
Set objHostTrans = Nothing
Exit Sub

ErrorHandler:
Call Parent.LogTrace(0, FormatErrMsg(MODULE_ID & ".Execute", Err, Err.Description))
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

