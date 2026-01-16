VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MW_PDS_Trans"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements ts_ECAncestor.TaskAncestor

Private Const MODULE_ID = "MW_PDS_Trans"

Dim Parent As tsEC.clsRunManager
'Dim Parent As Object
Dim TaskID As String
Dim objPU As ProcessingUnit
Dim objBatch As Batch
Dim sReceiv_vfei As Boolean

Private Sub TaskAncestor_Execute(Parameter As Variant)
    On Error GoTo ErrorHandler
    
    'Set objPU = Parameter
    Set objBatch = Parameter
   
    Call Parent.SubscribeVfeiMsg(Me, "", "PDS")
    
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
    Dim sVFEI_STRING As String
    Dim sTEMP_STRING_0 As String
    Dim sTEMP_STRING_1 As String
    Dim sTEMP_STRING_2 As String
    Dim sTEMP_INT1 As Integer
    Dim sOperator As String
    Dim objvfei_send As Message22
    Dim i As Integer
    Dim Recipe As String
    
    'Add forr Catch Objlot Info from VFEI
    
    Dim objLot As Lot
    Dim sTaskListName As String
    Dim stemp As String
    Dim TempStr As String
    Dim isFind As Boolean
    Dim isBatch As Boolean
    
    Dim blnIsSlot As Boolean
    
    Dim ACCEL_I_18_Value As String
    Dim ACCEL_I_43_Value As String
    
    On Error Resume Next
    'Yen add for MI-4 dual function
    For Each objLot In objBatch.LotList
        If objVfei.Var("CONTROLLER").Value <> objLot.LotID Then Exit Sub
    Next
    'Yen
    
    TempStr = objVfei.Var("LOT_ID").Value
    If Err.Number = 0 Then
        For Each objLot In objBatch.LotList
            If InStr(TempStr, objLot.LotID) <> 0 Then
                'If objLot.LotID = TempStr Then
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
    
    If isBatch Then
        Set objLot = objBatch.LotList(1)
    End If
    'End of Catch Objlot Info
    
    On Error Resume Next
    'For Check "Lot_ID"
    If InStr(objLot.LotID, objBatch.LotList(1).LotID) = 0 Then Exit Sub
    
    i = 1
    On Error Resume Next
    If sReceiv_vfei = True Then Exit Sub
    
    Set objvfei_send = New Message22
    
    sReceiv_vfei = True
    
    'added by Harlem
    Dim iSlot As Integer
    
    blnIsSlot = True
    iSlot = CInt(objVfei.Var("SLOT"))
    
    If Err.Number <> 0 Then
        Err.Clear
        iSlot = 1
        blnIsSlot = False
    End If
    
    Call Parent.LogTrace(0, MODULE_ID & ".ReceivedVefei ==> Slot = " & iSlot)
    
    'Senf VFEI
    objvfei_send.CommandID = "PDS"
    objvfei_send.MachineID = Parent.GetPuObject.EqDrvList(1).EqDrvID
    objvfei_send.MessageType = "E"
    objvfei_send.AddVar "PDS", "L", ""
    objvfei_send.AddVar "LOT_ID", "A", objLot.LotID
    objvfei_send.AddVar "PORT", "A", objLot.Port
    objvfei_send.AddVar "SLOT", "U4", iSlot
    objvfei_send.AddVar "BATCH_ID", "A", objVfei.Var("BATCH_ID")
    objvfei_send.AddVar "RECIPE_STEP", "A", objVfei.Var("RECIPE_STEP")
    objvfei_send.AddVar "RUN", "A", objVfei.Var("RUN")
    
    'Yen add for MI-4 dual function
    objvfei_send.AddVar "CONTROLLER", "A", objLot.LotID
    'Yen
    
    On Error GoTo ErrorHandler
    Do While GetToken2(Parent.GetTaskAttribute(TaskID, ""), Chr(0), i * 2) <> ""
        'Initeial
        sTEMP_STRING_1 = ""
        sTEMP_STRING_2 = ""
        
        On Error GoTo ErrorHandler
        
        sVFEI_STRING = GetToken2(Parent.GetTaskAttribute(TaskID, ""), Chr(0), i * 2 - 1)
        sTEMP_STRING_0 = GetToken2(Parent.GetTaskAttribute(TaskID, ""), Chr(0), i * 2)
        
        'Separate Parameter
        If InStr(sTEMP_STRING_0, "+") <> 0 Then
            sTEMP_STRING_1 = Left(sTEMP_STRING_0, InStr(sTEMP_STRING_0, "+") - 1)
            sTEMP_STRING_2 = Right(sTEMP_STRING_0, Len(sTEMP_STRING_0) - InStr(sTEMP_STRING_0, "+"))
            sOperator = "+"
        End If
        If InStr(sTEMP_STRING_0, "-") <> 0 Then
            sTEMP_STRING_1 = Left(sTEMP_STRING_0, InStr(sTEMP_STRING_0, "-") - 1)
            sTEMP_STRING_2 = Right(sTEMP_STRING_0, Len(sTEMP_STRING_0) - InStr(sTEMP_STRING_0, "-"))
            sOperator = "-"
        End If
        If InStr(sTEMP_STRING_0, "*") <> 0 Then
            sTEMP_STRING_1 = Left(sTEMP_STRING_0, InStr(sTEMP_STRING_0, "*") - 1)
            sTEMP_STRING_2 = Right(sTEMP_STRING_0, Len(sTEMP_STRING_0) - InStr(sTEMP_STRING_0, "*"))
            sOperator = "*"
        End If
        If InStr(sTEMP_STRING_0, "/") <> 0 Then
            sTEMP_STRING_1 = Left(sTEMP_STRING_0, InStr(sTEMP_STRING_0, "/") - 1)
            sTEMP_STRING_2 = Right(sTEMP_STRING_0, Len(sTEMP_STRING_0) - InStr(sTEMP_STRING_0, "/"))
            sOperator = "/"
        End If
        
        'Transfer sTempString to Variable String
        sTEMP_STRING_1 = GetParameterValue(objVfei, sTEMP_STRING_1)
        sTEMP_STRING_2 = GetParameterValue(objVfei, sTEMP_STRING_2)
        If sTEMP_STRING_1 = "" Or sTEMP_STRING_2 = "" Then GoTo Lable_Exit
        
        If GetCalValue(sTEMP_STRING_1, sTEMP_STRING_2, sOperator) Then
            Select Case sOperator
                Case "+"
                    objvfei_send.Var("PDS").AddVar sVFEI_STRING, "A", CDbl(sTEMP_STRING_1) + CDbl(sTEMP_STRING_2)
                Case "-"
                    objvfei_send.Var("PDS").AddVar sVFEI_STRING, "A", CDbl(sTEMP_STRING_1) - CDbl(sTEMP_STRING_2)
                Case "*"
                    objvfei_send.Var("PDS").AddVar sVFEI_STRING, "A", CDbl(sTEMP_STRING_1) * CDbl(sTEMP_STRING_2)
                Case "/"
                    objvfei_send.Var("PDS").AddVar sVFEI_STRING, "A", CDbl(sTEMP_STRING_1) / CDbl(sTEMP_STRING_2)
            End Select
        End If
Lable_Exit:
        i = i + 1
    Loop
    
    'Call Sub AMU_Delta
    If InStr(Parent.GetTaskAttribute(TaskID, "AMU_DELTA"), "Y") <> 0 Then
        Call AMU_Delta(objBatch.LotList(1).Recipe1, objVfei, objvfei_send)
    End If
    
    'Call Sub Energy_Delta
    If InStr(Parent.GetTaskAttribute(TaskID, "ENERGY_DELTA"), "Y") <> 0 Then
        Call Parent.LogTrace(0, MODULE_ID & ".ReceivedVefei ==> Calculate ENERGY_DELTA")
        Call Energy_Delta(objBatch.LotList(1).Recipe1, objVfei, objvfei_send)
    End If
    
    'Call Sub Recipe_PDS
    If InStr(Parent.GetTaskAttribute(TaskID, "RECIPE_PDS"), "Y") Then
        Call RECIPE_PDS(objBatch.LotList(1).Recipe1, objVfei, objvfei_send)
    End If
    
    'Call Sub ASCANS_Delta add by ShangChun #20070606
    'add blnIsSlot by Harlem 20100513
    If InStr(Parent.GetTaskAttribute(TaskID, "ASCANS_DELTA"), "Y") Then
        If blnIsSlot Then
            Call ASCANS_Delta(objBatch.LotList(1), objVfei, objvfei_send)
        End If
    End If
    ' ShangChun #20070606
    
    ' Slot-specific ACCEL_I handling with safe read
    If iSlot = 18 Then
        ACCEL_I_18_Value = SafePdsVar(objVfei, "ACCEL_I")
        If ACCEL_I_18_Value <> "" Then
            AddLotAttribute "ACCEL_I_18", ACCEL_I_18_Value
            Call Parent.LogTrace(0, MODULE_ID & ".ReceivedVefei ==> AddLotAttribute ACCCEL_I For Slot18 ")
        Else
            Call Parent.LogTrace(0, MODULE_ID & ".ReceivedVefei ==> ACCEL_I missing on slot 18, skip store")
        End If
    ElseIf iSlot = 43 Then
        ACCEL_I_43_Value = SafePdsVar(objVfei, "ACCEL_I")
        If ACCEL_I_43_Value <> "" Then
            AddLotAttribute "ACCEL_I_43", ACCEL_I_43_Value
            Call Parent.LogTrace(0, MODULE_ID & ".ReceivedVefei ==> AddLotAttribute ACCCEL_I For Slot43 ")
        Else
            Call Parent.LogTrace(0, MODULE_ID & ".ReceivedVefei ==> ACCEL_I missing on slot 43, skip store")
        End If
    ElseIf iSlot = 19 Then
        Call CalACCCEL_I(objBatch.LotList(1), objVfei, objvfei_send)
        Call Parent.LogTrace(0, MODULE_ID & ".ReceivedVefei ==> Calculate ACCCEL_I For Slot18 And Slot19")
    ElseIf iSlot = 44 Then
        Call CalACCCEL_I(objBatch.LotList(1), objVfei, objvfei_send)
        Call Parent.LogTrace(0, MODULE_ID & ".ReceivedVefei ==> Calculate ACCCEL_I For Slot43 And Slot44")
    End If
    
    Call Parent.SendVfei(objvfei_send, Me)
    GoTo Cleanup

ErrorHandler:
    Call Parent.LogTrace(0, FormatErrMsg(MODULE_ID & ".ReceivedVfei", Err, Err.Description))
Cleanup:
    sReceiv_vfei = False
    Exit Sub
End Sub

Private Property Let TaskAncestor_TaskID(ByVal RHS As String)
    TaskID = RHS
End Property

Private Property Get TaskAncestor_TaskID() As String
    TaskAncestor_TaskID = TaskID
End Property
'-----------------------------------20070606 add by ShangChun
'DF Want to set SPC Monitor of MI Scan no delta for VTN IMP layer recipe
' 1. VTN IMP Recipe (BF50K/55K)
' 2. 每個Lot收 slot 9/10/18/19總共4片wafer之scan no
' 3. 針對 (Slot 10- Slot9) 與 (Slot 19- Slot18)之scan no delta收值並上SPC monitor
' 4. 目前scan no (ASCANS)已於UEDA有收值
'-------------------------------------------------------------------------------
Private Sub ASCANS_Delta(objLot As Lot, objVfei As vfei.Message22, objvfei_send)
    Dim objEqpData  As EQP_DataItem
    Dim i   As Integer
    If Not (objVfei.Var("SLOT") = 10 Or objVfei.Var("SLOT") = 19 Or _
            objVfei.Var("SLOT") = 35 Or objVfei.Var("SLOT") = 44) Then Exit Sub
    
    Dim vASCANS As Integer
    Dim ASCANS_DELTA_VALUE As String
    
    On Error GoTo ErrorHandler
    
    If objVfei.Var("SLOT") = 10 Then
        For Each objEqpData In objLot.EQPData
            If objEqpData.EQParameter = "ASCANS" Then
                For i = 1 To objEqpData.dataList.Count
                    If objEqpData.dataList.Item(i).sLot = 9 Then
                        vASCANS = objEqpData.dataList.Item(i).Value
                        ASCANS_DELTA_VALUE = objVfei.Var("PDS").Var("ASCANS") - vASCANS
                        objvfei_send.Var("PDS").AddVar "ASCANS_DELTA", "A", CDbl(ASCANS_DELTA_VALUE)
                        Exit For
                    End If
                Next i
                Exit For
            End If
        Next objEqpData
    ElseIf objVfei.Var("SLOT") = 19 Then
        For Each objEqpData In objLot.EQPData
            If objEqpData.EQParameter = "ASCANS" Then
                For i = 1 To objEqpData.dataList.Count
                    If objEqpData.dataList.Item(i).sLot = 18 Then
                        vASCANS = objEqpData.dataList.Item(i).Value
                        ASCANS_DELTA_VALUE = objVfei.Var("PDS").Var("ASCANS") - vASCANS
                        objvfei_send.Var("PDS").AddVar "ASCANS_DELTA", "A", CDbl(ASCANS_DELTA_VALUE)
                        Exit For
                    End If
                Next i
                Exit For
            End If
        Next objEqpData
    ElseIf objVfei.Var("SLOT") = 35 Then
        For Each objEqpData In objLot.EQPData
            If objEqpData.EQParameter = "ASCANS" Then
                For i = 1 To objEqpData.dataList.Count
                    If objEqpData.dataList.Item(i).sLot = 34 Then
                        vASCANS = objEqpData.dataList.Item(i).Value
                        ASCANS_DELTA_VALUE = objVfei.Var("PDS").Var("ASCANS") - vASCANS
                        objvfei_send.Var("PDS").AddVar "ASCANS_DELTA", "A", CDbl(ASCANS_DELTA_VALUE)
                        Exit For
                    End If
                Next i
                Exit For
            End If
        Next objEqpData
    ElseIf objVfei.Var("SLOT") = 44 Then
        For Each objEqpData In objLot.EQPData
            If objEqpData.EQParameter = "ASCANS" Then
                For i = 1 To objEqpData.dataList.Count
                    If objEqpData.dataList.Item(i).sLot = 43 Then
                        vASCANS = objEqpData.dataList.Item(i).Value
                        ASCANS_DELTA_VALUE = objVfei.Var("PDS").Var("ASCANS") - vASCANS
                        objvfei_send.Var("PDS").AddVar "ASCANS_DELTA", "A", CDbl(ASCANS_DELTA_VALUE)
                        Exit For
                    End If
                Next i
                Exit For
            End If
        Next objEqpData
    End If
ErrorHandler:
    Call Parent.LogTrace(0, FormatErrMsg(MODULE_ID & ".ASCANS_Delta", Err, Err.Description))
End Sub

Private Sub AMU_Delta(sRecipe As String, objVfei As vfei.Message22, objvfei_send)
    Dim Recipe_Gas As String
    Dim AMU_Delta_Value As String
    Dim i As Integer
    i = 1

    Do While Not IsNumeric(Mid(sRecipe, i, 1))
        i = i + 1
    Loop
    Recipe_Gas = Left(sRecipe, i - 1)

    'Recipe_Gas = Left(sRecipe, 2)

    On Error GoTo AMU_Delta_Error_Lable
    
    AMU_Delta_Value = Parent.GetTaskAttribute(TaskID, "AMU_" + Recipe_Gas) - objVfei.Var("PDS").Var("AMU")
    
    objvfei_send.Var("PDS").AddVar "AMU_DELTA", "A", CDbl(AMU_Delta_Value)
    
AMU_Delta_Error_Lable:
End Sub

Private Sub Energy_Delta(sRecipe As String, objVfei As vfei.Message22, objvfei_send)
    Dim Energy_Value As String
    Dim Energy_Delta_Value As String
    Dim i As Integer
    
    On Error GoTo Energy_Delta_Error_Lable
        
    i = 1
    Do While (IsNumeric(Right(Left(sRecipe, InStr(sRecipe, "K") - 1), i)))
        i = i + 1
    Loop
        
    Energy_Value = Right(Left(sRecipe, InStr(sRecipe, "K") - 1), i - 1)
    
    If InStr(Parent.GetTaskAttribute(TaskID, "ENERGY_DELTA"), "K") <> 0 Then
        Energy_Delta_Value = Energy_Value - objVfei.Var("PDS").Var("ENERGY") / 1000
    Else
        Energy_Delta_Value = Energy_Value - objVfei.Var("PDS").Var("ENERGY")
    End If
    
    objvfei_send.Var("PDS").AddVar "ENERGY_DELTA", "A", CDbl(Energy_Delta_Value)
    
    Call Parent.LogTrace(0, MODULE_ID & ".Energy_Delta,  ENERGY_DELTA = " & CDbl(Energy_Delta_Value))
    
    Exit Sub
Energy_Delta_Error_Lable:

    Call Parent.LogTrace(0, MODULE_ID & ".Energy_Delta,  Error = " & Err.Description)

End Sub

Private Sub RECIPE_PDS(sRecipe As String, objVfei As vfei.Message22, objvfei_send)
    Dim PDS_NAME_VALUE As String
    Dim PDS_NAME As String
    Dim PDS_VALUE As String
    Dim i As Integer

    On Error GoTo RECIPE_PDS_Error_Lable
    
    PDS_NAME_VALUE = Parent.GetTaskAttribute(TaskID, "RECIPE_" + sRecipe)
    PDS_NAME = Left(PDS_NAME_VALUE, InStr(PDS_NAME_VALUE, "#") - 1)
    PDS_VALUE = Right(PDS_NAME_VALUE, Len(PDS_NAME_VALUE) - InStr(PDS_NAME_VALUE, "#"))

    objvfei_send.Var("PDS").AddVar PDS_NAME, "A", PDS_VALUE

RECIPE_PDS_Error_Lable:
End Sub

Private Function GetParameterValue(objNewVfei As vfei.Message22, sParameter As String) As String
    On Error Resume Next
    
    GetParameterValue = objNewVfei.Var("PDS").Var(sParameter)
    If Err.Number <> 0 Then
        GetParameterValue = ""
    Else
        'IF sParameter is "Time" Format (00:00:00) then Transfer to Second
        If InStr(sParameter, ":") Then
            GetParameterValue = CDbl(GetToken2(sParameter, ":", 1)) * 3600 + CDbl(GetToken2(sParameter, ":", 2)) * 60 + CDbl(GetToken2(sParameter, ":", 3))
            If Err.Number <> 0 Then
                GetParameterValue = ""
            End If
        End If
    End If
    
    Call Parent.LogTrace(0, MODULE_ID & ".GetParameterValue," & sParameter & " = " & GetParameterValue)
    
    Err.Clear
    Exit Function
End Function

Private Function GetCalValue(sParameter1 As String, sParameter2 As String, tOperator As String) As Boolean
    Dim sTempValue As Double
    On Error Resume Next
    
    Select Case tOperator
        Case "+"
            sTempValue = CDbl(sParameter1) + CDbl(sParameter2)
        Case "-"
            sTempValue = CDbl(sParameter1) - CDbl(sParameter2)
        Case "*"
            sTempValue = CDbl(sParameter1) * CDbl(sParameter2)
        Case "/"
            sTempValue = CDbl(sParameter1) / CDbl(sParameter2)
    End Select
    
    If Err.Number <> 0 Then
        GetCalValue = False
    Else
        GetCalValue = True
    End If
    
    Call Parent.LogTrace(0, MODULE_ID & ".GetCalValue,CalValue = " & sTempValue)
    
    Err.Clear
    Exit Function
End Function

'-----------------------------------20220704 add by YL
'DF Want to set SPC Monitor of MI Slot18 and Slot19 ACCEL_I dif
'-------------------------------------------------------------------------------
Private Sub CalACCCEL_I(objLot As Lot, objVfei As vfei.Message22, objvfei_send)
    Dim ACCEL_I_18_Value As String
    Dim ACCEL_I_D1819_Value As String
    Dim ACCEL_I_43_Value As String
    Dim ACCEL_I_D4344_Value As String
    Dim currAccel As String

    On Error GoTo CalACCCEL_I_Error_Lable
    
    currAccel = SafePdsVar(objVfei, "ACCEL_I")
    
    If objVfei.Var("SLOT") = "19" Then
        ACCEL_I_18_Value = GetLotAttributeValue("ACCEL_I_18")
        If ACCEL_I_18_Value = "" Or currAccel = "" Then Exit Sub
        
        ACCEL_I_D1819_Value = CStr(CDbl(currAccel) - CDbl(ACCEL_I_18_Value))
        objvfei_send.Var("PDS").AddVar "ACCEL_I_D18_19", "A", CDbl(ACCEL_I_D1819_Value)
        AddLotAttribute "ACCEL_I_D1819", ACCEL_I_D1819_Value
    ElseIf objVfei.Var("SLOT") = "44" Then
        ACCEL_I_43_Value = GetLotAttributeValue("ACCEL_I_43")
        If ACCEL_I_43_Value = "" Or currAccel = "" Then Exit Sub
        
        ACCEL_I_D4344_Value = CStr(CDbl(currAccel) - CDbl(ACCEL_I_43_Value))
        objvfei_send.Var("PDS").AddVar "ACCEL_I_D18_19", "A", CDbl(ACCEL_I_D4344_Value)
        AddLotAttribute "ACCEL_I_D1819", ACCEL_I_D4344_Value
    End If

CalACCCEL_I_Error_Lable:
End Sub

' Safe getter for PDS variable; missing ->
Private Function SafePdsVar(objMsg As vfei.Message22, varName As String) As String
    On Error Resume Next
    SafePdsVar = objMsg.Var("PDS").Var(varName)
    If Err.Number <> 0 Then
        SafePdsVar = ""
        Err.Clear
    End If
End Function

Private Sub AddLotAttribute(iName As String, iValue As String)
    On Error Resume Next
    objBatch.LotList(1).Attributes.Add(iName).Value = iValue
    If Err.Number <> 0 Then '代表屬性已存在
        objBatch.LotList(1).Attributes.Item(iName).Value = iValue
    End If
End Sub

Private Function GetLotAttributeValue(iName As String) As String
    On Error Resume Next
    GetLotAttributeValue = objBatch.LotList(1).Attributes(iName).Value
    If Err.Number <> 0 Then
        GetLotAttributeValue = ""
        Err.Clear
    End If
End Function

