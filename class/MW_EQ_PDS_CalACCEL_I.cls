VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MW_EQ_PDS_CalACCEL_I"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements ts_ECAncestor.TaskAncestor

Private Const MODULE_ID = "MW_EQ_PDS_CalACCEL_I"

Dim parent As tsEC.clsRunManager
'Dim Parent As Object
Dim TaskID As String
Dim objPU As ProcessingUnit
Dim objBatch As Batch
Dim sReceiv_vfei As Boolean
Private Sub TaskAncestor_Execute(Parameter As Variant)
    On Error GoTo ErrorHandler
    
    Set objBatch = Parameter
   
    Call parent.SubscribeVfeiMsg(Me, "", "PDS")
    
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
    On Error GoTo ErrorHandler

    Exit Sub
ErrorHandler:
    Call parent.LogTrace(0, FormatErrMsg(MODULE_ID & ".ReceivedHostMsg", Err, Err.Description))
End Sub

Private Sub TaskAncestor_ReceivedVfei(objVfei As vfei.Message22)

    Dim objvfei_send As Message22
    Dim i As Integer
    Dim Recipe As String
    Dim ACCEL_I_18_Value As String
    Dim ACCEL_I_19_Value As String
    Dim ACCEL_I_43_Value As String
    Dim ACCEL_I_44_Value As String
    
    Dim objLot As Lot
    Dim TempStr As String
    Dim isFind As Boolean
    Dim isBatch As Boolean
    
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
                    Call parent.LogTrace(0, "Received a Event " & objVfei.CommandID & "! But can't find associate Object")
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
     
    
    On Error Resume Next
    If sReceiv_vfei = True Then Exit Sub
    
    
    Set objvfei_send = New Message22
    
    sReceiv_vfei = True
    
    
    Dim iSlot As String
    
    iSlot = CInt(objVfei.Var("SLOT"))
    If Err.Number <> 0 Then Exit Sub
    
    If iSlot = "18" Then
        ACCEL_I_18_Value = objVfei.Var("PDS").Var("ACCEL_I")
        If Err.Number <> 0 Then
            sReceiv_vfei = False
            Exit Sub
        End If
        AddLotAttribute "ACCEL_I_18", ACCEL_I_18_Value
        Call parent.LogTrace(0, MODULE_ID & ".ReceivedVefei ==> AddLotAttribute ACCCEL_I For Slot18 ")
    ElseIf iSlot = "43" Then
        ACCEL_I_43_Value = objVfei.Var("PDS").Var("ACCEL_I")
        If Err.Number <> 0 Then
            sReceiv_vfei = False
            Exit Sub
        End If
        AddLotAttribute "ACCEL_I_43", ACCEL_I_43_Value
        Call parent.LogTrace(0, MODULE_ID & ".ReceivedVefei ==> AddLotAttribute ACCCEL_I For Slot43 ")
    ElseIf iSlot = "19" Then
        ACCEL_I_19_Value = objVfei.Var("PDS").Var("ACCEL_I")
        If Err.Number <> 0 Then
            sReceiv_vfei = False
            Exit Sub
        End If
        'Send VFEI
        objvfei_send.CommandID = "PDS"
        objvfei_send.MachineID = parent.GetPuObject.EqDrvList(1).EqDrvID
        objvfei_send.MessageType = "E"
        objvfei_send.AddVar "PDS", "L", ""
        objvfei_send.AddVar "LOT_ID", "A", objLot.LotID
        objvfei_send.AddVar "PORT", "A", objLot.Port
        objvfei_send.AddVar "SLOT", "U4", iSlot
        objvfei_send.AddVar "BATCH_ID", "A", objVfei.Var("BATCH_ID")
        objvfei_send.AddVar "RECIPE_STEP", "A", objVfei.Var("RECIPE_STEP")
        objvfei_send.AddVar "RUN", "A", objVfei.Var("RUN")
        objvfei_send.AddVar "CONTROLLER", "A", objLot.LotID
        objvfei_send.AddVar "CALCULATOR", "A", objLot.LotID
        
        Call CalACCCEL_I(objBatch.LotList(1), objVfei, objvfei_send)
        Call parent.LogTrace(0, MODULE_ID & ".ReceivedVefei ==> Calculate ACCCEL_I For Slot18 And Slot19")
        Call parent.SendVfei(objvfei_send, Me)
    ElseIf iSlot = "44" Then
        ACCEL_I_44_Value = objVfei.Var("PDS").Var("ACCEL_I")
        If Err.Number <> 0 Then
            sReceiv_vfei = False
            Exit Sub
        End If
        'Send VFEI
        objvfei_send.CommandID = "PDS"
        objvfei_send.MachineID = parent.GetPuObject.EqDrvList(1).EqDrvID
        objvfei_send.MessageType = "E"
        objvfei_send.AddVar "PDS", "L", ""
        objvfei_send.AddVar "LOT_ID", "A", objLot.LotID
        objvfei_send.AddVar "PORT", "A", objLot.Port
        objvfei_send.AddVar "SLOT", "U4", iSlot
        objvfei_send.AddVar "BATCH_ID", "A", objVfei.Var("BATCH_ID")
        objvfei_send.AddVar "RECIPE_STEP", "A", objVfei.Var("RECIPE_STEP")
        objvfei_send.AddVar "RUN", "A", objVfei.Var("RUN")
        objvfei_send.AddVar "CONTROLLER", "A", objLot.LotID
        objvfei_send.AddVar "CALCULATOR", "A", objLot.LotID
        
        Call CalACCCEL_I(objBatch.LotList(1), objVfei, objvfei_send)
        Call parent.LogTrace(0, MODULE_ID & ".ReceivedVefei ==> Calculate ACCCEL_I For Slot43 And Slot44")
        Call parent.SendVfei(objvfei_send, Me)
    End If
    
    
    
    '    Select Case iSlot
    '        Case "18"
    '            ACCEL_I_18_Value = objVfei.Var("PDS").Var("ACCEL_I")
    '            If Err.Number <> 0 Then
    '                bolContinue = False
    '                sReceiv_vfei = False
    '                Exit Sub
    '            End If
    '            AddLotAttribute "ACCEL_I_18", ACCEL_I_18_Value
    '            Call Parent.LogTrace(0, MODULE_ID & ".ReceivedVefei ==> AddLotAttribute ACCCEL_I For Slot18 ")
    '        Case "19"
    '            ACCEL_I_19_Value = objVfei.Var("PDS").Var("ACCEL_I")
    '            If Err.Number <> 0 Then
    '                bolContinue = False
    '                sReceiv_vfei = False
    '                Exit Sub
    '            End If
    '            'Senf VFEI
    '            objvfei_send.CommandID = "PDS"
    '            objvfei_send.MachineID = Parent.GetPuObject.EqDrvList(1).EqDrvID
    '            objvfei_send.MessageType = "E"
    '            objvfei_send.AddVar "PDS", "L", ""
    '            objvfei_send.AddVar "LOT_ID", "A", objLot.LotID
    '            objvfei_send.AddVar "PORT", "A", objLot.Port
    '            objvfei_send.AddVar "SLOT", "U4", iSlot
    '            objvfei_send.AddVar "BATCH_ID", "A", objVfei.Var("BATCH_ID")
    '            objvfei_send.AddVar "RECIPE_STEP", "A", objVfei.Var("RECIPE_STEP")
    '            objvfei_send.AddVar "RUN", "A", objVfei.Var("RUN")
    '            objvfei_send.AddVar "CONTROLLER", "A", objLot.LotID

    '            On Error GoTo ErrorHandler
    
    '            Call CalACCCEL_I(objBatch.LotList(1), objVfei, objvfei_send)
    '            Call Parent.LogTrace(0, MODULE_ID & ".ReceivedVefei ==> Calculate ACCCEL_I For Slot18 And Slot19")
    '            Call Parent.SendVfei(objvfei_send, Me)
    '        Case "43"
    '            ACCEL_I_43_Value = objVfei.Var("PDS").Var("ACCEL_I")
    '            If Err.Number <> 0 Then
    '                bolContinue = False
    '                sReceiv_vfei = False
    '                Exit Sub
    '            End If
    '            AddLotAttribute "ACCEL_I_43", ACCEL_I_43_Value
    '            Call Parent.LogTrace(0, MODULE_ID & ".ReceivedVefei ==> AddLotAttribute ACCCEL_I For Slot43 ")
    '        Case "44"
    '            ACCEL_I_44_Value = objVfei.Var("PDS").Var("ACCEL_I")
    '            If Err.Number <> 0 Then
    '                bolContinue = False
    '                sReceiv_vfei = False
    '                Exit Sub
    '            End If
    '            'Senf VFEI
    '            objvfei_send.CommandID = "PDS"
    '            objvfei_send.MachineID = Parent.GetPuObject.EqDrvList(1).EqDrvID
    '            objvfei_send.MessageType = "E"
    '            objvfei_send.AddVar "PDS", "L", ""
    '            objvfei_send.AddVar "LOT_ID", "A", objLot.LotID
    '            objvfei_send.AddVar "PORT", "A", objLot.Port
    '            objvfei_send.AddVar "SLOT", "U4", iSlot
    '            objvfei_send.AddVar "BATCH_ID", "A", objVfei.Var("BATCH_ID")
    '            objvfei_send.AddVar "RECIPE_STEP", "A", objVfei.Var("RECIPE_STEP")
    '            objvfei_send.AddVar "RUN", "A", objVfei.Var("RUN")
    '            objvfei_send.AddVar "CONTROLLER", "A", objLot.LotID

    '            On Error GoTo ErrorHandler
    
    '            Call CalACCCEL_I(objBatch.LotList(1), objVfei, objvfei_send)
    '            Call Parent.LogTrace(0, MODULE_ID & ".ReceivedVefei ==> Calculate ACCCEL_I For Slot43 And Slot44")
    '            Call Parent.SendVfei(objvfei_send, Me)
    '    End Select
    

    sReceiv_vfei = False
    
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
'-----------------------------------20220704 add by YL
'DF Want to set SPC Monitor of MI Slot18 and Slot19 ACCEL_I dif
'-------------------------------------------------------------------------------
Private Sub CalACCCEL_I(objLot As Lot, objVfei As vfei.Message22, objvfei_send)
    Dim ACCEL_I_18_Value As String
    Dim ACCEL_I_D1819_Value As String
    Dim ACCEL_I_43_Value As String
    Dim ACCEL_I_D4344_Value As String

    On Error GoTo CalACCCEL_I_Error_Lable
    
    If objVfei.Var("SLOT") = "19" Then
        ACCEL_I_18_Value = GetLotAttributeValue("ACCEL_I_18")
    
        ACCEL_I_D1819_Value = CStr(CDbl(objVfei.Var("PDS").Var("ACCEL_I")) - CDbl(ACCEL_I_18_Value))
    
        objvfei_send.Var("PDS").AddVar "ACCEL_I_D18_19", "A", CDbl(ACCEL_I_D1819_Value)
        
        AddLotAttribute "ACCEL_I_D1819", ACCEL_I_D1819_Value
        
    ElseIf objVfei.Var("SLOT") = "44" Then
        ACCEL_I_43_Value = GetLotAttributeValue("ACCEL_I_43")
    
        ACCEL_I_D4344_Value = CStr(CDbl(objVfei.Var("PDS").Var("ACCEL_I")) - CDbl(ACCEL_I_43_Value))
    
        objvfei_send.Var("PDS").AddVar "ACCEL_I_D18_19", "A", CDbl(ACCEL_I_D4344_Value)
        
        AddLotAttribute "ACCEL_I_D1819", ACCEL_I_D4344_Value
    End If

CalACCCEL_I_Error_Lable:

End Sub
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
End Function


