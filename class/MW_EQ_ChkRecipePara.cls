VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MW_EQ_ChkRecipePara"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit


Implements ts_ECAncestor.TaskAncestor

Private Const MODULE_ID = "MW_EQ_ChkRecipePara"

Dim Parent As tsEC.clsRunManager
Dim TaskID As String
Dim objBatch As Batch
Dim objSmif As SmifItem
Dim objPort As PortItem
Dim objLot As Lot

Private Sub Class_Terminate()
    Set objBatch = Nothing
    Set objPort = Nothing
    Set objSmif = Nothing
    Set objLot = Nothing
End Sub

Private Sub TaskAncestor_Execute(Parameter As Variant)
    Dim objEqDrv As EqDrv
    
    Set objBatch = Parameter
    
    Call Parent.SubscribeVfeiMsg(Me, Parent.GetPuObject.EqDrvList(1).EqDrvID, "Beamscan_Complete")
    
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
    Dim objnewVfei As New Message22
    Dim Parray(1 To 5) As String
    Dim sPPID As String
    Dim sPrePPID As String
    Dim scharPPID As String
    Dim sDopant As String
    Dim nEnergy As Single
    Dim nDose As Double
    Dim nRotations As String, nTilt As Integer
    Dim intCnt As Integer
    Dim bWait As Boolean
    Dim TempStrA As String
    Dim TempStrB As String
    Dim strVal As String
    Dim strErrDesc As String
    Dim strLogErr As String
    
    On Error GoTo ErrorHandler
    
'Yen add for MI-4 dual function
For Each objLot In objBatch.LotList
If (objVfei.CommandID = "Beamscan_Complete") Or (objVfei.CommandID = "GO_DISPLAY") Then
If (objVfei.Var("CONTROLLER").Value <> objLot.LotID) Then Exit Sub
End If
Next
'Yen
    
    'If (objVfei.CommandID <> "Beamscan_Complete") Or (objVfei.CommandID <> "QUERY_STATUS") Then Exit Sub
    Select Case objVfei.CommandID
        Case "Beamscan_Complete"
           Call Parent.LogTrace(0, "BeamScan_Complete")
           Call Parent.SubscribeVfeiMsg(Me, Parent.GetPuObject.EqDrvList(1).EqDrvID, "GO_DISPLAY")
           ' bWait = False
           ' TempStrA = Trim$(objVfei.Var("LLOT").Value)
           ' TempStrB = Trim$(objVfei.Var("RLOT").Value)
            
           ' If TempStrA <> "" And TempStrB <> "" Then
           '    For Each objLot In objBatch.LotList
           '       If objLot.LotID = TempStrA Then
           '          objLot.Chamber = "Beamscan_Complete"
           '       ElseIf objLot.LotID = TempStrB Then
           '          objLot.Chamber = "Beamscan_Complete"
           '       End If
           '    Next
           ' ElseIf TempStrA <> "" And TempStrB = "" Then
           '    For Each objLot In objBatch.LotList
           '        If objLot.LotID = TempStrA Then
           '           objLot.Chamber = "Beamscan_Complete"
           '        End If
           '    Next
           ' ElseIf TempStrA = "" And TempStrB <> "" Then
           '    For Each objLot In objBatch.LotList
           '        If objLot.LotID = TempStrB Then
           '           objLot.Chamber = "Beamscan_Complete"
           '        End If
           '    Next
           '' End If
      '
            'For Each objLot In objBatch.LotList
            '   If objLot.Chamber <> "Beamscan_Complete" Then
            '      bWait = True
            '      Exit For
            '   End If
            'Next
       
           'If bWait = False Then
           '   Call Parent.SubscribeVfeiMsg(Me, Parent.GetPuObject.EqDrvList(1).EqDrvID, "GO_DISPLAY")
           'End If
           
        Case "GO_DISPLAY"
           
            If objBatch.RunMode = "PR" Then
                Parray(1) = "FIL_I"
                Parray(2) = "ION_AMU"
                Parray(3) = "DOSE_EXPONENT"
                Parray(4) = "ENERGY"
                Parray(5) = "STATUS"
                With objnewVfei
                    .CommandID = "QUERY_STATUS"
                    .MachineID = objVfei.MachineID
                    .MessageType = "C"
                    .AddVar "PARAMETER", "A", Parray
                End With
                Call Parent.SendVfei(objnewVfei, Me)
            Else
                Call Parent.TaskComplete(TaskID, "")
                Set objBatch = Nothing
            End If
       
        Case "QUERY_STATUS"
            For Each objLot In objBatch.LotList
                sPPID = objLot.Recipe1
            Next
            'Filament I
            strVal = objVfei.Var("FIL_I").Value
            
            'Check ION_AMU
            strVal = CInt(objVfei.Var("ION_AMU").Value)
            strErrDesc = MFCheckAMU(sPPID, strVal, intCnt)
            
            strLogErr = Trim(strErrDesc)
            
            'Check Energy
            strVal = CInt(objVfei.Var("ENERGY").Value)
            strErrDesc = MFCheckEnergy(sPPID, strVal, intCnt)
            If Trim(strErrDesc) <> "" Then
                strLogErr = strLogErr & strErrDesc
            End If
            
            'Check Dose Exponent
            strVal = CInt(objVfei.Var("DOSE_EXPONENT").Value)
            strErrDesc = MFChkQVal(sPPID, strVal, intCnt)
            If Trim(strErrDesc) <> "" Then
                strLogErr = strLogErr & strErrDesc
            End If
            
            'If strLogErr <> "" Then
            '    Call AlarmReport(Parent, Me, "10011", strLogErr)
            'End If
            
            'Call Parent.TaskComplete(TaskID, "")
            'Set objBatch = Nothing
            
             If strLogErr = "" Then
                Call Parent.TaskComplete(TaskID, "OK")
                Set objBatch = Nothing
             Else
                Call AlarmReport(Parent, Me, "10011", strLogErr)
             End If
            
           
            
            
            
          '  sPPID = objLot.Recipe1
            'Get Dopant
          '  For i = 1 To Len(sPrePPID) Step 1
          '     scharPPID = Mid(sPrePPID, i, 1)
          '     If (scharPPID >= "A" And scharPPID <= "Z") Or (scharPPID >= "a" And scharPPID <= "z") Then
          '        sDopant = sDopant & scharPPID
          '     Else
          '        sPrePPID = Mid(sPrePPID, i + 1)
          '        Exit For
          '     End If
          '  Next i
            
           'Get Energy
           'For i = 1 To Len(sPrePPID) Step 1 'For i = Len(sPrePPID) To 1 Step -1
           '   scharPPID = Mid(sPrePPID, i, 1)
           '   If scharPPID < "0" Or scharPPID > "9" And scharPPID <> "." Then
           '      nEnergy = CSng(Mid(sPrePPID, i + 1))
           '      sPrePPID = Left(sPrePPID, i)
           '      Exit For
           '   End If
           'Next i
          '
          ' 'Get Dose
          ' nDose = Left(sPrePPID, 5)
          ' sPrePPID = Mid(sPrePPID, 6)
   '
   '        'Get Rotations,Tilt
   '        If Len(sPrePPID) = 0 Then
   '           nRotations = 0
   '           nTilt = 0
   '        ElseIf Len(sPrePPID) = 1 Then
   '           nRotations = Left(sPrePPID, 1)
   '           nTilt = 0
   '        ElseIf Len(sPrePPID) = 2 Then
   '           nRotations = 0
   '           nTilt = CInt(Right(sPrePPID, 1))
   '        ElseIf Len(sPrePPID) = 3 Then
   '           For i = 1 To Len(sPrePPID)
   '              scharPPID = Mid(sPrePPID, 1, 1)
   '              If scharPPID < "A" Or scharPPID > "Z" Then
   '                 If i = 2 Then
   '                    nRotations = 0
   '                    nTilt = CInt(Mid(sPrePPID, i, Len(sPrePPID) - i - 1))
   '                 ElseIf i = 3 Then
   '                    nRotations = Left(sPrePPID, 1)
   '                    nTilt = CInt(Mid(sPrePPID, i, Len(sPrePPID) - i - 1))
   '                 End If
   '              End If
   '           Next
   '        ElseIf Len(sPrePPID) = 4 Then
   '           nRotations = Left(sPrePPID, 1)
    '          nTilt = CInt(Mid(sPrePPID, 3, 2))
   '        End If
    End Select
    
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

Private Function MFCheckAMU(strPPID As String, strVal As String, _
                                        intCnt As Integer) As String
Dim strSpc As String
Dim strError As String
Dim strChar As String
Dim bolFlag As Boolean
Dim DblMaxVal As Double
Dim DblMinVal As Double
Dim DblAMU As Double
Dim DblAmuVal  As Double
Dim strDesc As String
Dim strErrMsg As String

    'Get the species length
    'Where As+ is As
    '      P+, P++, P+++ is P
    '      BF2+ is BF
    '      B+, B++, B+++ is B
    For intCnt = 1 To Len(strPPID)
        strChar = Mid(strPPID, intCnt, 1)
        If IsNumeric(strChar) Then
            Exit For
        End If
    Next
    
    strSpc = Mid(strPPID, 1, intCnt - 1)
    
    'Check for AMU
    Select Case strSpc
        Case "B"
            'Default AMU value for B
            DblAmuVal = 11
        Case "P"
            'Default AMU value for P
            DblAmuVal = 31
        Case "AS"
            'Default AMU value for AS
            DblAmuVal = 75
        Case "BF"
            'Default AMU value for BF
            DblAmuVal = 49
        Case "IN"
            DblAmuVal = 115
        Case "SB"
            DblAmuVal = 121
            
    End Select
   
    DblMaxVal = DblAmuVal + 1
    DblMinVal = DblAmuVal - 1
    
    If CDbl(strVal) > DblMaxVal Or CDbl(strVal) < DblMinVal Then
        strDesc = "AMU(" & DblAmuVal & ") ; PPID(" & strVal & ")"
        strErrMsg = strDesc & vbCrLf
        Call Parent.LogTrace(0, "CheckAMU:" & strDesc)
    End If
    
    MFCheckAMU = strDesc
    
End Function

'Private Function MFChkQVal(bolRotQ As Boolean, bolRotB As Boolean, strDose As String, _
                    strMantissa As String, strExponent As String) As String
                    
Private Function MFChkQVal(strPPID As String, strVal As String, intCnt As Integer) As String
'***************************************************************************************
'bolRotQ - rotation Q, if Q found in PPID then the value will be true else false
'bolRotB - rotation B, if B found in PPID then the value will be true else false
'strDose - Dosage value in PPID
'strMantissa - Mantissa, the value will pass out for checking
'strExponent - Exponent, the value will pass out for checking
'***************************************************************************************
Dim DblDose As Double
Dim strDose As String
Dim IntRotatePos As Integer
Dim bolRotateQ  As Boolean
Dim bolRotateB As Boolean
Dim IntPos As Integer
Dim IntTiltPos As Integer
Dim strCDos As String
Dim strExponent As String
Dim strDesc As String
Dim strErrMsg As String

Const constExp As Double = 11

    'default to 3 bcoz the gas has value B also
    IntPos = 3
        
    'B = Rotate twice, Q = Rotate 4 times(only for MI)
    If InStr(strPPID, "Q") <> 0 Then
        IntRotatePos = InStr(strPPID, "Q")
        bolRotateQ = True
    ElseIf InStr(Mid(strPPID, IntPos), "B") <> 0 Then
        IntRotatePos = InStr(Mid(strPPID, IntPos), "B") + 1
        bolRotateB = True
    End If

    'Tilt
    IntTiltPos = InStr(strPPID, "T")

    ' Get the Dosage Value
    If IntRotatePos = 0 And IntTiltPos = 0 Then
        strDose = Mid(strPPID, intCnt + 1)
    Else
        If (IntRotatePos = 0 And IntTiltPos <> 0) Or _
           (IntRotatePos > IntTiltPos And IntTiltPos <> 0) Then
            strDose = Mid(strPPID, intCnt + 1, IntTiltPos - intCnt - 1)
        ElseIf (IntRotatePos <> 0 And IntTiltPos = 0) Or _
               (IntTiltPos > IntRotatePos And IntRotatePos <> 0) Then
            strDose = Mid(strPPID, intCnt + 1, IntRotatePos - intCnt - 1)
        End If
    End If
    
    strExponent = constExp
    'if Q found, dosage has to divide by 4
    If bolRotateQ Then
        DblDose = CDbl(strDose) / 4
    'if B found, dosage has to devided by 2
    ElseIf bolRotateB Then
        DblDose = CDbl(strDose) / 2
    Else 'cannot found Q and B in PPID
        DblDose = CDbl(strDose)
    End If
    
    strCDos = CStr(DblDose)
    
    If Len(strCDos) > 4 Then
        strExponent = CStr(strExponent + Len(strCDos) - 4)
    End If
    
    If strVal <> strExponent Then
        strDesc = "Dose Exponent(" & strExponent & ") ; PPID(" & strVal & ")"
        strErrMsg = strDesc & vbCrLf
        Call Parent.LogTrace(0, "CheckDoseExponent:" & strDesc)
    End If
    MFChkQVal = strDesc
End Function

Private Function MFCheckEnergy(strPPID As String, strVal As String, intCnt As Integer) As String
Dim strErrDesc As String
Dim strEnergy As String
Dim bolMFlag As Boolean
Dim IntLenM As Integer
Dim IntMaxEnrgy As Integer
Dim IntMinEnrgy As Integer
Dim IntPos As Integer

    ' Get the Energy, which can be found in PPID,
    ' where K = KeV, V = eV and M = Mev
    ' Added by Quah for above condition
    If InStr(strPPID, "K") <> 0 Then
        IntPos = InStr(strPPID, "K")
    ElseIf InStr(strPPID, "V") <> 0 Then
        IntPos = InStr(strPPID, "V")
    Else
        IntPos = InStr(strPPID, "M")
        bolMFlag = True
    End If
    
    If IntPos = 0 Then
        strErrDesc = "can't find energy position " & strPPID & ")"
        MFCheckEnergy = strErrDesc & vbCrLf
        Call Parent.LogTrace(0, "CheckRecipeBodyError:" & strErrDesc)
    Else
        strEnergy = Mid(strPPID, intCnt, IntPos - intCnt)
        'condition apply to Energy from PPID where Energy type is "M"
        'where "_" are use instead of "." for decimal
        If bolMFlag Then
            IntLenM = InStr(strEnergy, "_")
            If IntLenM <> 0 Then
                strEnergy = Left(strEnergy, IntLenM - 1) & "." & _
                        Right(strEnergy, Len(strEnergy) - IntLenM)
            End If
        End If
        
        'pass the value back to the calling function
        intCnt = IntPos
        
    End If
    
    IntMaxEnrgy = CInt(strEnergy) + 1
    IntMinEnrgy = CInt(strEnergy) - 1
    
    If IntMinEnrgy > CDbl(strVal) Or IntMaxEnrgy < CDbl(strVal) Then
        strErrDesc = "Energy(" & strEnergy & "); PPID(" & strVal & ")"
        MFCheckEnergy = MFCheckEnergy & strErrDesc & vbCrLf
        Call Parent.LogTrace(0, "CheckEnergy:" & strErrDesc)
    End If

End Function



