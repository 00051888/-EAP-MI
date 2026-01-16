VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MW_PDS2_CalGasRatio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements ts_ECAncestor.TaskAncestor

Private Const MODULE_ID = "MW_PDS2_CalGasRatio"
Dim Parent As tsEC.clsRunManager
Dim TaskID As String
Dim objBatch As Batch
Dim vShowAlarm As Boolean

Private Sub TaskAncestor_Execute(Parameter As Variant)
Dim objEqDrv As EqDrv

On Error GoTo ErrorHandler
Set objBatch = Parameter
For Each objEqDrv In Parent.GetPuObject.EqDrvList
    Call Parent.SubscribeVfeiMsg(Me, objEqDrv.EqDrvID, "PDS2")
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
    Dim sWAFER_STIME As String
    Dim sWAFER_ETIME As String
    Dim sIMPTIME As String
    Dim sDOSE As String
    Dim sBEAMCURRENT As String
    Dim objVfei2 As Message22


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

        On Error Resume Next
        sSlot = objVfei.Var("SLOT").Value
        If CInt(sSlot) > 25 Then
            sSlot = CStr(CInt(sSlot) - 25)
        End If
        
        'CAL BEAMCURRENT
        sBEAMCURRENT = CStr(objVfei.Var("PDS").Var("BEAMCURRENT").Value)
        Call Parent.LogTrace(0, "[TRACE] BEAMCURRENT = " & sBEAMCURRENT)
        'CAL DOSE
        sDOSE = CStr(GetDoseValueFromPPID(objLot.Recipe1))
        Call Parent.LogTrace(0, "[TRACE] GetDoseValueFromPPID = " & sDOSE)
        'CAL IMPTIME
        sWAFER_STIME = CStr(objVfei.Var("PDS").Var("WAFER_START").Value)
        Call Parent.LogTrace(0, "[TRACE] WAFER_START = " & sWAFER_STIME)
        sWAFER_ETIME = CStr(objVfei.Var("PDS").Var("WAFER_END").Value)
        Call Parent.LogTrace(0, "[TRACE] WAFER_END = " & sWAFER_ETIME)
        sIMPTIME = CStr(CalculateTimeDifference(sWAFER_STIME, sWAFER_ETIME))
        Call Parent.LogTrace(0, "[TRACE] CalculateTimeDifference = " & sIMPTIME)
    
        Set objVfei2 = New Message22
        objVfei2.CommandID = "PDS"
        objVfei2.MachineID = Parent.GetPuObject.EqDrvList(1).EqDrvID
        objVfei2.MessageType = "E"
        objVfei2.AddVar "LOT_ID", "A", objLot.LotID
        objVfei2.AddVar "SLOT", "A", sSlot
        objVfei2.AddVar "CONTROLLER", "A", objLot.LotID
        objVfei2.AddVar "PDS", "L", "PDS"
        With objVfei2.Var("PDS")
            .AddVar "BEAMCURRENT", "A", sBEAMCURRENT
            .AddVar "IMPTIME", "A", sIMPTIME
            .AddVar "DOSE", "A", sDOSE
        End With
    
        Call Parent.SendVfei(objVfei2, Me)
  
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

Function GetDoseValueFromPPID(tPPID As String) As Double
    ' 宣告變數
    Dim double_Return As Double
    Dim iPos1 As Integer
    Dim iPos2 As Integer
    Dim iLength As Integer
    Dim tMant As String
    Dim tExpo As String
    Dim i As Integer

    On Error GoTo ErrorHandler
    
    ' 初始化返回值
    GetDoseValueFromPPID = 0
    
    ' 取得K的位置後一位
    iPos1 = InStr(tPPID, "K") + 1
    If iPos1 = 0 Then
        GoTo ErrorHandler
    End If
    
    ' 取得E的位置
    iPos2 = InStr(iPos1, tPPID, "E")
    If iPos2 = 0 Then
        GoTo ErrorHandler
    End If
    
    ' 計算Mantissa的長度
    iLength = iPos2 - iPos1
    If iLength <= 0 Then
        GoTo ErrorHandler
    End If
    
    ' 取得Mantissa值並在後面加上一個0
    tMant = Mid(tPPID, iPos1, iLength)
    tMant = tMant & "0"
    
    ' 取得Exponent?，只取前兩位數字的部分
    tExpo = Left(Mid(tPPID, iPos2 + 1), 2)
    
    ' 清理 Exponent 只保留數字部分
    Dim cleanedExpo As String
    cleanedExpo = ""
    For i = 1 To Len(tExpo)
        If IsNumeric(Mid(tExpo, i, 1)) Then
            cleanedExpo = cleanedExpo & Mid(tExpo, i, 1)
        End If
    Next i
    
    ' 檢查清理後的 Exponent 是否為有效的數字
    If Not IsNumeric(cleanedExpo) Then
        GoTo ErrorHandler
    End If
    
    ' 計算返回值，調整公式以滿足需求
    double_Return = (CDbl(tMant) * (10 ^ CDbl(cleanedExpo))) * 10000000
    
    ' 設置返回值
    GetDoseValueFromPPID = double_Return
    Exit Function
    
ErrorHandler:
    GetDoseValueFromPPID = -1 ' 回傳一個錯誤代碼或其他指示錯誤的值
End Function

Function CalculateTimeDifference(startTimeStr As String, endTimeStr As String) As Long
    ' 使用自定義函數將字串轉換為日期時間型態
    Dim startTime As Date
    Dim endTime As Date
    
    On Error GoTo ErrorHandler
    
    startTime = ConvertStringToDate(startTimeStr)
    Call Parent.LogTrace(0, "[TRACE] ConvertStringToDate = " & startTime)
    endTime = ConvertStringToDate(endTimeStr)
    Call Parent.LogTrace(0, "[TRACE] ConvertStringToDate = " & endTime)
    
    ' 計算時間差（以秒為單位）
    CalculateTimeDifference = DateDiff("s", startTime, endTime)
    Call Parent.LogTrace(0, "[TRACE] ConvertStringToDate out put = " & CalculateTimeDifference)
    
    Exit Function
    
ErrorHandler:
    CalculateTimeDifference = -1 ' 回傳一個錯誤代碼或其他指示錯誤的值
End Function

Function ConvertStringToDate(dateStr As String) As Date
    ' 解析 yymmddhhmmss 格式的字串
    Dim yearPart As Integer
    Dim monthPart As Integer
    Dim dayPart As Integer
    Dim hourPart As Integer
    Dim minutePart As Integer
    Dim secondPart As Integer
    
    ' 提取各部分
    yearPart = 2000 + CInt(Mid(dateStr, 1, 2))
    monthPart = CInt(Mid(dateStr, 3, 2))
    dayPart = CInt(Mid(dateStr, 5, 2))
    hourPart = CInt(Mid(dateStr, 7, 2))
    minutePart = CInt(Mid(dateStr, 9, 2))
    secondPart = CInt(Mid(dateStr, 11, 2))
    
    ' 組合成日期時間型態
    ConvertStringToDate = DateSerial(yearPart, monthPart, dayPart) + TimeSerial(hourPart, minutePart, secondPart)
End Function
