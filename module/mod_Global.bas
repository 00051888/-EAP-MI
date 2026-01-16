Attribute VB_Name = "mod_Global"
Option Explicit
Private Const ALARM_REPORT_AlarmID = 1
Private Const ALARM_REPORT_AlarmText = 2
Private Const ALARM_REPORT_ECKeyString = 3
Private Const ALARM_REPORT_OptionMode = 4
Private Const ALARM_REPORT_PPID1 = 5
Private Const ALARM_REPORT_PPID2 = 6
Private Const ALARM_REPORT_Description = 7
Public gv_Language As String

Public Enum AlarmMode
   LogOnly = 100
   OK = 1
   Rework_Continue = 2
   HoldLot_Continue = 3
   Rework_HoldLot_Continue = 4
   AckHold = 5
   Continue_HoldJob = 6
   AckHoldJob = 7
   Retry_HoldJob = 8
   Retry_Cancel = 9
   OK_Retry = 10
   Continue_Cancel = 11
End Enum
Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Declare Function GetPrivateProfileSection Lib "kernel32" Alias "GetPrivateProfileSectionA" (ByVal lpAppName As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Declare Function WritePrivateProfileString& Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal AppName$, ByVal KeyName$, ByVal keydefault$, ByVal FileName$)
Public Function AddToINI(ByVal SectionName As String, ByVal KeyName As String, ByVal KeyValue As String, ByVal IniFile As String) As Long

    AddToINI = WritePrivateProfileString(SectionName, KeyName, KeyValue, IniFile)
End Function


Public Function GetToken2(ByVal MessageBuf As String, ByVal Delimitor As String, ByVal Index As Long) As String
    Dim StartPos As Long, EndPos As Long
    Dim Count As Long
    
    ' Reset the search pointer
    StartPos = 0
    EndPos = 0
    Count = 0
    
    ' Get the first one char as the delimiter
    Delimitor = Mid(Delimitor, 1, 1)
    
    ' Because the end of string is equivalent to
    ' a delimitor, add on a delimitor to its tailer
    MessageBuf = MessageBuf & Delimitor
    
    ' Do loop to get the necessary token
    For EndPos = 1 To Len(MessageBuf)
        ' Search the delimiter in MessageBuf
        If Mid(MessageBuf, EndPos, 1) = Delimitor Then
            ' Delimiter is Found
            ' Check if there are continuous delimitors
            If (EndPos - StartPos) = 1 Then
                ' Continuous delimitors, nothing to do
            Else
                ' Something exists, increment to Index
                Count = Count + 1
                
                ' Check whether this that we want
                If Count = Index Then
                    ' yes, it is.
                    GetToken2 = Mid(MessageBuf, StartPos + 1, _
                        EndPos - StartPos - 1)
                    Exit Function
                End If
            End If
            
            ' adjust the StartPos
            StartPos = EndPos
        End If
    Next
    
    ' So far, token specified is Not Found yet
    GetToken2 = ""
    
End Function
Public Function FormatErrMsg(ModuleName As String, ErrorCode As Long, ErrorText As String) As String
    FormatErrMsg = ModuleName & " Error: (" & Format$(ErrorCode) & ") " & ErrorText
End Function
Public Function SendLotEventReport(Parent As tsEC.clsRunManager, objTask As Object, objLot As Lot) As Boolean
    Dim objHostTrans As HostTransaction
    
    
    
    Set objHostTrans = Parent.NewHostTransaction(objTask, "LOT_EVENT_TO_ECUI")
    With objHostTrans.Primary
        .Item("LOT_ID") = objLot.LotID
        .Item("CASSETTE") = objLot.Cassette
        .Item("STATUS") = objLot.Status
        .Item("PORT") = objLot.Port
        .Item("BATCH_ID") = objLot.BatchID
    End With
    objHostTrans.SubjectName = Parent.GetPuObject.UI_Subject
    objHostTrans.Send
    
End Function
Public Function AlarmReport(Parent As tsEC.clsRunManager, objTask As TaskAncestor, AlarmID As String, Optional AlarmText As String, Optional Description As String, Optional AlarmMode As AlarmMode) As String
    Static KeySerial As Long
    Dim sAlarmText As String
    Dim objHostTrans As HostTransaction
    
    
    If Len(AlarmText) = 0 Then
       sAlarmText = GetTaskAttributeFromMDB(Parent, objTask.TaskID, AlarmID)
    Else
       sAlarmText = AlarmText
    End If
    
    If AlarmMode = 0 Then
       AlarmMode = OK
    End If
    
    
    Call Parent.LogAlarm(Val(AlarmID), "SET", 0, sAlarmText)
    
    
    If AlarmMode <> LogOnly Then
       
       Set objHostTrans = Parent.NewHostTransaction(objTask, "ALARM_REPORT_TO_ECUI")
    
       KeySerial = KeySerial + 1
       If KeySerial >= 100000 Then KeySerial = 0
       With objHostTrans.Primary
            .Item("ALARM_ID") = AlarmID
            .Item("TEXT") = sAlarmText
            .Item("DESCRIPTION") = Description
            .Item("RECIPE1") = ""
            .Item("RECIPE2") = ""
            .Item("OPTION_MODE") = AlarmMode
            .Item("EC_KEY_STRING") = CStr(KeySerial)
            
       End With
       objHostTrans.SubjectName = Parent.GetPuObject.UI_Subject
       objHostTrans.Send

    
       AlarmReport = CStr(KeySerial)
    End If
    
End Function


Function GetLotListString(objBatch As Batch) As String
    Dim TempStr As String
    Dim objLot As Lot
    
    
    For Each objLot In objBatch.LotList
        TempStr = TempStr & objLot.LotID & ","
    Next
    GetLotListString = Left$(TempStr, Len(TempStr) - 1)
    
    
    

End Function

' Input  Parameter  : SectionName,FileName
' Output Parameter  : Setting     : That will be an array , each element is one line of ini file
'                                   For eaxmple , [GRID1=Lot,Lot1,100]
Public Sub GetIniSectionSetting(ByVal FileName As String, ByVal SectionName As String, Setting() As String)
  
    Dim sReturn As String
    Dim iStart As Integer, iEnd As Integer
    Dim i As Integer
    Dim iCount As Integer
   
    sReturn = Space(8000)
    Call GetPrivateProfileSection(SectionName, sReturn, 8000, FileName)
    sReturn = Trim$(sReturn)
    ReDim Setting(0 To 0)
   
    iCount = 0
    iStart = 1
   
    Do While iStart < Len(sReturn)
       iCount = iCount + 1
       If iCount = 1 Then
          ReDim Setting(1 To 1)
       Else
          ReDim Preserve Setting(1 To iCount)
       End If
       iEnd = InStr(iStart, sReturn, Chr(0))
       If iEnd = 0 Then iEnd = Len(sReturn)
       Setting(iCount) = Mid$(sReturn, iStart, iEnd - iStart)
       iStart = iEnd + 1
    Loop
    
  

End Sub
Public Function GetIniSetting(ByVal FileName As String, ByVal SectionName As String, ByVal KeyName As String) As String
    Dim ReturnValue As String
    Dim ReturnLen As Integer
    
    ReturnValue = Space(300)

   
    ReturnLen = GetPrivateProfileString(SectionName, KeyName, "", ReturnValue, 300, FileName)
    GetIniSetting = Left$(ReturnValue, ReturnLen)

End Function

Function Gettoken(ByVal str As String, dm As String, pos As Integer) As String
    Dim tmp_str As String
    Dim i As Integer
    Dim local_pos As Integer
    Dim no_dm As Boolean
    Dim tokencount As Integer
    
    no_dm = True
    Gettoken = ""
    local_pos = 0
    tmp_str = str
    tokencount = GetTokenCount(tmp_str, dm)
    If tokencount = 0 Then
        Exit Function
    End If
    While Len(tmp_str) <> 0
        i = InStr(tmp_str, dm)
        If i = 0 And no_dm = True Then
            Exit Function
        End If
        local_pos = local_pos + 1
        no_dm = False
        If pos = local_pos Then
            If i = 0 Then
                Gettoken = tmp_str
                Exit Function
            Else
                Gettoken = Left$(tmp_str, i - 1)
                Exit Function
            End If
        Else
            If pos > tokencount Then Exit Function
            tmp_str = Right$(tmp_str, Len(tmp_str) - i)
        End If
    Wend
End Function
Function GetTokenCount(ByVal str As String, dm As String) As Integer
    Dim tmp_str As String
    Dim i As Integer
    Dim no_dm As Boolean
    
    no_dm = True
    GetTokenCount = 0
    tmp_str = str
    While Len(tmp_str) <> 0
        i = InStr(tmp_str, dm)
        If i = 0 Then
            If no_dm = True Then
                GetTokenCount = 0
                Exit Function
            Else
                GetTokenCount = GetTokenCount + 1
                Exit Function
            End If
        End If
        no_dm = False
        tmp_str = Right$(tmp_str, Len(tmp_str) - i)
        GetTokenCount = GetTokenCount + 1
    Wend
    GetTokenCount = GetTokenCount + 1
End Function
Public Function Translate(Parent As tsEC.clsRunManager, sChinese As String, sEnglish As String) As String
    
    On Error Resume Next
    gv_Language = Parent.GetPuObject.Attributes("Language").Value
    If Err.Number <> 0 Then
        Translate = sChinese
    Else
        If gv_Language = "English" Then
            Translate = sEnglish
        Else
            Translate = sChinese
        End If
    End If
    Err.Clear
End Function
Public Function GetTaskAttributeFromMDB(Parent As tsEC.clsRunManager, sTaskID As String, KeyName As String) As String
    On Error Resume Next
    
    gv_Language = Parent.GetPuObject.Attributes("Language").Value
    If Err.Number <> 0 Then
        GetTaskAttributeFromMDB = Parent.GetTaskAttribute(sTaskID, KeyName)
    
    Else
        If gv_Language = "English" Then
            GetTaskAttributeFromMDB = Parent.GetTaskAttribute(sTaskID, KeyName & "_ENG")
            If Err.Number <> 0 Or GetTaskAttributeFromMDB = "" Then
                GetTaskAttributeFromMDB = Parent.GetTaskAttribute(sTaskID, KeyName)
            End If
        Else
            GetTaskAttributeFromMDB = Parent.GetTaskAttribute(sTaskID, KeyName)
        End If
    End If
    
    Err.Clear
End Function
Public Function ParseingString(ByVal Msg As String, ParseingWord As String) As String


    Dim TempStr As String
    Dim ReplyMsg As String

    Dim i As Integer

    ReplyMsg = ""
         
    For i = 1 To Len(Msg)
        TempStr = Mid$(Msg, i, 1)
        If TempStr = " " Then
            ReplyMsg = ReplyMsg & ParseingWord
        Else
            ReplyMsg = ReplyMsg & TempStr
            
        End If
        
    Next i

    ParseingString = ReplyMsg


End Function
