VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MS_MES_EQ_Parameter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements ts_ECAncestor.TaskAncestor

Private Const MODULE_ID = "MS_MES_EQ_Parameter"

Dim parent      As tsEC.clsRunManager
Dim TaskID      As String
Dim objBatch    As Batch
Dim objPU As ProcessingUnit


'Input EQ Parameter from EAP to MES (FAB8F use only)
'Primary=@PDSB04_U_
'     <#Master> [<ProcessingUnit/A17> <Parameter/A17> <ParameterValue/F>
'<ParameterString/A33>]
'Secondary=

Private Sub TaskAncestor_Execute(Parameter As Variant)
    Dim objHostTrans As HostTransaction
    Dim Item         As Variant
    Dim i            As Integer
    
    Dim sParameterString As String
    Dim sParameter_Total As Double
    Dim sParameter_Value As Double
    
    
    
    Dim objLot As Lot
    Dim objSPEC As MES_DataItem
    Dim objData As DataList
    
    Dim tmpcount As Integer
    
    
    
    
    On Error GoTo ErrorHandler
    Set objBatch = Parameter
  
    sParameterString = parent.GetTaskAttribute(TaskID, "PARAMETER")
    
    If sParameterString = "" Then
        Call parent.TaskComplete(TaskID, "")
    Else
        For Each objLot In objBatch.LotList
            For Each objSPEC In objLot.MESData
                tmpcount = 0
                sParameter_Total = 0
                If InStr(sParameterString, Trim(objSPEC.Parameter)) > 0 Then

                        For i = 1 To objSPEC.DataList.Count
                            tmpcount = tmpcount + 1
                            sParameter_Total = sParameter_Total + objSPEC.DataList.Item(i).Value
                        Next i
                        sParameter_Value = sParameter_Total / tmpcount
                        Call parent.LogTrace(0, MODULE_ID & ",LOT=" & objLot.LotID & ",PARAMETER= " & objSPEC.Parameter & " ,Value= " & sParameter_Value)
                        
                        Set objHostTrans = parent.NewHostTransaction(Me, "PDSB04_U_")
                        With objHostTrans.Primary.Item("Master").Add
                            .Item("ProcessingUnit").Value = parent.GetPuObject.PUID
                            .Item("Parameter").Value = objSPEC.Parameter
                            .Item("ParameterValue").Value = sParameter_Value
                        End With
                        objHostTrans.Send

                End If
            Next objSPEC
        Next objLot
    End If
   

    Call parent.TaskComplete(TaskID, "")
    

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
    On Error GoTo ErrorHandler
    
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


