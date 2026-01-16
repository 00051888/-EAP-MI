VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MS_MES_Info"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements ts_ECAncestor.TaskAncestor

Private Const MODULE_ID = "MS_MES_Info"

Dim Parent As Object
Dim TaskID As String
Dim objPU As ProcessingUnit
Dim objBatch As Batch
Dim objVfei As vfei.Message22
Dim sReceiv_vfei As Boolean
Dim count As Integer


Private Sub TaskAncestor_Execute(Parameter As Variant)
    Dim i As Integer
    
    On Error GoTo ErrorHandler
    
    'Set objPU = Parameter
    Set objBatch = Parameter
   
    count = 0
            
    For i = 1 To objBatch.LotList.count
        count = count + 1
         Set objVfei = New Message22
        objVfei.CommandID = "FDC_MES_Info"
        objVfei.MachineID = Parent.GetPuObject.EqDrvList(1).EqDrvID
        
        objVfei.AddVar "LOT_ID", "A", objBatch.LotList(i).LotID
        objVfei.AddVar "port", "A", objBatch.LotList(i).Port
        
        objVfei.AddVar "process", "A", objBatch.LotList(i).PROCESS
        objVfei.AddVar "product", "A", objBatch.LotList(i).Product
        objVfei.AddVar "route", "A", objBatch.LotList(i).Route
        objVfei.AddVar "process_step", "A", objBatch.LotList(i).Step
        
        objVfei.AddVar "recipe", "A", objBatch.LotList(i).Recipe1
        objVfei.AddVar "route_name", "A", objBatch.LotList(i).RouteName
        objVfei.AddVar "process_step_name", "A", objBatch.LotList(i).StepName
     '#### EDEN Add for UFDC 05/09/06 ####
        objVfei.AddVar "cassette", "A", objBatch.LotList(i).Cassette
        objVfei.AddVar "layer", "A", objBatch.LotList(i).Layer
        objVfei.AddVar "lottype", "A", objBatch.LotList(i).LotType
        objVfei.AddVar "recipe2", "A", objBatch.LotList(i).Recipe2
        
        objVfei.AddVar "mesrecipe1", "A", objBatch.LotList(i).Recipe1
        objVfei.AddVar "mesrecipe2", "A", objBatch.LotList(i).Recipe2
        'objVfei.AddVar "mesrecipefamily", "A",
        'objVfei.AddVar "recipetype", "A", objBatch.LotList(1)
     '####
     Call Parent.SendVfei(objVfei, Me)
    Set objVfei = Nothing
    Next
    Call Parent.TaskComplete(TaskID, "OK")
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
    count = count - 1
    If count = 0 Then Exit Sub
      
   
    
ErrorHandler:
    Call Parent.LogTrace(0, FormatErrMsg(MODULE_ID & ".ReceivedVfei", Err, Err.Description))

End Sub

Private Property Let TaskAncestor_TaskID(ByVal RHS As String)
    TaskID = RHS
End Property

Private Property Get TaskAncestor_TaskID() As String
    TaskAncestor_TaskID = TaskID
End Property

Private Sub AMU_Delta(sRecipe As String, objVfei As vfei.Message22, objvfei_send)
Dim Recipe_Gas As String
Dim AMU_Delta_Value As Double
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

If InStr(Parent.GetTaskAttribute(TaskID, "ENERGY_DELTA"), "/K") Then
    Energy_Delta_Value = Energy_Value - objVfei.Var("PDS").Var("ENERGY") / 1000
Else
    Energy_Delta_Value = Energy_Value - objVfei.Var("PDS").Var("ENERGY")
End If

objvfei_send.Var("PDS").AddVar "ENERGY_DELTA", "A", CDbl(Energy_Delta_Value)

Energy_Delta_Error_Lable:
End Sub



