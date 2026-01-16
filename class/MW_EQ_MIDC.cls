VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MW_EQ_MIDC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit


Implements ts_ECAncestor.TaskAncestor

Private Const MODULE_ID = "MW_EQ_MIDC"

Dim Parent As tsEC.clsRunManager
Dim TaskID As String
Dim objBatch As Batch
Dim objHostTrans As HostTransaction


Private Sub TaskAncestor_Execute(Parameter As Variant)
    Dim objEqDrv As EqDrv
    
    On Error GoTo ErrorHandler
    
    Set objBatch = Parameter
    
    For Each objEqDrv In Parent.GetPuObject.EqDrvList
       Call Parent.SubscribeVfeiMsg(Me, objEqDrv.EqDrvID, "Start_End_MIDC")
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
    On Error GoTo ErrorHandler
    Set objHostTrans = Parent.NewHostTransaction(Me, "MI_DC_START_END")
    
    If InStr(1, objVfei.Var("STATUS").Value, "START", 1) <> 0 Then
        
'[Begin]
'Destination = MIDC
'Primary=MI_DC_START_END <%Subject><%TID><PU><Recipe><LOT_ID><WAFER_ID><PRODUCT><STATUS>
'        <PORT_ID><PROCESS><ROUTE><STEP><RECIPE1><RECIPE2><MESRECIPE1><MESRECIPE2><LOT_TYPE>
'        <LAYER><CASSETTE_ID>
'Secondary=//MI_DC_START_END <%Status><%TID>
'[End]

        With objHostTrans.Primary
              .Item("LOT_ID").Value = objVfei.Var("LOT_ID").Value
              .Item("WAFER_ID").Value = objVfei.Var("WAFER_ID").Value
              .Item("PU") = Parent.PUID
              .Item("RECIPE").Value = objVfei.Var("RECIPE").Value
              .Item("PRODUCT").Value = objVfei.Var("PRODUCT").Value
              .Item("STATUS").Value = objVfei.Var("STATUS").Value
              
              ' add for uFDC output parameter #ShangChun 20070628 {~
              
              .Item("PORT_ID").Value = objVfei.Var("PORT_ID").Value
              .Item("PROCESS").Value = objVfei.Var("PROCESS").Value
              .Item("ROUTE").Value = objVfei.Var("ROUTE").Value
              .Item("STEP").Value = objVfei.Var("STEP").Value
              .Item("RECIPE1").Value = objVfei.Var("RECIPE1").Value
              .Item("RECIPE2").Value = objVfei.Var("RECIPE2").Value
              .Item("MESRECIPE1").Value = objVfei.Var("MESRECIPE1").Value
              .Item("MESRECIPE2").Value = objVfei.Var("MESRECIPE2").Value
              .Item("LOT_TYPE").Value = objVfei.Var("LOT_TYPE").Value
              .Item("LAYER").Value = objVfei.Var("LAYER").Value
              .Item("CASSETTE_ID").Value = objVfei.Var("CASSETTE_ID").Value
              
              ' ~}

        End With
        objHostTrans.Send
      
    Else
       With objHostTrans.Primary
                   .Item("STATUS").Value = objVfei.Var("STATUS").Value
                   .Item("PU") = Parent.PUID

        End With
        objHostTrans.Send
    
    
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













