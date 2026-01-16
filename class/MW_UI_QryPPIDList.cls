VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MW_UI_QryPPIDList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements ts_ECAncestor.TaskAncestor

Private Const MODULE_ID = "MW_UI_QryPPIDList"

Dim Parent As tsEC.clsRunManager
Dim TaskID As String

Dim objPU As ProcessingUnit
Dim objBatch As Batch

Private Sub TaskAncestor_Execute(Parameter As Variant)
    On Error GoTo ErrorHandler
    
    Set objPU = Parameter
    Call Parent.SubscribeHostMsg(Me, "QUERY_PPID_LIST")

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
   
    Dim i As Integer
    
    
    On Error GoTo ErrorHandler
    
    If objHostTrans.TransactionName <> "QUERY_PPID_LIST" Then Exit Sub
    '---------------------------------------------------------------------------------------------

    With objHostTrans.Secondary
        .Item("MES_PPID") = "PPID"
        .Item("ISFOUND") = "Y"
        For i = 1 To objPU.RecipeList.Count
            .Item("PPIDLIST").Add.Item("PPID") = objPU.RecipeList.Item(i)
        Next i
    End With
    objHostTrans.Reply
    '---------------------------------------------------------------------------------------------
  
    Exit Sub
ErrorHandler:
    Call Parent.LogTrace(0, FormatErrMsg(MODULE_ID & ".ReceivedHostMsg", Err, Err.Description))
End Sub

Private Sub TaskAncestor_ReceivedVfei(objVfei As vfei.Message22)
    On Error GoTo ErrorHandler
    
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



