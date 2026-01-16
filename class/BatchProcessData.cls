VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "BatchProcessData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Public BatchID As String
Public LotListData As New Collection
Public Sub AddLotList(iLot As String, iCassetteId As String, iRecipe As String, iPortID As String)
    Dim vLotListData As LotListData
    Set vLotListData = New LotListData
    vLotListData.LotID = iLot
    vLotListData.CassetteID = iCassetteId
    vLotListData.PortID = iPortID
    vLotListData.Recipe = iRecipe
    LotListData.Add vLotListData
End Sub







