VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LotListData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Public LotID As String
Public CassetteID As String
Public PortID As String
Public Recipe As String
Public LotEventCol As New Collection
Private WaferEventCol As New Collection

Public Sub AddLotEvent(iTID As String, iEventName As String, iTMST As String)
Dim vLotEventData As LotEventData

Set vLotEventData = New LotEventData
vLotEventData.TID = iTID
vLotEventData.EventName = iEventName
vLotEventData.EventTMST = iTMST
LotEventCol.Add vLotEventData
End Sub

Public Sub AddWaferEvent(iTID As String, iSlot As String, iChamber As String, iRecipe As String, iEventName As String, iTMST As String)
Dim vSlotChamberData As SlotChamberData

Set vSlotChamberData = New SlotChamberData
vSlotChamberData.TID = iTID
vSlotChamberData.sLotID = iSlot
vSlotChamberData.Chamber = iChamber
vSlotChamberData.Recipe = iRecipe
vSlotChamberData.EventName = iEventName
vSlotChamberData.EventTMST = iTMST
WaferEventCol.Add vSlotChamberData
End Sub

Public Function GetLotEventCol() As Collection
Set GetLotEventCol = LotEventCol
End Function

Public Function GetWaferEventCol() As Collection
Set GetWaferEventCol = WaferEventCol
End Function





