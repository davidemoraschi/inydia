VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "INyDIA_SourceTube"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private internal_RemainingVolume As Long
Private internal_NoOfQuotas As Integer
Private internal_LastQuotaVolume As Long
Private internal_IsTubeEmpty As Boolean
Public DetectedVolume As Long
Public ReqQuotaVolume As Long
Public UseIncompleteQuota As Boolean
Public PreviousTubesQuotas As Integer
Public TubeOrderNo As Integer
Public NoOfDispensedQuotas As Integer
Public MinLastQuota As Long
Public Property Get RemainingVolume()
    RemainingVolume = DetectedVolume - NoOfDispensedQuotas * ReqQuotaVolume
End Property
Public Property Get NoOfQuotas()
    NoOfQuotas = IIf(UseIncompleteQuota And DetectedVolume > 0 And LastQuotaVolume > MinLastQuota, Int(DetectedVolume / ReqQuotaVolume) + 1, Int(DetectedVolume / ReqQuotaVolume))
End Property
Public Property Get LastQuotaVolume()
    LastQuotaVolume = Abs((DetectedVolume Mod ReqQuotaVolume) * UseIncompleteQuota)
End Property
Public Property Get IsTubeEmpty()
    IsTubeEmpty = RemainingVolume <= MinLastQuota
End Property
        
Private Sub Class_Initialize()
    'Debug.Print "Class_Initialize"
        UseIncompleteQuota = True
        MinLastQuota = 100
End Sub
