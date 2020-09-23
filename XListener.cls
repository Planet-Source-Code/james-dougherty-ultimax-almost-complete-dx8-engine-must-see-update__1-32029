VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "XListener"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'BETA NOT FINISHED... STILL IN DEVELOPEMENT...
'NOT ALL FUNCTIONS WORK "RIGHT" YET(NOT THAT MANY THOUGH)

'PLEASE NOTE:
' THIS IS NOT ALL MY CODE... SO IT SHOULD BE TO THEIR RESPECTED
' OWNER... I KNOW 1 IS SIMON.

'FOR THAT SIMPLE FACT THIS IS FREEWARE... USE AND ABUSE IT...

Private DX As New DirectX8
Private DXSound As DirectSound8
Private LPosition As D3DVECTOR
Private sListener As DirectSound3DListener8

Public Function Listener() As DirectSound3DListener8
Set Listener = sListener
End Function

Public Function Initialize_Listener(hWnd As Long) As Boolean
Dim BufferDesc As DSBUFFERDESC
Dim Buffer As DirectSoundPrimaryBuffer8
On Error GoTo InitFailed

ErrStr = "DirectX DirectSoundCreate"
Set DXSound = DX.DirectSoundCreate("")
ErrStr = "DXSound3D SetCooperativeLevel"
DXSound.SetCooperativeLevel hWnd, DSSCL_PRIORITY
BufferDesc.lFlags = DSBCAPS_CTRL3D Or DSBCAPS_PRIMARYBUFFER
ErrStr = "DXSound3D CreatePrimarySoundBuffer"
Set Buffer = DXSound.CreatePrimarySoundBuffer(BufferDesc)
If Buffer Is Nothing Then GoTo InitFailed
ErrStr = "DXSound3D GetDirectSound3DListener"
Set sListener = Buffer.GetDirectSound3DListener
If sListener Is Nothing Then GoTo InitFailed
Initialize_Listener = True
Exit Function

InitFailed:
 ErrorToFile "ErrorD3DS", ErrStr
End Function

Public Sub Set_Listeners_Head_Position(Top_Of_Head_X As Single, Top_Of_Head_Y As Single, Top_Of_Head_Z As Single, _
                                       Front_Of_Head_X As Single, Front_Of_Head_Y As Single, Front_Of_Head_Z As Single)
On Error Resume Next
sListener.SetOrientation Top_Of_Head_X, Top_Of_Head_Y, Top_Of_Head_Z, Front_Of_Head_X, Front_Of_Head_Y, Front_Of_Head_Z, DS3D_IMMEDIATE
End Sub

Public Property Get Get_Listener_Position_X() As Single
Get_Listener_Position_X = LPosition.X
End Property

Public Property Get Get_Listener_Position_Y() As Single
Get_Listener_Position_Y = LPosition.Y
End Property

Public Property Get Get_Listener_Position_Z() As Single
Get_Listener_Position_Z = LPosition.Z
End Property

Public Property Get Get_Listener_PositionEX() As D3DVECTOR
Get_Listener_PositionEX = LPosition
End Property

Public Sub Set_Listener_Position(PositionX As Single, PositionY As Single, PositionZ As Single)
On Error Resume Next
sListener.SetPosition PositionX / 50, PositionY, PositionZ / 50, DS3D_DEFERRED
LPosition.X = PositionX / 50: LPosition.Y = PositionY: LPosition.Z = PositionZ / 50
End Sub

Public Sub Set_Listeners_Velocity(VelocityX As Single, VelocityY As Single, VelocityZ As Single)
On Error Resume Next
sListener.SetVelocity VelocityX, VelocityY, VelocityZ, DS3D_DEFERRED
End Sub

Public Sub Update_Listener()
On Error Resume Next
sListener.CommitDeferredSettings
End Sub

Public Sub Cleanup_Listener()
Set DXSound = Nothing
Set DX = Nothing
End Sub
