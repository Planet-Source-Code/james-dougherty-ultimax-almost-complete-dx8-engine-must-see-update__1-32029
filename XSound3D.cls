VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "XSound3D"
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

Private Type DX3DBuffers
  isLoaded As Boolean
  Buffer As DirectSoundSecondaryBuffer8
  s3DBuffer As DirectSound3DBuffer8
  Position As D3DVECTOR
End Type: Private sBuffer() As DX3DBuffers

Private DX As New DirectX8
Private DXSound As DirectSound8
Private SoundFolder As String
Private CurrentBuffer As Integer
Private HoldFileName As String

Public Function DirectXSound() As DirectSound8
On Local Error Resume Next
Set DirectXSound = DXSound
End Function

Public Function Initialize_3D_Sound_Engine(hWnd As Long)
On Local Error GoTo errOut

ErrStr = "DirectX DirectSoundCreate"
Set DXSound = DX.DirectSoundCreate("")
ErrStr = "DXSound3D SetCooperativeLevel"
DXSound.SetCooperativeLevel hWnd, DSSCL_PRIORITY
Initialize_3D_Sound_Engine = True
Exit Function

errOut:
 ErrorToFile "ErrorD3DS", ErrStr
 DoEvents
End Function

Public Sub Cleanup_3D_Sound_Engine(NumberOfBuffers As Integer)
On Local Error Resume Next
Static i As Integer
For i = 0 To NumberOfBuffers
 If sBuffer(i).Buffer.GetStatus = DSBSTATUS_PLAYING Or DSBSTATUS_LOOPING Then
  sBuffer(i).Buffer.Stop
 End If
 Set sBuffer(i).Buffer = Nothing
Next
Set DXSound = Nothing
Set DX = Nothing
End Sub

Public Sub Play_3D_Sound(Buffer As Integer, Optional Volume As Long = 100, Optional LoopIt As Byte = 0)
On Local Error GoTo errOut

ErrStr = "Play_3D_Sound"
If Volume < 100 Then Volume_Level Buffer, Volume
If sBuffer(Buffer).isLoaded Then sBuffer(Buffer).Buffer.Play LoopIt
Exit Sub

errOut:
 ErrorToFile "ErrorPlay3D", ErrStr
End Sub

Public Sub Create_3D_Buffers(AmountOfBuffers As Integer, DefaultFile As String)
On Local Error GoTo errOut

ErrStr = "DXSound3D Create_3D_Buffers"
  ReDim sBuffer(AmountOfBuffers)
  For AmountOfBuffers = 0 To AmountOfBuffers
    ErrStr = "CreateBuffers - DX7LoadSound"
    LoadSound AmountOfBuffers, DefaultFile
  Next AmountOfBuffers
Exit Sub

errOut:
 ErrorToFile "ErrorCB3D", ErrStr
End Sub

Public Sub Stop3DSound(Buffer As Integer)
On Local Error GoTo errOut

ErrStr = "StopSound"
If sBuffer(Buffer).Buffer.GetStatus = DSBSTATUS_PLAYING Or DSBSTATUS_LOOPING Then
 sBuffer(Buffer).Buffer.Stop
 sBuffer(Buffer).Buffer.SetCurrentPosition 0
End If
Exit Sub

errOut:
 ErrorToFile "ErrorStopSound3D", ErrStr
End Sub

Public Sub Pause3DSound(Buffer As Integer)
On Local Error GoTo errOut

ErrStr = "Pause3DSound"
If sBuffer(Buffer).Buffer.GetStatus = DSBSTATUS_PLAYING Or DSBSTATUS_LOOPING Then
 sBuffer(Buffer).Buffer.Stop
End If

Exit Sub

errOut:
 ErrorToFile "ErrorPauseSound3D", ErrStr
End Sub

Public Function IsPlaying(Buffer As Integer) As Long
On Local Error Resume Next
IsPlaying = sBuffer(Buffer).Buffer.GetStatus
End Function

Public Sub SetSoundDir(FolderPath As String)
On Local Error Resume Next
SoundFolder = FolderPath & "\"
End Sub

Public Sub LoadSound(Buffer As Integer, sFile As String, Optional MinDistance = DS3D_DEFAULTMINDISTANCE, Optional MaxDistance = DS3D_DEFAULTMAXDISTANCE, Optional ConeOutsideVolume As Long = DS3D_DEFAULTCONEOUTSIDEVOLUME, Optional Buffer_Start_PosX As Single = 0, Optional Buffer_Start_PosY As Single = 5, Optional Buffer_Start_PosZ As Single = 0)
On Local Error GoTo errOut
  Dim Filename As String
  Dim BufferDesc As DSBUFFERDESC
  
  BufferDesc.lFlags = DSBCAPS_CTRL3D Or DSBCAPS_CTRLFREQUENCY Or DSBCAPS_CTRLVOLUME
  
  ErrStr = "guid3DAlgorithm"
  BufferDesc.guid3DAlgorithm = GUID_DS3DALG_NO_VIRTUALIZATION
  Filename = SoundFolder & sFile
  ErrStr = "CreateSoundBufferFromFile"
  On Error Resume Next
  Set sBuffer(Buffer).Buffer = DXSound.CreateSoundBufferFromFile(Filename, BufferDesc)
  ErrStr = "Creates3DBufferFromFile"
  Set sBuffer(Buffer).s3DBuffer = sBuffer(Buffer).Buffer.GetDirectSound3DBuffer
   
  ErrStr = "Setting Defaults"
  sBuffer(Buffer).s3DBuffer.SetMinDistance MinDistance, DS3D_DEFERRED
  sBuffer(Buffer).s3DBuffer.SetMaxDistance MaxDistance, DS3D_DEFERRED
  sBuffer(Buffer).s3DBuffer.SetConeAngles DS3D_MINCONEANGLE, DS3D_MAXCONEANGLE, DS3D_DEFERRED
  sBuffer(Buffer).s3DBuffer.SetConeOutsideVolume ConeOutsideVolume, DS3D_DEFERRED
  
  HoldFileName = Filename
  sBuffer(Buffer).isLoaded = True
  Exit Sub

errOut:
 Debug.Print ErrStr
End Sub

Public Sub Volume_Level(Buffer As Integer, Volume As Long)
On Local Error GoTo errOut

ErrStr = "VolumeLevel"
  If Volume > 0 Then
    sBuffer(Buffer).Buffer.SetVolume (50 * Volume) - 5000
  Else
    sBuffer(Buffer).Buffer.SetVolume -5000
  End If
Exit Sub

errOut:
 Debug.Print ErrStr
End Sub

Public Property Get Position(Buffer As Integer) As D3DVECTOR
On Local Error Resume Next
sBuffer(Buffer).s3DBuffer.GetPosition Position
End Property

Public Property Let Position(Buffer As Integer, Position As D3DVECTOR)
On Local Error Resume Next
sBuffer(Buffer).s3DBuffer.SetPosition Position.X / 50, Position.Y, Position.Z / 50, DS3D_IMMEDIATE
End Property

Public Sub Set_Sound_Position(Buffer As Integer, X As Single, Y As Single, Z As Single)
On Local Error Resume Next
sBuffer(Buffer).Position.X = X
sBuffer(Buffer).Position.Y = Y
sBuffer(Buffer).Position.Z = Z
sBuffer(Buffer).s3DBuffer.SetPosition X / 50, Y, Z / 50, DS3D_IMMEDIATE
End Sub

Public Sub Update_3D_Sound_Angle_In_Radians(Buffer As Integer, X As Single, Y As Single, Z As Single)
On Local Error Resume Next
Dim tmpX As Single
Dim tmpY As Single
Dim tmpZ As Single

tmpX = 5 * Cos(3.141 * X / 180)
tmpY = Y
tmpZ = 5 * Sin(3.141 * Z / 180)
    
If sBuffer(Buffer).Buffer Is Nothing Then Exit Sub
sBuffer(Buffer).s3DBuffer.SetConeOrientation tmpX, tmpY, tmpZ, DS3D_IMMEDIATE
End Sub

Public Property Get Get_Minimum_Disance(Buffer As Integer) As Single
On Local Error Resume Next
Get_Minimum_Disance = sBuffer(Buffer).s3DBuffer.GetMinDistance
End Property

Public Sub Set_Minimum_Distance(Buffer As Integer, Minimum_Distance As Single)
On Local Error Resume Next
sBuffer(Buffer).s3DBuffer.SetMinDistance Minimum_Distance, DS3D_IMMEDIATE
End Sub

Public Property Get MinDistance(Buffer As Integer) As Single
On Error Resume Next
MinDistance = sBuffer(Buffer).s3DBuffer.GetMinDistance
End Property

Public Property Let MinDistance(Buffer As Integer, Distance As Single)
On Error Resume Next
sBuffer(Buffer).s3DBuffer.SetMinDistance Distance, DS3D_IMMEDIATE
End Property

Public Property Get MaxDistance(Buffer As Integer) As Single
On Error Resume Next
MaxDistance = sBuffer(Buffer).s3DBuffer.GetMaxDistance
End Property

Public Property Let MaxDistance(Buffer As Integer, Distance As Single)
On Error Resume Next
sBuffer(Buffer).s3DBuffer.SetMaxDistance Distance, DS3D_IMMEDIATE
End Property

Public Property Get Get_Maximum_Disance(Buffer As Integer) As Single
On Local Error Resume Next
Get_Maximum_Disance = sBuffer(Buffer).s3DBuffer.GetMaxDistance
End Property

Public Sub Set_Maximum_Distance(Buffer As Integer, Maximum_Distance As Single)
On Local Error Resume Next
sBuffer(Buffer).s3DBuffer.SetMaxDistance Maximum_Distance, DS3D_IMMEDIATE
End Sub

Public Property Get ConeOutsideVolume(Buffer As Integer) As Single
On Local Error Resume Next
ConeOutsideVolume = sBuffer(Buffer).s3DBuffer.GetConeOutsideVolume
End Property

Public Property Let ConeOutsideVolume(Buffer As Integer, Volume As Single)
On Local Error Resume Next
sBuffer(Buffer).s3DBuffer.SetConeOutsideVolume Volume, DS3D_IMMEDIATE
End Property

Public Property Get Get_Cone_Outside_Volume(Buffer As Integer) As Single
On Local Error Resume Next
Get_Cone_Outside_Volume = sBuffer(Buffer).s3DBuffer.GetConeOutsideVolume
End Property

Public Sub Set_Cone_Outside_Volume(Buffer As Integer, New_Volume As Single)
On Local Error Resume Next
sBuffer(Buffer).s3DBuffer.SetConeOutsideVolume New_Volume, DS3D_IMMEDIATE
End Sub

Public Property Get Velocity(Buffer As Integer) As D3DVECTOR
On Local Error Resume Next
sBuffer(Buffer).s3DBuffer.GetPosition Velocity
End Property

Public Property Let Velocity(Buffer As Integer, Velocity As D3DVECTOR)
On Local Error Resume Next
sBuffer(Buffer).s3DBuffer.SetPosition Velocity.X, Velocity.Y, Velocity.Z, DS3D_IMMEDIATE
End Property

Public Sub Set_Buffer_Velocity(Buffer As Integer, X As Single, Y As Single, Z As Single)
On Local Error Resume Next
sBuffer(Buffer).s3DBuffer.SetVelocity X, Y, Z, DS3D_IMMEDIATE
End Sub

Public Sub Duplicate_Sound_Buffer(DestBuffers() As DirectSoundSecondaryBuffer8, Current_Buffer As Integer)
On Local Error Resume Next
Dim i As Integer
For i = 0 To UBound(DestBuffers)
    Set DestBuffers(i) = DXSound.DuplicateSoundBuffer(sBuffer(Current_Buffer).Buffer)
Next
End Sub

Public Sub Save_Buffer_To_File(Buffer As Integer, Filename As String)
On Local Error Resume Next
sBuffer(Buffer).Buffer.SaveToFile Filename
End Sub

Public Function GetSpeakerConfig() As String
On Local Error Resume Next
Select Case DXSound.GetSpeakerConfig
 Case DSSPEAKER_HEADPHONE
  GetSpeakerConfig = "Stereo Headphone"
 Case DSSPEAKER_MONO
  GetSpeakerConfig = "Laptop Mono"
 Case DSSPEAKER_QUAD
  GetSpeakerConfig = "Quadraphonic Speakers"
 Case DSSPEAKER_SURROUND
  GetSpeakerConfig = "Surround Sound"
 Case DSSPEAKER_STEREO
  GetSpeakerConfig = "Desktop Stereo"
 Case Else
  GetSpeakerConfig = "Desktop Stereo"
End Select
End Function

Public Sub SetSpeakerConfig(SpeakerStyle As SpeakerStyle)
On Local Error Resume Next
Select Case SpeakerStyle
Case 0
  DXSound.SetSpeakerConfig DSSPEAKER_HEADPHONE
 Case 1
  DXSound.SetSpeakerConfig DSSPEAKER_STEREO
 Case 2
  DXSound.SetSpeakerConfig DSSPEAKER_MONO
 Case 3
  DXSound.SetSpeakerConfig DSSPEAKER_SURROUND
 Case 4
  DXSound.SetSpeakerConfig DSSPEAKER_QUAD
 Case Else
  DXSound.SetSpeakerConfig DSSPEAKER_STEREO
End Select
End Sub
