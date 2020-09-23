VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "XSound2D"
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

Private Type dxBuffers
  isLoaded As Boolean
  Buffer As DirectSoundSecondaryBuffer8
End Type

Public Enum SpeakerStyle
 Stereo_Headphone = 0
 Desktop_Stereo = 1
 Laptop_Mono = 2
 Surround_Sound = 3
 Quadraphonic_Speakers = 4
End Enum

Private DX As New DirectX8
Private DXSound As DirectSound8
Private SoundFolder As String
Private sBuffer() As dxBuffers
Private CurrentBuffer As Integer
 
Public Function Initialize_Sound_Engine(hWnd As Long) As Boolean
On Local Error GoTo errOut
ErrStr = "DirectX DirectSoundCreate"
Set DXSound = DX.DirectSoundCreate("")

ErrStr = "DXSound SetCooperativeLevel"
DXSound.SetCooperativeLevel hWnd, DSSCL_PRIORITY
Initialize_Sound_Engine = True
Exit Function

errOut:
 ErrorToFile "ErrorDS", ErrStr, Err.Description
End Function

Public Function IsPlaying(Buffer As Integer) As Long
On Local Error Resume Next
IsPlaying = sBuffer(Buffer).Buffer.GetStatus
End Function

Public Sub SetSoundDir(FolderPath As String)
On Local Error Resume Next
SoundFolder = FolderPath & "\"
End Sub

Public Sub CreateBuffers(AmountOfBuffers As Integer, DefaultFile As String)
On Local Error GoTo errOut

ReDim sBuffer(AmountOfBuffers)
For AmountOfBuffers = 0 To AmountOfBuffers
 ErrStr = "CreateBuffers - DX7LoadSound"
 LoadSound AmountOfBuffers, DefaultFile
Next AmountOfBuffers
Exit Sub

errOut:
 ErrorToFile "ErrorCB", ErrStr, Err.Description
End Sub

Public Sub PlaySound(Buffer As Integer, Optional Volume As Long = 75, Optional PanValue As Byte = 50, Optional FreqValue As Long = 44100, Optional LoopIt As Byte = 0)
On Local Error GoTo errOut

ErrStr = "PlaySoundWithPan - LoadSound"
If PanValue <> 50 And PanValue < 101 Then PanSound Buffer, PanValue
If Volume < 101 Then VolumeLevel Buffer, Volume
If FreqValue > 100 And FreqValue < 100001 Then FrequencyLevel Buffer, FreqValue
If sBuffer(Buffer).isLoaded Then sBuffer(Buffer).Buffer.Play LoopIt
Exit Sub

errOut:
 ErrorToFile "ErrorPlaySoundWP", ErrStr, Err.Description
End Sub

Public Sub StopSound(Buffer As Integer)
On Local Error GoTo errOut

ErrStr = "StopSound"
If sBuffer(Buffer).Buffer.GetStatus = DSBSTATUS_PLAYING Or DSBSTATUS_LOOPING Then
 sBuffer(Buffer).Buffer.Stop
 sBuffer(Buffer).Buffer.SetCurrentPosition 0
End If

Exit Sub

errOut:
 ErrorToFile "ErrorStopSound", ErrStr, Err.Description
End Sub

Public Sub PauseSound(Buffer As Integer)
On Local Error GoTo errOut

ErrStr = "PauseSound"
If sBuffer(Buffer).Buffer.GetStatus = DSBSTATUS_PLAYING Or DSBSTATUS_LOOPING Then
 sBuffer(Buffer).Buffer.Stop
End If
Exit Sub

errOut:
 ErrorToFile "ErrorPauseSound", ErrStr, Err.Description
End Sub

Public Sub PanSound(Buffer As Integer, PanValue As Byte)
On Local Error GoTo errOut

ErrStr = "PanValue"
Select Case PanValue
Case 0
 sBuffer(Buffer).Buffer.SetPan -10000
Case 100
 sBuffer(Buffer).Buffer.SetPan 10000
Case Else
 sBuffer(Buffer).Buffer.SetPan (100 * PanValue) - 5000
End Select
Exit Sub

errOut:
 ErrorToFile "ErrorPanSound2D", ErrStr, Err.Description
End Sub

Public Sub FrequencyLevel(Buffer As Integer, FreqValue As Long)
On Local Error GoTo errOut

ErrStr = "FreqValue"
Dim i As Long
i = FreqValue
If i <= 100 Then i = 101
If i >= 100000 Then i = 100000
If Not (sBuffer(Buffer).Buffer Is Nothing) Then sBuffer(Buffer).Buffer.SetFrequency FreqValue
Exit Sub

errOut:
 ErrorToFile "ErrorFrequencyLevel2D", ErrStr, Err.Description
End Sub

Public Sub VolumeLevel(Buffer As Integer, Volume As Long)
On Local Error GoTo errOut

ErrStr = "VolumeLevel"
If Volume > 0 Then
 sBuffer(Buffer).Buffer.SetVolume (50 * Volume) - 5000
Else
 sBuffer(Buffer).Buffer.SetVolume -5000
End If
Exit Sub

errOut:
 ErrorToFile "ErrorVolumeLvL2D", ErrStr, Err.Description
End Sub

Public Sub LoadSound(Buffer As Integer, sFile As String)
On Local Error GoTo errOut
Dim Filename As String
Dim BufferDesc As DSBUFFERDESC
  
BufferDesc.lFlags = DSBCAPS_CTRLFREQUENCY Or DSBCAPS_CTRLPAN _
                    Or DSBCAPS_CTRLVOLUME Or DSBCAPS_STATIC
  
Filename = SoundFolder & sFile
ErrStr = "CreateSoundBufferFromFile"
Set sBuffer(Buffer).Buffer = DXSound.CreateSoundBufferFromFile(Filename, BufferDesc)
sBuffer(Buffer).isLoaded = True
Exit Sub

errOut:
 ErrorToFile "ErrorLoadSound", ErrStr, Err.Description
End Sub

Public Sub Cleanup_Sound_Engine(NumberOfBuffers As Integer)
On Local Error Resume Next
Static i As Integer
For i = 0 To NumberOfBuffers
 Set sBuffer(i).Buffer = Nothing
Next
Set DXSound = Nothing
Set DX = Nothing
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
