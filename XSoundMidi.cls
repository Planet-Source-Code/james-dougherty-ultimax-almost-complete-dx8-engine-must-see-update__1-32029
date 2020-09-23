VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "XSoundMidi"
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

Public Enum SoundFX
 Distortion = 0
 Echo = 1
 Chorus = 2
 Flanger = 3
 Compressor = 4
 Gargle = 5
 ParamEQ = 6
 WavesReverb = 7
End Enum

Private Eng As New XEngine3D
Private Perf As DirectMusicPerformance8
Private Loader As DirectMusicLoader8
Private Segment As DirectMusicSegment8
Private Path As DirectMusicAudioPath8
Private State As DirectMusicSegmentState8
Private aEffects As Long
Private MidiFolder As String

Public Sub SetMidiDir(FolderPath As String)
On Local Error Resume Next
MidiFolder = FolderPath & "\"
Loader.SetSearchDirectory MidiFolder
End Sub

Public Function IsPlaying() As Boolean
On Local Error Resume Next
IsPlaying = Perf.IsPlaying(Segment, State)
End Function

Public Sub Initialize_Midi_Engine(hWnd As Long)
On Error GoTo FailedInit
Dim dmusAudio As DMUS_AUDIOPARAMS

ErrStr = "DirectMusicPerformanceCreate"
Set Perf = Eng.DirectX.DirectMusicPerformanceCreate
ErrStr = "DirectMusicLoaderCreate"
Set Loader = Eng.DirectX.DirectMusicLoaderCreate
    
ErrStr = "InitAudio"
Perf.InitAudio hWnd, DMUS_AUDIOF_ALL, dmusAudio, Nothing, DMUS_APATH_DYNAMIC_3D, 128
ErrStr = "CreateStandardAudioPath"
Set Path = Perf.CreateStandardAudioPath(DMUS_APATH_DYNAMIC_3D, 64, True)
Exit Sub
    
FailedInit:
 ErrorToFile "ErrorMIDI", ErrStr, PossibleCause
 End
End Sub

Public Sub Load_Midi(Filename As String)
On Local Error GoTo Failed
Dim tmpString As String

Set Segment = Nothing
tmpString = MidiFolder & Filename
Set Segment = Loader.LoadSegment(tmpString)
If (Right$(Filename, 4) = ".mid") Or (Right$(Filename, 4) = ".rmi") Or (Right$(Filename, 5) = ".midi") Then
 Segment.SetStandardMidiFile
End If
Segment.Download Perf.GetDefaultAudioPath
Exit Sub
Failed:
 ErrorToFile "ErrorLoadMIDI", ErrStr, PossibleCause
 End
End Sub

Public Sub Play_Midi(Loop_It As Boolean)
On Local Error Resume Next
If Not (Segment Is Nothing) Then
 Segment.SetRepeats Loop_It
 Perf.PlaySegmentEx Segment, DMUS_SEGF_DEFAULT, 0
 Set State = Perf.PlaySegmentEx(Segment, DMUS_SEGF_DEFAULT, 0, , Perf.GetDefaultAudioPath)
End If
End Sub

Public Sub Stop_Midi()
On Local Error Resume Next
If Not (Segment Is Nothing) Then Perf.StopEx Segment, 0, DMUS_SEGF_DEFAULT
End Sub

Private Sub Pan_Midi(Pan_Value As Single)
On Local Error Resume Next
If Segment Is Nothing Then Exit Sub
Dim Buffer As DirectSound3DBuffer8
    
If Pan_Value >= 10 Then Pan_Value = 10
If Pan_Value <= -10 Then Pan_Value = -10
    
If Not (Path Is Nothing) Then
 Set Buffer = Path.GetObjectinPath(DMUS_PCHANNEL_ALL, DMUS_PATH_BUFFER, 0, vbNullString, 0, "IID_IDirectSound3DBuffer")
 If Not (Buffer Is Nothing) Then
  Buffer.SetPosition Pan_Value, 0, 0, DS3D_IMMEDIATE
 End If
End If

Set Buffer = Nothing
End Sub

Public Sub VolumeLevel(Volume As Long)
On Local Error GoTo errOut

ErrStr = "VolumeLevel"
If Volume > 0 Then
 Perf.SetMasterVolume (50 * Volume) - 5000
Else
 Perf.SetMasterVolume -5000
End If
Exit Sub

errOut:
 ErrorToFile "ErrorVolumeLvL", ErrStr, Err.Description
End Sub

Public Sub Set_Tempo(Optional Tempo As Single = 30)
On Local Error GoTo errOut

If Tempo >= 30 Then Tempo = 30
If Tempo <= 0 Then Tempo = 0

ErrStr = "Set_Tempo"
Perf.SetMasterTempo (Tempo / 10)
Exit Sub

errOut:
 ErrorToFile "ErrorVolumeLvL", ErrStr, Err.Description
End Sub

Public Sub Set_Groove_Level(Groove As Integer)
On Local Error GoTo errOut

ErrStr = "Groove"
Perf.SetMasterGrooveLevel Groove
Exit Sub

errOut:
 ErrorToFile "ErrorGroove", ErrStr, Err.Description
End Sub

Public Sub Set_Audio_Effect(Effect As SoundFX)
On Local Error GoTo errOut
Dim DSEffects() As DSEFFECTDESC
Dim lResults() As Long
Dim lCount As Long
Dim Buffer As DirectSoundSecondaryBuffer8
    
If Segment Is Nothing Then Exit Sub
If Perf.IsPlaying(Segment, State) = True Then Perf.StopEx Segment, 0, 0
ReDim DSEffects(7)
ReDim lResults(7)
    
For lCount = 0 To 7
 Select Case Effect
  Case 0
   DSEffects(lCount).guidDSFXClass = DSFX_STANDARD_DISTORTION
  Case 1
   DSEffects(lCount).guidDSFXClass = DSFX_STANDARD_ECHO
  Case 2
   DSEffects(lCount).guidDSFXClass = DSFX_STANDARD_CHORUS
  Case 3
   DSEffects(lCount).guidDSFXClass = DSFX_STANDARD_FLANGER
  Case 4
   DSEffects(lCount).guidDSFXClass = DSFX_STANDARD_COMPRESSOR
  Case 5
   DSEffects(lCount).guidDSFXClass = DSFX_STANDARD_GARGLE
  Case 6
   DSEffects(lCount).guidDSFXClass = DSFX_STANDARD_PARAMEQ
  Case 7
   DSEffects(lCount).guidDSFXClass = DSFX_STANDARD_WAVES_REVERB
 End Select
Next

Set Buffer = Perf.GetDefaultAudioPath.GetObjectinPath(DMUS_PCHANNEL_ALL, DMUS_PATH_BUFFER, 0, GUID_ALL, 0, IID_DirectSoundSecondaryBuffer)
Perf.GetDefaultAudioPath.Activate False
Buffer.SetFX Effect, DSEffects, lResults
Perf.GetDefaultAudioPath.Activate True
Exit Sub

errOut:
 ErrorToFile "ErrorGroove", ErrStr, Err.Description
End Sub

Public Sub Cleanup_Midi_Engine()
On Local Error Resume Next
Perf.RemoveNotificationType DMUS_NOTIFY_ON_SEGMENT
If Not (Segment Is Nothing) Then Perf.StopEx Perf.GetDefaultAudioPath, 0, 0
Segment.Unload Perf.GetDefaultAudioPath
Set Segment = Nothing
Set Loader = Nothing
Perf.CloseDown
Set Path = Nothing
Set Perf = Nothing
Set Eng = Nothing
End Sub
