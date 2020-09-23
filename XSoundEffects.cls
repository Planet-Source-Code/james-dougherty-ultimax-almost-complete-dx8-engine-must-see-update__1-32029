VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "XSoundEffects"
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

Private Eng As New XEngine3D
Private Perf As DirectMusicPerformance8
Private Loader As DirectMusicLoader8
Private Segment As DirectMusicSegment8
Private Path As DirectMusicAudioPath8
Private State As DirectMusicSegmentState8
Private aEffects As Long
Private SounfFXFolder As String
Private HoldFileName As String
Private GoAhead As Boolean
Private DistFX As DirectSoundFXDistortion8
Private EchoFX As DirectSoundFXEcho8
Private ChorusFX As DirectSoundFXChorus8
Private FlanFX As DirectSoundFXFlanger8
Private CompFX As DirectSoundFXCompressor8
Private GargleFX As DirectSoundFXGargle8
Private PEQFX As DirectSoundFXParamEq8
Private WRevFX As DirectSoundFXWavesReverb8

Public Sub SetSoundFXDir(FolderPath As String)
On Local Error Resume Next
SounfFXFolder = FolderPath & "\"
Loader.SetSearchDirectory SounfFXFolder
End Sub

Public Function IsPlaying() As Boolean
On Local Error Resume Next
IsPlaying = Perf.IsPlaying(Segment, State)
End Function

Public Sub Initialize_SoundFX_Engine(hWnd As Long)
On Error GoTo FailedInit
Dim dmusAudio As DMUS_AUDIOPARAMS

ErrStr = "DirectMusicPerformanceCreate"
Set Perf = Eng.DirectX.DirectMusicPerformanceCreate
ErrStr = "DirectMusicLoaderCreate"
Set Loader = Eng.DirectX.DirectMusicLoaderCreate
    
ErrStr = "InitAudio"
Perf.InitAudio hWnd, DMUS_AUDIOF_EAX Or DMUS_AUDIOF_BUFFERS, dmusAudio, , DMUS_APATH_DYNAMIC_3D, 128
ErrStr = "CreateStandardAudioPath"
Set Path = Perf.CreateStandardAudioPath(DMUS_APATH_DYNAMIC_3D, 64, True)
Exit Sub
    
FailedInit:
 ErrorToFile "ErrorMIDI", ErrStr, PossibleCause
 End
End Sub

Public Sub Load_SoundFX(Wave_Filename As String)
On Local Error GoTo Failed
Dim Filename As String

ErrStr = "SetSearchDirectory"
Loader.SetSearchDirectory SounfFXFolder
ErrStr = "LoadSegment"
Filename = SounfFXFolder & Wave_Filename
Set Segment = Loader.LoadSegment(Filename)
ErrStr = "Download"
Segment.Download Perf.GetDefaultAudioPath
HoldFileName = Wave_Filename
GoAhead = True
Exit Sub

Failed:
 ErrorToFile "ErrorLoadFX", ErrStr, PossibleCause
 GoAhead = False
End Sub

Public Sub Play_SoundFX(Loop_It As Boolean)
On Local Error GoTo Failed
Dim Filename As String

Set Segment = Nothing
Loader.SetSearchDirectory SounfFXFolder
ErrStr = "Load_SoundFX"
Filename = SounfFXFolder & HoldFileName
Set Segment = Loader.LoadSegment(Filename)
Segment.Download Perf.GetDefaultAudioPath

If Not (Segment Is Nothing) And GoAhead = True Then
 ErrStr = "SetRepeats"
 Segment.SetRepeats Loop_It
 ErrStr = "PlaySegmentEx"
 Perf.PlaySegmentEx Segment, DMUS_SEGF_DEFAULT, 0, Nothing, Perf.GetDefaultAudioPath
 ErrStr = "State"
 Set State = Perf.PlaySegmentEx(Segment, DMUS_SEGF_DEFAULT, 0, Nothing, Perf.GetDefaultAudioPath)
End If
Exit Sub

Failed:
 ErrorToFile "ErrorLoadPlay", ErrStr, PossibleCause
End Sub

Public Sub Stop_SoundFX()
On Local Error Resume Next
If Not (Segment Is Nothing) Then Perf.StopEx Segment, 0, DMUS_SEGF_DEFAULT
End Sub

Private Sub Pan_SoundFX(Pan_Value As Single)
On Local Error Resume Next
If Segment Is Nothing Then Exit Sub
Dim Buffer As DirectSound3DBuffer8
    
If Pan_Value >= 10 Then Pan_Value = 10
If Pan_Value <= -10 Then Pan_Value = -10
    
If Not (Path Is Nothing) Then
 Set Buffer = Path.GetObjectinPath(DMUS_PCHANNEL_ALL, DMUS_PATH_BUFFER, 0, GUID_ALL, 0, IID_DirectSound3DBuffer)
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

Public Function Set_Audio_Effect(Effect As SoundFX) As Boolean
On Local Error GoTo errOut
Dim DSEffects() As DSEFFECTDESC
Dim lResults() As Long
Dim lCount As Long
Dim Buffer As DirectSoundSecondaryBuffer8
    
If Segment Is Nothing Then Exit Function
If Perf.IsPlaying(Segment, State) = True Then Perf.StopEx Segment, 0, 0
ReDim DSEffects(7)
ReDim lResults(7)
    
ErrStr = "SetEffect"
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

ErrStr = "SetBuffer"
Set Buffer = Perf.GetDefaultAudioPath.GetObjectinPath(DMUS_PCHANNEL_ALL, DMUS_PATH_BUFFER, 0, GUID_ALL, 0, IID_DirectSoundSecondaryBuffer)
ErrStr = "GetDefaultAudioPath1"
Perf.GetDefaultAudioPath.Activate False
ErrStr = "SetFX"
Buffer.SetFX Effect, DSEffects, lResults
ErrStr = "GetDefaultAudioPath2"
Perf.GetDefaultAudioPath.Activate True
Set_Audio_Effect = True
Exit Function

errOut:
 ErrorToFile "ErrorSetEffect", ErrStr, Err.Description
 Set_Audio_Effect = False
End Function

'Applies to all custom effect changes...
'Generates a generic error file and causes sound not to play?

Private Sub Change_Distortion_Settings(Edge_0_To_100 As Single, Gain_Neg60_To_0 As Single, _
                                      EQBandwidth_100_To_8000 As Single, EQCenterFrequency_100_To_8000 As Single, _
                                      LowpassCutoff_100_To_8000 As Single)
On Local Error Resume Next
If Not Set_Audio_Effect(True) Then Exit Sub
Dim NewFX As DSFXDISTORTION
    
With NewFX
 .fEdge = Edge_0_To_100
 .fGain = Gain_Neg60_To_0
 .fPostEQBandwidth = EQBandwidth_100_To_8000
 .fPostEQCenterFrequency = EQCenterFrequency_100_To_8000
 .fPreLowpassCutoff = LowpassCutoff_100_To_8000
End With

DistFX.SetAllParameters NewFX
End Sub

Private Sub Change_Echo_Settings(Feedback_0_To_100 As Single, Left_Delay_1_To_2000 As Single, _
                                Right_Delay_1_To_2000 As Single, Wet_Dry_Mix_0_To_100 As Single, PanDelay As Boolean)
If Not Set_Audio_Effect(True) Then Exit Sub
Dim NewFX As DSFXECHO
    
With NewFX
 .fFeedback = Feedback_0_To_100
 .fLeftDelay = Left_Delay_1_To_2000
 .fRightDelay = Right_Delay_1_To_2000
 .fWetDryMix = Wet_Dry_Mix_0_To_100
 .lPanDelay = PanDelay
End With

EchoFX.SetAllParameters NewFX
End Sub

Private Sub Change_Chorus_Settings(Feedback_Neg99_To_99 As Single, Delay_0_To_20 As Single, _
                                  Depth_1_To_100 As Single, Wet_Dry_Mix_0_To_100 As Single, _
                                  Frequency_0_To_10 As Single, Phase_0_To_4 As Long, _
                                  WaveForm_Sine As Boolean)
'If Not Set_Audio_Effect(True) Then Exit Sub
Dim NewFX As DSFXCHORUS
    
With NewFX
 .fFeedback = Feedback_Neg99_To_99
 .fDelay = Delay_0_To_20
 .fDepth = Depth_1_To_100
 .fWetDryMix = Wet_Dry_Mix_0_To_100
 .fFrequency = Frequency_0_To_10
 .lPhase = Phase_0_To_4
 If WaveForm_Sine = True Then
  .lWaveform = DSFX_WAVE_SIN
 Else
  .lWaveform = DSFX_WAVE_TRIANGLE
 End If
End With

ChorusFX.SetAllParameters NewFX
End Sub

Private Sub Change_Flanger_Settings(Feedback_Neg99_To_99 As Single, Delay_0_To_4 As Single, _
                                   Depth_1_To_100 As Single, Wet_Dry_Mix_0_To_100 As Single, _
                                   Frequency_0_To_10 As Single, Phase_0_To_4 As Long, _
                                   WaveForm_Sine As Boolean)
If Not Set_Audio_Effect(True) Then Exit Sub
Dim NewFX As DSFXFLANGER
    
With NewFX
 .fFeedback = Feedback_Neg99_To_99
 .fDelay = Delay_0_To_4
 .fDepth = Depth_1_To_100
 .fWetDryMix = Wet_Dry_Mix_0_To_100
 .fFrequency = Frequency_0_To_10
 .lPhase = Phase_0_To_4
 If WaveForm_Sine = True Then
  .lWaveform = DSFX_WAVE_SIN
 Else
  .lWaveform = DSFX_WAVE_TRIANGLE
 End If
End With

FlanFX.SetAllParameters NewFX
End Sub

Private Sub Change_Compressor_Settings(Attack_1_To_500 As Single, Gain_Neg60_To_60 As Single, _
                                      Predelay_0_To_4 As Single, Ratio_1_To_100 As Single, _
                                      Release_50_To_3000 As Single, Threshold_Neg60_To_0 As Long)
If Not Set_Audio_Effect(True) Then Exit Sub
Dim NewFX As DSFXCOMPRESSOR
    
With NewFX
 .fAttack = Attack_1_To_500
 .fGain = Gain_Neg60_To_60
 .fPredelay = Predelay_0_To_4
 .fRatio = Ratio_1_To_100
 .fRelease = Release_50_To_3000
 .fThreshold = Threshold_Neg60_To_0
End With

CompFX.SetAllParameters NewFX
End Sub

Private Sub Change_Gargle_Settings(Rate_Hz_1_To_1000 As Single, WaveForm_Square As Boolean)
If Not Set_Audio_Effect(True) Then Exit Sub
Dim NewFX As DSFXGARGLE
    
With NewFX
 .lRateHz = Rate_Hz_1_To_1000
 If WaveForm_Square = True Then
  .lWaveShape = DSFXGARGLE_WAVE_SQUARE
 Else
  .lWaveShape = DSFXGARGLE_WAVE_TRIANGLE
 End If
End With

GargleFX.SetAllParameters NewFX
End Sub

Private Sub Change_ParamEQ_Settings(Gain_Neg15_To_15 As Single, Bandwidth_1_TO_36 As Single, Center_80_TO_16000 As Single)
If Not Set_Audio_Effect(True) Then Exit Sub
Dim NewFX As DSFXPARAMEQ
    
With NewFX
 .fBandwidth = Bandwidth_1_TO_36
 .fCenter = Center_80_TO_16000
 .fGain = Gain_Neg15_To_15
End With

PEQFX.SetAllParameters NewFX
End Sub

Private Sub Change_Waves_Reverb_Settings(Gain_Neg96_To_0 As Single, Mix_Neg96_To_0 As Single, Time_1_TO_3000 As Single, Ratio_1_TO_999 As Single)
If Not Set_Audio_Effect(True) Then Exit Sub
Dim NewFX As DSFXWAVESREVERB
    
With NewFX
 .fInGain = Gain_Neg96_To_0
 .fHighFreqRTRatio = Ratio_1_TO_999
 .fReverbMix = Mix_Neg96_To_0
 .fReverbTime = Time_1_TO_3000
End With

WRevFX.SetAllParameters NewFX
End Sub

Public Sub Cleanup_SoundFX_Engine()
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
