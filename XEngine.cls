VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "XEngine3D"
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

Public Enum FilterType
 None = 0
 Dark_Map = 1
 Light_Map = 2
 Detail_Map = 3
 Gloss_Map = 4
 Detail_Modulate = 5
 Detail_Add = 6
 Add_Smooth_Saturation = 7
 Linear_Blend = 8
 Inverse_Light_Map = 9
End Enum

Public Enum TIMER_COMMAND
 TIMER_RESET = 1
 TIMER_start = 2
 TIMER_STOP = 3
 TIMER_ADVANCE = 4
 TIMER_GETABSOLUTETIME = 5
 TIMER_GETAPPTIME = 6
 TIMER_GETELLAPSEDTIME = 7
End Enum

Private Type TLVERTEX
    X As Single
    Y As Single
    Z As Single
    RHW As Single
    Color As Long
    Specular As Long
    tU As Single
    tV As Single
End Type
Private Const FVF = D3DFVF_XYZRHW Or D3DFVF_TEX1 Or D3DFVF_DIFFUSE Or D3DFVF_SPECULAR

Private DX8 As New DirectX8
Private D3DX As New D3DX8
Private D3D As Direct3D8
Private HoldCaps As D3DCAPS8
Private Math As New XMath
Private HoldD3DPPWIN As D3DPRESENT_PARAMETERS
Private HoldD3DPPFS As D3DPRESENT_PARAMETERS
Private d3dpp As D3DPRESENT_PARAMETERS
Private HoldWidth As Long
Private HoldHeight As Long
Private HoldhWnd As Long
Private FullS As Boolean
Private InitSucessfull As Boolean
Private vbFont As IFont
Private fnt As New StdFont
Private xFont As D3DXFont
Private TextInitialized As Boolean
Private RenderColor As Long
Private fps As Single
Private BackVerts(3) As TLVERTEX
Private Background_Picture As Direct3DTexture8
Private HasBackground As Boolean
Private DoDark As Boolean
Private DoDetail As Boolean
Private DoLight As Boolean
Private DoGloss As Boolean
Private DoDMod As Boolean
Private DoDAdd As Boolean
Private DoASat As Boolean
Private DoLBlend As Boolean
Private DoNone As Boolean
Private MapTexture As Direct3DTexture8
Private Declare Function GetClientRect Lib "user32.dll" (ByVal hWnd As Long, ByRef RECT As RECT) As Long

'|œœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœ|
'|œœ                     User Access                        œœ|
'|œœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœ|

Public Function DirectX() As DirectX8
 Set DirectX = DX8
End Function

Public Function Direct3DX() As D3DX8
 Set Direct3DX = D3DX
End Function

Public Function Direct3D() As Direct3D8
 Set Direct3D = D3D
End Function

Public Function Direct3DDevice() As Direct3DDevice8
 Set Direct3DDevice = D3DD
End Function

'|œœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœ|
'|œœ                     Main Setup                         œœ|
'|œœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœ|

Private Function Initialize_Win(hWnd As Long) As Boolean
 On Local Error GoTo errOut
 ErrStr = ""
 If DX8 Is Nothing Then Set DX8 = New DirectX8
 If D3DX Is Nothing Then Set D3DX = New D3DX8
 If D3D Is Nothing Then Set D3D = DX8.Direct3DCreate
 
 InitSucessfull = False
 
 Dim Mode As D3DDISPLAYMODE
 ErrStr = "GetAdapterDisplayMode"
 D3D.GetAdapterDisplayMode D3DADAPTER_DEFAULT, Mode
         
 d3dpp.Windowed = 1
 d3dpp.SwapEffect = D3DSWAPEFFECT_DISCARD
 d3dpp.BackBufferFormat = Mode.format
 d3dpp.BackBufferCount = 1
 d3dpp.EnableAutoDepthStencil = 1
 d3dpp.AutoDepthStencilFormat = D3DFMT_D16
 d3dpp.hDeviceWindow = hWnd
 
 ErrStr = "CreateDevice"
 Set D3DD = D3D.CreateDevice(D3DADAPTER_DEFAULT, D3DDEVTYPE_HAL, hWnd, D3DCREATE_SOFTWARE_VERTEXPROCESSING, d3dpp)
    
 D3DD.SetRenderState D3DRS_ZENABLE, D3DZB_TRUE
 D3DD.GetDeviceCaps HoldCaps
 HoldD3DPPWIN = d3dpp
 Initialize_Win = True
 FullS = False
 InitSucessfull = True
 Exit Function
 
errOut:
 ErrorToFile "ErrorWIN", ErrStr, PossibleCause
 End
End Function

Public Function Restore_Window_Device(hWnd As Long) As Boolean
 On Local Error GoTo errOut
 If DX8 Is Nothing Then Exit Function
 If D3DX Is Nothing Then Exit Function
 If D3D Is Nothing Then Exit Function
 
 d3dpp.BackBufferWidth = 0
 d3dpp.BackBufferHeight = 0
 d3dpp.hDeviceWindow = hWnd
 d3dpp.Windowed = 1
 Restore_Window_Device = True
 Exit Function
 
errOut:
 ErrorToFile "Error Restore", ErrStr, PossibleCause
 End
End Function

Private Function Initialize_FS(hWnd As Long, Width As Long, Height As Long) As Boolean
 On Local Error GoTo errOut
 ErrStr = ""
 Dim Mode As D3DDISPLAYMODE
 Dim rc As RECT
 Dim Ret As Long
 Dim DType As Long
 Dim Caps As D3DCAPS8
 Dim format As Long
 If DX8 Is Nothing Then Set DX8 = New DirectX8
 If D3DX Is Nothing Then Set D3DX = New D3DX8
 If D3D Is Nothing Then Set D3D = DX8.Direct3DCreate
    
 InitSucessfull = False
     
 ErrStr = "GetDeviceCaps"
 DType = D3DDEVTYPE_HAL
 D3D.GetDeviceCaps D3DADAPTER_DEFAULT, DType, Caps
 If Err.Number Then
  Err.Clear
  DType = D3DDEVTYPE_REF
  D3D.GetDeviceCaps D3DADAPTER_DEFAULT, DType, Caps
 End If
 
 ErrStr = "GetAdapterDisplayMode"
 D3D.GetAdapterDisplayMode D3DADAPTER_DEFAULT, Mode
 ErrStr = "Get_Display_Mode"
 format = Mode.format
 Ret = Get_Display_Mode(Width, Height, format)
 
 d3dpp.Windowed = 0
 d3dpp.SwapEffect = D3DSWAPEFFECT_FLIP
 d3dpp.BackBufferFormat = Mode.format
 d3dpp.BackBufferCount = 1
 d3dpp.BackBufferWidth = Width
 d3dpp.BackBufferHeight = Height
 d3dpp.EnableAutoDepthStencil = 1
 d3dpp.AutoDepthStencilFormat = D3DFMT_D16
    
 ErrStr = "CreateDevice"
 Set D3DD = D3D.CreateDevice(D3DADAPTER_DEFAULT, DType, hWnd, D3DCREATE_SOFTWARE_VERTEXPROCESSING, d3dpp)
 
 D3DD.SetRenderState D3DRS_ZENABLE, D3DZB_TRUE
 D3DD.GetDeviceCaps HoldCaps
 HoldD3DPPFS = d3dpp
 Initialize_FS = True
 HoldWidth = Width
 HoldHeight = Height
 FullS = True
 InitSucessfull = True
 Exit Function
 
errOut:
 ErrorToFile "ErrorFS", ErrStr, PossibleCause
 End
End Function

Private Function Get_Display_Mode(ByVal Width As Long, ByVal Height As Long, format As Long) As Long
Dim i  As Long
Dim Count As Long
Dim d3ddm As D3DDISPLAYMODE
Dim FoundMode As Boolean
    
i = 0
Count = D3D.GetAdapterModeCount(D3DADAPTER_DEFAULT) - 1
    
If Err.Number Then
    Get_Display_Mode = D3DERR_INVALIDCALL
    Exit Function
End If
    
For i = 0 To Count
        
Call D3D.EnumAdapterModes(D3DADAPTER_DEFAULT, i, d3ddm)
 If Err.Number Then
    Get_Display_Mode = Err.Number
    Exit Function
 End If
         
 If d3ddm.Width = Width And d3ddm.Height = Height Then
    If d3ddm.format = D3DFMT_R8G8B8 Or _
        d3ddm.format = D3DFMT_R5G6B5 Then
        FoundMode = True
        format = d3ddm.format
        Exit For
    End If
 End If
Next
        
If FoundMode Then
    Exit Function
Else
    Get_Display_Mode = D3DERR_INVALIDDEVICE
End If
    
End Function

Public Function Restore_FullScreen_Device(hWnd As Long) As Boolean
 On Local Error GoTo errOut
 Initialize_FS hWnd, HoldWidth, HoldHeight
 Setup_Default_Rendering_State
 Restore_FullScreen_Device = True
 Exit Function
    
errOut:
 ErrorToFile "Error Restore", ErrStr, PossibleCause
 End
End Function

Private Function Reset_Device(hWnd As Long) As Long
If FullS = True Then
 Restore_FullScreen_Device hWnd
Else
 Restore_Window_Device hWnd
End If
End Function

Public Function Initialize_EngineX8(hWnd As Long, FullScreen As Boolean, Optional Width As Long = 640, Optional Height As Long = 480) As Boolean
 On Local Error GoTo errOut
 ErrStr = ""
 If DX8 Is Nothing Then Set DX8 = New DirectX8
 If D3DX Is Nothing Then Set D3DX = New D3DX8
 
 ErrStr = "Direct3DCreate"
 Set D3D = DX8.Direct3DCreate()
 If D3D Is Nothing Then Exit Function
 
 HasBackground = False
 
 If FullS = True Then
  ErrStr = "Initialize_FS"
  Initialize_FS hWnd, Width, Height
 Else
  ErrStr = "Initialize_Win"
  Initialize_Win hWnd
 End If
 HoldhWnd = hWnd
 Initialize_EngineX8 = True
 Exit Function
 
errOut:
 ErrorToFile "ErrorInit", ErrStr, PossibleCause
 End
End Function

Public Sub Cleanup_XEngine8()
 Set Background_Picture = Nothing
 Set MapTexture = Nothing
 Set Math = Nothing
 Set xFont = Nothing
 Set D3DD = Nothing
 Set D3D = Nothing
 Set D3DX = Nothing
 Set DX8 = Nothing
End Sub

'|œœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœ|
'|œœ                     Rendering                          œœ|
'|œœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœ|

Public Function Start_Engine_Render(Optional Clear_Color As Long)
 On Local Error Resume Next
 If D3DD Is Nothing Then Exit Function
 Dim Res As Long
    
 Res = D3DD.TestCooperativeLevel
    
 If Res = D3DERR_DEVICELOST Then
    Exit Function
 ElseIf Res = D3DERR_DEVICENOTRESET Then
    Res = 0
    Res = Reset_Device(HoldhWnd)
    If Res Then Exit Function
 End If
 Update_FPS
 D3DD.Clear 0, ByVal 0, D3DCLEAR_TARGET Or D3DCLEAR_ZBUFFER, Clear_Color, 1, 0
 D3DD.BeginScene
 Render_Background_Picture
End Function

Public Function End_Engine_Render()
 On Local Error Resume Next
 D3DD.EndScene
 D3DD.Present ByVal 0, ByVal 0, 0, ByVal 0
End Function

Public Function Set_Background_Color(Red As Single, Green As Single, Blue As Single)
 Dim Color As D3DCOLORVALUE
 Color.r = Red: Color.g = Green: Color.b = Blue: Color.A = 1
 
 RenderColor = D3DCOLORVALUEtoLONG(Color)
End Function

Public Function Set_Background_Picture(Filename As String, Filter_Type As FilterType, Optional FilterMapFile As String, Optional Texture_U_Power = 1, Optional Texture_V_Power = 1)
Dim ScRect As RECT
Dim Screen_Height As Long
Dim Screen_Width As Long

Call GetClientRect(HoldhWnd, ScRect)
Screen_Height = ScRect.bottom - ScRect.Top
Screen_Width = ScRect.Right - ScRect.Left

BackVerts(0).X = 0: BackVerts(0).Y = Screen_Height
BackVerts(1).X = 0: BackVerts(1).Y = 0
BackVerts(2).X = Screen_Width: BackVerts(2).Y = Screen_Height
BackVerts(3).X = Screen_Width: BackVerts(3).Y = 0
BackVerts(0).tU = 0: BackVerts(0).tV = Texture_V_Power: BackVerts(0).RHW = 1: BackVerts(0).Color = &HFFFFFF
BackVerts(1).tU = 0: BackVerts(1).tV = 0: BackVerts(1).RHW = 1: BackVerts(1).Color = &HFFFFFF
BackVerts(2).tU = Texture_U_Power: BackVerts(2).tV = Texture_V_Power: BackVerts(2).RHW = 1: BackVerts(2).Color = &HFFFFFF
BackVerts(3).tU = Texture_U_Power: BackVerts(3).tV = 0: BackVerts(3).RHW = 1: BackVerts(3).Color = &HFFFFFF

Set Background_Picture = D3DX.CreateTextureFromFile(D3DD, Filename)

If Filter_Type = Dark_Map And FilterMapFile <> "" Then
 Set MapTexture = D3DX.CreateTextureFromFile(D3DD, FilterMapFile)
 DoDark = True
End If
If Filter_Type = Light_Map And FilterMapFile <> "" Then
 Set MapTexture = D3DX.CreateTextureFromFile(D3DD, FilterMapFile)
 DoLight = True
End If
If Filter_Type = Detail_Map And FilterMapFile <> "" Then
 Set MapTexture = D3DX.CreateTextureFromFile(D3DD, FilterMapFile)
 DoDetail = True
End If
If Filter_Type = Gloss_Map And FilterMapFile <> "" Then
 Set MapTexture = D3DX.CreateTextureFromFile(D3DD, FilterMapFile)
 DoGloss = True
End If
If Filter_Type = Detail_Modulate And FilterMapFile <> "" Then
 Set MapTexture = D3DX.CreateTextureFromFile(D3DD, FilterMapFile)
 DoDMod = True
End If
If Filter_Type = Detail_Add And FilterMapFile <> "" Then
 Set MapTexture = D3DX.CreateTextureFromFile(D3DD, FilterMapFile)
 DoDAdd = True
End If
If Filter_Type = Add_Smooth_Saturation And FilterMapFile <> "" Then
 Set MapTexture = D3DX.CreateTextureFromFile(D3DD, FilterMapFile)
 DoASat = True
End If
If Filter_Type = Linear_Blend And FilterMapFile <> "" Then
 Set MapTexture = D3DX.CreateTextureFromFile(D3DD, FilterMapFile)
 DoLBlend = True
End If
If Filter_Type = None Then
 Set MapTexture = Nothing
 DoNone = True
 DoDark = False
 DoDetail = False
 DoLight = False
 DoGloss = False
 DoDMod = False
 DoASat = False
 DoDAdd = False
 DoLBlend = False
End If
HasBackground = True
End Function

Private Sub Render_Background_Picture()
Dim matView As D3DMATRIX, matViewSave As D3DMATRIX
If D3DD Is Nothing Then Exit Sub

D3DD.GetTransform D3DTS_VIEW, matViewSave
matView = matViewSave
matView.m41 = 0: matView.m42 = -0.002: matView.m43 = 0
D3DD.SetTransform D3DTS_VIEW, matView
D3DD.SetRenderState D3DRS_ZENABLE, 0
D3DD.SetVertexShader FVF
D3DD.SetTexture 0, Background_Picture

If HasBackground = True Then
 If DoDark = True Then
  D3DD.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, BackVerts(0), Len(BackVerts(0))
  D3DD.SetTextureStageState 0, D3DTSS_TEXCOORDINDEX, 0
  D3DD.SetTextureStageState 0, D3DTSS_COLORARG1, D3DTA_TEXTURE
  D3DD.SetTextureStageState 0, D3DTSS_ALPHAOP, D3DTOP_SELECTARG1
  D3DD.SetTexture 1, MapTexture
  D3DD.SetTextureStageState 1, D3DTSS_TEXCOORDINDEX, 0
  D3DD.SetTextureStageState 1, D3DTSS_COLORARG1, D3DTA_TEXTURE
  D3DD.SetTextureStageState 1, D3DTSS_COLORARG2, D3DTA_CURRENT
  D3DD.SetTextureStageState 1, D3DTSS_COLOROP, D3DTOP_MODULATE
  D3DD.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, BackVerts(0), Len(BackVerts(0))
  D3DD.SetTexture 0, Nothing
  D3DD.SetTexture 1, Nothing
 End If
 If DoLight = True Then
  D3DD.SetTextureStageState 0, D3DTSS_TEXCOORDINDEX, 0
  D3DD.SetTextureStageState 0, D3DTSS_COLORARG1, D3DTA_TEXTURE
  D3DD.SetTextureStageState 0, D3DTSS_ALPHAOP, D3DTOP_SELECTARG1
  D3DD.SetTexture 1, MapTexture
  D3DD.SetTextureStageState 1, D3DTSS_TEXCOORDINDEX, 0
  D3DD.SetTextureStageState 1, D3DTSS_COLORARG1, D3DTA_TEXTURE
  D3DD.SetTextureStageState 1, D3DTSS_COLORARG2, D3DTA_CURRENT
  D3DD.SetTextureStageState 1, D3DTSS_COLOROP, D3DTOP_ADD
  D3DD.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, BackVerts(0), Len(BackVerts(0))
  D3DD.SetTexture 0, Nothing
  D3DD.SetTexture 1, Nothing
 End If
 If DoDetail = True Then
  D3DD.SetTextureStageState 0, D3DTSS_TEXCOORDINDEX, 0
  D3DD.SetTextureStageState 0, D3DTSS_COLORARG1, D3DTA_TEXTURE
  D3DD.SetTextureStageState 0, D3DTSS_ALPHAOP, D3DTOP_SELECTARG1
  D3DD.SetTexture 1, MapTexture
  D3DD.SetTextureStageState 1, D3DTSS_TEXCOORDINDEX, 0
  D3DD.SetTextureStageState 1, D3DTSS_COLORARG1, D3DTA_TEXTURE
  D3DD.SetTextureStageState 1, D3DTSS_COLORARG2, D3DTA_CURRENT
  D3DD.SetTextureStageState 1, D3DTSS_COLOROP, D3DTOP_ADDSIGNED
  D3DD.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, BackVerts(0), Len(BackVerts(0))
  D3DD.SetTexture 0, Nothing
  D3DD.SetTexture 1, Nothing
 End If
 If DoGloss = True Then '2 Pass's
  D3DD.SetTextureStageState 0, D3DTSS_TEXCOORDINDEX, 0
  D3DD.SetTextureStageState 0, D3DTSS_COLORARG1, D3DTA_TEXTURE
  D3DD.SetTextureStageState 0, D3DTSS_COLOROP, D3DTOP_SELECTARG1
  D3DD.SetTexture 1, MapTexture
  D3DD.SetTextureStageState 1, D3DTSS_TEXCOORDINDEX, 0
  D3DD.SetTextureStageState 1, D3DTSS_COLORARG1, D3DTA_DIFFUSE
  D3DD.SetTextureStageState 1, D3DTSS_COLORARG2, D3DTA_CURRENT
  D3DD.SetTextureStageState 1, D3DTSS_COLOROP, D3DTOP_MODULATE
  D3DD.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, BackVerts(0), Len(BackVerts(0))
  D3DD.SetTexture 0, MapTexture
  D3DD.SetTextureStageState 0, D3DTSS_TEXCOORDINDEX, 0
  D3DD.SetTextureStageState 0, D3DTSS_COLORARG1, D3DTA_TEXTURE
  D3DD.SetTextureStageState 0, D3DTSS_COLOROP, D3DTOP_SELECTARG1
  D3DD.SetTexture 1, MapTexture
  D3DD.SetTextureStageState 1, D3DTSS_TEXCOORDINDEX, 0
  D3DD.SetTextureStageState 1, D3DTSS_COLORARG1, D3DTA_TEXTURE
  D3DD.SetTextureStageState 1, D3DTSS_COLORARG2, D3DTA_CURRENT
  D3DD.SetTextureStageState 1, D3DTSS_COLOROP, D3DTOP_MODULATE
  D3DD.SetRenderState D3DRS_ALPHABLENDENABLE, 1
  D3DD.SetRenderState D3DRS_SRCBLEND, D3DBLEND_ONE
  D3DD.SetRenderState D3DRS_DESTBLEND, D3DBLEND_ONE
  D3DD.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, BackVerts(0), Len(BackVerts(0))
  D3DD.SetRenderState D3DRS_DESTBLEND, D3DBLEND_ZERO
  D3DD.SetRenderState D3DRS_SRCBLEND, D3DBLEND_ZERO
  D3DD.SetRenderState D3DRS_ALPHABLENDENABLE, 0
  D3DD.SetTexture 0, Nothing
  D3DD.SetTexture 1, Nothing
 End If
 If DoDMod = True Then
  D3DD.SetTextureStageState 0, D3DTSS_TEXCOORDINDEX, 0
  D3DD.SetTextureStageState 0, D3DTSS_COLORARG1, D3DTA_TEXTURE
  D3DD.SetTextureStageState 0, D3DTSS_COLOROP, D3DTOP_MODULATE
  D3DD.SetTextureStageState 0, D3DTSS_COLORARG2, D3DTA_DIFFUSE
  D3DD.SetTextureStageState 0, D3DTSS_ALPHAARG1, D3DTA_TEXTURE
  D3DD.SetTextureStageState 0, D3DTSS_ALPHAOP, D3DTOP_SELECTARG2
  D3DD.SetTextureStageState 0, D3DTSS_ALPHAARG2, D3DTA_DIFFUSE
  D3DD.SetTexture 1, MapTexture
  D3DD.SetTextureStageState 1, D3DTSS_TEXCOORDINDEX, 0
  D3DD.SetTextureStageState 1, D3DTSS_COLORARG1, D3DTA_TEXTURE
  D3DD.SetTextureStageState 1, D3DTSS_COLOROP, D3DTOP_MODULATE2X
  D3DD.SetTextureStageState 1, D3DTSS_COLORARG2, D3DTA_CURRENT
  D3DD.SetTextureStageState 1, D3DTSS_ALPHAARG1, D3DTA_TEXTURE
  D3DD.SetTextureStageState 1, D3DTSS_ALPHAOP, D3DTOP_SELECTARG1
  D3DD.SetTextureStageState 1, D3DTSS_ALPHAARG2, D3DTA_DIFFUSE
  D3DD.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, BackVerts(0), Len(BackVerts(0))
  D3DD.SetTexture 0, Nothing
  D3DD.SetTexture 1, Nothing
 End If
 If DoDAdd = True Then
  D3DD.SetTextureStageState 0, D3DTSS_TEXCOORDINDEX, 0
  D3DD.SetTextureStageState 0, D3DTSS_COLORARG1, D3DTA_TEXTURE
  D3DD.SetTextureStageState 0, D3DTSS_COLOROP, D3DTOP_MODULATE
  D3DD.SetTextureStageState 0, D3DTSS_COLORARG2, D3DTA_DIFFUSE
  D3DD.SetTextureStageState 0, D3DTSS_ALPHAARG1, D3DTA_TEXTURE
  D3DD.SetTextureStageState 0, D3DTSS_ALPHAOP, D3DTOP_SELECTARG2
  D3DD.SetTextureStageState 0, D3DTSS_ALPHAARG2, D3DTA_DIFFUSE
  D3DD.SetTexture 1, MapTexture
  D3DD.SetTextureStageState 1, D3DTSS_TEXCOORDINDEX, 0
  D3DD.SetTextureStageState 1, D3DTSS_COLORARG1, D3DTA_TEXTURE
  D3DD.SetTextureStageState 1, D3DTSS_COLOROP, D3DTOP_ADDSIGNED
  D3DD.SetTextureStageState 1, D3DTSS_COLORARG2, D3DTA_CURRENT
  D3DD.SetTextureStageState 1, D3DTSS_ALPHAARG1, D3DTA_TEXTURE
  D3DD.SetTextureStageState 1, D3DTSS_ALPHAOP, D3DTOP_SELECTARG1
  D3DD.SetTextureStageState 1, D3DTSS_ALPHAARG2, D3DTA_DIFFUSE
  D3DD.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, BackVerts(0), Len(BackVerts(0))
  D3DD.SetTexture 0, Nothing
  D3DD.SetTexture 1, Nothing
 End If
 If DoASat = True Then
  D3DD.SetTextureStageState 0, D3DTSS_TEXCOORDINDEX, 0
  D3DD.SetTextureStageState 0, D3DTSS_COLORARG1, D3DTA_TEXTURE
  D3DD.SetTextureStageState 0, D3DTSS_COLOROP, D3DTOP_MODULATE
  D3DD.SetTextureStageState 0, D3DTSS_COLORARG2, D3DTA_DIFFUSE
  D3DD.SetTextureStageState 0, D3DTSS_ALPHAARG1, D3DTA_TEXTURE
  D3DD.SetTextureStageState 0, D3DTSS_ALPHAOP, D3DTOP_SELECTARG2
  D3DD.SetTextureStageState 0, D3DTSS_ALPHAARG2, D3DTA_DIFFUSE
  D3DD.SetTexture 1, MapTexture
  D3DD.SetTextureStageState 1, D3DTSS_TEXCOORDINDEX, 0
  D3DD.SetTextureStageState 1, D3DTSS_COLORARG1, D3DTA_TEXTURE
  D3DD.SetTextureStageState 1, D3DTSS_COLOROP, D3DTOP_ADD
  D3DD.SetTextureStageState 1, D3DTSS_COLORARG2, D3DTA_CURRENT
  D3DD.SetTextureStageState 1, D3DTSS_ALPHAARG1, D3DTA_TEXTURE
  D3DD.SetTextureStageState 1, D3DTSS_ALPHAOP, D3DTOP_SELECTARG1
  D3DD.SetTextureStageState 1, D3DTSS_ALPHAARG2, D3DTA_DIFFUSE
  D3DD.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, BackVerts(0), Len(BackVerts(0))
  D3DD.SetTexture 0, Nothing
  D3DD.SetTexture 1, Nothing
 End If
 If DoLBlend = True Then
  D3DD.SetTextureStageState 0, D3DTSS_TEXCOORDINDEX, 0
  D3DD.SetTextureStageState 0, D3DTSS_COLORARG1, D3DTA_TEXTURE
  D3DD.SetTextureStageState 0, D3DTSS_COLOROP, D3DTOP_MODULATE
  D3DD.SetTextureStageState 0, D3DTSS_COLORARG2, D3DTA_DIFFUSE
  D3DD.SetTextureStageState 0, D3DTSS_ALPHAARG1, D3DTA_TEXTURE
  D3DD.SetTextureStageState 0, D3DTSS_ALPHAOP, D3DTOP_SELECTARG2
  D3DD.SetTextureStageState 0, D3DTSS_ALPHAARG2, D3DTA_DIFFUSE
  D3DD.SetTexture 1, MapTexture
  D3DD.SetTextureStageState 1, D3DTSS_TEXCOORDINDEX, 0
  D3DD.SetTextureStageState 1, D3DTSS_COLORARG1, D3DTA_TEXTURE
  D3DD.SetTextureStageState 1, D3DTSS_COLOROP, D3DTOP_BLENDDIFFUSEALPHA
  D3DD.SetTextureStageState 1, D3DTSS_COLORARG2, D3DTA_CURRENT
  D3DD.SetTextureStageState 1, D3DTSS_ALPHAARG1, D3DTA_TEXTURE
  D3DD.SetTextureStageState 1, D3DTSS_ALPHAOP, D3DTOP_SELECTARG1
  D3DD.SetTextureStageState 1, D3DTSS_ALPHAARG2, D3DTA_DIFFUSE
  D3DD.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, BackVerts(0), Len(BackVerts(0))
  D3DD.SetTexture 0, Nothing
  D3DD.SetTexture 1, Nothing
 End If
 If DoNone = True Then
  D3DD.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, BackVerts(0), Len(BackVerts(0))
  D3DD.SetTexture 0, Nothing
 End If
End If

D3DD.SetTransform D3DTS_VIEW, matViewSave
D3DD.SetRenderState D3DRS_ZENABLE, 1
End Sub

Public Sub Setup_Default_Rendering_State()
 D3DD.SetRenderState D3DRS_DITHERENABLE, 1
 D3DD.SetRenderState D3DRS_ZENABLE, 1
 D3DD.SetRenderState D3DRS_LIGHTING, 0
End Sub

Public Sub Set_Engines_Shade_Type(Shade_Type As CONST_D3DSHADEMODE)
 On Local Error Resume Next
 D3DD.SetRenderState D3DRS_SHADEMODE, Shade_Type
End Sub

Public Sub Set_Engines_Fill_Type(Fill_Type As CONST_D3DFILLMODE)
 On Local Error Resume Next
 D3DD.SetRenderState D3DRS_FILLMODE, Fill_Type
End Sub

'|œœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœ|
'|œœ                         FPS                            œœ|
'|œœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœ|

Private Function Update_FPS()
 Static i As Integer
 Static tEnd As Single
 Static tCurrent As Single
 Static HoldFPS As Single

 i = i + 1
 If i = 30 Then
  tCurrent = Timer
    If tCurrent <> tEnd Then
      HoldFPS = 30 / (Timer - tEnd)
      tEnd = Timer
      i = 0
      fps = HoldFPS
    End If
 End If
End Function

Public Property Get Get_FPS(Optional Format_FPS_String As Boolean = True) As String
 If Format_FPS_String = True Then
  Get_FPS = format$(fps, "###.00")
 Else
  Get_FPS = fps
 End If
End Property

Public Sub Show_FPS_Track(PosX As Integer, PosY As Integer, Red As Single, Green As Single, Blue As Single)
Dim Output As String
Output = String$(fps \ 2, "Ž")
Draw_Text Output, PosX, PosY, Red, Green, Blue
DoEvents
End Sub

'|œœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœ|
'|œœ                     Text Functions                     œœ|
'|œœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœ|

Public Sub Initialize_Text(Optional Font_Name As String = "Ariel", Optional Font_Size As Single = 8, Optional Bold As Boolean = False, _
                           Optional Underline As Boolean = False, Optional Italic As Boolean = False, Optional Strikethrough As Boolean = False)
fnt.Name = Font_Name
fnt.Size = Font_Size
fnt.Bold = Bold
fnt.Strikethrough = Strikethrough
fnt.Underline = Underline
fnt.Italic = Italic
Set vbFont = fnt
Set fnt = Nothing
Set xFont = D3DX.CreateFont(D3DD, vbFont.hFont)
TextInitialized = True
End Sub

Public Function Draw_Text(Text_To_Draw As String, PosX As Integer, PosY As Integer, Red As Single, Green As Single, Blue As Single)
On Local Error Resume Next
Dim TextRect As RECT
Dim Screen_Height As Long
Dim Screen_Width As Long
Dim col As D3DCOLORVALUE

If TextInitialized = False Then Exit Function

col.r = Red: col.g = Green: col.b = Blue: col.A = 1
Call GetClientRect(HoldhWnd, TextRect)
TextRect.Left = PosX
TextRect.Top = PosY
Screen_Height = TextRect.bottom - TextRect.Top
Screen_Width = TextRect.Right - TextRect.Left

xFont.Begin
D3DX.DrawText xFont, D3DCOLORVALUEtoLONG(col), Text_To_Draw, TextRect, 0
xFont.End
End Function

'|œœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœ|
'|œœ                     Misc. Functions                    œœ|
'|œœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœœ|

Public Property Get IsFullScreen() As Boolean
 IsFullScreen = FullS
End Property

Public Function Create_Texture(Filename As String) As Direct3DTexture8
On Local Error Resume Next
Select Case UCase(Right(Filename, 3))
Case "BMP"
Case "JPG"
  Dim tmpPic As IPictureDisp
  Set tmpPic = LoadPicture(Filename)
  Filename = Left(Filename, Len(Filename) - 3) & "bmp"
  SavePicture tmpPic, Filename
Case Else
  Exit Function
End Select
 Set Create_Texture = D3DX.CreateTextureFromFile(D3DD, Filename)
End Function

Public Function Create_TextureEX(Filename As String, Optional ColorKey As Long = &HFF000000) As Direct3DTexture8
On Local Error Resume Next
Select Case UCase(Right(Filename, 3))
Case "BMP"
Case "JPG"
  Dim tmpPic As IPictureDisp
  Set tmpPic = LoadPicture(Filename)
  Filename = Left(Filename, Len(Filename) - 3) & "bmp"
  SavePicture tmpPic, Filename
Case Else
  Exit Function
End Select
 Set Create_TextureEX = D3DX.CreateTextureFromFileEx(D3DD, Filename, D3DX_DEFAULT, _
                                                     D3DX_DEFAULT, D3DX_DEFAULT, 0, _
                                                     D3DFMT_UNKNOWN, D3DPOOL_MANAGED, _
                                                     D3DX_FILTER_POINT, D3DX_FILTER_POINT, _
                                                     ColorKey, ByVal 0, ByVal 0)
End Function

Private Function D3DCOLORVALUEtoLONG(cv As D3DCOLORVALUE) As Long
    Dim r As Long
    Dim g As Long
    Dim b As Long
    Dim A As Long
    Dim C As Long
    
    r = cv.r * 255
    g = cv.g * 255
    b = cv.b * 255
    A = cv.A * 255
    
    If A > 127 Then
        A = A - 128
        C = A * 2 ^ 24 Or &H80000000
        C = C Or r * 2 ^ 16
        C = C Or g * 2 ^ 8
        C = C Or b
    Else
        C = A * 2 ^ 24
        C = C Or r * 2 ^ 16
        C = C Or g * 2 ^ 8
        C = C Or b
    End If
    
    D3DCOLORVALUEtoLONG = C
End Function
