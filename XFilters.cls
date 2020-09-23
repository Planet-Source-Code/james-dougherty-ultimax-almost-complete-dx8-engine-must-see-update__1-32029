VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "XFilters"
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

Private FilterSelected(50) As Boolean
Private HoldValue As Long
Private mat As D3DMATRIX, matTextSave As D3DMATRIX

Public Sub Set_Default_Filter()
 With D3DD
  .SetRenderState D3DRS_EDGEANTIALIAS, 1
  .SetRenderState D3DRS_DITHERENABLE, 1
  .SetRenderState D3DRS_ZENABLE, 1
  .SetRenderState D3DRS_NORMALIZENORMALS, 1
 End With
End Sub

Public Sub Set_Soften_Filter()
D3DD.SetTextureStageState 0, D3DTSS_MINFILTER, D3DTEXF_LINEAR
End Sub

Public Sub Set_Dither_Filter()
D3DD.SetRenderState D3DRS_DITHERENABLE, 1
End Sub

Public Sub Set_Color_Fix_Filter()
 FilterSelected(0) = False
 With D3DD
  .SetTextureStageState 0, D3DTSS_COLOROP, D3DTOP_SELECTARG1
  .SetTextureStageState 0, D3DTSS_COLORARG1, D3DTA_TEXTURE
  .SetTextureStageState 0, D3DTSS_ALPHAOP, D3DTOP_SELECTARG1
  .SetTextureStageState 0, D3DTSS_ALPHAARG1, D3DTA_TEXTURE
 End With
 FilterSelected(0) = True
End Sub

Public Sub Set_Full_Bright_Filter()
 FilterSelected(1) = False
 With D3DD
  .SetTextureStageState 0, D3DTSS_TEXCOORDINDEX, 0
  .SetTextureStageState 0, D3DTSS_COLORARG1, D3DTA_TEXTURE
  .SetTextureStageState 0, D3DTSS_COLORARG2, D3DTA_CURRENT
  .SetTextureStageState 0, D3DTSS_COLOROP, D3DTOP_ADDSIGNED
  .SetRenderState D3DRS_SHADEMODE, 1
 End With
 FilterSelected(1) = True
End Sub

Public Sub Set_Anisotropic_Filter()
 FilterSelected(2) = False
 Dim Caps As D3DCAPS8
 With D3DD
  If Caps.MaxSimultaneousTextures < 2 Then
   .SetTextureStageState 0, D3DTSS_MAGFILTER, D3DTEXF_ANISOTROPIC
   .SetTextureStageState 0, D3DTSS_MINFILTER, D3DTEXF_ANISOTROPIC
   .SetTextureStageState 0, D3DTSS_MIPFILTER, D3DTEXF_ANISOTROPIC
   .SetTextureStageState 0, D3DTSS_MAXANISOTROPY, 1
  Else
   .SetTextureStageState 0, D3DTSS_MAGFILTER, D3DTEXF_ANISOTROPIC
   .SetTextureStageState 0, D3DTSS_MINFILTER, D3DTEXF_ANISOTROPIC
   .SetTextureStageState 0, D3DTSS_MIPFILTER, D3DTEXF_ANISOTROPIC
   .SetTextureStageState 0, D3DTSS_MAXANISOTROPY, 16
  End If
 End With
 FilterSelected(2) = True
End Sub

Public Sub Set_Bilinear_Filter()
 FilterSelected(3) = False
 With D3DD
  .SetTextureStageState 0, D3DTSS_MAGFILTER, D3DTEXF_LINEAR
  .SetTextureStageState 0, D3DTSS_MINFILTER, D3DTEXF_LINEAR
  .SetTextureStageState 0, D3DTSS_MIPFILTER, D3DTEXF_LINEAR
 End With
 FilterSelected(3) = True
End Sub

Public Sub Set_Point_Sampling_Filter()
 FilterSelected(4) = False
 With D3DD
  .SetTextureStageState 0, D3DTSS_MAGFILTER, D3DTEXF_LINEAR
  .SetTextureStageState 0, D3DTSS_MINFILTER, D3DTEXF_LINEAR
 End With
 FilterSelected(4) = True
End Sub

Public Sub Set_Mirror_X_Filter()
 FilterSelected(5) = False
 With D3DD
  .SetTextureStageState 0, D3DTSS_ADDRESSU, D3DTADDRESS_MIRROR
  .SetTextureStageState 0, D3DTSS_ADDRESSW, D3DTADDRESS_MIRROR
 End With
 FilterSelected(5) = True
End Sub

Public Sub Set_Mirror_Y_Filter()
 FilterSelected(6) = False
 With D3DD
  .SetTextureStageState 0, D3DTSS_ADDRESSV, D3DTADDRESS_MIRROR
  .SetTextureStageState 0, D3DTSS_ADDRESSW, D3DTADDRESS_MIRROR
 End With
 FilterSelected(6) = True
End Sub

Public Sub Set_Mirror_XY_Filter()
 FilterSelected(7) = False
 With D3DD
  .SetTextureStageState 0, D3DTSS_ADDRESSU, D3DTADDRESS_MIRROR
  .SetTextureStageState 0, D3DTSS_ADDRESSV, D3DTADDRESS_MIRROR
  .SetTextureStageState 0, D3DTSS_ADDRESSW, D3DTADDRESS_MIRROR
 End With
 FilterSelected(7) = True
End Sub

Public Sub Set_Clamp_Filter()
 FilterSelected(8) = False
 With D3DD
  .SetTextureStageState 0, D3DTSS_ADDRESSU, D3DTADDRESS_CLAMP
  .SetTextureStageState 0, D3DTSS_ADDRESSV, D3DTADDRESS_CLAMP
  .SetTextureStageState 0, D3DTSS_ADDRESSW, D3DTADDRESS_CLAMP
 End With
 FilterSelected(8) = True
End Sub

'Doesnt Work Right?
'NOTE---> need to fix <---ETON
Private Sub Set_Border_Filter(Red As Single, Green As Single, Blue As Single)
 Dim Color As D3DCOLORVALUE
 Color.r = Red: Color.g = Green: Color.b = Blue: Color.A = 1
 FilterSelected(9) = False
 With D3DD
  .SetTextureStageState 0, D3DTSS_ADDRESSU, D3DTADDRESS_BORDER
  .SetTextureStageState 0, D3DTSS_ADDRESSV, D3DTADDRESS_BORDER
  .SetTextureStageState 0, D3DTSS_ADDRESSW, D3DTADDRESS_BORDER
  .SetTextureStageState 0, D3DTSS_BORDERCOLOR, D3DCOLORVALUEtoLONG(Color)
 End With
 FilterSelected(9) = True
End Sub

Public Sub Set_Mirror_Once_Filter()
 FilterSelected(10) = False
 With D3DD
  .SetTextureStageState 0, D3DTSS_ADDRESSU, D3DTADDRESS_MIRRORONCE
  .SetTextureStageState 0, D3DTSS_ADDRESSV, D3DTADDRESS_MIRRORONCE
  .SetTextureStageState 0, D3DTSS_ADDRESSW, D3DTADDRESS_MIRRORONCE
 End With
 FilterSelected(10) = True
End Sub

Public Sub Set_Modulate_Alpha_Filter(Red As Single, Green As Single, Blue As Single)
 Dim AColor As D3DCOLORVALUE
 AColor.r = Red: AColor.g = Green: AColor.b = Blue: AColor.A = 1
 
 FilterSelected(11) = False
 With D3DD
  HoldValue = .GetRenderState(D3DRS_AMBIENT)
  .SetRenderState D3DRS_AMBIENT, D3DCOLORVALUEtoLONG(AColor)
  .SetTextureStageState 0, D3DTSS_COLORARG1, D3DTA_TEXTURE
  .SetTextureStageState 0, D3DTSS_COLORARG2, D3DTA_DIFFUSE
  .SetTextureStageState 0, D3DTSS_COLOROP, D3DTOP_MODULATE
  .SetTextureStageState 0, D3DTSS_ALPHAARG1, D3DTA_TEXTURE
  .SetTextureStageState 0, D3DTSS_ALPHAARG2, D3DTA_DIFFUSE
  .SetTextureStageState 0, D3DTSS_ALPHAOP, D3DTOP_MODULATE
 End With
 FilterSelected(11) = True
End Sub

Public Sub Clean_Rendering_Device()
If FilterSelected(0) = True Then
 With D3DD
  .SetTextureStageState 0, D3DTSS_COLOROP, D3DTEXF_NONE
  .SetTextureStageState 0, D3DTSS_COLORARG1, D3DTEXF_NONE
  .SetTextureStageState 0, D3DTSS_ALPHAOP, D3DTEXF_NONE
  .SetTextureStageState 0, D3DTSS_ALPHAARG1, D3DTEXF_NONE
 End With
End If
If FilterSelected(1) = True Then
 With D3DD
  .SetTextureStageState 0, D3DTSS_TEXCOORDINDEX, D3DTEXF_NONE
  .SetTextureStageState 0, D3DTSS_COLORARG1, D3DTEXF_NONE
  .SetTextureStageState 0, D3DTSS_COLORARG2, D3DTEXF_NONE
  .SetTextureStageState 0, D3DTSS_COLOROP, D3DTEXF_NONE
  .SetRenderState D3DRS_SHADEMODE, 1
 End With
End If
If FilterSelected(2) = True Then
  With D3DD
   .SetTextureStageState 0, D3DTSS_MAGFILTER, D3DTEXF_NONE
   .SetTextureStageState 0, D3DTSS_MINFILTER, D3DTEXF_NONE
   .SetTextureStageState 0, D3DTSS_MIPFILTER, D3DTEXF_NONE
   .SetTextureStageState 0, D3DTSS_MAXANISOTROPY, D3DTEXF_NONE
  End With
End If
If FilterSelected(3) = True Then
 With D3DD
  .SetTextureStageState 0, D3DTSS_MAGFILTER, D3DTEXF_NONE
  .SetTextureStageState 0, D3DTSS_MINFILTER, D3DTEXF_NONE
  .SetTextureStageState 0, D3DTSS_MIPFILTER, D3DTEXF_NONE
 End With
End If
If FilterSelected(4) = True Then
 With D3DD
  .SetTextureStageState 0, D3DTSS_MAGFILTER, D3DTEXF_NONE
  .SetTextureStageState 0, D3DTSS_MINFILTER, D3DTEXF_NONE
 End With
End If
If FilterSelected(5) = True Then
 With D3DD
  .SetTextureStageState 0, D3DTSS_ADDRESSU, D3DTEXF_NONE
  .SetTextureStageState 0, D3DTSS_ADDRESSW, D3DTEXF_NONE
 End With
End If
If FilterSelected(6) = True Then
 With D3DD
  .SetTextureStageState 0, D3DTSS_ADDRESSV, D3DTEXF_NONE
  .SetTextureStageState 0, D3DTSS_ADDRESSW, D3DTEXF_NONE
 End With
End If
If FilterSelected(7) = True Then
 With D3DD
  .SetTextureStageState 0, D3DTSS_ADDRESSU, D3DTEXF_NONE
  .SetTextureStageState 0, D3DTSS_ADDRESSV, D3DTEXF_NONE
  .SetTextureStageState 0, D3DTSS_ADDRESSW, D3DTEXF_NONE
 End With
End If
If FilterSelected(8) = True Then
 With D3DD
  .SetTextureStageState 0, D3DTSS_ADDRESSU, D3DTEXF_NONE
  .SetTextureStageState 0, D3DTSS_ADDRESSV, D3DTEXF_NONE
  .SetTextureStageState 0, D3DTSS_ADDRESSW, D3DTEXF_NONE
 End With
End If
If FilterSelected(9) = True Then
 With D3DD
  .SetTextureStageState 0, D3DTSS_ADDRESSU, D3DTEXF_NONE
  .SetTextureStageState 0, D3DTSS_ADDRESSV, D3DTEXF_NONE
  .SetTextureStageState 0, D3DTSS_ADDRESSW, D3DTEXF_NONE
  .SetTextureStageState 0, D3DTSS_BORDERCOLOR, 0
 End With
End If
If FilterSelected(10) = True Then
 With D3DD
  .SetTextureStageState 0, D3DTSS_ADDRESSU, D3DTEXF_NONE
  .SetTextureStageState 0, D3DTSS_ADDRESSV, D3DTEXF_NONE
  .SetTextureStageState 0, D3DTSS_ADDRESSW, D3DTEXF_NONE
 End With
End If
If FilterSelected(11) = True Then
 With D3DD
  '.SetRenderState D3DRS_AMBIENT, HoldValue
  .SetTextureStageState 0, D3DTSS_COLORARG1, D3DTEXF_NONE
  .SetTextureStageState 0, D3DTSS_COLORARG2, D3DTEXF_NONE
  .SetTextureStageState 0, D3DTSS_COLOROP, D3DTEXF_NONE
  .SetTextureStageState 0, D3DTSS_ALPHAARG1, D3DTEXF_NONE
  .SetTextureStageState 0, D3DTSS_ALPHAARG2, D3DTEXF_NONE
  .SetTextureStageState 0, D3DTSS_ALPHAOP, D3DTEXF_NONE
 End With
End If
End Sub

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
