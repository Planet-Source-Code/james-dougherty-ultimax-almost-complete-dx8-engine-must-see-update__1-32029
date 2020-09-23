VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "XSpriteXYZ"
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



'SPRITE'S STAY CLAMPED TO VIEWPORT..GREAT FOR HUD'S


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

Private Type tSprite
 SpriteVerts(3) As TLVERTEX
 Position As D3DVECTOR
 Transparent As Boolean
 sHeight As Single
 sWidth As Single
End Type: Private Sprite As tSprite

Private Eng As New XEngine3D
Private Math As New XMath
Private SpriteSurface As Direct3DTexture8
Private HasSurface As Boolean
Private Declare Function GetClientRect Lib "user32.dll" (ByVal hWnd As Long, ByRef RECT As RECT) As Long

Public Property Get Get_Sprite_Position_X() As Single
 Get_Sprite_Position_X = Sprite.Position.X
End Property

Public Property Get Get_Sprite_Position_Y() As Single
 Get_Sprite_Position_Y = Sprite.Position.Y
End Property

Public Property Get Get_Sprite_Position_Z() As Single
 Get_Sprite_Position_Z = Sprite.Position.Z
End Property

Public Property Get Get_Sprite_PositionEX() As D3DVECTOR
 Get_Sprite_PositionEX = Sprite.Position
End Property

Public Property Get Get_Sprites_Width() As Single
 Get_Sprites_Width = Sprite.sWidth
End Property

Public Property Get Get_Sprites_Height() As Single
 Get_Sprites_Height = Sprite.sHeight
End Property

Public Sub Position_Sprite(X As Single, Y As Single, Z As Single)
Dim W, H As Single
W = Sprite.sWidth: H = Sprite.sHeight
Sprite.Position = Math.Vector3(X, Y, Z)

With Sprite
 .SpriteVerts(0).X = X:     .SpriteVerts(0).Y = Y + H: .SpriteVerts(0).Z = Z
 .SpriteVerts(1).X = X:     .SpriteVerts(1).Y = Y:     .SpriteVerts(1).Z = Z
 .SpriteVerts(2).X = X + W: .SpriteVerts(2).Y = Y + H: .SpriteVerts(2).Z = Z
 .SpriteVerts(3).X = X + W: .SpriteVerts(3).Y = Y:     .SpriteVerts(3).Z = Z
 .SpriteVerts(0).tU = 0: .SpriteVerts(0).tV = 1: .SpriteVerts(0).RHW = 1: .SpriteVerts(0).Color = &HFFFFFF
 .SpriteVerts(1).tU = 0: .SpriteVerts(1).tV = 0: .SpriteVerts(1).RHW = 1: .SpriteVerts(1).Color = &HFFFFFF
 .SpriteVerts(2).tU = 1: .SpriteVerts(2).tV = 1: .SpriteVerts(2).RHW = 1: .SpriteVerts(2).Color = &HFFFFFF
 .SpriteVerts(3).tU = 1: .SpriteVerts(3).tV = 0: .SpriteVerts(3).RHW = 1: .SpriteVerts(3).Color = &HFFFFFF
End With

End Sub

Public Sub Scale_Sprite(Scale_X As Single, Scale_Y As Single)
Dim W As Single, H As Single
Dim X As Single, Y As Single, Z As Single
W = Scale_X: H = Scale_Y
Sprite.sWidth = Scale_X: Sprite.sHeight = Scale_Y
X = Sprite.Position.X: Y = Sprite.Position.Y: Z = Sprite.Position.Z

With Sprite
 .SpriteVerts(0).X = X:     .SpriteVerts(0).Y = Y + H: .SpriteVerts(0).Z = Z
 .SpriteVerts(1).X = X:     .SpriteVerts(1).Y = Y:     .SpriteVerts(1).Z = Z
 .SpriteVerts(2).X = X + W: .SpriteVerts(2).Y = Y + H: .SpriteVerts(2).Z = Z
 .SpriteVerts(3).X = X + W: .SpriteVerts(3).Y = Y:     .SpriteVerts(3).Z = Z
 .SpriteVerts(0).tU = 0: .SpriteVerts(0).tV = 1: .SpriteVerts(0).RHW = 1: .SpriteVerts(0).Color = &HFFFFFF
 .SpriteVerts(1).tU = 0: .SpriteVerts(1).tV = 0: .SpriteVerts(1).RHW = 1: .SpriteVerts(1).Color = &HFFFFFF
 .SpriteVerts(2).tU = 1: .SpriteVerts(2).tV = 1: .SpriteVerts(2).RHW = 1: .SpriteVerts(2).Color = &HFFFFFF
 .SpriteVerts(3).tU = 1: .SpriteVerts(3).tV = 0: .SpriteVerts(3).RHW = 1: .SpriteVerts(3).Color = &HFFFFFF
End With

End Sub

Public Sub Enable_Transparency(Enabled As Boolean)
If Enabled = True Then
 Sprite.Transparent = True
Else
 Sprite.Transparent = False
End If
End Sub

Public Function Create_New_Sprite(Filename As String, Sprite_Width As Single, Sprite_Height As Single, _
                                  StartX As Single, StartY As Single, StartZ As Single) As Long

With Sprite
 .SpriteVerts(0).X = StartX:                .SpriteVerts(0).Y = StartY + Sprite_Height: .SpriteVerts(0).Z = StartZ
 .SpriteVerts(1).X = StartX:                .SpriteVerts(1).Y = StartY:                 .SpriteVerts(1).Z = StartZ
 .SpriteVerts(2).X = StartX + Sprite_Width: .SpriteVerts(2).Y = StartY + Sprite_Height: .SpriteVerts(2).Z = StartZ
 .SpriteVerts(3).X = StartX + Sprite_Width: .SpriteVerts(3).Y = StartY:                 .SpriteVerts(3).Z = StartZ
 .SpriteVerts(0).tU = 0: .SpriteVerts(0).tV = 1: .SpriteVerts(0).RHW = 1: .SpriteVerts(0).Color = &HFFFFFF
 .SpriteVerts(1).tU = 0: .SpriteVerts(1).tV = 0: .SpriteVerts(1).RHW = 1: .SpriteVerts(1).Color = &HFFFFFF
 .SpriteVerts(2).tU = 1: .SpriteVerts(2).tV = 1: .SpriteVerts(2).RHW = 1: .SpriteVerts(2).Color = &HFFFFFF
 .SpriteVerts(3).tU = 1: .SpriteVerts(3).tV = 0: .SpriteVerts(3).RHW = 1: .SpriteVerts(3).Color = &HFFFFFF
 
 .Position = Math.Vector3(StartX, StartY, StartZ)
 .sHeight = Sprite_Height: .sWidth = Sprite_Width
End With

 Set SpriteSurface = Eng.Direct3DX.CreateTextureFromFileEx(D3DD, Filename, D3DX_DEFAULT, _
                                               D3DX_DEFAULT, D3DX_DEFAULT, 0, _
                                               D3DFMT_UNKNOWN, D3DPOOL_MANAGED, _
                                               D3DX_FILTER_POINT, D3DX_FILTER_POINT, _
                                               &HFF000000, ByVal 0, ByVal 0)
 HasSurface = True
End Function

Public Sub Render_Sprite()
If D3DD Is Nothing Then Exit Sub

If HasSurface = True Then
  D3DD.SetRenderState D3DRS_ALPHABLENDENABLE, 1
  D3DD.SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCALPHA
  D3DD.SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA
  If Sprite.Transparent = True Then
   D3DD.SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCCOLOR
   D3DD.SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCCOLOR
  End If
  D3DD.SetVertexShader FVF
  D3DD.SetTexture 0, SpriteSurface
  D3DD.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, Sprite.SpriteVerts(0), Len(Sprite.SpriteVerts(0))
  D3DD.SetTexture 0, Nothing
  D3DD.SetRenderState D3DRS_DESTBLEND, 0
  D3DD.SetRenderState D3DRS_SRCBLEND, 0
  D3DD.SetRenderState D3DRS_ALPHABLENDENABLE, 0
Else
  Exit Sub
End If
  
End Sub

Public Sub Cleanup_Sprite_Engine()
Set SpriteSurface = Nothing
Set Math = Nothing
Set Eng = Nothing
End Sub
