VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "XMiscObjects"
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

Private Type TVERTEX
 v() As D3DVERTEX
 Postion As D3DVECTOR
 Color As Long
 tU As Single
 tV As Single
End Type
Private Const D3DFVF_CUSTOMVERTEX = (D3DFVF_XYZ Or D3DFVF_DIFFUSE Or D3DFVF_TEX1)
Private Const PI = 3.14159265358979

Private Eng As New XEngine3D
Private VB As Direct3DVertexBuffer8
Private Cylinder_Texture As Direct3DTexture8
Private HasCylTexture As Boolean
Private Gear_Texture As Direct3DTexture8
Private HasGearTexture As Boolean

Public Sub Create_Cylinder(Optional Texture_Filename As String, Optional Length As Single = 1)
pCreateCylinder Length
If Texture_Filename <> "" Then
 Set Cylinder_Texture = Eng.Direct3DX.CreateTextureFromFile(D3DD, Texture_Filename)
 HasCylTexture = True
End If
End Sub

Private Sub pCreateCylinder(Optional Length As Single = 1)
Dim i As Long
Dim Vertices(100) As TVERTEX
Dim VertexSizeInBytes As Long
Dim theta As Single
    
VertexSizeInBytes = Len(Vertices(0))
    
For i = 0 To 49
 theta = (2 * PI * i) / (50 - 1)
 Vertices(2 * i + 0).Postion = vec3(Sin(theta), -Length, Cos(theta))
 Vertices(2 * i + 0).Color = &HFFFFFFFF
 Vertices(2 * i + 0).tU = i / (50 - 1)
 Vertices(2 * i + 0).tV = 1
 Vertices(2 * i + 1).Postion = vec3(Sin(theta), Length, Cos(theta))
 Vertices(2 * i + 1).Color = &HFF808080
 Vertices(2 * i + 1).tU = i / (50 - 1)
 Vertices(2 * i + 1).tV = 0
Next

Set VB = D3DD.CreateVertexBuffer(VertexSizeInBytes * 50 * 2, _
                                 0, D3DFVF_CUSTOMVERTEX, D3DPOOL_DEFAULT)
If VB Is Nothing Then Exit Sub

D3DVertexBuffer8SetData VB, 0, VertexSizeInBytes * 100, 0, Vertices(0)
End Sub

Public Sub Render_Cylinder()
Dim smat As D3DMATRIX
Dim mPos As D3DMATRIX
Dim v As TVERTEX
Dim sizeOfVertex As Long

D3DD.GetTransform D3DTS_WORLD, smat
mPos = smat
mPos.m11 = 1: mPos.m22 = 1: mPos.m33 = 1: mPos.m44 = 1
mPos.m41 = 0: mPos.m42 = 10: mPos.m43 = 0
D3DD.SetTransform D3DTS_WORLD, mPos
 
If HasCylTexture = True Then
 D3DD.SetTexture 0, Cylinder_Texture
End If
sizeOfVertex = Len(v)
D3DD.SetStreamSource 0, VB, sizeOfVertex
D3DD.SetVertexShader D3DFVF_CUSTOMVERTEX
D3DD.DrawPrimitive D3DPT_TRIANGLESTRIP, 0, (4 * 25) - 2

D3DD.SetTransform D3DTS_WORLD, smat
End Sub

Public Sub Create_Gear(Optional Texture_Filename As String, Optional GearDepth As Single = 0.1)
pCreateGear GearDepth
If Texture_Filename <> "" Then
 Set Gear_Texture = Eng.Direct3DX.CreateTextureFromFile(D3DD, Texture_Filename)
 HasGearTexture = True
End If
End Sub

Private Sub pCreateGear(Optional GearDepth As Single = 0.1)
Dim i As Long
Dim Vertices(100) As TVERTEX
Dim VertexSizeInBytes As Long
Dim theta As Single
    
VertexSizeInBytes = Len(Vertices(0))
    
For i = 0 To 49
 theta = (40 * PI * i) / (50 - 1)
 Vertices(2 * i + 0).Postion = vec3(Sin(theta), -GearDepth, Cos(theta))
 Vertices(2 * i + 0).Color = &HFFFFFFFF
 Vertices(2 * i + 0).tU = i / (50 - 1)
 Vertices(2 * i + 0).tV = 1
 Vertices(2 * i + 1).Postion = vec3(Sin(theta), GearDepth, Cos(theta))
 Vertices(2 * i + 1).Color = &HFF808080
 Vertices(2 * i + 1).tU = i / (50 - 1)
 Vertices(2 * i + 1).tV = 0
Next

Set VB = D3DD.CreateVertexBuffer(VertexSizeInBytes * 50 * 2, _
                                 0, D3DFVF_CUSTOMVERTEX, D3DPOOL_DEFAULT)
If VB Is Nothing Then Exit Sub

D3DVertexBuffer8SetData VB, 0, VertexSizeInBytes * 100, 0, Vertices(0)
End Sub

Public Sub Render_Gear()
Dim smat As D3DMATRIX
Dim mPos As D3DMATRIX
Dim v As TVERTEX
Dim sizeOfVertex As Long

D3DD.GetTransform D3DTS_WORLD, smat
mPos = smat
mPos.m11 = 1: mPos.m22 = 1: mPos.m33 = 1: mPos.m44 = 1
mPos.m41 = -10: mPos.m42 = 10: mPos.m43 = 0
D3DD.SetTransform D3DTS_WORLD, mPos
 
If HasGearTexture = True Then
 D3DD.SetTexture 0, Gear_Texture
End If

sizeOfVertex = Len(v)
D3DD.SetStreamSource 0, VB, sizeOfVertex
D3DD.SetVertexShader D3DFVF_CUSTOMVERTEX
D3DD.DrawPrimitive D3DPT_TRIANGLESTRIP, 0, (4 * 25) - 2

D3DD.SetTransform D3DTS_WORLD, smat
End Sub

Public Sub Cleanup_Misc_Objects()
Set Cylinder_Texture = Nothing
Set Gear_Texture = Nothing
Set VB = Nothing
Set Eng = Nothing
End Sub

Private Function vec3(X As Single, Y As Single, Z As Single) As D3DVECTOR
    vec3.X = X
    vec3.Y = Y
    vec3.Z = Z
End Function
