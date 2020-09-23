VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "XObjects"
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

Public Enum FVFType
 FVFT_D3DFVF_VERTEX = 0
 FVFT_D3DFVF_MIRRORVERTEX = 1
 FVFT_D3DFVF_COLORVERTEX = 2
 Custom = 3
End Enum

Private Type tBillboard
 BMesh As D3DXMesh
 Position As D3DVECTOR
 SetPos As Boolean
 Rotation As D3DVECTOR
 SetRot As Boolean
 Texture As Direct3DTexture8
 mRot As D3DMATRIX
End Type

Private Type tObjects
 mMesh As D3DXMesh
 Object_Name As String
 Position As D3DVECTOR
 SetPos As Boolean
 Rotation As D3DVECTOR
 SetRot As Boolean
 ScaleO As D3DVECTOR
 ScaleObject As Boolean
 Texture As Direct3DTexture8
 EnviroMapped As Boolean
 EnvTextY As Single
 Material As D3DMATERIAL8
 NewMaterial As Boolean
 Billboard As tBillboard
 Transparent As Boolean
 SetFilter As Boolean
 Mirror As Boolean
 SphereCenter As D3DVECTOR
 SphereRadius As Single
 HasSphere As Boolean
 ComputedClipPlane As Boolean
 Enhanced As Boolean
 EnhancedSegments As Single
 DarkMapAnimated As Boolean
 AnimationSpeed As Integer
 GlassEffect As Boolean
End Type: Private Mesh() As tObjects

Private Type CUSTOMVERTEX
 X As Single
 Y As Single
 Z As Single
 tU As Single
 tV As Single
End Type

Private Type MESHTOOL
 VertB As Direct3DVertexBuffer8
 NumVertices As Long
 Vertices() As CUSTOMVERTEX
End Type

Private Const PI = 3.14159265358979
Private Const PIdiv180 = 1.74532925199433E-02

Private Eng As New XEngine3D
Private Cam As New XCamera
Private HNumMesh As Long
Private mPos As D3DMATRIX
Private mRot As D3DMATRIX
Private mScale As D3DMATRIX
Private EMat As D3DMATRIX
Private sEMat As D3DMATRIX
Private WorldMat As D3DMATRIX
Private BillboardMatrix As D3DMATRIX
Private DoDark As Boolean
Private DoDetail As Boolean
Private DoLight As Boolean
Private DoGloss As Boolean
Private DoDMod As Boolean
Private DoDAdd As Boolean
Private DoASat As Boolean
Private DoLBlend As Boolean
Private DoILM As Boolean
Private MapTexture As Direct3DTexture8
Private UseOdd As Boolean
Private ObjectCount As Long
Private ClipPlanes() As D3DPLANE
Private numClipPlanes As Long
Private Const D3DFVF_MIRRORVERTEX = (D3DFVF_XYZ Or D3DFVF_NORMAL Or D3DFVF_DIFFUSE)
Private Const D3DFVF_VERTEX = D3DFVF_XYZ Or D3DFVF_NORMAL Or D3DFVF_TEX1
Private Const D3DFVF_COLORVERTEX = (D3DFVF_XYZ Or D3DFVF_DIFFUSE Or D3DFVF_TEX1)

Public Sub Compute_Bounding_Sphere(Index As Integer)
Eng.Direct3DX.ComputeBoundingSphereFromMesh Mesh(Index).mMesh, Mesh(Index).SphereCenter, Mesh(Index).SphereRadius
Mesh(Index).HasSphere = True
End Sub

Public Sub Compute_Bounding_Box(Index As Integer, MinExt As D3DVECTOR, MaxExt As D3DVECTOR)
Eng.Direct3DX.ComputeBoundingBoxFromMesh Mesh(Index).mMesh, MinExt, MaxExt
End Sub

Public Property Get Get_Number_Of_Objects() As Long
Get_Number_Of_Objects = Len(ObjectCount)
End Property

Public Property Get Get_Mesh(Index As Integer) As D3DXMesh
Set Get_Mesh = Mesh(Index).mMesh
End Property

Public Property Get Get_Object_Name(Index As Integer) As String
If Mesh(Index).Object_Name = "" Then Mesh(Index).Object_Name = "Default"
Get_Object_Name = Mesh(Index).Object_Name
End Property

Public Sub Set_Object_Name(Index As Integer, Name As String)
Mesh(Index).Object_Name = Name
End Sub

Public Property Get Get_Object_Scale_X(Index As Integer) As Single
Get_Object_Scale_X = Mesh(Index).ScaleO.X
End Property

Public Property Get Get_Object_Scale_Y(Index As Integer) As Single
Get_Object_Scale_Y = Mesh(Index).ScaleO.Y
End Property

Public Property Get Get_Object_Scale_Z(Index As Integer) As Single
Get_Object_Scale_Z = Mesh(Index).ScaleO.Z
End Property

Public Property Get Get_Object_ScaleEX(Index As Integer) As D3DVECTOR
Get_Object_ScaleEX = Mesh(Index).ScaleO
End Property

Public Property Get Get_Object_Position_X(Index As Integer) As Single
Get_Object_Position_X = Mesh(Index).Position.X
End Property

Public Property Get Get_Object_Position_Y(Index As Integer) As Single
Get_Object_Position_Y = Mesh(Index).Position.Y
End Property

Public Property Get Get_Object_Position_Z(Index As Integer) As Single
Get_Object_Position_Z = Mesh(Index).Position.Z
End Property

Public Property Get Get_Object_PositionEX(Index As Integer) As D3DVECTOR
Get_Object_PositionEX = Mesh(Index).Position
End Property

Public Sub Set_Object_Position(Index As Integer, X As Single, Y As Single, Z As Single)
Mesh(Index).Position = vec3(X, Y, Z)
Mesh(Index).SetPos = True
End Sub

Public Property Get Get_Object_Rotation_X(Index As Integer) As Single
Get_Object_Rotation_X = Mesh(Index).Rotation.X
End Property

Public Property Get Get_Object_Rotation_Y(Index As Integer) As Single
Get_Object_Rotation_Y = Mesh(Index).Rotation.Y
End Property

Public Property Get Get_Object_Rotation_Z(Index As Integer) As Single
Get_Object_Rotation_Z = Mesh(Index).Rotation.Z
End Property

Public Property Get Get_Object_RotationEX(Index As Integer) As D3DVECTOR
Get_Object_RotationEX = Mesh(Index).Rotation
End Property

Public Sub Rotate_Object(Index As Integer, X As Single, Y As Single, Z As Single)
Mesh(Index).Rotation = vec3(X, Y, Z)
Mesh(Index).SetRot = True
End Sub

Public Sub Scale_Object(Index As Integer, X As Single, Y As Single, Z As Single)
Mesh(Index).ScaleO = vec3(X, Y, Z)
Mesh(Index).ScaleObject = True
End Sub

Public Sub Enable_Object_Transparency(Index As Integer, Enabled As Boolean)
If Enabled = True Then
 Mesh(Index).Transparent = True
Else
 Mesh(Index).Transparent = False
End If
End Sub

Public Sub Enable_Glass_Effect(Index As Integer, Enabled As Boolean)
If Enabled = True Then
 Mesh(Index).GlassEffect = True
Else
 Mesh(Index).GlassEffect = False
End If
End Sub

Public Sub Apply_Object_Filter(Index As Integer, Filter_Type As FilterType, FilterMap_Filename As String)
Mesh(Index).SetFilter = True
If Filter_Type = Dark_Map And FilterMap_Filename <> "" Then
 Set MapTexture = Eng.Direct3DX.CreateTextureFromFile(D3DD, FilterMap_Filename)
 DoDark = True
End If
If Filter_Type = Light_Map And FilterMap_Filename <> "" Then
 Set MapTexture = Eng.Direct3DX.CreateTextureFromFile(D3DD, FilterMap_Filename)
 DoLight = True
End If
If Filter_Type = Detail_Map And FilterMap_Filename <> "" Then
 Set MapTexture = Eng.Direct3DX.CreateTextureFromFile(D3DD, FilterMap_Filename)
 DoDetail = True
End If
If Filter_Type = Gloss_Map And FilterMap_Filename <> "" Then
 Set MapTexture = Eng.Direct3DX.CreateTextureFromFile(D3DD, FilterMap_Filename)
 DoGloss = True
End If
If Filter_Type = Detail_Modulate And FilterMap_Filename <> "" Then
 Set MapTexture = Eng.Direct3DX.CreateTextureFromFile(D3DD, FilterMap_Filename)
 DoDMod = True
End If
If Filter_Type = Detail_Add And FilterMap_Filename <> "" Then
 Set MapTexture = Eng.Direct3DX.CreateTextureFromFile(D3DD, FilterMap_Filename)
 DoDAdd = True
End If
If Filter_Type = Add_Smooth_Saturation And FilterMap_Filename <> "" Then
 Set MapTexture = Eng.Direct3DX.CreateTextureFromFile(D3DD, FilterMap_Filename)
 DoASat = True
End If
If Filter_Type = Linear_Blend And FilterMap_Filename <> "" Then
 Set MapTexture = Eng.Direct3DX.CreateTextureFromFile(D3DD, FilterMap_Filename)
 DoLBlend = True
End If
If Filter_Type = Inverse_Light_Map And FilterMap_Filename <> "" Then
 Set MapTexture = Eng.Direct3DX.CreateTextureFromFile(D3DD, FilterMap_Filename)
 DoILM = True
End If
If Filter_Type = None Then
 Set MapTexture = Nothing
 DoILM = False
 DoLBlend = False
 DoDAdd = False
 DoASat = False
 DoDark = False
 DoDMod = False
 DoDetail = False
 DoLight = False
 DoGloss = False
End If
End Sub

Public Sub Animate_Dark_Map(Index As Integer, Animate As Boolean, Optional Animation_Speed As Integer = 1)
If Animate = True Then
 Mesh(Index).DarkMapAnimated = True
 Mesh(Index).AnimationSpeed = Animation_Speed
Else
 Mesh(Index).DarkMapAnimated = False
End If
End Sub

Private Sub Start_Render_Filter(Index As Integer)
Static i As Integer
If Mesh(Index).SetFilter = True Then
 If DoDark = True And Mesh(Index).DarkMapAnimated = False Then
  D3DD.SetTextureStageState 0, D3DTSS_TEXCOORDINDEX, 0
  D3DD.SetTextureStageState 0, D3DTSS_COLORARG1, D3DTA_TEXTURE
  D3DD.SetTextureStageState 0, D3DTSS_ALPHAOP, D3DTOP_SELECTARG1
  D3DD.SetTexture 1, MapTexture
  D3DD.SetTextureStageState 1, D3DTSS_TEXCOORDINDEX, 0
  D3DD.SetTextureStageState 1, D3DTSS_COLORARG1, D3DTA_TEXTURE
  D3DD.SetTextureStageState 1, D3DTSS_COLORARG2, D3DTA_CURRENT
  D3DD.SetTextureStageState 1, D3DTSS_COLOROP, D3DTOP_MODULATE
 Else
  i = i + Mesh(Index).AnimationSpeed
  D3DD.SetTextureStageState 0, D3DTSS_TEXCOORDINDEX, 0
  D3DD.SetTextureStageState 0, D3DTSS_COLORARG1, D3DTA_TEXTURE
  D3DD.SetTextureStageState 0, D3DTSS_ALPHAOP, D3DTOP_SELECTARG1
  D3DD.SetTexture 1, MapTexture
  D3DD.SetTextureStageState 1, D3DTSS_TEXCOORDINDEX, 0
  D3DD.SetTextureStageState 1, D3DTSS_COLORARG1, D3DTA_TEXTURE
  D3DD.SetTextureStageState 1, D3DTSS_COLORARG2, D3DTA_CURRENT
  If i < 40 Then
   D3DD.SetTextureStageState 1, D3DTSS_COLOROP, D3DTOP_MODULATE
  ElseIf i < 80 Then
   D3DD.SetTextureStageState 1, D3DTSS_COLOROP, D3DTOP_MODULATE2X
  ElseIf i < 120 Then
   D3DD.SetTextureStageState 1, D3DTSS_COLOROP, D3DTOP_MODULATE4X
  ElseIf i < 160 Then
   D3DD.SetTextureStageState 1, D3DTSS_COLOROP, D3DTOP_MODULATE2X
  ElseIf i < 200 Then
   D3DD.SetTextureStageState 1, D3DTSS_COLOROP, D3DTOP_MODULATE
   i = 0
  End If
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
 End If
 If DoGloss = True Then 'Pass 1...Pass 2 is in Render_Object
  D3DD.SetTextureStageState 0, D3DTSS_TEXCOORDINDEX, 0
  D3DD.SetTextureStageState 0, D3DTSS_COLORARG1, D3DTA_TEXTURE
  D3DD.SetTextureStageState 0, D3DTSS_COLOROP, D3DTOP_SELECTARG1
  D3DD.SetTexture 1, MapTexture
  D3DD.SetTextureStageState 1, D3DTSS_TEXCOORDINDEX, 0
  D3DD.SetTextureStageState 1, D3DTSS_COLORARG1, D3DTA_DIFFUSE
  D3DD.SetTextureStageState 1, D3DTSS_COLORARG2, D3DTA_CURRENT
  D3DD.SetTextureStageState 1, D3DTSS_COLOROP, D3DTOP_MODULATE
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
 End If
 If DoILM = True Then
  D3DD.SetTextureStageState 0, D3DTSS_TEXCOORDINDEX, 0
  D3DD.SetTextureStageState 0, D3DTSS_COLORARG1, D3DTA_TEXTURE Or D3DTA_COMPLEMENT
  D3DD.SetTextureStageState 0, D3DTSS_COLOROP, D3DTOP_MODULATE
  D3DD.SetTextureStageState 0, D3DTSS_COLORARG2, D3DTA_DIFFUSE
  D3DD.SetTextureStageState 0, D3DTSS_ALPHAARG1, D3DTA_TEXTURE
  D3DD.SetTextureStageState 0, D3DTSS_ALPHAOP, D3DTOP_SELECTARG2
  D3DD.SetTextureStageState 0, D3DTSS_ALPHAARG2, D3DTA_DIFFUSE
  D3DD.SetTexture 1, MapTexture
  D3DD.SetTextureStageState 1, D3DTSS_TEXCOORDINDEX, 0
  D3DD.SetTextureStageState 1, D3DTSS_COLORARG1, D3DTA_TEXTURE Or D3DTA_COMPLEMENT
  D3DD.SetTextureStageState 1, D3DTSS_COLOROP, D3DTOP_MODULATE
  D3DD.SetTextureStageState 1, D3DTSS_COLORARG2, D3DTA_CURRENT
  D3DD.SetTextureStageState 1, D3DTSS_ALPHAARG1, D3DTA_TEXTURE
  D3DD.SetTextureStageState 1, D3DTSS_ALPHAOP, D3DTOP_SELECTARG1
  D3DD.SetTextureStageState 1, D3DTSS_ALPHAARG2, D3DTA_DIFFUSE
 End If
End If
End Sub

Private Sub End_Render_Filter(Index As Integer)
If Mesh(Index).SetFilter = True Then
 D3DD.SetTextureStageState 1, D3DTSS_COLORARG1, D3DTOP_DISABLE Or D3DTOP_DISABLE
 D3DD.SetTextureStageState 1, D3DTSS_COLOROP, D3DTOP_DISABLE
 D3DD.SetTexture 0, Nothing
 D3DD.SetTexture 1, Nothing
End If
End Sub

Private Function Change_Material(Index As Integer, DiffuseRed As Single, DiffuseGreen As Single, DiffuseBlue As Single, _
                                Optional AmbientRed As Single = 0.5, Optional AmbientGreen As Single = 0.5, Optional AmbientBlue As Single = 0.5, _
                                Optional SpecularRed As Single = 0, Optional SpecularGreen As Single = 0, Optional SpecularBlue As Single = 0, _
                                Optional EmissiveRed As Single = 0, Optional EmissiveGreen As Single = 0, Optional EmissiveBlue As Single = 0, Optional Power As Single = 1)
Dim Mater As D3DMATERIAL8
Mater.diffuse.r = DiffuseRed: Mater.diffuse.g = DiffuseGreen
Mater.diffuse.b = DiffuseBlue: Mater.diffuse.A = 1
Mater.Ambient.r = AmbientRed: Mater.Ambient.g = AmbientGreen
Mater.Ambient.b = AmbientBlue: Mater.Ambient.A = 1
Mater.Specular.r = SpecularRed: Mater.Specular.g = SpecularGreen
Mater.Specular.b = SpecularBlue: Mater.Specular.A = 1
Mater.emissive.r = EmissiveRed: Mater.emissive.g = EmissiveGreen
Mater.emissive.b = EmissiveBlue: Mater.emissive.A = 1
Mater.Power = Power
Mesh(Index).Material = Mater
Mesh(Index).NewMaterial = True
End Function

Public Sub Initialize_Objects_Engine(Default_Num_Objects As Integer)
On Local Error Resume Next
ReDim Mesh(Default_Num_Objects)
HNumMesh = Default_Num_Objects
For Default_Num_Objects = 0 To Default_Num_Objects
  Create_Box Default_Num_Objects, 0, 0, 0
  Mesh(Default_Num_Objects).SetPos = False
Next Default_Num_Objects

D3DD.SetRenderState D3DRS_CULLMODE, D3DCULL_NONE
End Sub

Public Sub Render_Object(Index As Integer, Camera As XCamera)
Dim i As Long
Dim NumMat As Long
Dim Material As D3DMATERIAL8
Dim smat As D3DMATRIX
Dim smat2 As D3DMATRIX
Dim mat As D3DMATRIX
Dim mScaled As D3DMATRIX
Dim HoldRender1 As Long
Dim HoldRender2 As Long
If D3DD Is Nothing Then Exit Sub

Update_Clip_Planes Index, Camera
If Sphere_Visible(Index) = False Then Exit Sub

D3DD.GetTransform D3DTS_WORLD, smat2

If Mesh(Index).SetRot = True Then
 D3DD.GetTransform D3DTS_WORLD, smat
 mRot = smat
 D3DXMatrixRotationYawPitchRoll mRot, Mesh(Index).Rotation.Y, Mesh(Index).Rotation.X, Mesh(Index).Rotation.Z
 D3DD.SetTransform D3DTS_WORLD, mRot
End If

If Mesh(Index).SetPos = True Then
 D3DD.GetTransform D3DTS_WORLD, smat
 mPos = smat
 mPos.m11 = 1: mPos.m22 = 1: mPos.m33 = 1: mPos.m44 = 1
 mPos.m41 = Mesh(Index).Position.X: mPos.m42 = Mesh(Index).Position.Y: mPos.m43 = Mesh(Index).Position.Z
 D3DD.SetTransform D3DTS_WORLD, mPos
End If

If Mesh(Index).ScaleObject = True Then
 D3DD.GetTransform D3DTS_WORLD, smat
 mScale = smat
 D3DXMatrixScaling mScale, Mesh(Index).ScaleO.X, Mesh(Index).ScaleO.Y, Mesh(Index).ScaleO.Z
 D3DXMatrixMultiply smat, mScale, smat
 D3DD.SetTransform D3DTS_WORLD, smat
End If

If Mesh(Index).Transparent = True Then
 D3DD.SetRenderState D3DRS_ALPHABLENDENABLE, 1
 D3DD.SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCCOLOR
 D3DD.SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCCOLOR
End If

If Mesh(Index).GlassEffect = True Then
 HoldRender1 = D3DD.GetRenderState(D3DRS_CULLMODE)
 HoldRender2 = D3DD.GetRenderState(D3DRS_ZWRITEENABLE)
 D3DD.SetRenderState D3DRS_CULLMODE, D3DCULL_NONE
 D3DD.SetRenderState D3DRS_ZWRITEENABLE, False
End If

If Mesh(Index).Enhanced = True Then: D3DD.SetRenderState D3DRS_PATCHSEGMENTS, Mesh(Index).EnhancedSegments: D3DD.SetRenderState D3DRS_PATCHEDGESTYLE, 1
If Mesh(Index).Mirror = True Then Start_Mirror_Render Camera
Render_Enviroment_Map Index, Mesh(Index).EnvTextY

If Mesh(Index).NewMaterial = True Then
 D3DD.SetMaterial Mesh(Index).Material
 If Not Mesh(Index).Texture Is Nothing Then
  D3DD.SetTexture 0, Mesh(Index).Texture
 End If
Else
 Material.diffuse.r = 1
 Material.diffuse.g = 1
 Material.diffuse.b = 1
 Material.diffuse.A = 1
 Material.Specular = Material.diffuse
 Material.Ambient = Material.diffuse
 D3DD.SetMaterial Material
 If Not Mesh(Index).Texture Is Nothing Then
  D3DD.SetTexture 0, Mesh(Index).Texture
 End If
End If

Start_Render_Filter Index

NumMat = 1
D3DD.SetRenderState D3DRS_ZBIAS, 1

For i = 0 To NumMat
 Mesh(Index).mMesh.DrawSubset i
Next

If Mesh(Index).Mirror = True Then End_Mirror_Render

If DoGloss = True Then
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
 For i = 0 To NumMat
  Mesh(Index).mMesh.DrawSubset i
 Next
 D3DD.SetRenderState D3DRS_DESTBLEND, D3DBLEND_ZERO
 D3DD.SetRenderState D3DRS_SRCBLEND, D3DBLEND_ZERO
 D3DD.SetRenderState D3DRS_ALPHABLENDENABLE, 0
End If
 
D3DD.SetRenderState D3DRS_ZBIAS, 0
End_Render_Filter Index

If Mesh(Index).Enhanced = True Then: D3DD.SetRenderState D3DRS_PATCHSEGMENTS, 0

D3DD.SetTransform D3DTS_WORLD, smat2
If Mesh(Index).GlassEffect = True Then
 D3DD.SetRenderState D3DRS_CULLMODE, HoldRender1
 D3DD.SetRenderState D3DRS_ZWRITEENABLE, HoldRender2
End If
D3DD.SetRenderState D3DRS_DESTBLEND, D3DBLEND_ZERO
D3DD.SetRenderState D3DRS_SRCBLEND, D3DBLEND_ZERO
D3DD.SetRenderState D3DRS_ALPHABLENDENABLE, 0
If Mesh(Index).EnviroMapped = True Then
 D3DD.SetTransform D3DTS_TEXTURE0, sEMat
 D3DD.SetTextureStageState 0, D3DTSS_TEXTURETRANSFORMFLAGS, D3DTEXF_NONE
 D3DD.SetTextureStageState 0, D3DTSS_TEXCOORDINDEX, D3DTEXF_NONE
End If
End Sub

Public Sub Cleanup_Objects_Engine()
Dim i As Integer
Set MapTexture = Nothing
For i = 0 To HNumMesh
 Set Mesh(i).mMesh = Nothing
 Set Mesh(i).Texture = Nothing
 Set Mesh(i).Billboard.BMesh = Nothing
 Set Mesh(i).Billboard.Texture = Nothing
Next
Set Cam = Nothing
Set Eng = Nothing
End Sub

Public Sub Enable_Enviroment_Mapping(Index As Integer, Enabled As Boolean, Optional Texture_Y_Position = 0.75)
If Enabled = True Then
 Mesh(Index).EnviroMapped = True
 Mesh(Index).EnvTextY = Texture_Y_Position
End If
End Sub

Private Sub Render_Enviroment_Map(Index As Integer, Optional TextureYPos As Single = 0.75)
If Mesh(Index).EnviroMapped = True Then
 D3DD.GetTransform D3DTS_TEXTURE0, sEMat
 EMat = sEMat
 EMat.m11 = 0.03: EMat.m12 = 0#: EMat.m22 = -0.07: EMat.m33 = 0.1
 EMat.m41 = 0.7: EMat.m42 = TextureYPos

 D3DD.SetTransform D3DTS_TEXTURE0, EMat
 D3DD.SetTextureStageState 0, D3DTSS_TEXTURETRANSFORMFLAGS, D3DTTFF_COUNT2
 D3DD.SetTextureStageState 0, D3DTSS_TEXCOORDINDEX, D3DTSS_TCI_CAMERASPACEPOSITION
End If
End Sub

'Doesnt Work Right
Private Sub Create_Mirror(Index As Integer, Width As Single, Height As Single)
Set Mesh(Index).mMesh = Nothing
Set Mesh(Index).mMesh = pCreate2DWall(Height, Width, 1, 1)
Mesh(Index).Mirror = True
End Sub

Private Sub Start_Mirror_Render(Camera As XCamera)
Dim rMat As D3DMATRIX
Dim Plane As D3DPLANE
Dim A As D3DVECTOR
Dim b As D3DVECTOR
Dim C As D3DVECTOR
Dim D As D3DVECTOR

A.X = -1.5: A.Y = 1.5: A.Z = 3#
b.X = 1.5: b.Y = 1.5: b.Z = 3#
C.X = -1.5: C.Y = -1.5: C.Z = 3#
D.X = 1.5: D.Y = -1.5: D.Z = 3#

D3DD.GetTransform D3DTS_WORLD, WorldMat
D3DXPlaneFromPoints Plane, A, b, C
D3DXMatrixReflect rMat, Plane
D3DD.SetTransform D3DTS_WORLD, rMat

D3DD.SetRenderState D3DRS_CULLMODE, D3DCULL_CW
D3DXPlaneFromPoints Plane, b, A, Camera.Get_Camera_PositionEX
D3DD.SetClipPlane 0, Plane
D3DXPlaneFromPoints Plane, D, b, Camera.Get_Camera_PositionEX
D3DD.SetClipPlane 1, Plane
D3DXPlaneFromPoints Plane, C, D, Camera.Get_Camera_PositionEX
D3DD.SetClipPlane 2, Plane
D3DXPlaneFromPoints Plane, A, C, Camera.Get_Camera_PositionEX
D3DD.SetClipPlane 3, Plane

D3DD.SetRenderState D3DRS_CLIPPLANEENABLE, D3DCLIPPLANE0 Or D3DCLIPPLANE1 Or D3DCLIPPLANE2 Or D3DCLIPPLANE3
End Sub

Private Sub End_Mirror_Render()
D3DD.SetTransform D3DTS_WORLD, WorldMat
D3DD.SetRenderState D3DRS_CLIPPLANEENABLE, 0
D3DD.SetRenderState D3DRS_CULLMODE, D3DCULL_CCW

D3DD.SetRenderState D3DRS_ALPHABLENDENABLE, True
D3DD.SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCALPHA
D3DD.SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA
End Sub

Private Sub Create_Sphere(Index As Integer, Radius As Single, Optional Smoothness As Long = 30, Optional Texture_Filename As String)
On Local Error Resume Next
Set Mesh(Index).mMesh = Nothing
Set Mesh(Index).mMesh = Eng.Direct3DX.CreateSphere(D3DD, Radius, Smoothness, Smoothness, Nothing)
Change_FVF Mesh(Index).mMesh, D3DFVF_VERTEX
If Texture_Filename <> "" Then
 Set Mesh(Index).Texture = Eng.Direct3DX.CreateTextureFromFile(D3DD, Texture_Filename)
 UseOdd = True
End If
ObjectCount = ObjectCount + 1
End Sub

Private Sub Create_Teapot(Index As Integer, Optional Texture_Filename As String)
On Local Error Resume Next
Set Mesh(Index).mMesh = Nothing
Set Mesh(Index).mMesh = Eng.Direct3DX.CreateTeapot(D3DD, Nothing)
Change_FVF Mesh(Index).mMesh, D3DFVF_VERTEX
If Texture_Filename <> "" Then
 Set Mesh(Index).Texture = Eng.Direct3DX.CreateTextureFromFile(D3DD, Texture_Filename)
 UseOdd = True
End If
ObjectCount = ObjectCount + 1
End Sub

Private Sub Create_Torus(Index As Integer, OuterRadius As Single, InnerRadius As Single, Optional Smoothness As Long = 30, Optional Texture_Filename As String)
On Local Error Resume Next
Set Mesh(Index).mMesh = Nothing
Set Mesh(Index).mMesh = Eng.Direct3DX.CreateTorus(D3DD, InnerRadius, OuterRadius, Smoothness, Smoothness, Nothing)
Change_FVF Mesh(Index).mMesh, D3DFVF_VERTEX
If Texture_Filename <> "" Then
 Set Mesh(Index).Texture = Eng.Direct3DX.CreateTextureFromFile(D3DD, Texture_Filename)
 UseOdd = True
End If
ObjectCount = ObjectCount + 1
End Sub

Private Sub Create_Cylinder(Index As Integer, OuterRadius As Single, InnerRadius As Single, Length As Single, Optional Smoothness As Long = 30, Optional Texture_Filename As String)
On Local Error Resume Next
Set Mesh(Index).mMesh = Nothing
Set Mesh(Index).mMesh = Eng.Direct3DX.CreateCylinder(D3DD, InnerRadius, OuterRadius, Length, Smoothness, Smoothness, Nothing)
'Change_FVF Mesh(Index).mMesh, D3DFVF_VERTEX
If Texture_Filename <> "" Then
 Set Mesh(Index).Texture = Eng.Direct3DX.CreateTextureFromFile(D3DD, Texture_Filename)
 UseOdd = True
End If
ObjectCount = ObjectCount + 1
End Sub

Private Sub Scale_Down_Texture(Mesh As D3DXMesh, Amount As Single)
Dim i As Long
Dim TextureTool As MESHTOOL
MESHTOOL_INIT TextureTool, Mesh

For i = 0 To TextureTool.NumVertices - 1
TextureTool.Vertices(i).tU = TextureTool.Vertices(i).tU * Amount
TextureTool.Vertices(i).tV = TextureTool.Vertices(i).tV * Amount
Next
    
D3DVertexBuffer8SetData TextureTool.VertB, 0, Len(TextureTool.Vertices(0)) * TextureTool.NumVertices, 0, TextureTool.Vertices(0)
MESHTOOL_DESTROY TextureTool
End Sub

Public Sub Flip_Mesh_Normals(Mesh As D3DXMesh)
On Local Error Resume Next
Dim Count As Long
Dim Size As Long
Dim i As Long
Dim Verts() As D3DVERTEX
Dim VB As Direct3DVertexBuffer8

Set VB = Mesh.GetVertexBuffer()
    
Size = Eng.Direct3DX.GetFVFVertexSize(Mesh.GetFVF())
Count = Mesh.GetNumVertices()

ReDim Verts(Count)
        
If Mesh.GetFVF() = D3DFVF_VERTEX Then
 D3DVertexBuffer8GetData VB, 0, Size * Count, 0, Verts(0)
 For i = 0 To Count - 1
  Verts(i).nx = -Verts(i).nx
  Verts(i).ny = -Verts(i).ny
  Verts(i).nz = -Verts(i).nz
 Next
 D3DVertexBuffer8SetData VB, 0, Size * Count, 0, Verts(0)
Else
  Exit Sub
End If
End Sub

Public Sub Change_FVF(Mesh As D3DXMesh, FVF_Type As FVFType, Optional FVF As Long)
Dim tempMesh As D3DXMesh
On Local Error GoTo FVFERR

ErrStr = "Error Clone Mesh FVF"
If Mesh Is Nothing Then Exit Sub

If FVF_Type = FVFT_D3DFVF_COLORVERTEX Then
 Set tempMesh = Mesh.CloneMeshFVF(D3DXMESH_MANAGED, D3DFVF_COLORVERTEX, D3DD)
 Set Mesh = tempMesh
ElseIf FVF_Type = FVFT_D3DFVF_MIRRORVERTEX Then
 Set tempMesh = Mesh.CloneMeshFVF(D3DXMESH_MANAGED, D3DFVF_MIRRORVERTEX, D3DD)
 Set Mesh = tempMesh
ElseIf FVF_Type = FVFT_D3DFVF_VERTEX Then
 Set tempMesh = Mesh.CloneMeshFVF(D3DXMESH_MANAGED, D3DFVF_VERTEX, D3DD)
 Set Mesh = tempMesh
ElseIf FVF_Type = Custom Then
 Set tempMesh = Mesh.CloneMeshFVF(D3DXMESH_MANAGED, FVF, D3DD)
 Set Mesh = tempMesh
End If
 Set tempMesh = Nothing
Exit Sub

FVFERR:
ErrorToFile "ErrorFVF", ErrStr, PossibleCause
End Sub

Public Sub Create_Box(Index As Integer, Width As Single, Height As Single, Depth As Single, Optional Texture_Filename As String)
On Local Error Resume Next
Set Mesh(Index).mMesh = Nothing
Set Mesh(Index).mMesh = pCreateBox(Width, Height, Depth)
If Texture_Filename <> "" Then
 Set Mesh(Index).Texture = Eng.Direct3DX.CreateTextureFromFile(D3DD, Texture_Filename)
End If
ObjectCount = ObjectCount + 1
End Sub

Private Function pCreateBox(Width As Single, Height As Single, Depth As Single) As D3DXMesh
'Private
Dim rMesh As D3DXMesh
Dim vertexbuffer As Direct3DVertexBuffer8
Dim Verts(28) As D3DVERTEX
Dim indices(36) As Integer
Dim W As Single, D As Single, h1 As Single, H2 As Single
W = Width / 2
H2 = Height / 2
h1 = -Height / 2
D = Depth / 2
    
Set rMesh = Eng.Direct3DX.CreateMeshFVF(4 * 6, 6 * 6, D3DXMESH_MANAGED, D3DFVF_VERTEX, D3DD)
With Verts(0): .X = -W: .Y = H2: .Z = -D: .nz = 1: .tU = 0: .tV = 0: End With
With Verts(1): .X = W: .Y = H2: .Z = -D: .nz = 1: .tU = 1: .tV = 0: End With
With Verts(2): .X = W: .Y = h1: .Z = -D: .nz = 1: .tU = 1: .tV = 1: End With
With Verts(3): .X = -W: .Y = h1: .Z = -D: .nz = 1: .tU = 0: .tV = 1: End With
indices(0) = 0: indices(1) = 1: indices(2) = 2
indices(3) = 0: indices(4) = 2: indices(5) = 3
With Verts(4): .X = -W: .Y = h1: .Z = D: .nz = -1: .tU = 0: .tV = 1: End With
With Verts(5): .X = W: .Y = h1: .Z = D: .nz = -1: .tU = 1: .tV = 1: End With
With Verts(6): .X = W: .Y = H2: .Z = D: .nz = -1: .tU = 1: .tV = 0: End With
With Verts(7): .X = -W: .Y = H2: .Z = D: .nz = -1: .tU = 0: .tV = 0: End With
indices(6) = 4: indices(7) = 5: indices(8) = 6
indices(9) = 4: indices(10) = 6: indices(11) = 7
With Verts(8): .X = W: .Y = h1: .Z = -D: .nx = -1: .tU = 0: .tV = 0: End With
With Verts(9): .X = W: .Y = h1: .Z = D: .nx = -1: .tU = 1: .tV = 0: End With
With Verts(10): .X = W: .Y = H2: .Z = D: .nx = -1: .tU = 1: .tV = 1: End With
With Verts(11): .X = W: .Y = H2: .Z = -D: .nx = -1: .tU = 0: .tV = 1: End With
indices(12) = 8: indices(13) = 9: indices(14) = 10
indices(15) = 8: indices(16) = 10: indices(17) = 11
With Verts(16): .X = -W: .Y = H2: .Z = -D: .nx = 1: .tU = 0: .tV = 1: End With
With Verts(17): .X = -W: .Y = H2: .Z = D: .nx = 1: .tU = 1: .tV = 1: End With
With Verts(18): .X = -W: .Y = h1: .Z = D: .nx = 1: .tU = 1: .tV = 0: End With
With Verts(19): .X = -W: .Y = h1: .Z = -D: .nx = 1: .tU = 0: .tV = 0: End With
indices(18) = 16: indices(19) = 17: indices(20) = 18
indices(21) = 16: indices(22) = 18: indices(23) = 19
With Verts(20): .X = -W: .Y = H2: .Z = -D: .ny = -1: .tU = 0: .tV = 0: End With
With Verts(21): .X = -W: .Y = H2: .Z = D: .ny = -1: .tU = 1: .tV = 0: End With
With Verts(22): .X = W: .Y = H2: .Z = D: .ny = -1: .tU = 1: .tV = 1: End With
With Verts(23): .X = W: .Y = H2: .Z = -D: .ny = -1: .tU = 0: .tV = 1: End With
indices(24) = 20: indices(25) = 21: indices(26) = 22
indices(27) = 20: indices(28) = 22: indices(29) = 23
With Verts(24): .X = W: .Y = h1: .Z = -D: .ny = 1: .tU = 0: .tV = 1: End With
With Verts(25): .X = W: .Y = h1: .Z = D: .ny = 1: .tU = 1: .tV = 1: End With
With Verts(26): .X = -W: .Y = h1: .Z = D: .ny = 1: .tU = 1: .tV = 0: End With
With Verts(27): .X = -W: .Y = h1: .Z = -D: .ny = 1: .tU = 0: .tV = 0: End With
indices(30) = 24: indices(31) = 25: indices(32) = 26
indices(33) = 24: indices(34) = 26: indices(35) = 27
D3DXMeshVertexBuffer8SetData rMesh, 0, Len(Verts(0)) * 28, 0, Verts(0)
D3DXMeshIndexBuffer8SetData rMesh, 0, Len(indices(0)) * 36, 0, indices(0)
        
Set pCreateBox = rMesh
End Function

Public Sub Create_Floor_Or_Ceiling(Index As Integer, Width As Single, Depth As Single, Optional Texture_Filename As String)
On Local Error Resume Next
Set Mesh(Index).mMesh = Nothing
Set Mesh(Index).mMesh = pCreateFloor(Width, Depth)
If Texture_Filename <> "" Then
 Set Mesh(Index).Texture = Eng.Direct3DX.CreateTextureFromFile(D3DD, Texture_Filename)
End If
ObjectCount = ObjectCount + 1
End Sub

Private Function pCreateFloor(Width As Single, Depth As Single) As D3DXMesh
'Private
Dim rMesh As D3DXMesh
Dim vertexbuffer As Direct3DVertexBuffer8
Dim Verts(8) As D3DVERTEX
Dim indices(12) As Integer
Dim W As Single, H As Single, H2 As Single
W = Width / 2
H = Depth / 2
H2 = -Depth / 2
    
Set rMesh = Eng.Direct3DX.CreateMeshFVF(8, 12, D3DXMESH_MANAGED, D3DFVF_VERTEX, D3DD)
With Verts(0): .X = -W: .Y = 0: .Z = -H2: .nz = 1: .tU = 0: .tV = 1: End With
With Verts(1): .X = W: .Y = 0: .Z = -H2: .nz = 1: .tU = 1: .tV = 1: End With
With Verts(2): .X = W: .Y = 0: .Z = -H: .nz = 1: .tU = 1: .tV = 0: End With
With Verts(3): .X = -W: .Y = 0: .Z = -H: .nz = 1: .tU = 0: .tV = 0: End With
indices(0) = 0: indices(1) = 1: indices(2) = 2
indices(3) = 0: indices(4) = 2: indices(5) = 3
With Verts(4): .X = -W: .Y = 0: .Z = H: .nz = -1: .tU = 0: .tV = 0: End With
With Verts(5): .X = W: .Y = 0: .Z = H: .nz = -1: .tU = 1: .tV = 0: End With
With Verts(6): .X = W: .Y = 0: .Z = H2: .nz = -1: .tU = 1: .tV = 1: End With
With Verts(7): .X = -W: .Y = 0: .Z = H2: .nz = -1: .tU = 0: .tV = 1: End With
indices(6) = 4: indices(7) = 5: indices(8) = 6
indices(9) = 4: indices(10) = 6: indices(11) = 7
D3DXMeshVertexBuffer8SetData rMesh, 0, Len(Verts(0)) * 8, 0, Verts(0)
D3DXMeshIndexBuffer8SetData rMesh, 0, Len(indices(0)) * 12, 0, indices(0)
        
Set pCreateFloor = rMesh
End Function

Public Sub Create_3D_Wall(Index As Integer, Width As Single, Height As Single, Depth As Single, Optional Texture_Filename As String, Optional Texture_U_Power = 1, Optional Texture_V_Power = 1)
On Local Error Resume Next
Set Mesh(Index).mMesh = Nothing
Set Mesh(Index).mMesh = pCreate3DWall(Width, Height, Depth, Texture_U_Power, Texture_V_Power)
If Texture_Filename <> "" Then
 Set Mesh(Index).Texture = Eng.Direct3DX.CreateTextureFromFile(D3DD, Texture_Filename)
End If
ObjectCount = ObjectCount + 1
End Sub

Private Function pCreate3DWall(Width As Single, Height As Single, Depth As Single, Optional Texture_U_Power = 1, Optional Texture_V_Power = 1) As D3DXMesh
'Private
'Same as box except you can change tU and tV coords
'which means You can use filters on it :)
Dim rMesh As D3DXMesh
Dim vertexbuffer As Direct3DVertexBuffer8
Dim Verts(28) As D3DVERTEX
Dim indices(36) As Integer
Dim W As Single, D As Single, h1 As Single, H2 As Single
W = Width / 2
H2 = Height / 2
h1 = -Height / 2
D = Depth / 2
    
Set rMesh = Eng.Direct3DX.CreateMeshFVF(4 * 6, 6 * 6, D3DXMESH_MANAGED, D3DFVF_VERTEX, D3DD)
With Verts(0): .X = -W: .Y = H2: .Z = -D: .nz = 1: .tU = 0: .tV = 0: End With
With Verts(1): .X = W: .Y = H2: .Z = -D: .nz = 1: .tU = Texture_U_Power: .tV = 0: End With
With Verts(2): .X = W: .Y = h1: .Z = -D: .nz = 1: .tU = Texture_U_Power: .tV = Texture_V_Power: End With
With Verts(3): .X = -W: .Y = h1: .Z = -D: .nz = 1: .tU = 0: .tV = Texture_V_Power: End With
indices(0) = 0: indices(1) = 1: indices(2) = 2
indices(3) = 0: indices(4) = 2: indices(5) = 3
With Verts(4): .X = -W: .Y = h1: .Z = D: .nz = -1: .tU = 0: .tV = Texture_V_Power: End With
With Verts(5): .X = W: .Y = h1: .Z = D: .nz = -1: .tU = Texture_U_Power: .tV = Texture_V_Power: End With
With Verts(6): .X = W: .Y = H2: .Z = D: .nz = -1: .tU = Texture_U_Power: .tV = 0: End With
With Verts(7): .X = -W: .Y = H2: .Z = D: .nz = -1: .tU = 0: .tV = 0: End With
indices(6) = 4: indices(7) = 5: indices(8) = 6
indices(9) = 4: indices(10) = 6: indices(11) = 7
With Verts(8): .X = W: .Y = h1: .Z = -D: .nx = -1: .tU = 0: .tV = 0: End With
With Verts(9): .X = W: .Y = h1: .Z = D: .nx = -1: .tU = Texture_U_Power: .tV = 0: End With
With Verts(10): .X = W: .Y = H2: .Z = D: .nx = -1: .tU = Texture_U_Power: .tV = Texture_V_Power: End With
With Verts(11): .X = W: .Y = H2: .Z = -D: .nx = -1: .tU = 0: .tV = Texture_V_Power: End With
indices(12) = 8: indices(13) = 9: indices(14) = 10
indices(15) = 8: indices(16) = 10: indices(17) = 11
With Verts(16): .X = -W: .Y = H2: .Z = -D: .nx = 1: .tU = 0: .tV = Texture_V_Power: End With
With Verts(17): .X = -W: .Y = H2: .Z = D: .nx = 1: .tU = Texture_U_Power: .tV = Texture_V_Power: End With
With Verts(18): .X = -W: .Y = h1: .Z = D: .nx = 1: .tU = Texture_U_Power: .tV = 0: End With
With Verts(19): .X = -W: .Y = h1: .Z = -D: .nx = 1: .tU = 0: .tV = 0: End With
indices(18) = 16: indices(19) = 17: indices(20) = 18
indices(21) = 16: indices(22) = 18: indices(23) = 19
With Verts(20): .X = -W: .Y = H2: .Z = -D: .ny = -1: .tU = 0: .tV = 0: End With
With Verts(21): .X = -W: .Y = H2: .Z = D: .ny = -1: .tU = Texture_U_Power: .tV = 0: End With
With Verts(22): .X = W: .Y = H2: .Z = D: .ny = -1: .tU = Texture_U_Power: .tV = Texture_V_Power: End With
With Verts(23): .X = W: .Y = H2: .Z = -D: .ny = -1: .tU = 0: .tV = Texture_V_Power: End With
indices(24) = 20: indices(25) = 21: indices(26) = 22
indices(27) = 20: indices(28) = 22: indices(29) = 23
With Verts(24): .X = W: .Y = h1: .Z = -D: .ny = 1: .tU = 0: .tV = Texture_V_Power: End With
With Verts(25): .X = W: .Y = h1: .Z = D: .ny = 1: .tU = Texture_U_Power: .tV = Texture_V_Power: End With
With Verts(26): .X = -W: .Y = h1: .Z = D: .ny = 1: .tU = Texture_U_Power: .tV = 0: End With
With Verts(27): .X = -W: .Y = h1: .Z = -D: .ny = 1: .tU = 0: .tV = 0: End With
indices(30) = 24: indices(31) = 25: indices(32) = 26
indices(33) = 24: indices(34) = 26: indices(35) = 27
D3DXMeshVertexBuffer8SetData rMesh, 0, Len(Verts(0)) * 28, 0, Verts(0)
D3DXMeshIndexBuffer8SetData rMesh, 0, Len(indices(0)) * 36, 0, indices(0)
        
Set pCreate3DWall = rMesh
End Function

Public Sub Create_2D_Wall(Index As Integer, Height As Single, Width As Single, Optional Texture_Filename As String, Optional Texture_U_Power = 1, Optional Texture_V_Power = 1)
On Local Error Resume Next
Set Mesh(Index).mMesh = Nothing
Set Mesh(Index).mMesh = pCreate2DWall(Height, Width, Texture_U_Power, Texture_V_Power)
If Texture_Filename <> "" Then
 Set Mesh(Index).Texture = Eng.Direct3DX.CreateTextureFromFile(D3DD, Texture_Filename)
End If
ObjectCount = ObjectCount + 1
End Sub

Private Function pCreate2DWall(Height As Single, Width As Single, Optional Texture_U_Power = 1, Optional Texture_V_Power = 1) As D3DXMesh
'Private
Dim rMesh As D3DXMesh
Dim vertexbuffer As Direct3DVertexBuffer8
Dim Verts(8) As D3DVERTEX
Dim indices(12) As Integer
Dim W As Single, H As Single, W2 As Single
H = Height / 2
W = Width / 2
W2 = -Width / 2
    
Set rMesh = Eng.Direct3DX.CreateMeshFVF(8, 12, D3DXMESH_MANAGED, D3DFVF_VERTEX, D3DD)
With Verts(0): .X = -H: .Y = W2: .Z = -0: .nz = 1: .tU = 0: .tV = Texture_V_Power: End With
With Verts(1): .X = H: .Y = W2: .Z = -0: .nz = 1: .tU = Texture_U_Power: .tV = Texture_V_Power: End With
With Verts(2): .X = H: .Y = W: .Z = -0: .nz = 1: .tU = Texture_U_Power: .tV = 0: End With
With Verts(3): .X = -H: .Y = W: .Z = -0: .nz = 1: .tU = 0: .tV = 0: End With
indices(0) = 0: indices(1) = 1: indices(2) = 2
indices(3) = 0: indices(4) = 2: indices(5) = 3
With Verts(4): .X = -H: .Y = W: .Z = 0: .nz = -1: .tU = 0: .tV = 0: End With
With Verts(5): .X = H: .Y = W: .Z = 0: .nz = -1: .tU = Texture_U_Power: .tV = 0: End With
With Verts(6): .X = H: .Y = W2: .Z = 0: .nz = -1: .tU = Texture_U_Power: .tV = Texture_V_Power: End With
With Verts(7): .X = -H: .Y = W2: .Z = 0: .nz = -1: .tU = 0: .tV = Texture_V_Power: End With
indices(6) = 4: indices(7) = 5: indices(8) = 6
indices(9) = 4: indices(10) = 6: indices(11) = 7
D3DXMeshVertexBuffer8SetData rMesh, 0, Len(Verts(0)) * 8, 0, Verts(0)
D3DXMeshIndexBuffer8SetData rMesh, 0, Len(indices(0)) * 12, 0, indices(0)
        
Set pCreate2DWall = rMesh
End Function

Public Sub Create_Billboard(Index As Integer, Width As Single, Height As Single, Texture_Filename As String, Optional Texture_U_Power = 1, Optional Texture_V_Power = 1)
On Local Error Resume Next
Set Mesh(Index).Billboard.BMesh = Nothing
Set Mesh(Index).Billboard.BMesh = pCreateBillboard(Height, Width, Texture_U_Power, Texture_V_Power)
If Texture_Filename <> "" Then
 Set Mesh(Index).Billboard.Texture = Eng.Direct3DX.CreateTextureFromFileEx(D3DD, Texture_Filename, 256, 256, D3DX_DEFAULT, 0, D3DFMT_A1R5G5B5, D3DPOOL_MANAGED, D3DX_DEFAULT, D3DX_DEFAULT, &HFF000000, ByVal 0, ByVal 0)
End If
ObjectCount = ObjectCount + 1
End Sub

Public Sub Set_Billboard_Position(Index As Integer, X As Single, Y As Single, Z As Single)
Mesh(Index).Billboard.Position = vec3(X, Y, Z)
Mesh(Index).Billboard.SetPos = True
End Sub

Public Sub Render_Billboard(Index As Integer, Camera As XCamera)
On Local Error Resume Next
Dim i As Long
Dim NumMat As Long
Dim CamOX As Single
Dim CamOZ As Single
Dim Material As D3DMATERIAL8
Dim Caps As D3DCAPS8
Dim smat As D3DMATRIX
Dim smat2 As D3DMATRIX

Update_Clip_Planes Index, Camera
If Sphere_Visible(Index) = False Then Exit Sub

D3DD.SetRenderState D3DRS_ALPHABLENDENABLE, 1
D3DD.SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCALPHA
D3DD.SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA

If ((Caps.AlphaCmpCaps And D3DPCMPCAPS_GREATEREQUAL) = D3DPCMPCAPS_GREATEREQUAL) Then
 D3DD.SetRenderState D3DRS_ALPHATESTENABLE, 1
 D3DD.SetRenderState D3DRS_ALPHAREF, &H8&
 D3DD.SetRenderState D3DRS_ALPHAFUNC, D3DCMP_GREATEREQUAL
End If

CamOX = Camera.Get_Camera_Orientation_X * 2
CamOZ = Camera.Get_Camera_Orientation_Z * 2

D3DD.GetTransform D3DTS_WORLD, smat2
D3DD.GetTransform D3DTS_WORLD, smat
Mesh(Index).Billboard.mRot = smat
D3DXMatrixRotationYawPitchRoll Mesh(Index).Billboard.mRot, -((Camera.Get_Camera_Orientation_Y) * (CamOX / CamOZ) - 110), 0, 0
D3DD.SetTransform D3DTS_WORLD, Mesh(Index).Billboard.mRot

If Mesh(Index).Billboard.SetPos = True Then
 D3DD.GetTransform D3DTS_WORLD, smat
 BillboardMatrix = smat
 BillboardMatrix.m11 = 1: BillboardMatrix.m22 = 1: BillboardMatrix.m33 = 1: BillboardMatrix.m44 = 1
 BillboardMatrix.m41 = Mesh(Index).Billboard.Position.X
 BillboardMatrix.m42 = Mesh(Index).Billboard.Position.Y
 BillboardMatrix.m43 = Mesh(Index).Billboard.Position.Z
 D3DD.SetTransform D3DTS_WORLD, BillboardMatrix
End If

Material.diffuse.r = 1: Material.diffuse.g = 1: Material.diffuse.b = 1
Material.Ambient = Material.diffuse
NumMat = 1
D3DD.SetMaterial Material
D3DD.SetTexture 0, Mesh(Index).Billboard.Texture
For i = 0 To NumMat
 Mesh(Index).Billboard.BMesh.DrawSubset i
Next

D3DD.SetTransform D3DTS_WORLD, smat2
D3DD.SetRenderState D3DRS_SRCBLEND, 0
D3DD.SetRenderState D3DRS_DESTBLEND, 0
D3DD.SetRenderState D3DRS_ALPHATESTENABLE, 0
D3DD.SetRenderState D3DRS_ALPHABLENDENABLE, 0
End Sub

Private Function pCreateBillboard(Height As Single, Width As Single, Optional Texture_U_Power = 1, Optional Texture_V_Power = 1) As D3DXMesh
'Private
Dim rMesh As D3DXMesh
Dim vertexbuffer As Direct3DVertexBuffer8
Dim Verts(8) As D3DVERTEX
Dim indices(12) As Integer
Dim W As Single, H As Single, W2 As Single
H = Height / 2
W = Width / 2
W2 = -Width / 2
    
Set rMesh = Eng.Direct3DX.CreateMeshFVF(8, 12, D3DXMESH_MANAGED, D3DFVF_VERTEX, D3DD)
With Verts(0): .X = -H: .Y = W2: .Z = -0: .nz = 1: .tU = 0: .tV = Texture_V_Power: End With
With Verts(1): .X = H: .Y = W2: .Z = -0: .nz = 1: .tU = Texture_U_Power: .tV = Texture_V_Power: End With
With Verts(2): .X = H: .Y = W: .Z = -0: .nz = 1: .tU = Texture_U_Power: .tV = 0: End With
With Verts(3): .X = -H: .Y = W: .Z = -0: .nz = 1: .tU = 0: .tV = 0: End With
indices(0) = 0: indices(1) = 1: indices(2) = 2
indices(3) = 0: indices(4) = 2: indices(5) = 3
With Verts(4): .X = -H: .Y = W: .Z = 0: .nz = -1: .tU = 0: .tV = 0: End With
With Verts(5): .X = H: .Y = W: .Z = 0: .nz = -1: .tU = Texture_U_Power: .tV = 0: End With
With Verts(6): .X = H: .Y = W2: .Z = 0: .nz = -1: .tU = Texture_U_Power: .tV = Texture_V_Power: End With
With Verts(7): .X = -H: .Y = W2: .Z = 0: .nz = -1: .tU = 0: .tV = Texture_V_Power: End With
indices(6) = 4: indices(7) = 5: indices(8) = 6
indices(9) = 4: indices(10) = 6: indices(11) = 7
D3DXMeshVertexBuffer8SetData rMesh, 0, Len(Verts(0)) * 8, 0, Verts(0)
D3DXMeshIndexBuffer8SetData rMesh, 0, Len(indices(0)) * 12, 0, indices(0)
        
Set pCreateBillboard = rMesh
End Function

Public Sub Compute_Normals(Mesh As D3DXMesh)
On Local Error Resume Next
Dim bm As D3DXBaseMesh
Set bm = Mesh
Eng.Direct3DX.ComputeNormals bm
Set Mesh = bm
Set bm = Nothing
End Sub

Public Function Check_Camera_To_Object_Collision(Index As Integer, CameraPos As D3DVECTOR, ObjectRadius As Single, Optional OffsetX = 0, Optional OffsetZ = 0) As Boolean
If CameraPos.X > (Mesh(Index).Position.X + OffsetX) - ObjectRadius _
   And CameraPos.X < (Mesh(Index).Position.X + OffsetX) + ObjectRadius _
   And CameraPos.Z > (Mesh(Index).Position.Z + OffsetZ) - ObjectRadius _
   And CameraPos.Z < (Mesh(Index).Position.Z + OffsetZ) + ObjectRadius Then
 Check_Camera_To_Object_Collision = True
Else
 Check_Camera_To_Object_Collision = False
End If
End Function

Public Sub Tesselate_Mesh(Index As Integer, NumSegments As Single)
On Local Error Resume Next
Dim tmpMesh As D3DXMesh
Dim bAdjacency As D3DXBuffer

If NumSegments <= 0 Then NumSegments = 0
If NumSegments >= 1.9 Then NumSegments = 1.9

Set tmpMesh = Nothing
Set tmpMesh = Mesh(Index).mMesh

Change_FVF tmpMesh, D3DFVF_VERTEX
Compute_Normals tmpMesh
Eng.Direct3DX.TesselateMesh tmpMesh, Nothing, NumSegments, True

Set Mesh(Index).mMesh = tmpMesh
Set tmpMesh = Nothing
Mesh(Index).Enhanced = True
Mesh(Index).EnhancedSegments = NumSegments
End Sub

Public Sub Tesselate_Mesh2(Index As Integer, NumSegments As Single)
On Local Error Resume Next
Dim VertB As Direct3DVertexBuffer8
Dim NumVertices As Long
Dim Vertices() As CUSTOMVERTEX
Dim tmpMesh As D3DXMesh

Set tmpMesh = Nothing
Set tmpMesh = Mesh(Index).mMesh

Set VertB = tmpMesh.GetVertexBuffer
NumVertices = tmpMesh.GetNumVertices
ReDim Vertices(NumVertices)
D3DVertexBuffer8GetData VertB, 0, Len(Vertices(0)) * NumVertices, 0, Vertices(0)

'Add Vertices and faces here

D3DVertexBuffer8SetData VertB, 0, Len(Vertices(0)) * NumVertices, 0, Vertices(0)
Set Mesh(Index).mMesh = tmpMesh
Set tmpMesh = Nothing
Set VertB = Nothing
ReDim Vertices(0)
End Sub

Private Function Sphere_Visible(Index As Integer) As Boolean
Dim i As Long
Dim Distance As Single
    
For i = 0 To numClipPlanes - 1
 Distance = D3DXMATH_PLANE.D3DXPlaneDotCoord(ClipPlanes(i), Mesh(Index).SphereCenter)
 If Distance < -1 * Mesh(Index).SphereRadius Then
  Sphere_Visible = False
  Exit Function
 End If
Next

Sphere_Visible = True
End Function

Public Sub Compute_Clip_Planes(Index As Integer, Compute As Boolean)
If Compute = True Then
 Mesh(Index).ComputedClipPlane = True
 If Mesh(Index).HasSphere = False Then
  Compute_Bounding_Sphere Index
 End If
Else
 Mesh(Index).ComputedClipPlane = False
End If
End Sub

Private Sub Update_Clip_Planes(Index As Integer, Camera As XCamera)
If Mesh(Index).ComputedClipPlane = True And Mesh(Index).HasSphere = True Then
 pComputeClipPlanes Camera.Get_Camera_OrientationEX, Camera.Get_Camera_PositionEX, vec3(0, Camera.Get_Camera_Eye_Level, 0), _
                    Camera.Get_Camera_Field_Of_View, Camera.Get_Camera_Near, Camera.Get_Camera_Far, Camera.Get_Camera_Aspect
End If
End Sub

Private Sub pComputeClipPlanes(veye As D3DVECTOR, vat As D3DVECTOR, vUp As D3DVECTOR, FOV As Single, front As Single, back As Single, Aspect As Single)
Dim vDir As D3DVECTOR
Dim vright As D3DVECTOR
Dim vFrontCenter As D3DVECTOR
Dim vFrontUp As D3DVECTOR
Dim vFrontRight As D3DVECTOR
Dim vBackCenter As D3DVECTOR
Dim vBackRight As D3DVECTOR
Dim vbackLeft As D3DVECTOR
Dim vBackRightTop As D3DVECTOR
Dim vBackLeftTop As D3DVECTOR
Dim vBackRightBot As D3DVECTOR
Dim vBackLeftBot As D3DVECTOR
Dim DX As Single
Dim dy As Single
    
D3DXVec3Subtract vDir, vat, veye
D3DXVec3Normalize vDir, vDir
D3DXVec3Normalize vUp, vUp
D3DXVec3Cross vright, vDir, vUp
    
DX = Tan(FOV / 2) * back
dy = DX * Aspect
        
D3DXVec3Scale vBackCenter, vDir, back
D3DXVec3Add vBackCenter, vBackCenter, veye
D3DXVec3Scale vBackRight, vright, DX
D3DXVec3Add vBackRight, vBackCenter, vBackRight
D3DXVec3Scale vbackLeft, vright, -DX
D3DXVec3Add vbackLeft, vBackCenter, vbackLeft
D3DXVec3Scale vBackRightTop, vUp, dy
D3DXVec3Add vBackRightTop, vBackRight, vBackRightTop
D3DXVec3Scale vBackLeftTop, vUp, dy
D3DXVec3Add vBackLeftTop, vBackRight, vBackLeftTop
D3DXVec3Scale vBackRightBot, vUp, -dy
D3DXVec3Add vBackRightBot, vBackRight, vBackRightBot
D3DXVec3Scale vBackLeftBot, vUp, -dy
D3DXVec3Add vBackLeftBot, vBackRight, vBackLeftBot
D3DXVec3Scale vFrontCenter, vDir, front
D3DXVec3Add vFrontCenter, vFrontCenter, veye
D3DXVec3Scale vFrontRight, vright, DX
D3DXVec3Add vFrontRight, vFrontCenter, vFrontRight
D3DXVec3Scale vFrontUp, vUp, dy
D3DXVec3Add vFrontUp, vFrontCenter, vFrontUp
ReDim ClipPlanes(6)
numClipPlanes = 6
D3DXPlaneFromPointNormal ClipPlanes(0), veye, vDir
Dim vnegdir As D3DVECTOR
D3DXVec3Scale vnegdir, vDir, -1
D3DXPlaneFromPointNormal ClipPlanes(1), vBackCenter, vnegdir
D3DXPlaneFromPoints ClipPlanes(2), veye, vBackRightTop, vBackRightBot
D3DXPlaneFromPoints ClipPlanes(3), veye, vBackLeftTop, vBackLeftBot
D3DXPlaneFromPoints ClipPlanes(4), veye, vBackLeftTop, vBackRightTop
D3DXPlaneFromPoints ClipPlanes(5), veye, vBackRightBot, vBackLeftBot
numClipPlanes = 4
End Sub

Private Function vec3(X As Single, Y As Single, Z As Single) As D3DVECTOR
vec3.X = X
vec3.Y = Y
vec3.Z = Z
End Function

Private Function HeightField(Hill_Height As Single, X As Single, Y As Single) As Single
HeightField = Hill_Height * (Cos(X / 20 + 0.2) * Cos(Y / 15 - 0.2) + 1#)
End Function

Function FtoDW(f As Single) As Long
Dim buf As D3DXBuffer
Dim l As Long
Set buf = Eng.Direct3DX.CreateBuffer(4)
Eng.Direct3DX.BufferSetData buf, 0, 4, 1, f
Eng.Direct3DX.BufferGetData buf, 0, 4, 1, l
FtoDW = l
End Function

Private Sub MESHTOOL_INIT(Mt As MESHTOOL, m As D3DXMesh)
Set Mt.VertB = m.GetVertexBuffer
Mt.NumVertices = m.GetNumVertices
ReDim Mt.Vertices(Mt.NumVertices)
D3DVertexBuffer8GetData Mt.VertB, 0, Mt.NumVertices * Len(Mt.Vertices(0)), 0, Mt.Vertices(0)
End Sub
    
Private Sub MESHTOOL_DESTROY(Mt As MESHTOOL)
Set Mt.VertB = Nothing
ReDim Mt.Vertices(0)
End Sub
