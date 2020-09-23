VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "X3DXFiles"
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

Private Type CUSTOMVERTEX
 X As Single
 Y As Single
 Z As Single
 tU As Single
 tV As Single
End Type

Private Type tmesh
 Index As Integer
 Mesh As D3DXMesh
 PMesh As D3DXPMesh
 PMeshGenerated As Boolean
 Mesh_Name As String
 Position As D3DVECTOR
 SetPos As Boolean
 Rotation As D3DVECTOR
 SetRot As Boolean
 ScaleMesh As D3DVECTOR
 SetScale As Boolean
 Material As D3DMATERIAL8
 NewMaterial As Boolean
 MeshMaterials() As D3DMATERIAL8
 MeshTextures() As Direct3DTexture8
 NumMaterials As Long
 Texture As Direct3DTexture8
 NewText As Boolean
 MatAdjacency As D3DXBuffer
 Adjacency As D3DXBuffer
 Transparent As Boolean
 SphereCenter As D3DVECTOR
 SphereRadius As Single
 SpherePosition As D3DVECTOR
 HasSphere As Boolean
 Enhanced As Boolean
 EnhancedSegments As Single
 SetFilter As Boolean
 ComputedClipPlane As Boolean
 DarkMapAnimated As Boolean
 AnimationSpeed As Integer
 GlassEffect As Boolean
End Type: Private XMesh As tmesh

Private TextFolder As String
Private XFolder As String
Private Eng As New XEngine3D
Private Math As New XMath
Private mPos As D3DMATRIX
Private mRot As D3DMATRIX
Private mScale As D3DMATRIX
Private ObjectCount As Long
Private DoDark As Boolean
Private DoDetail As Boolean
Private DoLight As Boolean
Private DoGloss As Boolean
Private DoDMod As Boolean
Private DoDAdd As Boolean
Private DoASat As Boolean
Private DoLBlend As Boolean
Private DoILM As Boolean
Private ClipPlanes() As D3DPLANE
Private numClipPlanes As Long
Private MapTexture As Direct3DTexture8
Private Const D3DFVF_MIRRORVERTEX = (D3DFVF_XYZ Or D3DFVF_NORMAL Or D3DFVF_DIFFUSE)
Private Const D3DFVF_VERTEX = D3DFVF_XYZ Or D3DFVF_NORMAL Or D3DFVF_TEX1
Private Const D3DFVF_COLORVERTEX = (D3DFVF_XYZ Or D3DFVF_DIFFUSE Or D3DFVF_TEX1)

Public Property Get Get_Object_Name() As String
If XMesh.Mesh_Name = "" Then XMesh.Mesh_Name = "Default"
Get_Object_Name = XMesh.Mesh_Name
End Property

Public Sub Set_Object_Name(Name As String)
XMesh.Mesh_Name = Name
End Sub

Public Property Get Get_Mesh() As D3DXMesh
Set Get_Mesh = XMesh.Mesh
End Property

Public Property Get Get_Number_Of_Objects() As Long
Get_Number_Of_Objects = Len(ObjectCount)
End Property

Public Property Get Get_Object_Position_X() As Single
Get_Object_Position_X = XMesh.Position.X
End Property

Public Property Get Get_Object_Position_Y() As Single
Get_Object_Position_Y = XMesh.Position.Y
End Property

Public Property Get Get_Object_Position_Z() As Single
Get_Object_Position_Z = XMesh.Position.Z
End Property

Public Property Get Get_Object_PositionEX() As D3DVECTOR
Get_Object_PositionEX = XMesh.Position
End Property

Public Property Get Get_Object_Scale_X() As Single
Get_Object_Scale_X = XMesh.ScaleMesh.X
End Property

Public Property Get Get_Object_Scale_Y() As Single
Get_Object_Scale_Y = XMesh.ScaleMesh.Y
End Property

Public Property Get Get_Object_Scale_Z() As Single
Get_Object_Scale_Z = XMesh.ScaleMesh.Z
End Property

Public Property Get Get_Object_ScaleEX() As D3DVECTOR
Get_Object_ScaleEX = XMesh.ScaleMesh
End Property

Public Sub Position_Mesh(PosX As Single, PosY As Single, PosZ As Single)
XMesh.Position = Math.Vector3(PosX, PosY, PosZ)
XMesh.SetPos = True
End Sub

Public Property Get Get_Object_Rotation_X() As Single
Get_Object_Rotation_X = XMesh.Rotation.X
End Property

Public Property Get Get_Object_Rotation_Y() As Single
Get_Object_Rotation_Y = XMesh.Rotation.Y
End Property

Public Property Get Get_Object_Rotation_Z() As Single
Get_Object_Rotation_Z = XMesh.Rotation.Z
End Property

Public Property Get Get_Object_RotationEX() As D3DVECTOR
Get_Object_RotationEX = XMesh.Rotation
End Property

Public Sub Rotate_Mesh(X As Single, Y As Single, Z As Single)
On Local Error Resume Next
XMesh.Rotation = Math.Vector3(X, Y, Z)
XMesh.SetRot = True
End Sub

Public Sub Scale_Mesh(ScaleX As Single, ScaleY As Single, ScaleZ As Single)
On Local Error Resume Next
XMesh.ScaleMesh = Math.Vector3(ScaleX, ScaleY, ScaleZ)
XMesh.SetScale = True
End Sub

Public Function Get_Face_Count() As Long
On Local Error Resume Next
Get_Face_Count = XMesh.Mesh.GetNumFaces
End Function

Public Function Get_Vertices_Count() As Long
On Local Error Resume Next
Get_Vertices_Count = XMesh.Mesh.GetNumVertices
End Function

Public Property Get Get_Sphere_Center() As D3DVECTOR
If XMesh.HasSphere = False Then Exit Property
Get_Sphere_Center = XMesh.SphereCenter
End Property

Private Sub Convert_To_Progressive_Mesh()
If XMesh.Mesh Is Nothing Then Exit Sub
Dim tmpMesh As D3DXMesh
Set tmpMesh = XMesh.Mesh.CloneMeshFVF(D3DXMESH_SYSTEMMEM, D3DFVF_XYZ Or D3DFVF_NORMAL Or D3DFVF_TEX1, D3DD)
Compute_Normals tmpMesh
Set XMesh.Mesh = Nothing
Set XMesh.Mesh = tmpMesh
'Set XMesh.PMesh = Eng.Direct3DX.GeneratePMesh(XMesh.Mesh, XMesh.Adjacency.GetBufferPointer, Nothing, Nothing, 1, D3DXMESHSIMP_VERTEX)
'Set XMesh.Mesh = Nothing
'XMesh.PMesh.SetNumVertices XMesh.PMesh.GetNumVertices
'XMesh.PMeshGenerated = True
End Sub

Public Function Get_Sphere_Radius() As Single
Dim VB1 As Direct3DVertexBuffer8
Dim NumVertices1 As Long
Dim Vertices1() As CUSTOMVERTEX
Dim i As Long
Dim CX As Single
Dim CY As Single
Dim CZ As Single
Dim MX As Single
Dim MY As Single
Dim MZ As Single
Dim X As Single
Dim Y As Single
Dim Z As Single
Dim Center As D3DVECTOR
Dim Radius As Single
Dim Position As D3DVECTOR
Dim mat As D3DMATRIX

Set VB1 = XMesh.Mesh.GetVertexBuffer
NumVertices1 = XMesh.Mesh.GetNumVertices
ReDim Vertices1(NumVertices1)
D3DVertexBuffer8GetData VB1, 0, NumVertices1 * Len(Vertices1(0)), 0, Vertices1(0)

D3DD.GetTransform D3DTS_WORLD, mat

CX = 0: CY = 0: CZ = 0
For i = 0 To NumVertices1
 X = Vertices1(i).X
 Y = Vertices1(i).Y
 Z = Vertices1(i).Z
 
 Position.X = X: Position.Y = Y: Position.Z = Z
 
 CX = CX + X * mat.m11 + Y * mat.m21 + Z * mat.m31 + mat.m41
 CY = CY + X * mat.m12 + Y * mat.m22 + Z * mat.m32 + mat.m42
 CZ = CZ + X * mat.m13 + Y * mat.m23 + Z * mat.m33 + mat.m43
Next

XMesh.SpherePosition = Position
Center.X = CX / NumVertices1
Center.Y = CY / NumVertices1
Center.Z = CZ / NumVertices1

For i = 0 To NumVertices1
 X = Vertices1(i).X
 Y = Vertices1(i).Y
 Z = Vertices1(i).Z
 
 MX = (X * mat.m11 + Y * mat.m21 + Z * mat.m31 + mat.m41) - Center.X
 MY = (X * mat.m11 + Y * mat.m21 + Z * mat.m31 + mat.m41) - Center.Y
 MZ = (X * mat.m11 + Y * mat.m21 + Z * mat.m31 + mat.m41) - Center.Z

 Radius = Sqr(MX * MX + MY * MY + MZ * MZ)
Next
Get_Sphere_Radius = Radius
XMesh.HasSphere = True
XMesh.SphereCenter = Center
XMesh.SphereRadius = Radius
Set VB1 = Nothing
End Function

Private Function GetDist3D(Position2 As D3DVECTOR, Position1 As D3DVECTOR) As Single
Math.Get_Distance_3D Position1, Position2
End Function

Public Function Check_Camera_To_Object_Collision(CameraPos As D3DVECTOR, ObjectRadius As Single, Optional OffsetX = 0, Optional OffsetZ = 0) As Boolean
If CameraPos.X > (XMesh.Position.X + OffsetX) - ObjectRadius _
   And CameraPos.X < (XMesh.Position.X + OffsetX) + ObjectRadius _
   And CameraPos.Z > (XMesh.Position.Z + OffsetZ) - ObjectRadius _
   And CameraPos.Z < (XMesh.Position.Z + OffsetZ) + ObjectRadius Then
 Check_Camera_To_Object_Collision = True
Else
 Check_Camera_To_Object_Collision = False
End If
End Function

Public Sub Compute_Bounding_Box(MinExt As D3DVECTOR, MaxExt As D3DVECTOR)
Eng.Direct3DX.ComputeBoundingBoxFromMesh XMesh.Mesh, MinExt, MaxExt
End Sub

Public Sub Tesselate_Mesh(NumSegments As Single)
On Local Error Resume Next
Dim tmpMesh As D3DXMesh

If NumSegments <= 0 Then NumSegments = 0
If NumSegments >= 1.9 Then NumSegments = 1.9

Set tmpMesh = Nothing
Set tmpMesh = XMesh.Mesh
Set XMesh.Mesh = Nothing

Change_FVF tmpMesh, D3DFVF_VERTEX
Compute_Normals tmpMesh
Set tmpMesh = Eng.Direct3DX.TesselateMesh(XMesh.Mesh, XMesh.Adjacency.GetBufferPointer, NumSegments, False)

Set XMesh.Mesh = tmpMesh
Set tmpMesh = Nothing
XMesh.Enhanced = True
XMesh.EnhancedSegments = NumSegments
End Sub

Public Function Change_Default_Material(DiffuseRed As Single, DiffuseGreen As Single, DiffuseBlue As Single, Optional Alpha As Single = 1#, _
                                        Optional AmbientRed As Single = 0.5, Optional AmbientGreen As Single = 0.5, Optional AmbientBlue As Single = 0.5, _
                                        Optional SpecularRed As Single = 0, Optional SpecularGreen As Single = 0, Optional SpecularBlue As Single = 0, _
                                        Optional EmissiveRed As Single = 0, Optional EmissiveGreen As Single = 0, Optional EmissiveBlue As Single = 0, Optional Power As Single = 1)
Dim Mater As D3DMATERIAL8
Dim i As Long
On Local Error Resume Next
Mater.diffuse.r = DiffuseRed: Mater.diffuse.g = DiffuseGreen
Mater.diffuse.b = DiffuseBlue: Mater.diffuse.A = Alpha
Mater.Ambient.r = AmbientRed: Mater.Ambient.g = AmbientGreen
Mater.Ambient.b = AmbientBlue: Mater.Ambient.A = Alpha
Mater.Specular.r = SpecularRed: Mater.Specular.g = SpecularGreen
Mater.Specular.b = SpecularBlue: Mater.Specular.A = Alpha
Mater.emissive.r = EmissiveRed: Mater.emissive.g = EmissiveGreen
Mater.emissive.b = EmissiveBlue: Mater.emissive.A = Alpha
Mater.Power = Power
XMesh.Material = Mater
XMesh.NewMaterial = True
End Function

Public Sub Enable_Object_Transparency(Enabled As Boolean)
If Enabled = True Then
 XMesh.Transparent = True
Else
 XMesh.Transparent = False
End If
End Sub

Public Sub Enable_Glass_Effect(Enabled As Boolean)
If Enabled = True Then
 XMesh.GlassEffect = True
Else
 XMesh.GlassEffect = False
End If
End Sub

Public Sub Change_Default_Texture(Filename As String)
On Local Error Resume Next
Set XMesh.Texture = Eng.Direct3DX.CreateTextureFromFile(D3DD, Filename)
XMesh.NewText = True
End Sub

Public Sub Set_Texture_Directory_For_X_Files(Path As String)
 TextFolder = Path & "\"
End Sub

Public Sub Set_Directory_For_X_Files(Path As String)
 XFolder = Path & "\"
End Sub

Public Sub Apply_Object_Filter(Index As Integer, Filter_Type As FilterType, FilterMap_Filename As String)
XMesh.SetFilter = True
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

Public Sub Animate_Dark_Map(Animate As Boolean, Optional Animation_Speed As Integer = 1)
If Animate = True Then
 XMesh.DarkMapAnimated = True
 XMesh.AnimationSpeed = Animation_Speed
Else
 XMesh.DarkMapAnimated = False
End If
End Sub

Private Sub Start_Render_Filter(Index As Integer)
Static i As Integer
If XMesh.SetFilter = True Then
 If DoDark = True And XMesh.DarkMapAnimated = False Then
  D3DD.SetTextureStageState 0, D3DTSS_TEXCOORDINDEX, 0
  D3DD.SetTextureStageState 0, D3DTSS_COLORARG1, D3DTA_TEXTURE
  D3DD.SetTextureStageState 0, D3DTSS_ALPHAOP, D3DTOP_SELECTARG1
  D3DD.SetTexture 1, MapTexture
  D3DD.SetTextureStageState 1, D3DTSS_TEXCOORDINDEX, 0
  D3DD.SetTextureStageState 1, D3DTSS_COLORARG1, D3DTA_TEXTURE
  D3DD.SetTextureStageState 1, D3DTSS_COLORARG2, D3DTA_CURRENT
  D3DD.SetTextureStageState 1, D3DTSS_COLOROP, D3DTOP_MODULATE
 ElseIf DoDark = True And XMesh.DarkMapAnimated = True Then
  i = i + XMesh.AnimationSpeed
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

Public Sub Compute_Normals(Mesh As D3DXMesh)
On Local Error Resume Next
Dim bm As D3DXBaseMesh
Set bm = Mesh
Set Mesh = Nothing
Eng.Direct3DX.ComputeNormals bm
Set Mesh = bm
Set bm = Nothing
End Sub

Public Sub Flip_Mesh_Normals()
On Local Error Resume Next
Dim Count As Long
Dim Size As Long
Dim i As Long
Dim Verts() As D3DVERTEX
Dim VB As Direct3DVertexBuffer8

Set VB = XMesh.Mesh.GetVertexBuffer()
    
Size = Eng.Direct3DX.GetFVFVertexSize(XMesh.Mesh.GetFVF())
Count = XMesh.Mesh.GetNumVertices()

ReDim Verts(Count)
        
If XMesh.Mesh.GetFVF() = D3DFVF_VERTEX Then
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

Public Function Load_X_File(XFilename As String, Optional TransparentColor As Long = -1)
On Local Error Resume Next
Dim strTexName As String
Dim i As Long
Set XMesh.Mesh = Nothing
Set XMesh.Adjacency = Nothing
Set XMesh.Mesh = Eng.Direct3DX.LoadMeshFromX(XFolder & XFilename, D3DXMESH_MANAGED, _
                                             D3DD, XMesh.Adjacency, XMesh.MatAdjacency, XMesh.NumMaterials)
If XMesh.Mesh Is Nothing Then Exit Function

ReDim XMesh.MeshMaterials(XMesh.NumMaterials)
ReDim XMesh.MeshTextures(XMesh.NumMaterials)

For i = 0 To XMesh.NumMaterials - 1
 Eng.Direct3DX.BufferGetMaterial XMesh.MatAdjacency, i, XMesh.MeshMaterials(i)
 XMesh.MeshMaterials(i).Ambient = XMesh.MeshMaterials(i).diffuse
 strTexName = Eng.Direct3DX.BufferGetTextureName(XMesh.MatAdjacency, i)
 If strTexName <> "" Then
  If TransparentColor < 0 Then
   Set XMesh.MeshTextures(i) = Eng.Direct3DX.CreateTextureFromFile(D3DD, TextFolder & strTexName)
  Else
   Set XMesh.MeshTextures(i) = Eng.Direct3DX.CreateTextureFromFileEx(D3DD, TextFolder & strTexName, 0, 0, 0, 0, D3DFMT_A1R5G5B5, D3DPOOL_DEFAULT, D3DX_DEFAULT, D3DX_DEFAULT, &HFF000000, ByVal 0, ByVal 0)
   XMesh.Transparent = True
  End If
 End If
Next

XMesh.Position.X = 0
XMesh.Position.Y = 0
XMesh.Position.Z = 0
ObjectCount = ObjectCount + 1
End Function

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

Public Sub Render_X_Mesh(Camera As XCamera)
On Local Error Resume Next
Dim i As Long
Dim smat As D3DMATRIX
Dim smat2 As D3DMATRIX
Dim mat As D3DMATRIX
Dim mm As D3DMATRIX
Dim vec As D3DVECTOR
Dim q As D3DQUATERNION
Dim HoldRender1 As Long
Dim HoldRender2 As Long
If D3DD Is Nothing Then Exit Sub

Update_Clip_Planes Camera
If Sphere_Visible = False Then Exit Sub

D3DD.GetTransform D3DTS_WORLD, smat2

If XMesh.SetRot = True Then
 D3DD.GetTransform D3DTS_WORLD, smat
 mRot = smat
 D3DXMatrixRotationYawPitchRoll mRot, XMesh.Rotation.Y, XMesh.Rotation.X, XMesh.Rotation.Z
 D3DD.SetTransform D3DTS_WORLD, mRot
End If

If XMesh.SetPos = True Then
 D3DD.GetTransform D3DTS_WORLD, smat
 mPos = smat
 mPos.m11 = 1: mPos.m22 = 1: mPos.m33 = 1: mPos.m44 = 1
 mPos.m41 = XMesh.Position.X: mPos.m42 = XMesh.Position.Y: mPos.m43 = XMesh.Position.Z
 D3DD.SetTransform D3DTS_WORLD, mPos
 D3DXMatrixInverse mPos, 0, mPos
End If
    
If XMesh.SetScale = True Then
 D3DD.GetTransform D3DTS_WORLD, smat
 mScale = smat
 D3DXMatrixScaling mScale, XMesh.ScaleMesh.X, XMesh.ScaleMesh.Y, XMesh.ScaleMesh.Z
 D3DXMatrixMultiply smat, mScale, smat
 D3DD.SetTransform D3DTS_WORLD, smat
End If

If XMesh.Transparent = True Then
 D3DD.SetRenderState D3DRS_ALPHABLENDENABLE, 1
 D3DD.SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCCOLOR
 D3DD.SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCCOLOR
End If

If XMesh.GlassEffect = True Then
 HoldRender1 = D3DD.GetRenderState(D3DRS_CULLMODE)
 HoldRender2 = D3DD.GetRenderState(D3DRS_ZWRITEENABLE)
 D3DD.SetRenderState D3DRS_CULLMODE, D3DCULL_NONE
 D3DD.SetRenderState D3DRS_ZWRITEENABLE, False
End If

If XMesh.Enhanced = True Then: D3DD.SetRenderState D3DRS_PATCHSEGMENTS, XMesh.EnhancedSegments: D3DD.SetRenderState D3DRS_PATCHEDGESTYLE, 1

 If XMesh.NewMaterial = True Then
   For i = 0 To XMesh.NumMaterials - 1
    D3DD.SetMaterial XMesh.Material
    D3DD.SetTexture 0, XMesh.MeshTextures(i)
    If XMesh.PMeshGenerated = True Then
     XMesh.PMesh.DrawSubset i
    Else
     XMesh.Mesh.DrawSubset i
    End If
    D3DD.SetTexture 0, Nothing
   Next
 ElseIf XMesh.NewText = True Then
    For i = 0 To XMesh.NumMaterials - 1
    D3DD.SetMaterial XMesh.MeshMaterials(i)
    D3DD.SetTexture 0, XMesh.Texture
    If XMesh.PMeshGenerated = True Then
     XMesh.PMesh.DrawSubset i
    Else
     XMesh.Mesh.DrawSubset i
    End If
    D3DD.SetTexture 0, Nothing
   Next
 ElseIf XMesh.NewMaterial = True And XMesh.NewText = True Then
   For i = 0 To XMesh.NumMaterials - 1
    D3DD.SetMaterial XMesh.Material
    D3DD.SetTexture 0, XMesh.Texture
    If XMesh.PMeshGenerated = True Then
     XMesh.PMesh.DrawSubset i
    Else
     XMesh.Mesh.DrawSubset i
    End If
    D3DD.SetTexture 0, Nothing
   Next
 Else
   For i = 0 To XMesh.NumMaterials - 1
    D3DD.SetMaterial XMesh.MeshMaterials(i)
    D3DD.SetTexture 0, XMesh.MeshTextures(i)
    If XMesh.PMeshGenerated = True Then
     XMesh.PMesh.DrawSubset i
    Else
     XMesh.Mesh.DrawSubset i
    End If
    D3DD.SetTexture 0, Nothing
   Next
 End If

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
 For i = 0 To XMesh.NumMaterials - 1
  If XMesh.PMeshGenerated = True Then
   XMesh.PMesh.DrawSubset i
  Else
   XMesh.Mesh.DrawSubset i
  End If
 Next
 D3DD.SetRenderState D3DRS_DESTBLEND, D3DBLEND_ZERO
 D3DD.SetRenderState D3DRS_SRCBLEND, D3DBLEND_ZERO
 D3DD.SetRenderState D3DRS_ALPHABLENDENABLE, 0
End If

If XMesh.Enhanced = True Then: D3DD.SetRenderState D3DRS_PATCHSEGMENTS, 0

D3DD.SetTransform D3DTS_WORLD, smat2
If XMesh.Transparent = True Then
 D3DD.SetRenderState D3DRS_DESTBLEND, D3DBLEND_ZERO
 D3DD.SetRenderState D3DRS_SRCBLEND, D3DBLEND_ZERO
 D3DD.SetRenderState D3DRS_ALPHABLENDENABLE, 0
End If
If XMesh.GlassEffect = True Then
 D3DD.SetRenderState D3DRS_CULLMODE, HoldRender1
 D3DD.SetRenderState D3DRS_ZWRITEENABLE, HoldRender2
End If
End Sub

Private Function Sphere_Visible() As Boolean
Dim i As Long
Dim Distance As Single
    
For i = 0 To numClipPlanes - 1
 Distance = D3DXMATH_PLANE.D3DXPlaneDotCoord(ClipPlanes(i), XMesh.SphereCenter)
 If Distance < -1 * XMesh.SphereRadius Then
  Sphere_Visible = False
  Exit Function
 End If
Next

Sphere_Visible = True
End Function

Public Sub Compute_Clip_Planes(Compute As Boolean)
If Compute = True Then
 XMesh.ComputedClipPlane = True
 If XMesh.HasSphere = False Then
  Get_Sphere_Radius
 End If
Else
 XMesh.ComputedClipPlane = False
End If
End Sub

Private Sub Update_Clip_Planes(Camera As XCamera)
If XMesh.ComputedClipPlane = True And XMesh.HasSphere = True Then
 pComputeClipPlanes Camera.Get_Camera_OrientationEX, Camera.Get_Camera_PositionEX, Math.Vector3(0, Camera.Get_Camera_Eye_Level, 0), _
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

Public Sub Cleanup_Geometry_Engine()
On Local Error Resume Next
Set MapTexture = Nothing
ReDim XMesh.MeshTextures(0)
Set XMesh.MeshTextures(0) = Nothing
ReDim XMesh.MeshMaterials(0)
Set XMesh.PMesh = Nothing
Set XMesh.Mesh = Nothing
Set XMesh.MatAdjacency = Nothing
Set XMesh.Adjacency = Nothing
Set XMesh.Texture = Nothing
Set Math = Nothing
Set Eng = Nothing
End Sub
