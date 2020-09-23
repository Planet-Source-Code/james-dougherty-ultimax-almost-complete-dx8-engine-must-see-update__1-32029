VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "XEnviroment"
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



'BILLBOARD DOES NOT WORK RIGHT

Private Const TerrainSize = 25
Private Const TerrainColor = &HFFFFFFFF

Private Type TerrainVertex
 v As D3DVECTOR
 Color As Long
 tU As Single
 tV As Single
End Type

Private Type LandVertex
 Position As D3DVECTOR
 Color As Long
 tU As Single
 tV As Single
End Type

Private Type CUSTOMVERTEX
 X As Single
 Y As Single
 Z As Single
 tU As Single
 tV As Single
End Type
Private Const D3DFVF_COLORVERTEX = (D3DFVF_XYZ Or D3DFVF_DIFFUSE Or D3DFVF_TEX1)

Private Type tCloud
 Clouds As D3DXMesh
 Cloud_Texture As Direct3DTexture8
 Cloud_Height As D3DVECTOR
 CloudMat As D3DMATRIX
 Animate As Boolean
 RenderClouds As Boolean
End Type: Private Cloud As tCloud

Private Type tWater
 mWater As D3DXMesh
 Water_Texture As Direct3DTexture8
 Water_Height As D3DVECTOR
 WaterMat As D3DMATRIX
 Animate As Boolean
 RenderWater As Boolean
 Reflect As Boolean
 DuelTexture As Boolean
 Custom_Texture As Direct3DTexture8
 Use_Custom_Texture As Boolean
End Type: Private Water As tWater

Private Type tWater2
 WaterTexture As Direct3DTexture8
 NumWaterVertices As Long
 NumWaterIndices As Long
 WaterIB As Direct3DIndexBuffer8
 WaterVB As Direct3DVertexBuffer8
 PlaneWater As D3DPLANE
 grVerts() As TerrainVertex
 grVerts2() As TerrainVertex
 LHeight As Single
 SizeX As Single
 SizeY As Single
 FlatSizeX As Single
 FlatSizeY As Single
End Type: Private Water2 As tWater2

Private Type tTerrain
 GroundTexture As Direct3DTexture8
 GroundTexture2 As Direct3DTexture8
 NumGroundVertices As Long
 NumGroundIndices As Long
 GroundIB As Direct3DIndexBuffer8
 GroundVB As Direct3DVertexBuffer8
 PlaneGround As D3DPLANE
 grVerts() As TerrainVertex
 grVerts2() As TerrainVertex
 LHeight As Single
 SizeX As Single
 SizeY As Single
 FlatSizeX As Single
 FlatSizeY As Single
End Type: Private Terrain As tTerrain

Private Type tFlare
 Flare(0 To 3) As D3DXMesh
 Flare_Texture(0 To 3) As Direct3DTexture8
 Position(0 To 3) As D3DVECTOR
 Rotation(0 To 3) As D3DVECTOR
 RenderFlares As Boolean
End Type: Private Flare As tFlare

Private Type tSun
 Sun As D3DXMesh
 Sun_Texture As Direct3DTexture8
 Position As D3DVECTOR
 SetPos As Boolean
 Rotation As D3DVECTOR
 SunMatrix As D3DMATRIX
 SunPos As D3DMATRIX
 SunRot As D3DMATRIX
End Type: Private Sun As tSun

Private Type MESHTOOL
 VertB As Direct3DVertexBuffer8
 NumVertices As Long
 Vertices() As CUSTOMVERTEX
End Type

Private Eng As New XEngine3D
Private Cam As New XCamera
Private VertexBuffers(0 To 41) As Direct3DVertexBuffer8
Private SkySphere As D3DXMesh
Private sMeshMaterials() As D3DMATERIAL8
Private sMeshTextures() As Direct3DTexture8
Private lMeshMaterials() As D3DMATERIAL8
Private lMeshTextures() As Direct3DTexture8
Private lNumMaterials As Long
Private sNumMaterials As Long
Private TerrainTexture As Direct3DTexture8
Private TerrainTexture2 As Direct3DTexture8
Private ChangeTerrainTexture As Boolean
Private mRot As D3DMATRIX
Private mRot2 As D3DMATRIX
Private Plane As D3DPLANE
Private Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function GetCurrentTime Lib "kernel32" Alias "GetTickCount" () As Long

Public Sub Initialize_Enviroment_Sphere(XFile As String, TexturePath As String)
On Local Error Resume Next
Dim lmtrlBuffer As D3DXBuffer
Dim smtrlBuffer As D3DXBuffer
Dim strTexName As String
Dim i As Long

Set SkySphere = Eng.Direct3DX.LoadMeshFromX(XFile, D3DXMESH_MANAGED, _
                                            D3DD, Nothing, smtrlBuffer, sNumMaterials)
                                   
If SkySphere Is Nothing Then Exit Sub
ReDim sMeshMaterials(sNumMaterials)
ReDim sMeshTextures(sNumMaterials)

For i = 0 To sNumMaterials - 1
 Eng.Direct3DX.BufferGetMaterial smtrlBuffer, i, sMeshMaterials(i)
 sMeshMaterials(i).Ambient = sMeshMaterials(i).diffuse
 strTexName = Eng.Direct3DX.BufferGetTextureName(smtrlBuffer, i)
 If strTexName <> "" Then
  Set sMeshTextures(i) = Eng.Direct3DX.CreateTextureFromFile(D3DD, TexturePath & strTexName)
 End If
Next
    
DXUtil_Timer TIMER_start
Set smtrlBuffer = Nothing
End Sub

Public Sub Render_Enviroment_Sphere(Optional SphereHeight As Single = -50, Optional Animate_Sphere As Boolean = False, Optional Speed As Single = 5)
Dim i As Long
Dim matView As D3DMATRIX, matViewSave As D3DMATRIX
Dim mat As D3DMATRIX, matTextSave As D3DMATRIX
Dim smat As D3DMATRIX
Dim AnimaMat As D3DMATRIX
Dim smat2 As D3DMATRIX
Dim Time1 As Single
If D3DD Is Nothing Then Exit Sub

D3DD.GetTransform D3DTS_VIEW, matViewSave
matView = matViewSave
matView.m41 = 0: matView.m42 = SphereHeight: matView.m43 = 0
D3DD.SetTransform D3DTS_VIEW, matView
D3DD.SetRenderState D3DRS_ZENABLE, 0

D3DD.GetTransform D3DTS_WORLD, smat
mRot = smat
D3DXMatrixRotationYawPitchRoll mRot, 0, 1.5, 0
D3DD.SetTransform D3DTS_WORLD, mRot

If Animate_Sphere = True Then
 Time1 = DXUtil_Timer(TIMER_GETAPPTIME)
 D3DD.GetTransform D3DTS_WORLD, smat2
 AnimaMat = smat2
 D3DXMatrixRotationAxis AnimaMat, vec3(0, 0, 1), 3.14 + (Time1) / Speed
 D3DXMatrixMultiply smat2, AnimaMat, smat2
 D3DD.SetTransform D3DTS_WORLD, smat2
End If

For i = 0 To sNumMaterials - 1
    D3DD.SetMaterial sMeshMaterials(i)
    D3DD.SetTexture 0, sMeshTextures(i)
    SkySphere.DrawSubset i
    D3DD.SetTexture 0, Nothing
Next

D3DD.SetTransform D3DTS_VIEW, matViewSave
D3DD.SetTransform D3DTS_WORLD, smat
D3DD.SetRenderState D3DRS_ZENABLE, 1
End Sub

Public Sub Create_Clouds(Texture_Filename As String, Width As Single, Depth As Single, Cloud_Height As Single, Optional Texture_U_Power As Single = 1, Optional Texture_V_Power As Single = 1)
On Local Error Resume Next
Set Cloud.Clouds = Nothing
Set Cloud.Cloud_Texture = Eng.Direct3DX.CreateTextureFromFileEx(D3DD, Texture_Filename, D3DX_DEFAULT, _
                                                                D3DX_DEFAULT, D3DX_DEFAULT, 0, _
                                                                D3DFMT_UNKNOWN, D3DPOOL_MANAGED, _
                                                                D3DX_FILTER_POINT, D3DX_FILTER_POINT, _
                                                                &HFF000000, ByVal 0, ByVal 0)
Set Cloud.Clouds = pCreate_Water_Or_Clouds(Width, Depth, Texture_U_Power, Texture_V_Power)
Cloud.Cloud_Height = vec3(0, Cloud_Height, 0)
Cloud.RenderClouds = True
End Sub

Public Sub Render_Clouds()
Dim i As Long
Dim NumMat As Long
Dim Material As D3DMATERIAL8
Dim smat As D3DMATRIX

If Cloud.RenderClouds = True Then
 D3DD.SetRenderState D3DRS_ALPHABLENDENABLE, 1
 D3DD.SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCALPHA
 D3DD.SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA
 D3DD.GetTransform D3DTS_WORLD, smat
 Cloud.CloudMat = smat
 Cloud.CloudMat.m11 = 1
 Cloud.CloudMat.m22 = 1
 Cloud.CloudMat.m33 = 1
 Cloud.CloudMat.m44 = 1
 Cloud.CloudMat.m41 = Cloud.Cloud_Height.X
 Cloud.CloudMat.m42 = Cloud.Cloud_Height.Y
 Cloud.CloudMat.m43 = Cloud.Cloud_Height.Z
 D3DD.SetTransform D3DTS_WORLD, Cloud.CloudMat

 NumMat = 1
 Material.diffuse.r = 1
 Material.diffuse.g = 1
 Material.diffuse.b = 1
 Material.diffuse.A = 0.1
 Material.Ambient = Material.diffuse
 
 D3DD.SetMaterial Material
 D3DD.SetTexture 0, Cloud.Cloud_Texture
 D3DD.SetTextureStageState 0, D3DTSS_ALPHAARG1, D3DTA_TEXTURE
 D3DD.SetTextureStageState 0, D3DTSS_ALPHAOP, D3DTOP_SELECTARG1
 D3DD.SetTextureStageState 0, D3DTSS_ALPHAARG2, D3DTA_DIFFUSE
 
 For i = 0 To NumMat
  Cloud.Clouds.DrawSubset i
 Next
 
 D3DD.SetTexture 0, Nothing
 D3DD.SetTransform D3DTS_WORLD, smat
 D3DD.SetRenderState D3DRS_DESTBLEND, 0
 D3DD.SetRenderState D3DRS_SRCBLEND, 0
 D3DD.SetRenderState D3DRS_ALPHABLENDENABLE, 0
End If

End Sub

Public Sub Create_Water(Texture_Filename As String, Width As Single, Depth As Single, Water_Height As Single, Optional Enable_Duel_Texturing As Boolean = False, Optional Custom_2nd_Texture As String = "", Optional Texture_U_Power As Single = 1, Optional Texture_V_Power As Single = 1)
On Local Error Resume Next
Set Water.mWater = Nothing
Set Water.Water_Texture = Eng.Direct3DX.CreateTextureFromFileEx(D3DD, Texture_Filename, D3DX_DEFAULT, _
                                                                D3DX_DEFAULT, D3DX_DEFAULT, 0, _
                                                                D3DFMT_UNKNOWN, D3DPOOL_MANAGED, _
                                                                D3DX_FILTER_POINT, D3DX_FILTER_POINT, _
                                                                &HFF000000, ByVal 0, ByVal 0)
Set Water.mWater = pCreate_Water_Or_Clouds(Width, Depth, Texture_U_Power, Texture_V_Power)
Water.Water_Height = vec3(0, Water_Height, 0)

If Enable_Duel_Texturing = True Then: Water.DuelTexture = True

If Custom_2nd_Texture <> "" Then
 Water.Use_Custom_Texture = True
 Set Water.Custom_Texture = Eng.Direct3DX.CreateTextureFromFileEx(D3DD, Texture_Filename, D3DX_DEFAULT, _
                                                                  D3DX_DEFAULT, D3DX_DEFAULT, 0, _
                                                                  D3DFMT_UNKNOWN, D3DPOOL_MANAGED, _
                                                                  D3DX_FILTER_POINT, D3DX_FILTER_POINT, _
                                                                  &HFF000000, ByVal 0, ByVal 0)
End If

Plane = D3DPLANE4(0, 5, 0, 0)
Water.RenderWater = True
End Sub

Private Sub Enable_Water_Reflection(Enabled As Boolean)
If Enabled = True Then
 Water.Reflect = True
Else
 Water.Reflect = False
End If
End Sub

Public Sub Render_Water()
Dim i As Long
Dim NumMat As Long
Dim Refmat As D3DMATRIX
Dim Material As D3DMATERIAL8
Dim smat As D3DMATRIX
Dim smat2 As D3DMATRIX
Dim mat As D3DMATRIX
Dim ThisTime As Single
Dim Caps As D3DCAPS8
Dim Mater As D3DMATERIAL8

If Water.RenderWater = True Then
 ThisTime = DXUtil_Timer(TIMER_GETAPPTIME)
 
 D3DD.SetRenderState D3DRS_ALPHABLENDENABLE, 1
 D3DD.SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCALPHA
 D3DD.SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA
 D3DD.SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCCOLOR
 D3DD.SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCCOLOR

 D3DD.GetTransform D3DTS_WORLD, smat
 Water.WaterMat = smat
 Water.WaterMat.m11 = 1
 Water.WaterMat.m22 = 1
 Water.WaterMat.m33 = 1
 Water.WaterMat.m44 = 1
 Water.WaterMat.m41 = Water.Water_Height.X
 Water.WaterMat.m42 = Water.Water_Height.Y + Sin(ThisTime) * 0.2
 Water.WaterMat.m43 = Water.Water_Height.Z
 D3DD.SetTransform D3DTS_WORLD, Water.WaterMat

 NumMat = 1
 Mater.diffuse.r = 1: Mater.diffuse.g = 1: Mater.diffuse.b = 1: Mater.diffuse.A = 0.5
 Mater.Ambient = Mater.diffuse
 D3DD.SetMaterial Mater

 D3DD.SetTexture 0, Water.Water_Texture
 D3DD.SetTextureStageState 0, D3DTSS_ALPHAOP, D3DTOP_SELECTARG1
 D3DD.SetTextureStageState 0, D3DTSS_ALPHAARG1, D3DTA_TEXTURE
 
 If Water.DuelTexture = True Then
  If Water.Use_Custom_Texture = True Then
   D3DD.SetTexture 1, Water.Custom_Texture
  Else
   D3DD.SetTexture 1, Water.Water_Texture
  End If
  D3DD.SetTextureStageState 1, D3DTSS_COLOROP, D3DTOP_MODULATE
  D3DD.SetTextureStageState 1, D3DTSS_COLORARG1, D3DTA_TEXTURE
  D3DD.SetTextureStageState 1, D3DTSS_COLORARG2, D3DTA_CURRENT
  D3DD.SetTextureStageState 1, D3DTSS_ALPHAOP, D3DTOP_SELECTARG1
  D3DD.SetTextureStageState 1, D3DTSS_ALPHAARG1, D3DTA_TEXTURE
 
  D3DD.GetTransform D3DTS_TEXTURE1, smat2
  mat = smat2
  mat.m11 = 0.07
  mat.m32 = 0.07
  mat.m41 = Sin(ThisTime) / 8: mat.m42 = (Cos(ThisTime) / 10) - (ThisTime / 10)
  D3DD.SetTransform D3DTS_TEXTURE1, mat
  D3DD.SetTextureStageState 1, D3DTSS_TEXCOORDINDEX, D3DTSS_TCI_CAMERASPACEPOSITION
  D3DD.SetTextureStageState 1, D3DTSS_TEXTURETRANSFORMFLAGS, D3DTTFF_COUNT2
 End If

 For i = 0 To NumMat
  Water.mWater.DrawSubset i
 Next
 
 D3DD.SetTexture 0, Nothing
 D3DD.SetTexture 1, Nothing
 D3DD.SetTextureStageState 0, D3DTSS_ALPHAOP, D3DTOP_DISABLE
 D3DD.SetTextureStageState 1, D3DTSS_COLOROP, D3DTOP_DISABLE
 D3DD.SetTextureStageState 1, D3DTSS_ALPHAOP, D3DTOP_DISABLE
 
 If Water.DuelTexture = True Then
  D3DD.SetTransform D3DTS_TEXTURE1, smat2
  D3DD.SetTextureStageState 1, D3DTSS_TEXTURETRANSFORMFLAGS, D3DTEXF_NONE
  D3DD.SetTextureStageState 1, D3DTSS_TEXCOORDINDEX, D3DTEXF_NONE
 End If

 D3DD.SetRenderState D3DRS_SRCBLEND, 0
 D3DD.SetRenderState D3DRS_DESTBLEND, 0
 D3DD.SetRenderState D3DRS_ALPHABLENDENABLE, 0
 D3DD.SetTransform D3DTS_WORLD, smat
End If
End Sub

Private Sub Update_Ripples()
Dim WaterT As MESHTOOL
Dim i As Long
MESHTOOL_INIT WaterT, Water.mWater
For i = 0 To WaterT.NumVertices - 1
 'WaterT.Vertices(i).Y = HeightField(8, WaterT.Vertices(i).X, WaterT.Vertices(i).Z)
Next
MESHTOOL_DESTROY WaterT
End Sub

Private Sub Scale_Mesh(Mesh As D3DXMesh, Amount As Single)
Dim i As Long
Dim TextureTool As MESHTOOL
MESHTOOL_INIT TextureTool, Mesh

For i = 0 To TextureTool.NumVertices - 1
TextureTool.Vertices(i).X = TextureTool.Vertices(i).X * Amount
TextureTool.Vertices(i).Y = TextureTool.Vertices(i).Y + Rnd(1) + Rnd(1) + Rnd(1) * (15 / 2)
TextureTool.Vertices(i).Z = TextureTool.Vertices(i).Z * Amount
Next
    
D3DVertexBuffer8SetData TextureTool.VertB, 0, Len(TextureTool.Vertices(0)) * TextureTool.NumVertices, 0, TextureTool.Vertices(0)
MESHTOOL_DESTROY TextureTool
End Sub

Public Sub Create_Terrain(Texture_Filename As String, Optional TerrainWidth As Single = 500, Optional TerrainHeight As Single = 500, Optional Hilliness As Single = 7, Optional Texture_Tile_Size As Integer = 40)
On Local Error GoTo errOut

ErrStr = "Create_Terrain"
pCreate_Terrain Texture_Filename, TerrainWidth, TerrainHeight, Hilliness, Texture_Tile_Size
Terrain.LHeight = Hilliness
Terrain.SizeX = TerrainWidth
Terrain.SizeY = TerrainHeight
Set Terrain.GroundTexture = Nothing
If Texture_Filename <> "" Then
 ErrStr = "Set Terrain.GroundTexture"
 Set Terrain.GroundTexture = Eng.Direct3DX.CreateTextureFromFile(D3DD, Texture_Filename)
End If
Exit Sub

errOut:
 ErrorToFile "ErrorTerrain", ErrStr, PossibleCause
End Sub

Public Sub Create_Flat_Terrain(Texture_Filename As String, Optional TerrainWidth As Single = 500, Optional TerrainHeight As Single = 500, Optional Texture_Tile_Size As Integer = 40)
On Local Error GoTo errOut

ErrStr = "Create_Terrain"
pCreate_Terrain Texture_Filename, TerrainWidth, TerrainHeight, 0, Texture_Tile_Size
Terrain.FlatSizeX = TerrainWidth
Terrain.FlatSizeY = TerrainHeight
Set Terrain.GroundTexture = Nothing
If Texture_Filename <> "" Then
 ErrStr = "Set Terrain.GroundTexture"
 Set Terrain.GroundTexture = Eng.Direct3DX.CreateTextureFromFile(D3DD, Texture_Filename)
End If
Exit Sub

errOut:
 ErrorToFile "ErrorTerrain", ErrStr, PossibleCause
End Sub

Private Function Create_Custom_Landscape(Heightmap As String, Texture_Filename As String, Optional Texture_Tile_Size As Integer = 40) As Boolean
'On Error GoTo errOut
Dim i As Long
Dim v As TerrainVertex
Dim vtx As Long
Dim pHeight As Single
Dim indices() As Integer
Dim Z1 As Long, X1 As Long
Dim ZZ As Long, XX As Long
Dim Map As Direct3DTexture8
Dim DESC As D3DSURFACE_DESC
Dim LRect As D3DLOCKED_RECT
Dim X As Long
Dim Y As Long

If Heightmap <> "" Then
 ErrStr = "Set Map"
 Set Map = Eng.Direct3DX.CreateTextureFromFile(D3DD, Heightmap)
End If
If Texture_Filename <> "" Then
 ErrStr = "Set Terrain.GroundTexture"
 Set Terrain.GroundTexture = Eng.Direct3DX.CreateTextureFromFile(D3DD, Texture_Filename)
End If

Terrain.NumGroundVertices = (TerrainSize + 1) * (TerrainSize + 1)
Terrain.NumGroundIndices = (TerrainSize * TerrainSize) * 48

Set Terrain.GroundVB = Nothing
Set Terrain.GroundIB = Nothing
Terrain.PlaneGround = D3DPLANE4(0, 1, 0, 0)
Set Terrain.GroundVB = D3DD.CreateVertexBuffer(Terrain.NumGroundVertices * Len(v), _
                       0, D3DFVF_COLORVERTEX, D3DPOOL_MANAGED)
ReDim Terrain.grVerts(TerrainSize * TerrainSize * TerrainSize * 2)

Map.GetLevelDesc 0, DESC
Map.LockRect 0, LRect, 0, 0

i = 0
XX = 0
ZZ = 0
For X = 0 To DESC.Width
 For Y = 0 To DESC.Height
  
  If LRect.pBits = &H0& Then: pHeight = 0
  If LRect.pBits = &H808080 Then: pHeight = 2
  If LRect.pBits = &HC0C0C0 Then: pHeight = 4
  If LRect.pBits = &HE0E0E0 Then: pHeight = 6
  If LRect.pBits = &HFFFFFF Then: pHeight = 8
 
 Next Y
Next X

For ZZ = 0 To TerrainSize
   For XX = 0 To TerrainSize
    Terrain.grVerts(i).v.X = DESC.Width * ((XX / TerrainSize) - 0.5)
    Terrain.grVerts(i).v.Z = DESC.Height * ((ZZ / TerrainSize) - 0.5)
    Terrain.grVerts(i).v.Y = HeightField(7 / 4, CSng(DESC.Width) + Terrain.grVerts(i).v.X, CSng(DESC.Height) + Terrain.grVerts(i).v.Z)
    Terrain.grVerts(i).Color = TerrainColor
    Terrain.grVerts(i).tU = XX * (Texture_Tile_Size / TerrainSize)
    Terrain.grVerts(i).tV = ZZ * (Texture_Tile_Size / TerrainSize)
    i = i + 1
   Next
  Next
  
D3DVertexBuffer8SetData Terrain.GroundVB, 0, Len(v) * (TerrainSize + 1) * (TerrainSize + 1), 0, Terrain.grVerts(0)
        
ReDim indices(Terrain.NumGroundIndices * 4)
Set Terrain.GroundIB = D3DD.CreateIndexBuffer(Terrain.NumGroundIndices * 2, _
                       0, D3DFMT_INDEX16, D3DPOOL_MANAGED)
                       
i = 0
For Z1 = 0 To TerrainSize - 1
 For X1 = 0 To TerrainSize - 1
  vtx = X1 + Z1 * (TerrainSize + 1)
  indices(i) = vtx + 1: i = i + 1
  indices(i) = vtx + 0: i = i + 1
  indices(i) = vtx + 0 + (TerrainSize + 1): i = i + 1
  indices(i) = vtx + 1: i = i + 1
  indices(i) = vtx + 0 + (TerrainSize + 1): i = i + 1
  indices(i) = vtx + 1 + (TerrainSize + 1): i = i + 1
 Next
Next
D3DIndexBuffer8SetData Terrain.GroundIB, 0, 2 * Terrain.NumGroundIndices, 0, indices(0)

Map.UnlockRect 0
Create_Custom_Landscape = True
Set Map = Nothing
Exit Function
errOut:
 Create_Custom_Landscape = False
End Function

Public Function Get_Terrain_Height(X As Single, Z As Single) As Single
Get_Terrain_Height = HeightField(Terrain.LHeight, X, Z)
End Function

Private Function pCreate_Terrain(Texture_Filename As String, Optional TerrainWidth As Single = 500, Optional TerrainHeight As Single = 500, Optional Hilliness As Single = 7, Optional Texture_Tile_Size As Integer = 40) As Boolean
On Local Error GoTo errOut
Dim i As Long
Dim v As TerrainVertex
Dim vtx As Long
Dim indices() As Integer
Dim Z As Long, X As Long
Dim ZZ As Long, XX As Long
    
Terrain.NumGroundVertices = (TerrainSize + 1) * (TerrainSize + 1)
Terrain.NumGroundIndices = (TerrainSize * TerrainSize) * 48

Set Terrain.GroundVB = Nothing
Set Terrain.GroundIB = Nothing
Terrain.PlaneGround = D3DPLANE4(0, 1, 0, 0)
Set Terrain.GroundVB = D3DD.CreateVertexBuffer(Terrain.NumGroundVertices * Len(v), _
                                               0, D3DFVF_COLORVERTEX, D3DPOOL_MANAGED)
ReDim Terrain.grVerts(TerrainSize * TerrainSize * TerrainSize * 2.5)
i = 0
For ZZ = 0 To TerrainSize
 For XX = 0 To TerrainSize
  Terrain.grVerts(i).v.X = TerrainWidth * ((XX / TerrainSize) - 0.5)
  Terrain.grVerts(i).v.Z = TerrainHeight * ((ZZ / TerrainSize) - 0.5)
  Terrain.grVerts(i).v.Y = HeightField(Hilliness, Terrain.grVerts(i).v.X, Terrain.grVerts(i).v.Z)
  Terrain.grVerts(i).Color = TerrainColor
  Terrain.grVerts(i).tU = XX * (Texture_Tile_Size / TerrainSize)
  Terrain.grVerts(i).tV = ZZ * (Texture_Tile_Size / TerrainSize)
  i = i + 1
 Next
Next
D3DVertexBuffer8SetData Terrain.GroundVB, 0, Len(v) * (TerrainSize + 1) * (TerrainSize + 1), 0, Terrain.grVerts(0)
        
ReDim indices(Terrain.NumGroundIndices * 4)
Set Terrain.GroundIB = D3DD.CreateIndexBuffer(Terrain.NumGroundIndices * 2, _
                       0, D3DFMT_INDEX16, D3DPOOL_MANAGED)
i = 0
For Z = 0 To TerrainSize - 1
 For X = 0 To TerrainSize - 1
  vtx = X + Z * (TerrainSize + 1)
  indices(i) = vtx + 1: i = i + 1
  indices(i) = vtx + 0: i = i + 1
  indices(i) = vtx + 0 + (TerrainSize + 1): i = i + 1
  indices(i) = vtx + 1: i = i + 1
  indices(i) = vtx + 0 + (TerrainSize + 1): i = i + 1
  indices(i) = vtx + 1 + (TerrainSize + 1): i = i + 1
 Next
Next
D3DIndexBuffer8SetData Terrain.GroundIB, 0, 2 * Terrain.NumGroundIndices, 0, indices(0)

pCreate_Terrain = True
Exit Function

errOut:
 pCreate_Terrain = False
End Function

Public Sub Render_Terrain()
Dim i As Long
Dim v As TerrainVertex
Dim SecondTexture As String

D3DD.SetTexture 0, Terrain.GroundTexture
D3DD.SetTextureStageState 0, D3DTSS_TEXCOORDINDEX, 0
D3DD.SetTextureStageState 0, D3DTSS_COLORARG1, D3DTA_TEXTURE
D3DD.SetTextureStageState 0, D3DTSS_ALPHAOP, D3DTOP_SELECTARG1

If SecondTexture <> "" Then
 Set Terrain.GroundTexture2 = Eng.Direct3DX.CreateTextureFromFile(D3DD, SecondTexture)
 D3DD.SetTexture 1, Terrain.GroundTexture2
 D3DD.SetTextureStageState 1, D3DTSS_COLORARG1, D3DTA_TEXTURE
 'D3DD.SetTextureStageState 1, D3DTSS_COLORARG2, D3DTA_CURRENT
 D3DD.SetTextureStageState 1, D3DTSS_COLOROP, D3DTA_CURRENT
End If

D3DD.SetVertexShader D3DFVF_COLORVERTEX
D3DD.SetStreamSource 0, Terrain.GroundVB, Len(v)
D3DD.SetIndices Terrain.GroundIB, 0
D3DD.DrawIndexedPrimitive D3DPT_TRIANGLELIST, _
                          0, Terrain.NumGroundVertices, _
                          0, (Terrain.NumGroundIndices / 3)
End Sub

Public Sub Cleanup_Enviroment_Engine()
Dim i As Long
Set Sun.Sun = Nothing
Set Sun.Sun_Texture = Nothing
For i = 0 To 3
 Set Flare.Flare(i) = Nothing
 Set Flare.Flare_Texture(i) = Nothing
Next
Set Terrain.GroundTexture = Nothing
Set Terrain.GroundTexture2 = Nothing
Set Terrain.GroundVB = Nothing
Set Terrain.GroundIB = Nothing
Set Water2.WaterTexture = Nothing
Set Water2.WaterVB = Nothing
Set Water2.WaterIB = Nothing
ReDim sMeshMaterials(0)
ReDim lMeshMaterials(0)
ReDim sMeshTextures(0)
Set sMeshTextures(0) = Nothing
ReDim lMeshTextures(0)
Set lMeshTextures(0) = Nothing
Set Water.Water_Texture = Nothing
Set Water.Custom_Texture = Nothing
Set Cloud.Cloud_Texture = Nothing
Set TerrainTexture2 = Nothing
Set TerrainTexture = Nothing
Set SkySphere = Nothing
Set Cam = Nothing
Set Eng = Nothing
End Sub

Public Sub Enable_Fog(FogEnabled As Boolean, Optional FogStart As Single = 1, Optional FogEnd As Single = 100, Optional Color As Long = &HE0E0E0)
If FogEnabled Then
 D3DD.SetRenderState D3DRS_FOGENABLE, 1
 D3DD.SetRenderState D3DRS_FOGCOLOR, Color
 D3DD.SetRenderState D3DRS_FOGVERTEXMODE, D3DFOG_LINEAR
 D3DD.SetRenderState D3DRS_RANGEFOGENABLE, 1
 D3DD.SetRenderState D3DRS_FOGSTART, FtoDW(FogStart)
 D3DD.SetRenderState D3DRS_FOGEND, FtoDW(FogEnd)
Else
 D3DD.SetRenderState D3DRS_FOGENABLE, 0
 D3DD.SetRenderState D3DRS_RANGEFOGENABLE, 0
 D3DD.SetRenderState D3DRS_FOGSTART, 0
 D3DD.SetRenderState D3DRS_FOGEND, 0
End If
End Sub

Private Function HeightField(Hill_Height As Single, X As Single, Z As Single) As Single
HeightField = Hill_Height * (Cos(X / 20 + 0.8) * Cos(Z / 15 - 0.8) + 1#)
End Function

Private Function pCreate_Water_Or_Clouds(Width As Single, Depth As Single, Optional tU As Single = 1, Optional tV As Single = 1) As D3DXMesh
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
With Verts(0): .X = -W: .Y = 0: .Z = -H2: .nz = 1: .tU = 0: .tV = tV: End With
With Verts(1): .X = W: .Y = 0: .Z = -H2: .nz = 1: .tU = tU: .tV = tV: End With
With Verts(2): .X = W: .Y = 0: .Z = -H: .nz = 1: .tU = tU: .tV = 0: End With
With Verts(3): .X = -W: .Y = 0: .Z = -H: .nz = 1: .tU = 0: .tV = 0: End With
indices(0) = 0: indices(1) = 1: indices(2) = 2
indices(3) = 0: indices(4) = 2: indices(5) = 3
With Verts(4): .X = -W: .Y = 0: .Z = H: .nz = -1: .tU = 0: .tV = 0: End With
With Verts(5): .X = W: .Y = 0: .Z = H: .nz = -1: .tU = tU: .tV = 0: End With
With Verts(6): .X = W: .Y = 0: .Z = H2: .nz = -1: .tU = tU: .tV = tV: End With
With Verts(7): .X = -W: .Y = 0: .Z = H2: .nz = -1: .tU = 0: .tV = tV: End With
indices(6) = 4: indices(7) = 5: indices(8) = 6
indices(9) = 4: indices(10) = 6: indices(11) = 7
D3DXMeshVertexBuffer8SetData rMesh, 0, Len(Verts(0)) * 8, 0, Verts(0)
D3DXMeshIndexBuffer8SetData rMesh, 0, Len(indices(0)) * 12, 0, indices(0)
        
Set pCreate_Water_Or_Clouds = rMesh
End Function

Public Sub Create_Sun(Sun_Texture_Filename As String, Sun_Height As Single, Sun_Width As Single)
Set Sun.Sun = Nothing
Set Sun.Sun = pCreateSun(Sun_Width, Sun_Height)
If Sun_Texture_Filename <> "" Then
 Set Sun.Sun_Texture = Eng.Direct3DX.CreateTextureFromFileEx(D3DD, Sun_Texture_Filename, D3DX_DEFAULT, _
                                                                D3DX_DEFAULT, D3DX_DEFAULT, 0, _
                                                                D3DFMT_UNKNOWN, D3DPOOL_MANAGED, _
                                                                D3DX_FILTER_POINT, D3DX_FILTER_POINT, _
                                                                &HFF000000, ByVal 0, ByVal 0)
End If
End Sub

Public Sub Set_Sun_Position(X As Single, Y As Single, Z As Single)
Sun.Position = vec3(X, Y, Z)
Sun.SetPos = True
End Sub

Public Sub Render_Sun(Camera As XCamera)
On Local Error Resume Next
Dim i As Long
Dim NumMat As Long
Dim Material As D3DMATERIAL8
Dim Caps As D3DCAPS8
Dim smat As D3DMATRIX
Dim smat2 As D3DMATRIX
Dim CamOX As Single
Dim CamOZ As Single

CamOX = Camera.Get_Camera_Orientation_X
CamOZ = Camera.Get_Camera_Orientation_Z

D3DD.GetTransform D3DTS_WORLD, smat2
D3DD.GetTransform D3DTS_WORLD, smat
Sun.SunRot = smat
D3DXMatrixRotationYawPitchRoll Sun.SunRot, -(Camera.Get_Camera_Orientation_Y * (CamOX / CamOZ) - 140), 0, 0
Sun.Rotation = vec3(0, -(Camera.Get_Camera_Orientation_Y * (CamOX / CamOZ) - 140), 0)
D3DD.SetTransform D3DTS_WORLD, Sun.SunRot

If Sun.SetPos = True Then
 D3DD.GetTransform D3DTS_WORLD, smat
 Sun.SunMatrix = smat
 Sun.SunMatrix.m11 = 1: Sun.SunMatrix.m22 = 1: Sun.SunMatrix.m33 = 1: Sun.SunMatrix.m44 = 1
 Sun.SunMatrix.m41 = Sun.Position.X
 Sun.SunMatrix.m42 = Sun.Position.Y
 Sun.SunMatrix.m43 = Sun.Position.Z
 D3DD.SetTransform D3DTS_WORLD, Sun.SunMatrix
End If

D3DD.SetRenderState D3DRS_ALPHABLENDENABLE, 1
D3DD.SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCALPHA
D3DD.SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA
D3DD.SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCCOLOR
D3DD.SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCCOLOR

Material.diffuse.r = 1: Material.diffuse.g = 1: Material.diffuse.b = 1: Material.diffuse.A = 0.5
Material.Ambient = Material.diffuse
NumMat = 1

D3DD.SetMaterial Material
D3DD.SetTexture 0, Sun.Sun_Texture
For i = 0 To NumMat
 Sun.Sun.DrawSubset i
Next

D3DD.SetRenderState D3DRS_DESTBLEND, 0
D3DD.SetRenderState D3DRS_SRCBLEND, 0
D3DD.SetRenderState D3DRS_ALPHABLENDENABLE, 0
D3DD.SetTransform D3DTS_WORLD, smat2
End Sub

Private Sub Create_Lens_Flare1(Flare_Texture1_Filename As String, Flare1_Height As Single, Flare1_Width As Single)
Set Flare.Flare(0) = Nothing
Set Flare.Flare(0) = pCreateFlare(Flare1_Width, Flare1_Height)
If Flare_Texture1_Filename <> "" Then
 Set Flare.Flare_Texture(0) = Eng.Direct3DX.CreateTextureFromFileEx(D3DD, Flare_Texture1_Filename, D3DX_DEFAULT, _
                                                                D3DX_DEFAULT, D3DX_DEFAULT, 0, _
                                                                D3DFMT_UNKNOWN, D3DPOOL_MANAGED, _
                                                                D3DX_FILTER_POINT, D3DX_FILTER_POINT, _
                                                                &HFF000000, ByVal 0, ByVal 0)
End If
End Sub

Private Sub Create_Lens_Flare2(Flare_Texture2_Filename As String, Flare2_Height As Single, Flare2_Width As Single)
Set Flare.Flare(1) = Nothing
Set Flare.Flare(1) = pCreateFlare(Flare2_Width, Flare2_Height)
If Flare_Texture2_Filename <> "" Then
 Set Flare.Flare_Texture(1) = Eng.Direct3DX.CreateTextureFromFileEx(D3DD, Flare_Texture2_Filename, D3DX_DEFAULT, _
                                                                D3DX_DEFAULT, D3DX_DEFAULT, 0, _
                                                                D3DFMT_UNKNOWN, D3DPOOL_MANAGED, _
                                                                D3DX_FILTER_POINT, D3DX_FILTER_POINT, _
                                                                &HFF000000, ByVal 0, ByVal 0)
End If
End Sub

Private Sub Create_Lens_Flare3(Flare_Texture3_Filename As String, Flare3_Height As Single, Flare3_Width As Single)
Set Flare.Flare(2) = Nothing
Set Flare.Flare(2) = pCreateFlare(Flare3_Width, Flare3_Height)
If Flare_Texture3_Filename <> "" Then
 Set Flare.Flare_Texture(2) = Eng.Direct3DX.CreateTextureFromFileEx(D3DD, Flare_Texture3_Filename, D3DX_DEFAULT, _
                                                                D3DX_DEFAULT, D3DX_DEFAULT, 0, _
                                                                D3DFMT_UNKNOWN, D3DPOOL_MANAGED, _
                                                                D3DX_FILTER_POINT, D3DX_FILTER_POINT, _
                                                                &HFF000000, ByVal 0, ByVal 0)
End If
End Sub

Private Sub Create_Lens_Flare4(Flare_Texture4_Filename As String, Flare4_Height As Single, Flare4_Width As Single)
Set Flare.Flare(3) = Nothing
Set Flare.Flare(3) = pCreateFlare(Flare4_Width, Flare4_Height)
If Flare_Texture4_Filename <> "" Then
 Set Flare.Flare_Texture(3) = Eng.Direct3DX.CreateTextureFromFileEx(D3DD, Flare_Texture4_Filename, D3DX_DEFAULT, _
                                                                D3DX_DEFAULT, D3DX_DEFAULT, 0, _
                                                                D3DFMT_UNKNOWN, D3DPOOL_MANAGED, _
                                                                D3DX_FILTER_POINT, D3DX_FILTER_POINT, _
                                                                &HFF000000, ByVal 0, ByVal 0)
End If
End Sub

Private Function pCreateFlare(Width As Single, Height As Single) As D3DXMesh
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
With Verts(0): .X = -H: .Y = W2: .Z = -0: .nz = 1: .tU = 0: .tV = 1: End With
With Verts(1): .X = H: .Y = W2: .Z = -0: .nz = 1: .tU = 1: .tV = 1: End With
With Verts(2): .X = H: .Y = W: .Z = -0: .nz = 1: .tU = 1: .tV = 0: End With
With Verts(3): .X = -H: .Y = W: .Z = -0: .nz = 1: .tU = 0: .tV = 0: End With
indices(0) = 0: indices(1) = 1: indices(2) = 2
indices(3) = 0: indices(4) = 2: indices(5) = 3
With Verts(4): .X = -H: .Y = W: .Z = 0: .nz = -1: .tU = 0: .tV = 0: End With
With Verts(5): .X = H: .Y = W: .Z = 0: .nz = -1: .tU = 1: .tV = 0: End With
With Verts(6): .X = H: .Y = W2: .Z = 0: .nz = -1: .tU = 1: .tV = 1: End With
With Verts(7): .X = -H: .Y = W2: .Z = 0: .nz = -1: .tU = 0: .tV = 1: End With
indices(6) = 4: indices(7) = 5: indices(8) = 6
indices(9) = 4: indices(10) = 6: indices(11) = 7
D3DXMeshVertexBuffer8SetData rMesh, 0, Len(Verts(0)) * 8, 0, Verts(0)
D3DXMeshIndexBuffer8SetData rMesh, 0, Len(indices(0)) * 12, 0, indices(0)
        
Set pCreateFlare = rMesh
End Function

Private Function pCreateSun(Width As Single, Height As Single) As D3DXMesh
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
With Verts(0): .X = -H: .Y = W2: .Z = -0: .nz = 1: .tU = 0: .tV = 1: End With
With Verts(1): .X = H: .Y = W2: .Z = -0: .nz = 1: .tU = 1: .tV = 1: End With
With Verts(2): .X = H: .Y = W: .Z = -0: .nz = 1: .tU = 1: .tV = 0: End With
With Verts(3): .X = -H: .Y = W: .Z = -0: .nz = 1: .tU = 0: .tV = 0: End With
indices(0) = 0: indices(1) = 1: indices(2) = 2
indices(3) = 0: indices(4) = 2: indices(5) = 3
With Verts(4): .X = -H: .Y = W: .Z = 0: .nz = -1: .tU = 0: .tV = 0: End With
With Verts(5): .X = H: .Y = W: .Z = 0: .nz = -1: .tU = 1: .tV = 0: End With
With Verts(6): .X = H: .Y = W2: .Z = 0: .nz = -1: .tU = 1: .tV = 1: End With
With Verts(7): .X = -H: .Y = W2: .Z = 0: .nz = -1: .tU = 0: .tV = 1: End With
indices(6) = 4: indices(7) = 5: indices(8) = 6
indices(9) = 4: indices(10) = 6: indices(11) = 7
D3DXMeshVertexBuffer8SetData rMesh, 0, Len(Verts(0)) * 8, 0, Verts(0)
D3DXMeshIndexBuffer8SetData rMesh, 0, Len(indices(0)) * 12, 0, indices(0)
        
Set pCreateSun = rMesh
End Function

Private Function vec2(X As Single, Y As Single) As D3DVECTOR2
vec2.X = X
vec2.Y = Y
End Function

Private Function vec3(X As Single, Y As Single, Z As Single) As D3DVECTOR
vec3.X = X
vec3.Y = Y
vec3.Z = Z
End Function

Private Function FtoDW(f As Single) As Long
Dim buf As D3DXBuffer
Dim l As Long
Set buf = Eng.Direct3DX.CreateBuffer(4)
Eng.Direct3DX.BufferSetData buf, 0, 4, 1, f
Eng.Direct3DX.BufferGetData buf, 0, 4, 1, l
FtoDW = l
End Function

Private Function LONGtoD3DCOLORVALUE(Color As Long) As D3DCOLORVALUE
Dim A As Long, r As Long, g As Long, b As Long
        
If Color < 0 Then
 A = ((Color And (&H7F000000)) / (2 ^ 24)) Or &H80&
Else
 A = Color / (2 ^ 24)
End If
r = (Color And &HFF0000) / (2 ^ 16)
g = (Color And &HFF00&) / (2 ^ 8)
b = (Color And &HFF&)
    
LONGtoD3DCOLORVALUE.A = A / 255
LONGtoD3DCOLORVALUE.r = r / 255
LONGtoD3DCOLORVALUE.g = g / 255
LONGtoD3DCOLORVALUE.b = b / 255
End Function

Private Function D3DPLANE4(A As Single, b As Single, C As Single, D As Single) As D3DPLANE
D3DPLANE4.A = A
D3DPLANE4.b = b
D3DPLANE4.C = C
D3DPLANE4.D = D
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

Private Function DXUtil_Timer(command As TIMER_COMMAND) As Single

    On Local Error Resume Next

    Static m_bTimerInitialized  As Boolean
    Static m_bUsingQPF         As Boolean
    Static m_fSecsPerTick  As Single
    Static m_fBaseTime    As Single
    Static m_fStopTime     As Single
    Static m_fLastTime As Single
    
    Dim fTime As Single

    ' Initialize the timer
    If (False = m_bTimerInitialized) Then
        m_bTimerInitialized = True
    End If

    fTime = Timer

    ' Reset the timer
    If (command = TIMER_RESET) Then
        m_fBaseTime = fTime
        m_fStopTime = 0
        m_fLastTime = 0
        DXUtil_Timer = 0
        Exit Function
    End If

    ' Return the current time
    If (command = TIMER_GETAPPTIME) Then
        DXUtil_Timer = fTime - m_fBaseTime
        Exit Function
    End If

    ' Start the timer
    If (command = TIMER_start) Then
        m_fBaseTime = m_fBaseTime + fTime - m_fStopTime
        m_fLastTime = m_fLastTime + fTime - m_fStopTime
        m_fStopTime = 0
    End If
    
    ' Stop the timer
    If (command = TIMER_STOP) Then
        m_fStopTime = fTime
    End If
     
    ' Advance the timer by 1/10th second
    If (command = TIMER_ADVANCE) Then
        m_fBaseTime = m_fBaseTime + fTime - (m_fStopTime + 0.1)
    End If
    
    
    ' Return ellapsed time
    If (command = TIMER_GETELLAPSEDTIME) Then
        DXUtil_Timer = fTime - m_fLastTime
        m_fLastTime = fTime
        If DXUtil_Timer < 0 Then DXUtil_Timer = 0
        Exit Function
    End If
    
    
    DXUtil_Timer = fTime
End Function

