VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "XDecal"
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

Private Type tDecal
Decal As D3DXMesh
Position As D3DVECTOR
SetPos As Boolean
Rotation As D3DVECTOR
Rotate As Boolean
Transparent As Boolean
DecalSurface As Direct3DTexture8
End Type: Private Decal As tDecal

Private Eng As New XEngine3D
Private Math As New XMath
Private HasSurface As Boolean
Private mPos As D3DMATRIX
Private mRot As D3DMATRIX
Private Declare Function GetClientRect Lib "user32.dll" (ByVal hWnd As Long, ByRef RECT As RECT) As Long

Public Property Get Get_Decal_Position_X() As Single
 Get_Decal_Position_X = Decal.Position.X
End Property

Public Property Get Get_Decal_Position_Y() As Single
 Get_Decal_Position_Y = Decal.Position.Y
End Property

Public Property Get Get_Decal_Position_Z() As Single
 Get_Decal_Position_Z = Decal.Position.Z
End Property

Public Property Get Get_Decal_PositionEX() As D3DVECTOR
 Get_Decal_PositionEX = Decal.Position
End Property

Public Sub Position_Decal(X As Single, Y As Single, Z As Single)
Decal.Position = Math.Vector3(X, Y, Z)
Decal.SetPos = True
End Sub

Public Property Get Get_Decal_Rotation_X() As Single
 Get_Decal_Rotation_X = Decal.Rotation.X
End Property

Public Property Get Get_Decal_Rotation_Y() As Single
 Get_Decal_Rotation_Y = Decal.Rotation.Y
End Property

Public Property Get Get_Decal_Rotation_Z() As Single
 Get_Decal_Rotation_Z = Decal.Rotation.Z
End Property

Public Property Get Get_Decal_RotationEX() As D3DVECTOR
 Get_Decal_RotationEX = Decal.Rotation
End Property

Public Sub Rotate_Decal(X As Single, Y As Single, Z As Single)
Decal.Rotation = Math.Vector3(X, Y, Z)
Decal.Rotate = True
End Sub

Public Function Check_Camera_To_Decal_Collision(CameraPos As D3DVECTOR, DecalRadius As Single, Optional OffsetX = 0, Optional OffsetZ = 0) As Boolean
If CameraPos.X > (Decal.Position.X + OffsetX) - DecalRadius _
   And CameraPos.X < (Decal.Position.X + OffsetX) + DecalRadius _
   And CameraPos.Z > (Decal.Position.Z + OffsetZ) - DecalRadius _
   And CameraPos.Z < (Decal.Position.Z + OffsetZ) + DecalRadius Then
 Check_Camera_To_Decal_Collision = True
Else
 Check_Camera_To_Decal_Collision = False
End If
End Function

Public Sub Enable_Transparency(Enabled As Boolean)
If Enabled = True Then
 Decal.Transparent = True
Else
 Decal.Transparent = False
End If
End Sub

Public Sub Create_Decal(Texture_Filename As String, Decal_Width As Single, Decal_Height As Single, Optional Texture_U_Power = 1, Optional Texture_V_Power = 1)
On Local Error Resume Next
Set Decal.Decal = Nothing
Set Decal.Decal = pCreateDecal(Decal_Width, Decal_Height, Texture_U_Power, Texture_V_Power)
If Texture_Filename <> "" Then
 Set Decal.DecalSurface = Eng.Direct3DX.CreateTextureFromFileEx(D3DD, Texture_Filename, D3DX_DEFAULT, _
                                                                D3DX_DEFAULT, D3DX_DEFAULT, 0, _
                                                                D3DFMT_UNKNOWN, D3DPOOL_MANAGED, _
                                                                D3DX_FILTER_POINT, D3DX_FILTER_POINT, _
                                                                &HFF000000, ByVal 0, ByVal 0)
End If
HasSurface = True
End Sub

Private Function pCreateDecal(Width As Single, Height As Single, Optional Texture_U_Power = 1, Optional Texture_V_Power = 1) As D3DXMesh
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
        
Set pCreateDecal = rMesh
End Function

Public Sub Render_Decal()
Dim i As Long
Dim NumMat As Long
Dim Material As D3DMATERIAL8
Dim smat As D3DMATRIX
Dim smat2 As D3DMATRIX
Dim mat As D3DMATRIX
If D3DD Is Nothing Then Exit Sub

If HasSurface = True Then
 D3DD.GetTransform D3DTS_WORLD, smat2

 If Decal.Rotate = True Then
  D3DD.GetTransform D3DTS_WORLD, smat
  mRot = smat
  D3DXMatrixRotationYawPitchRoll mRot, Decal.Rotation.Y, Decal.Rotation.X, Decal.Rotation.Z
  D3DD.SetTransform D3DTS_WORLD, mRot
 End If

 If Decal.SetPos = True Then
  D3DD.GetTransform D3DTS_WORLD, smat
  mPos = smat
  mPos.m11 = 1: mPos.m22 = 1: mPos.m33 = 1: mPos.m44 = 1
  mPos.m41 = Decal.Position.X: mPos.m42 = Decal.Position.Y: mPos.m43 = Decal.Position.Z
  D3DD.SetTransform D3DTS_WORLD, mPos
 End If

  D3DD.SetRenderState D3DRS_ALPHABLENDENABLE, 1
  D3DD.SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCALPHA
  D3DD.SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA
  If Decal.Transparent = True Then
   D3DD.SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCCOLOR
   D3DD.SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCCOLOR
  End If
  
  Material.diffuse.r = 1: Material.diffuse.g = 1: Material.diffuse.b = 0.5
  Material.Ambient = Material.diffuse
  NumMat = 1
  D3DD.SetMaterial Material
  D3DD.SetTexture 0, Decal.DecalSurface
  For i = 0 To NumMat
   Decal.Decal.DrawSubset i
  Next
  
  D3DD.SetTexture 0, Nothing
  D3DD.SetTransform D3DTS_WORLD, smat2
  D3DD.SetRenderState D3DRS_DESTBLEND, 0
  D3DD.SetRenderState D3DRS_SRCBLEND, 0
  D3DD.SetRenderState D3DRS_ALPHABLENDENABLE, 0
Else
  Exit Sub
End If
End Sub

Public Sub Cleanup_Decal_Engine()
 Set Decal.DecalSurface = Nothing
 Set Decal.Decal = Nothing
 Set Math = Nothing
 Set Eng = Nothing
End Sub
