Attribute VB_Name = "Collision2"
Option Explicit

'NOTE: THESE ARE JUST TEST'S......THATS IT..
'FOR COLLISION DETECTION LOOK IN CLASS MODULES...

Private Type tCollisionCollection
 Hit As Long
 TriFaceID As Long
 A As Single
 b As Single
 Dist As Single
End Type

Private CResult() As tCollisionCollection
Private Mesh() As D3DXMesh
Public Count As Long
Private Maxsize As Long
Private Const GrowSize = 10

Public Function GetCount() As Long
GetCount = Count
End Function

Public Function Find_Nearest_Object() As Long
Dim q As Long, mindist As Single, i As Long
q = -1
mindist = 1E+38

For i = 0 To Count - 1
 If CResult(i).Dist < mindist Then
  q = i
  mindist = CResult(i).Dist
 End If
Next

Find_Nearest_Object = q
End Function

Public Sub Retrieve_Data(ListIndex As Long, ByRef A As Single, ByRef b As Single, ByRef Dist As Single, ByRef TriFaceID As Long)
A = CResult(ListIndex).A
b = CResult(ListIndex).b
Dist = CResult(ListIndex).Dist
TriFaceID = CResult(ListIndex).TriFaceID
End Sub

Public Function RayPick(Mesh As D3DXMesh, vOrig As D3DVECTOR, vDir As D3DVECTOR, Radius As Single) As Boolean
Destroy
Dim World As D3DMATRIX
D3DD.GetTransform D3DTS_WORLD, World
RayPick = RayPickEx(Mesh, World, vOrig, vDir, Radius)
End Function

Private Function RayPickEx(Mesh As D3DXMesh, worldmatrix As D3DMATRIX, vOrig As D3DVECTOR, vDir As D3DVECTOR, Radius As Single) As Boolean
Dim NewWorldMatrix As D3DMATRIX
Dim InvWorldMatrix As D3DMATRIX
Dim currentMatrix As D3DMATRIX
Dim i As Long, det As Single, bHit As Boolean
Dim vNewDir As D3DVECTOR, vNewOrig As D3DVECTOR
Dim DistSq As Single
Dim MaxDistSq As Single
        
If Mesh Is Nothing Then Exit Function
        
D3DD.GetTransform D3DTS_WORLD, currentMatrix
D3DXMatrixMultiply NewWorldMatrix, currentMatrix, worldmatrix
D3DXMatrixInverse InvWorldMatrix, det, NewWorldMatrix
    
Call D3DXVec3TransformNormal(vNewDir, vDir, InvWorldMatrix)
Call D3DXVec3TransformCoord(vNewOrig, vOrig, InvWorldMatrix)
                           
MaxDistSq = D3DXVec3Length(vNewDir) + Radius
MaxDistSq = MaxDistSq * MaxDistSq
DistSq = D3DXVec3LengthSq(vNewOrig)
    
Dim tmpMesh As D3DXMesh
Dim tmpResult As tCollisionCollection

If DistSq < MaxDistSq Then
 Set tmpMesh = Mesh
 bHit = False
 tmpResult.Hit = 0
 If Not tmpMesh Is Nothing Then
  Call D3DXVec3Scale(vDir, vDir, 1000)
  XE.Direct3DX.Intersect tmpMesh, vNewOrig, vDir, tmpResult.Hit, tmpResult.TriFaceID, tmpResult.A, tmpResult.b, tmpResult.Dist
 End If
        
 If tmpResult.Hit <> 0 Then
  AddCollisionItem tmpMesh, tmpResult
  tmpResult.Hit = 0
 End If
 bHit = True
End If
    
RayPickEx = bHit
End Function

Public Sub AddCollisionItem(InMesh As D3DXMesh, InItem As tCollisionCollection)
If Maxsize = 0 Then
 ReDim CResult(GrowSize)
 ReDim Mesh(GrowSize)
 Maxsize = GrowSize
ElseIf Count >= Maxsize Then
 ReDim Preserve CResult(Maxsize + GrowSize)
 ReDim Preserve Mesh(Maxsize + GrowSize)
 Maxsize = Maxsize + GrowSize
End If
    
Set Mesh(Count) = InMesh
CResult(Count) = InItem
                             
Count = Count + 1
End Sub

Public Function Destroy()
ReDim Mesh(0)
Count = 0
Maxsize = 0
End Function
