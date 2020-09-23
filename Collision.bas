Attribute VB_Name = "Collision"
Option Explicit

'NOTE: THESE ARE JUST TEST'S......THATS IT..
'FOR COLLISION DETECTION LOOK IN CLASS MODULES...

Private Type CUSTOMVERTEX
 X As Single
 Y As Single
 Z As Single
 tU As Single
 tV As Single
End Type
Private Const D3DFVF_COLORVERTEX = (D3DFVF_XYZ Or D3DFVF_DIFFUSE Or D3DFVF_TEX1)

Private Type tCollision
 Sphere1 As D3DVECTOR
 Sphere2 As D3DVECTOR
 Sphere1Radius As Single
 Sphere2Radius As Single
 Distance As Double
End Type
Private Collision As tCollision

Private VB1 As Direct3DVertexBuffer8
Private NumVertices1 As Long
Private Vertices1() As CUSTOMVERTEX

Private Function GetDist3D(Sphere2 As D3DVECTOR, Sphere1 As D3DVECTOR) As Single

'Note: For realistic and more accurate detection
'      avoid using (^ 2) or the srq() function..

'Big differance..and a lot faster..maybe like 5 - 10 frames faster

GetDist3D = (Sphere2.X - Sphere1.X) * (Sphere2.X - Sphere1.X) + (Sphere2.Y - Sphere1.Y) * (Sphere2.Y - Sphere1.Y) + (Sphere2.Z - Sphere1.Z) * (Sphere2.Z - Sphere1.Z)
End Function

Public Function CheckSphereCollision(Object1Position As D3DVECTOR, Sphere1Radius As Single, Object2Position As D3DVECTOR, Sphere2Radius As Single, _
                                     Optional CSphereOffsetX As Single = 0, Optional CSphereOffsetY As Single = 0, Optional CSphereOffsetZ As Single = 0) As Boolean
Collision.Sphere1.X = Object1Position.X + CSphereOffsetX
Collision.Sphere1.Y = Object1Position.Y + CSphereOffsetY
Collision.Sphere1.Z = Object1Position.Z + CSphereOffsetZ
Collision.Sphere1Radius = Sphere1Radius

Collision.Sphere2.X = Object2Position.X
Collision.Sphere2.Y = Object2Position.Y
Collision.Sphere2.Z = Object2Position.Z
Collision.Sphere2Radius = Sphere2Radius

Collision.Distance = GetDist3D(Collision.Sphere1, Collision.Sphere2)

If (Sphere1Radius + Sphere2Radius) = Collision.Distance Then
    CheckSphereCollision = True
ElseIf (Sphere1Radius + Sphere2Radius) < Collision.Distance Then
    CheckSphereCollision = False
ElseIf (Sphere1Radius + Sphere2Radius) > Collision.Distance Then
    CheckSphereCollision = True
End If

End Function

Private Function GetDist3D2(Position2 As D3DVECTOR, Position1 As D3DVECTOR) As Single
GetDist3D2 = (Position2.X - Position1.X) * (Position2.X - Position1.X) + (Position2.Y - Position1.Y) * (Position2.Y - Position1.Y) + (Position2.Z - Position1.Z) * (Position2.Z - Position1.Z)
End Function

Public Function CheckCollision(Object As D3DXMesh, Object2Position As D3DVECTOR, Sphere2Radius As Single, _
                               Optional CSphereOffsetX As Single = 0, Optional CSphereOffsetY As Single = 0, Optional CSphereOffsetZ As Single = 0) As Boolean
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

Set VB1 = Object.GetVertexBuffer
NumVertices1 = Object.GetNumVertices
ReDim Vertices1(NumVertices1)
D3DVertexBuffer8GetData VB1, 0, NumVertices1 * Len(Vertices1(0)), 0, Vertices1(0)

D3DD.GetTransform D3DTS_WORLD, mat

CX = 0: CY = 0: CZ = 0
For i = 0 To NumVertices1
 X = Vertices1(i).X + CSphereOffsetX
 Y = Vertices1(i).Y + CSphereOffsetY
 Z = Vertices1(i).Z + CSphereOffsetZ
 
 Position.X = X: Position.Y = Y: Position.Z = Z
 
 CX = CX + X * mat.m11 + Y * mat.m21 + Z * mat.m31 + mat.m41
 CY = CY + X * mat.m12 + Y * mat.m22 + Z * mat.m32 + mat.m42
 CZ = CZ + X * mat.m13 + Y * mat.m23 + Z * mat.m33 + mat.m43
Next

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

 Radius = MX * MX + MY * MY + MZ * MZ
Next

Collision.Distance = GetDist3D2(Position, Object2Position)

If ((Radius + CSphereOffsetX) + Sphere2Radius) = Collision.Distance Then
    CheckCollision = True
ElseIf ((Radius + CSphereOffsetY) + Sphere2Radius) < Collision.Distance Then
    CheckCollision = False
ElseIf ((Radius + CSphereOffsetZ) + Sphere2Radius) > Collision.Distance Then
    CheckCollision = True
End If

End Function

Public Sub Cleanup_Collision_Engine()
'Set VB = Nothing
ReDim Vertices(0)
End Sub

Public Function Collided(Main_Position As D3DVECTOR, ObjectPos As D3DVECTOR, ByVal ObjectRadius As Single, Optional OffsetX = 0, Optional OffsetZ = 0) As Boolean
If Main_Position.X > (ObjectPos.X + OffsetX) - ObjectRadius _
   And Main_Position.X < (ObjectPos.X + OffsetX) + ObjectRadius _
   And Main_Position.Z > (ObjectPos.Z + OffsetZ) - ObjectRadius _
   And Main_Position.Z < (ObjectPos.Z + OffsetZ) + ObjectRadius Then
 Collided = True
End If
End Function

Public Function CheckBoxCollision(Object As D3DXMesh, Object2Position As D3DVECTOR, Sphere2Radius As Single) As Boolean
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
Dim Box As D3DBOX

Set VB1 = Object.GetVertexBuffer
NumVertices1 = Object.GetNumVertices
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

 Radius = MX * MX + MY * MY + MZ * MZ
Next

Box.Left = MX
Box.Right = MX * MX
Box.Top = MY
Box.bottom = MY * MY
Box.front = MZ
Box.back = MZ * MZ

If Object2Position.X > Box.Left _
   And Object2Position.X < Box.Right _
   And Object2Position.Z > Box.front _
   And Object2Position.Z < Box.back Then
 CheckBoxCollision = True
End If

'Collision.Distance = GetDist3D2(position, Object2Position)

'If (Radius + Sphere2Radius) = Collision.Distance Then
'    CheckCollision = True
'ElseIf (Radius + Sphere2Radius) < Collision.Distance Then
'    CheckCollision = False
'ElseIf (Radius + Sphere2Radius) > Collision.Distance Then
'    CheckCollision = True
'End If

End Function
