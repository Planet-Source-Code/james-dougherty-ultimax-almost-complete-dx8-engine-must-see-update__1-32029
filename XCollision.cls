VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "XCollision"
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

Private Type tCollision
 Sphere1 As D3DVECTOR
 Sphere2 As D3DVECTOR
 Sphere1Radius As Single
 Sphere2Radius As Single
 Distance As Double
End Type: Private Collision As tCollision

Private Function GetDist3D(Sphere2 As D3DVECTOR, Sphere1 As D3DVECTOR) As Single

'Note: For realistic and more accurate detection
'      avoid using (^ 2) or the srq() function..

'Big differance..and a lot faster..maybe like 5 - 10 frames faster

GetDist3D = (Sphere2.X - Sphere1.X) * (Sphere2.X - Sphere1.X) + (Sphere2.Y - Sphere1.Y) * (Sphere2.Y - Sphere1.Y) + (Sphere2.Z - Sphere1.Z) * (Sphere2.Z - Sphere1.Z)
End Function

Public Function Check_Sphere_To_Sphere_Collision(Object1Position As D3DVECTOR, Sphere1Radius As Single, Object2Position As D3DVECTOR, Sphere2Radius As Single, _
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
    Check_Sphere_To_Sphere_Collision = True
ElseIf (Sphere1Radius + Sphere2Radius) < Collision.Distance Then
    Check_Sphere_To_Sphere_Collision = False
ElseIf (Sphere1Radius + Sphere2Radius) > Collision.Distance Then
    Check_Sphere_To_Sphere_Collision = True
End If

End Function

Public Function Check_Camera_To_Object_Collision(Main_Position As D3DVECTOR, ObjectPos As D3DVECTOR, ByVal ObjectRadius As Single, Optional OffsetX = 0, Optional OffsetZ = 0) As Boolean
If Main_Position.X > (ObjectPos.X + OffsetX) - ObjectRadius _
   And Main_Position.X < (ObjectPos.X + OffsetX) + ObjectRadius _
   And Main_Position.Z > (ObjectPos.Z + OffsetZ) - ObjectRadius _
   And Main_Position.Z < (ObjectPos.Z + OffsetZ) + ObjectRadius Then
 Check_Camera_To_Object_Collision = True
Else
 Check_Camera_To_Object_Collision = False
End If
End Function
