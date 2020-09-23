VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "XCamera"
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

Private Type tCamera
 Position As D3DVECTOR
 Orientation As D3DVECTOR
 TurnSpeed As Single
 MoveSpeed As Single
 RunSpeed As Single
 PitchSpeed As Single
 StrafeSpeed As Single
 SlideOffWallSpeed As Single
 Angle As Single
 Degree As Single
 Pitch As Single
 MinPitch As Single
 MaxPitch As Single
 Running As Boolean
 Jump As Boolean
 Eye As Single
 FOV As Single
 Aspect As Single
 Near As Single
 Far As Single
End Type: Private Cam As tCamera

Private Math As XMath
Private matView As D3DMATRIX
Private matRotation As D3DMATRIX
Private matPitch As D3DMATRIX
Private matLook As D3DMATRIX
Private matPos As D3DMATRIX
Private matWorld As D3DMATRIX
Private matProj As D3DMATRIX
Private matZ As D3DMATRIX
Private PI As Single
Private g_90d As Single
Private g_180d As Single
Private g_270d As Single
Private g_360d As Single

Public Sub Initialize_Camera_Engine()
Set Math = New XMath
PI = Math.PI
g_90d = Math.Degrees_90
g_180d = Math.Degrees_180
g_270d = Math.Degrees_270
g_360d = Math.Degrees_360
 With Cam
 .TurnSpeed = g_90d / -70
 .MoveSpeed = 0.4
 .RunSpeed = 0.6
 .StrafeSpeed = 0.2
 .SlideOffWallSpeed = 0.1
 .PitchSpeed = 0.0099
 .Eye = 8
 .Position = Math.Vector3(0, Cam.Eye, 0)
 .Orientation = Math.Vector3(.Angle, .Pitch, 1)
 .Orientation.X = -1
 .Orientation.Y = 0
 .Angle = 0
 .Pitch = 0
 .MinPitch = -1.5
 .MaxPitch = -0.55
 .FOV = PI / 4
 .Aspect = 1
 .Near = 0.1
 .Far = 10000
 .Running = False
End With
Set Math = Nothing
End Sub

Public Property Get Get_Camera_Field_Of_View() As Single
Get_Camera_Field_Of_View = Cam.FOV
End Property

Public Sub Set_Camera_Field_Of_View(Optional Field_Of_View As Single = 0.785398163397448)
Cam.FOV = Field_Of_View
End Sub

Public Property Get Get_Camera_Aspect() As Single
Get_Camera_Aspect = Cam.Aspect
End Property

Public Sub Set_Camera_Aspect(Optional Aspect As Single = 1)
Cam.Aspect = Aspect
End Sub

Public Property Get Get_Camera_Near() As Single
Get_Camera_Near = Cam.Near
End Property

Public Sub Set_Camera_Near(Optional Near As Single = 0.1)
Cam.Near = Near
End Sub

Public Property Get Get_Camera_Far() As Single
Get_Camera_Far = Cam.Far
End Property

Public Sub Set_Camera_Far(Optional Far As Single = 10000)
Cam.Far = Far
End Sub

Public Property Get Get_Camera_Eye_Level() As Single
Get_Camera_Eye_Level = Cam.Position.Y
End Property

Public Function Set_Camera_Eye_Level(Optional Height As Single = 8)
Cam.Position.Y = Height
Cam.Eye = Height
End Function

Public Property Get Get_Min_Pitch() As Single
Get_Min_Pitch = Cam.MinPitch
End Property

Public Sub Set_Min_Pitch(Optional Value As Single = -1.5)
Cam.MinPitch = Value
End Sub

Public Property Get Get_Max_Pitch() As Single
Get_Max_Pitch = Cam.MaxPitch
End Property

Public Sub Set_Max_Pitch(Optional Value As Single = -0.55)
Cam.MaxPitch = Value
End Sub

Public Property Get Get_Pitch_Speed() As Single
Get_Pitch_Speed = Cam.PitchSpeed
End Property

Public Sub Set_Pitch_Speed(Optional Speed As Single = 0.0099)
Cam.PitchSpeed = Speed
End Sub

Public Property Get Get_Walking_Speed() As Single
Get_Walking_Speed = Cam.MoveSpeed
End Property

Public Sub Set_Walking_Speed(Optional Speed As Single = 0.2)
Cam.MoveSpeed = Speed
End Sub

Public Property Get Get_Running_Speed() As Single
Get_Running_Speed = Cam.RunSpeed
End Property

Public Sub Set_Running_Speed(Optional Speed As Single = 0.4)
Cam.RunSpeed = Speed
End Sub

Public Property Get Get_Slide_Off_Wall_Speed() As Single
Get_Slide_Off_Wall_Speed = Cam.SlideOffWallSpeed
End Property

Public Sub Set_Slide_Off_Wall_Speed(Optional Speed As Single = 0.1)
Cam.SlideOffWallSpeed = Speed
End Sub

Public Property Get Get_Camera_Position_X() As Single
Get_Camera_Position_X = Cam.Position.X
End Property

Public Property Get Get_Camera_Position_Y() As Single
Get_Camera_Position_Y = Cam.Position.Y
End Property

Public Property Get Get_Camera_Position_Z() As Single
Get_Camera_Position_Z = Cam.Position.Z
End Property

Public Property Get Get_Camera_PositionEX() As D3DVECTOR
Get_Camera_PositionEX = Cam.Position
End Property

Public Property Get Get_Camera_Orientation_X() As Single
Get_Camera_Orientation_X = Cam.Orientation.X
End Property

Public Property Get Get_Camera_Orientation_Y() As Single
Get_Camera_Orientation_Y = Cam.Orientation.Y
End Property

Public Property Get Get_Camera_Orientation_Z() As Single
Get_Camera_Orientation_Z = Cam.Orientation.Z
End Property

Public Property Get Get_Camera_OrientationEX() As D3DVECTOR
Get_Camera_OrientationEX = Cam.Orientation
End Property

Public Property Get Get_Camera_Pitch() As Single
Get_Camera_Pitch = Cam.Orientation.X
End Property

Public Function Run()
Cam.Running = True
End Function

Public Function Walk_Forward() As D3DVECTOR
Dim matView As D3DMATRIX, matTrans As D3DMATRIX, matRot As D3DMATRIX, matRotX As D3DMATRIX, matRotY As D3DMATRIX
Dim Translation As D3DVECTOR, vec As D3DVECTOR
On Local Error Resume Next
If Cam.Running = True Then
 D3DXMatrixRotationY matRotation, -Cam.Orientation.Y
 D3DXMatrixMultiply matRot, matPitch, matRotation
 D3DXMatrixTranslation matTrans, 0, 0, 2 * Cam.RunSpeed
 D3DXMatrixMultiply matView, matTrans, matRot
 D3DXVec3TransformCoord Translation, Translation, matView
 D3DXVec3Add Cam.Position, Translation, Cam.Position
 Cam.Running = False
Else
 D3DXMatrixRotationY matRotation, -Cam.Orientation.Y
 D3DXMatrixMultiply matRot, matPitch, matRotation
 D3DXMatrixTranslation matTrans, 0, 0, 2 * Cam.MoveSpeed
 D3DXMatrixMultiply matView, matTrans, matRot
 D3DXVec3TransformCoord Translation, Translation, matView
 D3DXVec3Add Cam.Position, Translation, Cam.Position
End If
Walk_Forward = Translation
If Cam.Jump = False Then
 vec = Cam.Position
 vec.Y = Cam.Eye
 Cam.Position = vec
End If
End Function

Public Function Walk_Backward()
Dim matView As D3DMATRIX, matTrans As D3DMATRIX, matRot As D3DMATRIX, matRotX As D3DMATRIX, matRotY As D3DMATRIX
Dim Translation As D3DVECTOR, vec As D3DVECTOR
On Local Error Resume Next
If Cam.Running = True Then
 D3DXMatrixRotationY matRotation, -Cam.Orientation.Y
 D3DXMatrixMultiply matRot, matPitch, matRotation
 D3DXMatrixTranslation matTrans, 0, 0, -2 * Cam.RunSpeed
 D3DXMatrixMultiply matView, matTrans, matRot
 D3DXVec3TransformCoord Translation, Translation, matView
 D3DXVec3Add Cam.Position, Translation, Cam.Position
 Cam.Running = False
Else
 D3DXMatrixRotationY matRotation, -Cam.Orientation.Y
 D3DXMatrixMultiply matRot, matPitch, matRotation
 D3DXMatrixTranslation matTrans, 0, 0, -2 * Cam.MoveSpeed
 D3DXMatrixMultiply matView, matTrans, matRot
 D3DXVec3TransformCoord Translation, Translation, matView
 D3DXVec3Add Cam.Position, Translation, Cam.Position
End If
If Cam.Jump = False Then
 vec = Cam.Position
 vec.Y = Cam.Eye
 Cam.Position = vec
End If
End Function

Public Function Turn_Left()
Cam.Orientation.Y = Cam.Orientation.Y - Cam.TurnSpeed
If Cam.Orientation.Y > g_360d Then
 Cam.Orientation.Y = 0 + (Cam.Orientation.Y - g_360d)
End If
End Function

Public Function Turn_Right()
Cam.Orientation.Y = Cam.Orientation.Y + Cam.TurnSpeed
If Cam.Orientation.Y < 0 Then
 Cam.Orientation.Y = g_360d - (-Cam.Orientation.Y)
End If
End Function

Public Function Look_Up()
If Cam.Orientation.X >= Cam.MaxPitch Then
 Cam.Orientation.X = Cam.Orientation.X
Else
 Cam.Orientation.X = Cam.Orientation.X + Cam.PitchSpeed
End If
End Function

Public Function Look_Down()
If Cam.Orientation.X <= Cam.MinPitch Then
 Cam.Orientation.X = Cam.Orientation.X
Else
 Cam.Orientation.X = Cam.Orientation.X - Cam.PitchSpeed
End If
End Function

Public Sub Strafe_Left()
Dim matView As D3DMATRIX, matTrans As D3DMATRIX, matRot As D3DMATRIX, matRotX As D3DMATRIX, matRotY As D3DMATRIX
Dim Translation As D3DVECTOR
D3DXMatrixRotationY matRotation, -Cam.Orientation.Y
D3DXMatrixMultiply matRot, matPitch, matRotation
D3DXMatrixTranslation matTrans, -2 * Cam.StrafeSpeed, 0, 0
D3DXMatrixMultiply matView, matTrans, matRot
D3DXVec3TransformCoord Translation, Translation, matView
D3DXVec3Add Cam.Position, Translation, Cam.Position
End Sub

Public Sub Strafe_Right()
Dim matView As D3DMATRIX, matTrans As D3DMATRIX, matRot As D3DMATRIX, matRotX As D3DMATRIX, matRotY As D3DMATRIX
Dim Translation As D3DVECTOR
D3DXMatrixRotationY matRotation, -Cam.Orientation.Y
D3DXMatrixMultiply matRot, matPitch, matRotation
D3DXMatrixTranslation matTrans, 2 * Cam.StrafeSpeed, 0, 0
D3DXMatrixMultiply matView, matTrans, matRot
D3DXVec3TransformCoord Translation, Translation, matView
D3DXVec3Add Cam.Position, Translation, Cam.Position
End Sub

Public Sub Slide_Off_Wall_Left()
Dim matView As D3DMATRIX, matTrans As D3DMATRIX, matRot As D3DMATRIX, matRotX As D3DMATRIX, matRotY As D3DMATRIX
Dim Translation As D3DVECTOR
D3DXMatrixRotationY matRotation, -Cam.Orientation.Y
D3DXMatrixMultiply matRot, matPitch, matRotation
D3DXMatrixTranslation matTrans, -0.5 * Cam.StrafeSpeed, 0, 0
D3DXMatrixMultiply matView, matTrans, matRot
D3DXVec3TransformCoord Translation, Translation, matView
D3DXVec3Add Cam.Position, Translation, Cam.Position
End Sub

Public Sub Slide_Off_Wall_Right()
Dim matView As D3DMATRIX, matTrans As D3DMATRIX, matRot As D3DMATRIX, matRotX As D3DMATRIX, matRotY As D3DMATRIX
Dim Translation As D3DVECTOR
D3DXMatrixRotationY matRotation, -Cam.Orientation.Y
D3DXMatrixMultiply matRot, matPitch, matRotation
D3DXMatrixTranslation matTrans, 0.5 * Cam.StrafeSpeed, 0, 0
D3DXMatrixMultiply matView, matTrans, matRot
D3DXVec3TransformCoord Translation, Translation, matView
D3DXVec3Add Cam.Position, Translation, Cam.Position
End Sub

Public Sub Free_Rotate(X As Single, Y As Single, Z As Single)
On Error Resume Next
Dim vec As D3DVECTOR
vec.X = -X: vec.Y = -Y: vec.Z = -Z
D3DXVec3Add Cam.Orientation, Cam.Orientation, vec
End Sub

Public Sub Start_Camera_Update()
On Local Error Resume Next
Dim vec As D3DVECTOR
D3DXMatrixIdentity matWorld
D3DD.SetTransform D3DTS_WORLD, matWorld

If Cam.Jump = False Then
 vec = Cam.Position
 vec.Y = Cam.Eye
 Cam.Position = vec
End If
End Sub

Public Sub End_Camera_Update()
On Local Error Resume Next

D3DXMatrixRotationY matRotation, Cam.Orientation.Y
If Cam.Orientation.X <= Cam.MinPitch Or Cam.Orientation.X >= Cam.MaxPitch Then
 Cam.Orientation.X = Cam.Orientation.X
Else
 D3DXMatrixRotationX matPitch, Cam.Orientation.X
End If
D3DXMatrixRotationX matZ, Cam.Orientation.Z
D3DXMatrixMultiply matLook, matRotation, matPitch
D3DXMatrixMultiply matLook, matLook, matZ
D3DXMatrixTranslation matPos, -Cam.Position.X, -Cam.Position.Y, -Cam.Position.Z
D3DXMatrixMultiply matView, matPos, matLook
D3DD.SetTransform D3DTS_VIEW, matView
D3DXMatrixPerspectiveFovLH matProj, Cam.FOV, Cam.Aspect, Cam.Near, Cam.Far
D3DD.SetTransform D3DTS_PROJECTION, matProj

End Sub
