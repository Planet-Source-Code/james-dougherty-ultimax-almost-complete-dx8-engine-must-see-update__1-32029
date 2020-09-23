VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "XLighting"
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

Private Const NUM_LIGHTS = 512

Private Type tLight
  Index As Integer
  Name As String
  Enabled As Boolean
  Range As Single
  Direction As D3DVECTOR
  Position As D3DVECTOR
  Light_Type As CONST_D3DLIGHTTYPE
  Lights As D3DLIGHT8
  OutterConeRadius As Single
  InnerConeRadius As Single
  Falloff As Single
  Attenuation0 As Single
  Attenuation1 As Single
  Attenuation2 As Single
  AmbientColorR As Single
  AmbientColorG As Single
  AmbientColorB As Single
  DiffuseColorR As Single
  DiffuseColorG As Single
  DiffuseColorB As Single
  SpecularColorR As Single
  SpecularColorG As Single
  SpecularColorB As Single
End Type: Private Light(NUM_LIGHTS) As tLight

Private Eng As New XEngine3D

Public Sub Cleanup_Lighting_Engine()
 Set Eng = Nothing
End Sub

Public Property Get Get_Light_Name(Light_Index As Integer) As String
 On Local Error Resume Next
 Get_Light_Name = Light(Light_Index).Name
End Property

Public Property Get Get_Light_Direction_X(Light_Index As Integer) As Single
On Local Error Resume Next
Get_Light_Direction_X = Light(Light_Index).Lights.Direction.X
End Property

Public Property Get Get_Light_Direction_Y(Light_Index As Integer) As Single
On Local Error Resume Next
Get_Light_Direction_Y = Light(Light_Index).Lights.Direction.Y
End Property

Public Property Get Get_Light_Direction_Z(Light_Index As Integer) As Single
On Local Error Resume Next
Get_Light_Direction_Z = Light(Light_Index).Lights.Direction.Z
End Property

Public Property Get Get_Light_DirectionEX(Light_Index As Integer) As D3DVECTOR
On Local Error Resume Next
Get_Light_DirectionEX = Light(Light_Index).Lights.Direction
End Property

Public Property Get Get_Light_Position_X(Light_Index As Integer) As Single
On Local Error Resume Next
Get_Light_Position_X = Light(Light_Index).Lights.Position.X
End Property

Public Property Get Get_Light_Position_Y(Light_Index As Integer) As Single
On Local Error Resume Next
Get_Light_Position_Y = Light(Light_Index).Lights.Position.Y
End Property

Public Property Get Get_Light_Position_Z(Light_Index As Integer) As Single
On Local Error Resume Next
Get_Light_Position_Z = Light(Light_Index).Lights.Position.Z
End Property

Public Property Get Get_Light_PositionEX(Light_Index As Integer) As D3DVECTOR
On Local Error Resume Next
Get_Light_PositionEX = Light(Light_Index).Lights.Position
End Property

Public Property Get Get_Light_Ambient_Color_Red(Light_Index As Integer) As Single
On Local Error Resume Next
Get_Light_Ambient_Color_Red = Light(Light_Index).Lights.Ambient.r
End Property

Public Property Get Get_Light_Ambient_Color_Green(Light_Index As Integer) As Single
On Local Error Resume Next
Get_Light_Ambient_Color_Green = Light(Light_Index).Lights.Ambient.g
End Property

Public Property Get Get_Light_Ambient_Color_Blue(Light_Index As Integer) As Single
On Local Error Resume Next
Get_Light_Ambient_Color_Blue = Light(Light_Index).Lights.Ambient.b
End Property

Public Property Get Get_Light_Ambient_ColorEX(Light_Index As Integer) As D3DCOLORVALUE
On Local Error Resume Next
Get_Light_Ambient_ColorEX = Light(Light_Index).Lights.Ambient
End Property

Public Property Get Get_Light_Diffuse_Color_Red(Light_Index As Integer) As Single
On Local Error Resume Next
Get_Light_Diffuse_Color_Red = Light(Light_Index).Lights.diffuse.r
End Property

Public Property Get Get_Light_Diffuse_Color_Green(Light_Index As Integer) As Single
On Local Error Resume Next
Get_Light_Diffuse_Color_Green = Light(Light_Index).Lights.diffuse.g
End Property

Public Property Get Get_Light_Diffuse_Color_Blue(Light_Index As Integer) As Single
On Local Error Resume Next
Get_Light_Diffuse_Color_Blue = Light(Light_Index).Lights.diffuse.b
End Property

Public Property Get Get_Light_Diffuse_ColorEX(Light_Index As Integer) As D3DCOLORVALUE
On Local Error Resume Next
Get_Light_Diffuse_ColorEX = Light(Light_Index).Lights.diffuse
End Property

Public Property Get Get_Light_Specular_Color_Red(Light_Index As Integer) As Single
On Local Error Resume Next
Get_Light_Specular_Color_Red = Light(Light_Index).Lights.Specular.r
End Property

Public Property Get Get_Light_Specular_Color_Green(Light_Index As Integer) As Single
On Local Error Resume Next
Get_Light_Specular_Color_Green = Light(Light_Index).Lights.Specular.g
End Property

Public Property Get Get_Light_Specular_Color_Blue(Light_Index As Integer) As Single
On Local Error Resume Next
Get_Light_Specular_Color_Blue = Light(Light_Index).Lights.Specular.b
End Property

Public Property Get Get_Light_Specular_ColorEX(Light_Index As Integer) As D3DCOLORVALUE
On Local Error Resume Next
Get_Light_Specular_ColorEX = Light(Light_Index).Lights.Specular
End Property

Public Property Get Get_SpotLight_Folloff(Light_Index As Integer) As Single
On Local Error Resume Next
Get_SpotLight_Folloff = Light(Light_Index).Lights.Falloff
End Property

Public Sub Set_Light_Position(Light_Index As Integer, Position As D3DVECTOR)
On Local Error Resume Next
 Light(Light_Index).Lights.Position = Position
 Light(Light_Index).Position.X = Position.X
 Light(Light_Index).Position.Y = Position.Y
 Light(Light_Index).Position.Z = Position.Z
 Eng.Direct3DDevice.SetLight Light_Index, Light(Light_Index).Lights
End Sub

Public Sub Set_Light_Ambient_Color(Light_Index As Integer, r As Single, g As Single, b As Single)
On Local Error Resume Next
 Light(Light_Index).Lights.Ambient = ColorValue4(r, g, b, 1)
 Light(Light_Index).AmbientColorR = r
 Light(Light_Index).AmbientColorG = g
 Light(Light_Index).AmbientColorB = b
 Eng.Direct3DDevice.SetLight Light_Index, Light(Light_Index).Lights
End Sub

Public Sub Set_Light_Diffuse_Color(Light_Index As Integer, r As Single, g As Single, b As Single)
On Local Error Resume Next
 Light(Light_Index).Lights.diffuse = ColorValue4(r, g, b, 1)
 Light(Light_Index).DiffuseColorR = r
 Light(Light_Index).DiffuseColorG = g
 Light(Light_Index).DiffuseColorB = b
 Eng.Direct3DDevice.SetLight Light_Index, Light(Light_Index).Lights
End Sub

Public Sub Set_Light_Specular_Color(Light_Index As Integer, r As Single, g As Single, b As Single)
On Local Error Resume Next
 Light(Light_Index).Lights.Specular = ColorValue4(r, g, b, 1)
 Light(Light_Index).SpecularColorR = r
 Light(Light_Index).SpecularColorG = g
 Light(Light_Index).SpecularColorB = b
 Eng.Direct3DDevice.SetLight Light_Index, Light(Light_Index).Lights
End Sub

Public Sub Set_Light_Direction_Position(Light_Index As Integer, X As Single, Y As Single, Z As Single)
On Local Error Resume Next
 Light(Light_Index).Lights.Direction = Vector(X, Y, Z)
 Light(Light_Index).Direction.X = X
 Light(Light_Index).Direction.Y = Y
 Light(Light_Index).Direction.Z = Z
 Eng.Direct3DDevice.SetLight Light_Index, Light(Light_Index).Lights
End Sub

Public Sub Set_Directional_Light_Cone_Radius(Light_Index As Integer, Inner_Cone As Single, Outter_Cone As Single)
On Local Error Resume Next
 Light(Light_Index).Lights.Phi = Outter_Cone
 Light(Light_Index).Lights.theta = Inner_Cone
 Eng.Direct3DDevice.SetLight Light_Index, Light(Light_Index).Lights
End Sub

Public Sub Set_Light_Falloff(Light_Index As Integer, Falloff As Single)
On Local Error Resume Next
 Light(Light_Index).Lights.Falloff = Falloff
 Light(Light_Index).Falloff = Falloff
 Eng.Direct3DDevice.SetLight Light_Index, Light(Light_Index).Lights
End Sub

Public Sub Set_Light_Attenuation(Light_Index As Integer, a1 As Single, a2 As Single, a3 As Single)
On Local Error Resume Next
 Light(Light_Index).Lights.Attenuation0 = a1
 Light(Light_Index).Lights.Attenuation1 = a2
 Light(Light_Index).Lights.Attenuation2 = a3
 Light(Light_Index).Attenuation0 = a1
 Light(Light_Index).Attenuation1 = a2
 Light(Light_Index).Attenuation2 = a3
End Sub

Public Sub Set_Light_Range(Light_Index As Integer, Range As Single)
On Local Error Resume Next
 Light(Light_Index).Lights.Range = Range
 Light(Light_Index).Range = Range
 Eng.Direct3DDevice.SetLight Light_Index, Light(Light_Index).Lights
End Sub

Public Function Create_Directional_Light(Light_Index As Integer, Light_Name As String, Direction As D3DVECTOR, Range As Single, DiffuseR As Single, DiffuseG As Single, DiffuseB As Single, Optional AmbientR As Single = 0, Optional AmbientG As Single = 0, Optional AmbientB As Single = 0) As Integer
On Local Error Resume Next
Create_Directional_Light = Light_Index
Light(Light_Index).Enabled = True

With Light(Light_Index)
    .Lights.Type = D3DLIGHT_DIRECTIONAL
    .Lights.Ambient = ColorValue4(AmbientR, AmbientG, AmbientB, 1)
    .Lights.diffuse = ColorValue4(DiffuseR, DiffuseG, DiffuseB, 1)
    .Lights.Specular = .Lights.diffuse
    .Lights.Direction = Direction
    .Lights.Range = Range
    .Lights.Attenuation0 = 1
    .Lights.Attenuation1 = 1
    .Lights.Attenuation2 = 0
    .Name = Light_Name
    .Index = Light_Index
    .AmbientColorR = AmbientR
    .AmbientColorG = AmbientG
    .AmbientColorB = AmbientB
    .DiffuseColorR = DiffuseR
    .DiffuseColorG = DiffuseG
    .DiffuseColorB = DiffuseB
    .SpecularColorR = DiffuseR
    .SpecularColorG = DiffuseG
    .SpecularColorB = DiffuseB
    .Light_Type = D3DLIGHT_DIRECTIONAL
    .Range = Range
    .Enabled = True
    .Direction = Direction
End With
    
Eng.Direct3DDevice.SetLight Light_Index, Light(Light_Index).Lights
Eng.Direct3DDevice.LightEnable Light_Index, True
Eng.Direct3DDevice.SetRenderState D3DRS_LIGHTING, 1
End Function

Public Function Create_Point_Light(Light_Index As Integer, Light_Name As String, Position As D3DVECTOR, Range As Single, DiffuseR As Single, DiffuseG As Single, DiffuseB As Single, AmbientR As Single, AmbientG As Single, AmbientB As Single) As Integer
On Local Error Resume Next
Create_Point_Light = Light_Index
Light(Light_Index).Enabled = True

With Light(Light_Index)
    .Lights.Type = D3DLIGHT_POINT
    .Lights.Position = Position
    .Lights.Ambient = ColorValue4(AmbientR, AmbientG, AmbientB, 1)
    .Lights.diffuse = ColorValue4(DiffuseR, DiffuseG, DiffuseB, 1)
    .Lights.Specular = .Lights.diffuse
    .Lights.Range = Range
    .Lights.Attenuation0 = 1
    .Lights.Attenuation1 = 1
    .Lights.Attenuation2 = 0
    .Name = Light_Name
    .Index = Light_Index
    .AmbientColorR = AmbientR
    .AmbientColorG = AmbientG
    .AmbientColorB = AmbientB
    .DiffuseColorR = DiffuseR
    .DiffuseColorG = DiffuseG
    .DiffuseColorB = DiffuseB
    .SpecularColorR = DiffuseR
    .SpecularColorG = DiffuseG
    .SpecularColorB = DiffuseB
    .Position = Position
    .Light_Type = D3DLIGHT_POINT
    .Range = Range
    .Enabled = True
End With

Eng.Direct3DDevice.SetLight Light_Index, Light(Light_Index).Lights
Eng.Direct3DDevice.LightEnable Light_Index, True
Eng.Direct3DDevice.SetRenderState D3DRS_LIGHTING, 1
End Function

Public Function Create_Spot_Light(Light_Index As Integer, Light_Name As String, Direction As D3DVECTOR, Position As D3DVECTOR, Outter_Radius As Single, Inner_Radius As Single, Falloff As Single, DiffuseR As Single, DiffuseG As Single, DiffuseB As Single, AmbientR As Single, AmbientG As Single, AmbientB As Single) As Integer
On Local Error Resume Next
Create_Spot_Light = Light_Index
Light(Light_Index).Enabled = True

With Light(Light_Index)
    .Lights.Type = D3DLIGHT_SPOT
    .Lights.Direction = Direction
    .Lights.Position = Position
    .Lights.Falloff = Falloff
    .Lights.Phi = Outter_Radius
    .Lights.theta = Inner_Radius
    .Lights.Ambient = ColorValue4(AmbientR, AmbientG, AmbientB, 1)
    .Lights.diffuse = ColorValue4(DiffuseR, DiffuseG, DiffuseB, 1)
    .Lights.Specular = .Lights.diffuse
    .Lights.Attenuation0 = 1
    .Lights.Attenuation1 = 1
    .Lights.Attenuation2 = 0
    .Name = Light_Name
    .Index = Light_Index
    .AmbientColorR = AmbientR
    .AmbientColorG = AmbientG
    .AmbientColorB = AmbientB
    .DiffuseColorR = DiffuseR
    .DiffuseColorG = DiffuseG
    .DiffuseColorB = DiffuseB
    .SpecularColorR = DiffuseR
    .SpecularColorG = DiffuseG
    .SpecularColorB = DiffuseB
    .Position = Position
    .Direction = Direction
    .Light_Type = D3DLIGHT_SPOT
    .InnerConeRadius = Inner_Radius
    .OutterConeRadius = Outter_Radius
    .Enabled = True
End With

Eng.Direct3DDevice.SetLight Light_Index, Light(Light_Index).Lights
Eng.Direct3DDevice.LightEnable Light_Index, True
Eng.Direct3DDevice.SetRenderState D3DRS_LIGHTING, 1
End Function

Public Sub Enable_Disable_Light(Light_Index As Integer, Enabled As Boolean)
On Local Error Resume Next
Eng.Direct3DDevice.LightEnable Light_Index, Enabled
Light(Light_Index).Enabled = Enabled
End Sub

Private Function ColorValue4(r As Single, g As Single, b As Single, A As Single) As D3DCOLORVALUE
On Local Error Resume Next
 Dim C As D3DCOLORVALUE
 C.A = A
 C.r = r
 C.g = g
 C.b = b
 ColorValue4 = C
End Function

Private Function Vector(X, Y, Z) As D3DVECTOR
On Local Error Resume Next
 Vector.X = X
 Vector.Y = Y
 Vector.Z = Z
End Function
