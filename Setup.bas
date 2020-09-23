Attribute VB_Name = "Setup"
Option Explicit

'Setup ENGINE And Render

Public Const Num_Buffers = 10
Public Const Num_Buffers3D = 10
Public Const Num_Objects = 20

Public XE As New XEngine3D
Public XI As New XInput
Public XF As New XFilters
Public XL As New XLighting
Public XS2D As New XSound2D
Public XS3D As New XSound3D
Public XX As New X3DXFiles
Public XO As New XObjects
Public XS As New XSpriteXYZ
Public XC As New XCamera
Public XEN As New XEnviroment
Public XM As New XMouse
Public XD As New XDecal
Public XList As New XListener
Public XMidi As New XSoundMidi
Public XFX As New XSoundEffects
Public Math As New XMath

Public Weeds(0 To 20) As New X3DXFiles
Public EndLoop As Boolean
Public OK As Boolean
Public GetFPS As String

Private Sub CheckKeyPresses()
Static i As Long
Dim OhHmm As D3DVECTOR
Dim Pos As D3DVECTOR
Dim Pos2 As D3DVECTOR
Dim Hit As Boolean

If XI.KeyState(X_Escape) Then
 EndLoop = True
End If
XC.Start_Camera_Update
Math.Vector3_Normalize XC.Get_Camera_PositionEX
Math.Vector3_Normalize XX.Get_Object_PositionEX

If XI.KeyState(X_Up) Then
 If Not XX.Check_Camera_To_Object_Collision(XC.Get_Camera_PositionEX, 3.5, 1) _
    And Not XO.Check_Camera_To_Object_Collision(0, XC.Get_Camera_PositionEX, 8, -1, -1) _
    And Not XD.Check_Camera_To_Decal_Collision(XC.Get_Camera_PositionEX, 2, -1, -1) Then
  XC.Walk_Forward
  XS2D.PlaySound 0, 100, , 25000, 0
 Else
  XC.Slide_Off_Wall_Left
  XS2D.PlaySound 0, 100, , 22000, 0
 End If
End If

If XI.KeyState(X_Down) Then
 If Not XX.Check_Camera_To_Object_Collision(XC.Get_Camera_PositionEX, 3.5, 1) _
    And Not XO.Check_Camera_To_Object_Collision(0, XC.Get_Camera_PositionEX, 8, -1, -1) _
    And Not XD.Check_Camera_To_Decal_Collision(XC.Get_Camera_PositionEX, 2, -1, -1) Then
  XC.Walk_Backward
  XS2D.PlaySound 0, 100, , 25000, 0
 Else
  XC.Slide_Off_Wall_Left
  XS2D.PlaySound 0, 100, , 25000, 0
 End If
End If

If XI.KeyState(X_Left) Then XC.Strafe_Left: XS2D.PlaySound 0, 100, , 25000, 0
If XI.KeyState(X_Right) Then XC.Strafe_Right: XS2D.PlaySound 0, 100, , 25000, 0
If XI.KeyState(X_RShift) Then XC.Run
If XI.KeyState(X_RShift) And XI.KeyState(X_Up) Or _
   XI.KeyState(X_RShift) And XI.KeyState(X_Down) Then
 XS2D.PlaySound 0, 100, , 30000, 0
End If

If XI.KeyState(X_A) Then XC.Look_Up
If XI.KeyState(X_Z) Then XC.Look_Down
XM.Update_Mouse
XC.Free_Rotate XM.Get_Mouse_Position_Y / 100, XM.Get_Mouse_Position_X / 100, XM.Get_Mouse_Position_Z / 100
XC.End_Camera_Update

Math.Vector3_Normalize XC.Get_Camera_PositionEX
XC.Set_Camera_Eye_Level 2 + XEN.Get_Terrain_Height(XC.Get_Camera_Position_X, XC.Get_Camera_Position_Z)
If XC.Get_Camera_Eye_Level < 6 Then XC.Set_Camera_Eye_Level 6
If XC.Get_Camera_Eye_Level > 10 Then XC.Set_Camera_Eye_Level 10
End Sub

Public Sub Render()
Dim i As Integer
Do
 DoEvents
 CheckKeyPresses
 XE.Start_Engine_Render &HFF707070
 XEN.Render_Enviroment_Sphere , True, 50
 XEN.Render_Terrain
 XEN.Render_Water
 XO.Render_Object 0, XC
 XD.Render_Decal
 XO.Render_Billboard 1, XC
 XX.Render_X_Mesh XC
 XS.Render_Sprite
 GetFPS = XE.Draw_Text(XE.Get_FPS & " FPS", 5, 460, 0, 1, 0.5)
 XE.Show_FPS_Track 85, 460, 0, 1, 0.5
 XE.End_Engine_Render
Loop Until EndLoop = True
Cleanup
Unload frmMain
End Sub

Public Sub Cleanup()
XFX.Cleanup_SoundFX_Engine
Cleanup_Collision_Engine
XMidi.Cleanup_Midi_Engine
XList.Cleanup_Listener
XD.Cleanup_Decal_Engine
XM.Cleanup_Mouse_Engine
XEN.Cleanup_Enviroment_Engine
XO.Cleanup_Objects_Engine
XS.Cleanup_Sprite_Engine
XX.Cleanup_Geometry_Engine
XS2D.Cleanup_Sound_Engine Num_Buffers
XS3D.Cleanup_3D_Sound_Engine Num_Buffers3D
XL.Cleanup_Lighting_Engine
XI.Cleanup_Input_Engine
XE.Cleanup_XEngine8

Set Math = Nothing
Set XFX = Nothing
Set XMidi = Nothing
Set XList = Nothing
Set XD = Nothing
Set XM = Nothing
Set XEN = Nothing
Set XC = Nothing
Set XS = Nothing
Set XO = Nothing
Set XX = Nothing
Set XS3D = Nothing
Set XS2D = Nothing
Set XL = Nothing
Set XF = Nothing
Set XI = Nothing
Set XE = Nothing
End Sub

