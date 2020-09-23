VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "XMouse"
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

Public Enum TypeMouseButton
 Left_Button = 0
 Right_Button = 1
 Middle_Button = 2
End Enum

Private XEngine As New XEngine3D
Private DI8 As DirectInput8
Private DIDEV8 As DirectInputDevice8
Private IMouseState As DIMOUSESTATE
Private MouseS As Direct3DSurface8
Private Showing As Boolean
Private HasSurface As Boolean

Public Function MouseState() As DIMOUSESTATE
On Local Error Resume Next
MouseState = IMouseState
DIDEV8.GetDeviceStateMouse IMouseState
End Function

Public Property Get Mouse_Button_Is_Down(Button As TypeMouseButton) As Boolean
On Local Error Resume Next
Mouse_Button_Is_Down = IMouseState.Buttons(Button)
DIDEV8.GetDeviceStateMouse IMouseState
End Property

Public Property Get Get_Mouse_Position_X() As Long
On Local Error Resume Next
Get_Mouse_Position_X = IMouseState.lX
End Property

Public Property Get Get_Mouse_Position_Y() As Long
On Local Error Resume Next
Get_Mouse_Position_Y = IMouseState.lY
End Property

Public Property Get Get_Mouse_Position_Z() As Long
On Local Error Resume Next
Get_Mouse_Position_Z = IMouseState.lZ
End Property

Public Function Initialize_Mouse_Engine(hWnd As Long) As Boolean
On Error GoTo errOut
ErrStr = "DirectInputCreate"
Set DI8 = XEngine.DirectX.DirectInputCreate()
ErrStr = "CreateDevice"
Set DIDEV8 = DI8.CreateDevice("GUID_SysMouse")
If DIDEV8 Is Nothing Then Exit Function
DIDEV8.SetCommonDataFormat DIFORMAT_MOUSE
ErrStr = "SetCooperativeLevel"
DIDEV8.SetCooperativeLevel hWnd, DISCL_NONEXCLUSIVE Or DISCL_BACKGROUND
ErrStr = "Acquire"
DIDEV8.Acquire
Initialize_Mouse_Engine = True
Exit Function
errOut:
 ErrorToFile "ErrorMOUSE", ErrStr, Err.Description
End Function

Public Sub Cleanup_Mouse_Engine()
On Local Error Resume Next
If HasSurface = True Then Set MouseS = Nothing
If Showing = False Then D3DD.ShowCursor True
If Not DIDEV8 Is Nothing Then
 DIDEV8.Unacquire
End If
Set XEngine = Nothing
Set DIDEV8 = Nothing
Set DI8 = Nothing
End Sub

Public Sub Update_Mouse()
On Local Error Resume Next
DIDEV8.GetDeviceStateMouse IMouseState
End Sub
