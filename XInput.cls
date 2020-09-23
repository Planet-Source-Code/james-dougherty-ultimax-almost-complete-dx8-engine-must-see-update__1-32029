VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "XInput"
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

Public Enum TypeInputKeyX
 X_1 = DIK_1
 X_2 = DIK_2
 X_3 = DIK_3
 X_4 = DIK_4
 X_5 = DIK_5
 X_6 = DIK_6
 X_7 = DIK_7
 X_8 = DIK_8
 X_9 = DIK_9
 X_0 = DIK_0
 X_A = DIK_A
 X_B = DIK_B
 X_C = DIK_C
 X_D = DIK_D
 X_E = DIK_E
 X_F = DIK_F
 X_G = DIK_G
 X_H = DIK_H
 X_I = DIK_I
 X_J = DIK_J
 X_K = DIK_K
 X_L = DIK_L
 X_M = DIK_M
 X_N = DIK_N
 X_O = DIK_O
 X_P = DIK_P
 X_Q = DIK_Q
 X_R = DIK_R
 X_S = DIK_S
 X_T = DIK_T
 X_U = DIK_U
 X_V = DIK_V
 X_W = DIK_W
 X_X = DIK_X
 X_Y = DIK_Y
 X_Z = DIK_Z
 X_F1 = DIK_F1
 X_F2 = DIK_F2
 X_F3 = DIK_F3
 X_F4 = DIK_F4
 X_F5 = DIK_F5
 X_F6 = DIK_F6
 X_F7 = DIK_F7
 X_F8 = DIK_F8
 X_F9 = DIK_F9
 X_F10 = DIK_F10
 X_F11 = DIK_F11
 X_F12 = DIK_F12
 X_Num1 = DIK_NUMPAD1
 X_Num2 = DIK_NUMPAD2
 X_Num3 = DIK_NUMPAD3
 X_Num4 = DIK_NUMPAD4
 X_Num5 = DIK_NUMPAD5
 X_Num6 = DIK_NUMPAD6
 X_Num7 = DIK_NUMPAD7
 X_Num8 = DIK_NUMPAD8
 X_Num9 = DIK_NUMPAD9
 X_Num0 = DIK_NUMPAD0
 X_NumEnter = DIK_NUMPADENTER
 X_Up = DIK_UP
 X_Down = DIK_DOWN
 X_Left = DIK_LEFT
 X_Right = DIK_RIGHT
 X_Escape = DIK_ESCAPE
 X_Enter = DIK_RETURN
 X_LShift = DIK_LSHIFT
 X_RShift = DIK_RSHIFT
 X_LControl = DIK_LCONTROL
 X_RControl = DIK_RCONTROL
 X_Space = DIK_SPACE
 X_Insert = DIK_INSERT
 X_Delete = DIK_DELETE
 X_Home = DIK_HOME
 X_End = DIK_END
 X_PageUp = DIK_PRIOR
 X_PageDown = DIK_NEXT
 X_BackSpace = DIK_BACKSPACE
 X_Add = DIK_ADD
 X_Subtract = DIK_SUBTRACT
 X_Period = DIK_PERIOD
 X_Tab = DIK_TAB
End Enum

Private XEngine As New XEngine3D
Private DI8 As DirectInput8
Private DIDEV8 As DirectInputDevice8
Private IKeyState As DIKEYBOARDSTATE

Public Function XInput8() As DirectInput8
 Set XInput8 = DI8
End Function

Public Function KeyState(Key As TypeInputKeyX) As Boolean
 On Local Error Resume Next
 KeyState = IKeyState.Key(Key)
 DIDEV8.GetDeviceStateKeyboard IKeyState
End Function

Public Function Initialize_Input_Engine(hWnd As Long) As Boolean
 On Local Error GoTo errOut
 ErrStr = ""
 
 ErrStr = "DirectInputCreate"
 Set DI8 = XEngine.DirectX.DirectInputCreate()
        
 ErrStr = "CreateDevice"
 Set DIDEV8 = DI8.CreateDevice("GUID_SysKeyboard")
 DIDEV8.SetCommonDataFormat DIFORMAT_KEYBOARD
 DIDEV8.SetCooperativeLevel hWnd, DISCL_BACKGROUND Or DISCL_NONEXCLUSIVE
 DIDEV8.Acquire
 Initialize_Input_Engine = True
 Exit Function
 
errOut:
 ErrorToFile "ErrorInitDI", ErrStr, PossibleCause
 End
End Function

Public Sub Cleanup_Input_Engine()
 Set XEngine = Nothing
 DIDEV8.Unacquire
 Set DIDEV8 = Nothing
 Set DI8 = Nothing
End Sub
