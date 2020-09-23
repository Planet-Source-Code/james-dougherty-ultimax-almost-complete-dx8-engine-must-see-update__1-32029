VERSION 5.00
Begin VB.Form frmSplash 
   BorderStyle     =   0  'None
   ClientHeight    =   7230
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8835
   LinkTopic       =   "Form1"
   Picture         =   "frmSplash.frx":0000
   ScaleHeight     =   482
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   589
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "(C) 2001-2002 James E. Dougherty"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   300
      Left            =   0
      TabIndex        =   1
      Top             =   6900
      Width           =   4245
   End
   Begin VB.Label lblStatus 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Loading"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   1320
      TabIndex        =   0
      Top             =   240
      Width           =   975
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'BETA NOT FINISHED... STILL IN DEVELOPEMENT...
'NOT ALL FUNCTIONS WORK "RIGHT" YET(NOT THAT MANY THOUGH)

'PLEASE NOTE:
' THIS IS NOT ALL MY CODE... SO IT SHOULD BE TO THEIR RESPECTED
' OWNER... I KNOW 1 IS SIMON.

'FOR THAT SIMPLE FACT THIS IS FREEWARE... USE AND ABUSE IT...

Public Sub Initialize_Setup()
Dim i As Integer
frmMain.Show
DoEvents
frmSplash.Show
DoEvents

lblStatus.Caption = "Initializing EngineX8"
OK = XE.Initialize_EngineX8(frmMain.hWnd, False)
If Not (OK) Then End
lblStatus.Caption = "Initializing Input Engine"
OK = XI.Initialize_Input_Engine(frmMain.hWnd)
If Not (OK) Then End
lblStatus.Caption = "Initializing 2D Sound Engine"
OK = XS2D.Initialize_Sound_Engine(frmMain.hWnd)
If Not (OK) Then End
lblStatus.Caption = "Initializing 3D Sound Engine"
XS3D.Initialize_3D_Sound_Engine frmMain.hWnd
lblStatus.Caption = "Initializing Mouse Engine"
XM.Initialize_Mouse_Engine frmMain.hWnd
lblStatus.Caption = "Initializing Camera Engine"
XC.Initialize_Camera_Engine

XMidi.Initialize_Midi_Engine frmMain.hWnd
XMidi.SetMidiDir App.Path & "\Sounds"
XMidi.Load_Midi "canyon.mid"
XMidi.Set_Audio_Effect Chorus
XMidi.VolumeLevel 80
XMidi.Play_Midi False

lblStatus.Caption = "Initializing Filter"
XF.Set_Anisotropic_Filter
DoEvents

lblStatus.Caption = "Creating Ship"
XX.Set_Directory_For_X_Files App.Path & "\Objects"
XX.Set_Texture_Directory_For_X_Files App.Path & "\Textures"
XX.Load_X_File "FighterShip.x"
XX.Position_Mesh 10, 7, 0
XX.Scale_Mesh 0.01, 0.01, 0.01
XX.Enable_Object_Transparency True
XX.Enable_Glass_Effect True

lblStatus.Caption = "Initializing Enviroment Sphere"
XEN.Initialize_Enviroment_Sphere App.Path & "\Objects\SkySphere.x", App.Path & "\Textures\"
lblStatus.Caption = "Initializing Terrain"
XEN.Create_Terrain App.Path & "\Textures\CoolGrass.bmp", 1000, 1000, 4
lblStatus.Caption = "Creating Water System"
XEN.Create_Water App.Path & "\Textures\Water3.bmp", 1000, 1000, 4, False
XEN.Enable_Fog False, 1, 100, &HE0E0E0

lblStatus.Caption = "Initializing Object Engine"
XO.Initialize_Objects_Engine Num_Objects
lblStatus.Caption = "Creating Box"
XO.Create_3D_Wall 0, 15, 15, 15, App.Path & "\Textures\tree35S.tga"
lblStatus.Caption = "Seting Box's Position"
XO.Set_Object_Position 0, 0, 10, 50
XO.Scale_Object 0, 0.5, 0.5, 0.5
lblStatus.Caption = "Initializing Enviroment Mapping"
XO.Apply_Object_Filter 0, Dark_Map, App.Path & "\Textures\Particle.bmp"
XO.Enable_Object_Transparency 0, True
XO.Enable_Glass_Effect 0, True
XO.Set_Object_Name 0, "Transparent Cube"

lblStatus.Caption = "Creating Billboard System"
XO.Create_Billboard 1, 5, 5, App.Path & "\Textures\tree35S.tga"
lblStatus.Caption = "Seting Billboards Position"
XO.Set_Billboard_Position 1, 0, 8, 0

lblStatus.Caption = "Initializing Text"
XE.Initialize_Text , 10, True

lblStatus.Caption = "Setting 2D Sound Directory"
XS2D.SetSoundDir App.Path & "\Sounds"
lblStatus.Caption = "Creating 2D Buffers"
XS2D.CreateBuffers Num_Buffers, "Foot2.wav"
lblStatus.Caption = "Loading 2D Sound0"
XS2D.LoadSound 0, "Foot2.wav"
DoEvents

XD.Create_Decal App.Path & "\Textures\Logo.bmp", 2, 2
XD.Position_Decal 0, 9, 10
XD.Rotate_Decal 0, 2, 0
XD.Enable_Transparency True

lblStatus.Caption = "Setting Render State"
D3DD.SetRenderState D3DRS_LIGHTING, 0
EndLoop = False
lblStatus.Caption = "Done Loading..."
DoEvents
Unload Me
Render
End Sub

Private Sub Form_Load()
Initialize_Setup
End Sub

