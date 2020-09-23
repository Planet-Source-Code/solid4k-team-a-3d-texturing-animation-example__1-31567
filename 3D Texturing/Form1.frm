VERSION 5.00
Object = "{08216199-47EA-11D3-9479-00AA006C473C}#2.1#0"; "RMCONTROL.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Texture Test Thing..."
   ClientHeight    =   6495
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7230
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6495
   ScaleWidth      =   7230
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   615
      Left            =   120
      TabIndex        =   1
      Top             =   0
      Width           =   6975
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "Form1.frx":0000
         Left            =   2160
         List            =   "Form1.frx":0010
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   210
         Width           =   975
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Start"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "Size:"
         Height          =   255
         Left            =   1680
         TabIndex        =   4
         Top             =   240
         Width           =   375
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   1800
      Top             =   1920
   End
   Begin RMControl7.RMCanvas RM 
      Height          =   5655
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   6975
      _ExtentX        =   12303
      _ExtentY        =   9975
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim MS_Sphere As Direct3DRMMeshBuilder3 ' The Mesh (body) of our sphere
Dim FR_Sphere As Direct3DRMFrame3           ' The Frame (layer) for our sphere

' Note: A Mesh is the 3D of our object, like skin over our body

Dim Sphere_Tex As Direct3DRMTexture3     ' This is that metal texture's source
Dim LT_Light As Direct3DRMLight                 ' Helps us see the sphere

Const Sin5 = 8.715574E-02!
Const Cos5 = 0.9961947!
' I'll try to explain the above Const's;
' They are used as long, but exact numbers for degrees
' In our case, it will turn us 5 degrees

Public Sub DX_Init()    ' Initiate (setup) our scene
With RM
    .StartWindowed      ' Tells RMControl we're starting our "movie"
    .SceneFrame.SetSceneBackgroundRGB 0, 0, 0  ' The black bg color
    .Viewport.SetBack 400   ' How far we can see back without objects disappearing
    .CameraFrame.SetPosition Nothing, 0, 0, 0   ' How we see the world
End With
Set FR_Sphere = RM.D3DRM.CreateFrame(RM.SceneFrame) ' Create a frame
Set MS_Sphere = RM.D3DRM.CreateMeshBuilder()                ' Create a Mesh
End Sub
Public Sub DX_MakeObjects() ' Makes our 3D Stuff
Set Sphere_Tex = RM.D3DRM.LoadTexture(App.Path & "\texture.bmp") ' Load Texture
Set LT_Light = RM.D3DRM.CreateLightRGB(D3DRMLIGHT_AMBIENT, 1, 1, 1) ' Create a light, and its color
MS_Sphere.LoadFromFile App.Path & "\SPHERE.X", 0, 0, Nothing, Nothing ' Loads the mesh file for a sphere
MS_Sphere.SetTexture Sphere_Tex ' Wraps our texture around the sphere
MS_Sphere.Translate 0, 0, 0             ' I dunno
MS_Sphere.ScaleMesh 2, 2, 2           ' This makes our object 2 times bigger than it was
FR_Sphere.SetPosition Nothing, 0, 0, 10 ' Puts our 3D Sphere 10 "feet" away and in the center.
'Note: 0 is the center, "-number" is left, and "number" is right
FR_Sphere.AddVisual MS_Sphere ' This makes the 3D Sphere with its texture visible

FR_Sphere.AddLight LT_Light ' This adds the light to the sphere
End Sub

Private Sub Combo1_Click()
' This is pretty self expainitory
If Combo1.ListIndex = 0 Then
    MS_Sphere.ScaleMesh 1, 1, 1
ElseIf Combo1.ListIndex = 1 Then
    MS_Sphere.ScaleMesh 2, 2, 2
ElseIf Combo1.ListIndex = 2 Then
    MS_Sphere.ScaleMesh 3, 3, 3
ElseIf Combo1.ListIndex = 3 Then
    MS_Sphere.ScaleMesh 4, 4, 4
Else
    MS_Sphere.ScaleMesh 5, 5, 5
End If
DX_Init
DX_MakeObjects
End Sub

Private Sub Command1_Click()
RM.SetFocus
Select Case Command1.Caption
Case "Start"
    Command1.Caption = "Stop"
    FR_Sphere.SetRotation Nothing, 0, 1, 0, 0.1 ' This will put a spin on the ball and keep it spinning
Case "Stop"
    Command1.Caption = "Start"
    FR_Sphere.SetRotation Nothing, 0, 0, 0, 0   ' Stops the spinning
End Select
End Sub

Private Sub Command2_Click()

End Sub

Private Sub Form_Load()
DX_Init
DX_MakeObjects
End Sub

Private Sub RM_KeyDown(keyCode As Integer, Shift As Integer)
If keyCode = vbKeyLeft Then
    FR_Sphere.SetOrientation FR_Sphere, -Sin5, 0, Cos5, 0, 1, 0 ' Rotates the ball left
End If
If keyCode = vbKeyRight Then
    FR_Sphere.SetOrientation FR_Sphere, Sin5, 0, Cos5, 0, 1, 0 ' Rotates the ball right
End If
If keyCode = vbKeyUp Then
    FR_Sphere.SetOrientation FR_Sphere, 0, -Sin5, Cos5, 0, Cos5, 0 ' Rotates the ball up
End If
If keyCode = vbKeyDown Then
    FR_Sphere.SetOrientation FR_Sphere, 0, Sin5, Cos5, 0, Cos5, 0 ' Rotates the ball down
End If
End Sub

Private Sub Timer1_Timer()
RM.Update ' Keeps our 3D Scene updated
End Sub
