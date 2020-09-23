VERSION 5.00
Begin VB.Form frmDemo3 
   BackColor       =   &H00AB663D&
   BorderStyle     =   1  'Fest Einfach
   Caption         =   " Wizzard style demo form build with Art2GUI"
   ClientHeight    =   5160
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8925
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   11.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5160
   ScaleWidth      =   8925
   StartUpPosition =   1  'Fenstermitte
   Begin Demo_ucArt2Gui.ucArt2Gui ucArt2Gui 
      Height          =   870
      Index           =   4
      Left            =   8295
      TabIndex        =   14
      Top             =   15
      Width           =   795
      _ExtentX        =   1402
      _ExtentY        =   1535
      BackColor       =   6697728
      ForeColor       =   16744448
      Design          =   2
      RadGradWidth    =   76
      CenterX         =   124
      CenterY         =   58
      GradientAreas   =   0
   End
   Begin Demo_ucArt2Gui.ucArt2Gui ucArt2Gui 
      Height          =   420
      Index           =   3
      Left            =   8370
      TabIndex        =   13
      Top             =   4530
      Width           =   300
      _ExtentX        =   529
      _ExtentY        =   741
      BackColor       =   11232829
      Design          =   4
      Shape           =   8
      RadGradWidth    =   84
      CenterX         =   6
      GradientAreas   =   0
   End
   Begin Demo_ucArt2Gui.ucArt2Gui ucArt2Gui 
      Height          =   420
      Index           =   2
      Left            =   420
      TabIndex        =   12
      Top             =   4530
      Width           =   300
      _ExtentX        =   529
      _ExtentY        =   741
      BackColor       =   11232829
      Design          =   4
      Shape           =   8
      ShapeLT         =   -1
      ShapeRB         =   -1
      RadGradWidth    =   84
      CenterX         =   6
      GradientAreas   =   0
   End
   Begin VB.Frame Frame 
      BackColor       =   &H00AB663D&
      Caption         =   " Give us some info "
      ForeColor       =   &H00FFFFFF&
      Height          =   360
      Left            =   195
      TabIndex        =   11
      Top             =   3300
      Width           =   2445
   End
   Begin VB.TextBox Text 
      Appearance      =   0  '2D
      BackColor       =   &H00E0E0E0&
      Height          =   390
      Index           =   2
      Left            =   5865
      TabIndex        =   10
      Top             =   2685
      Width           =   2565
   End
   Begin VB.TextBox Text 
      Appearance      =   0  '2D
      BackColor       =   &H00E0E0E0&
      Height          =   390
      Index           =   1
      Left            =   5865
      TabIndex        =   9
      Top             =   2190
      Width           =   2565
   End
   Begin VB.TextBox Text 
      Appearance      =   0  '2D
      BackColor       =   &H00E0E0E0&
      Height          =   390
      Index           =   0
      Left            =   5865
      TabIndex        =   0
      Text            =   " Some user data ..."
      Top             =   1710
      Width           =   2565
   End
   Begin Demo_ucArt2Gui.ucArt2Gui ucArt2Gui 
      Height          =   495
      Index           =   0
      Left            =   120
      TabIndex        =   3
      Top             =   4500
      Width           =   8730
      _ExtentX        =   15399
      _ExtentY        =   873
      BackColor       =   11232829
      Design          =   5
      Shape           =   6
      RadGradWidth    =   76
      CenterX         =   124
      CenterY         =   58
      GradientAreas   =   2
      GradArea-Position1=   0,401
      GradArea-Color2 =   11232829
   End
   Begin Demo_ucArt2Gui.ucArt2Gui ucArt2Gui1 
      Height          =   60
      Index           =   0
      Left            =   15
      TabIndex        =   1
      Top             =   900
      Width           =   8910
      _ExtentX        =   15716
      _ExtentY        =   106
      BackColor       =   11232829
      Design          =   5
      RadGradWidth    =   76
      CenterX         =   124
      CenterY         =   58
      GradientAreas   =   4
      GradArea-Position1=   0,354
      GradArea-Color2 =   9868950
      GradArea-Position2=   0,552
      GradArea-GradGamma2=   0,33
      GradArea-Position3=   0,732
      GradArea-Color4 =   11232829
   End
   Begin Demo_ucArt2Gui.ucArt2Gui ucArt2Gui1 
      Height          =   60
      Index           =   1
      Left            =   0
      TabIndex        =   2
      Top             =   4245
      Width           =   8910
      _ExtentX        =   15716
      _ExtentY        =   106
      BackColor       =   11232829
      Design          =   5
      RadGradWidth    =   76
      CenterX         =   124
      CenterY         =   58
      GradientAreas   =   4
      GradArea-Position1=   0,354
      GradArea-Color2 =   9868950
      GradArea-Position2=   0,552
      GradArea-GradGamma2=   0,33
      GradArea-Position3=   0,732
      GradArea-Color4 =   11232829
   End
   Begin Demo_ucArt2Gui.ucArt2Gui ucArt2Gui 
      Height          =   3165
      Index           =   1
      Left            =   1170
      TabIndex        =   4
      Top             =   1020
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   5583
      BackColor       =   11232829
      ForeColor       =   16777215
      HiliteColor     =   13736273
      RadGradWidth    =   208
      CenterX         =   124
      CenterY         =   58
      GradientAreas   =   0
   End
   Begin Demo_ucArt2Gui.ucArt2Gui ucArt2Gui 
      Height          =   870
      Index           =   5
      Left            =   6690
      TabIndex        =   15
      Top             =   30
      Width           =   660
      _ExtentX        =   1164
      _ExtentY        =   1535
      BackColor       =   6697728
      ForeColor       =   16744448
      RadGradWidth    =   76
      CenterX         =   124
      CenterY         =   58
      GradientAreas   =   0
   End
   Begin VB.Image Image1 
      Height          =   720
      Left            =   315
      Picture         =   "frmDemo3.frx":0000
      Top             =   1335
      Width           =   720
   End
   Begin VB.Label Label 
      Alignment       =   2  'Zentriert
      BackStyle       =   0  'Transparent
      Caption         =   "ART 2 GUI"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   885
      Index           =   1
      Left            =   7560
      TabIndex        =   16
      Top             =   30
      Width           =   630
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Enter rest of data"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      Index           =   2
      Left            =   3510
      TabIndex        =   8
      Top             =   2760
      Width           =   2160
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Enter more data"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      Index           =   1
      Left            =   3510
      TabIndex        =   7
      Top             =   2250
      Width           =   2160
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Enter data"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      Index           =   0
      Left            =   3510
      TabIndex        =   6
      Top             =   1755
      Width           =   2160
   End
   Begin VB.Label Label 
      BackStyle       =   0  'Transparent
      Caption         =   "YOU ARE ON STEP 2"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Index           =   0
      Left            =   405
      TabIndex        =   5
      Top             =   180
      Width           =   4575
   End
   Begin VB.Shape Shape 
      BackColor       =   &H00663300&
      BackStyle       =   1  'Undurchsichtig
      BorderStyle     =   0  'Transparent
      Height          =   900
      Index           =   0
      Left            =   15
      Top             =   15
      Width           =   8925
   End
   Begin VB.Shape Shape 
      BackColor       =   &H00663300&
      BackStyle       =   1  'Undurchsichtig
      BorderStyle     =   0  'Transparent
      Height          =   840
      Index           =   1
      Left            =   30
      Top             =   4320
      Width           =   8925
   End
End
Attribute VB_Name = "frmDemo3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

