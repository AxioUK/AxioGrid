VERSION 5.00
Begin VB.Form AxGridAbout 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   1320
   ClientLeft      =   3615
   ClientTop       =   3690
   ClientWidth     =   4530
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1320
   ScaleWidth      =   4530
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox picLogo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   3540
      Picture         =   "axGridAbout.frx":0000
      ScaleHeight     =   240
      ScaleWidth      =   870
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   1005
      Width           =   870
   End
   Begin VB.CommandButton OKButton 
      Appearance      =   0  'Flat
      Caption         =   "OK"
      Height          =   300
      Left            =   135
      TabIndex        =   0
      Top             =   960
      Width           =   720
   End
   Begin VB.Image Image2 
      Height          =   480
      Left            =   1170
      Picture         =   "axGridAbout.frx":0B42
      Stretch         =   -1  'True
      Top             =   750
      Visible         =   0   'False
      Width           =   510
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   90
      Picture         =   "axGridAbout.frx":1074
      Stretch         =   -1  'True
      Top             =   90
      Width           =   510
   End
   Begin VB.Label lblV2 
      BackStyle       =   0  'Transparent
      Caption         =   ".0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   240
      Left            =   4080
      TabIndex        =   7
      Top             =   450
      Width           =   225
   End
   Begin VB.Label lblV1 
      BackStyle       =   0  'Transparent
      Caption         =   "2"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   36
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   945
      Left            =   3555
      TabIndex        =   5
      Top             =   -15
      Width           =   600
   End
   Begin VB.Label lblProd 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "AxGrid "
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   525
      Left            =   795
      TabIndex        =   2
      Top             =   105
      Width           =   1800
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Coded by:"
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   9
         Charset         =   255
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   180
      Index           =   1
      Left            =   2640
      TabIndex        =   4
      Top             =   1050
      Width           =   825
   End
   Begin VB.Label lblVersion 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "v0.00"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   165
      Left            =   2445
      TabIndex        =   3
      Top             =   390
      Width           =   375
   End
   Begin VB.Label lblInfo 
      BackStyle       =   0  'Transparent
      Caption         =   "Another Xtended Grid with ADO Features"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   330
      Left            =   75
      TabIndex        =   1
      Top             =   630
      Width           =   3345
      WordWrap        =   -1  'True
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H00FF0000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00C0FFC0&
      BorderStyle     =   6  'Inside Solid
      BorderWidth     =   2
      FillColor       =   &H00004000&
      Height          =   735
      Left            =   3540
      Top             =   150
      Width           =   870
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00FFC0C0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFC0C0&
      BorderStyle     =   6  'Inside Solid
      Height          =   450
      Left            =   30
      Top             =   30
      Width           =   4470
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H0080C0FF&
      BackStyle       =   1  'Opaque
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   1320
      Left            =   0
      Top             =   0
      Width           =   4530
   End
End
Attribute VB_Name = "AxGridAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Resize()
lblVersion.Left = (lblProd.Left + lblProd.Width) - 150
End Sub

Private Sub OKButton_Click()
   Unload Me
End Sub
