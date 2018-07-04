VERSION 5.00
Begin VB.Form frmWelcome 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   Caption         =   "Welcome Screen"
   ClientHeight    =   9105
   ClientLeft      =   2340
   ClientTop       =   1080
   ClientWidth     =   15780
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   DrawStyle       =   1  'Dash
   Icon            =   "frmSplash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9105
   ScaleWidth      =   15780
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer1 
      Interval        =   5000
      Left            =   480
      Top             =   360
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   """The Best Of the Best"""
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   36
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   975
      Left            =   120
      TabIndex        =   1
      Top             =   8040
      Width           =   8655
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   """Welcome to Auto Expo"""
      BeginProperty Font 
         Name            =   "Brush Script MT"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   1335
      Left            =   5280
      TabIndex        =   0
      Top             =   2160
      Width           =   9855
   End
   Begin VB.Image SplashImg 
      Height          =   9105
      Left            =   0
      Picture         =   "frmSplash.frx":000C
      Stretch         =   -1  'True
      Top             =   0
      Width           =   15855
   End
End
Attribute VB_Name = "frmWelcome"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub Form_Load()
    SplashImg.Width = Me.Width
    SplashImg.Height = Me.Height
End Sub

Private Sub Timer1_Timer()
 Unload Me
    Load login
    login.Show
End Sub
