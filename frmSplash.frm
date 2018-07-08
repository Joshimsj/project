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
      Left            =   0
      Top             =   8640
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Welcomes You !!!!!"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   36
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   1095
      Left            =   5160
      TabIndex        =   3
      Top             =   6000
      Width           =   7215
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Auto Expo"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   36
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   1935
      Left            =   7320
      TabIndex        =   2
      Top             =   4200
      Width           =   2895
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "The Best"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   36
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   1095
      Left            =   6600
      TabIndex        =   1
      Top             =   3360
      Width           =   3975
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "The Best Of "
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   36
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   975
      Left            =   6000
      TabIndex        =   0
      Top             =   2400
      Width           =   4935
   End
   Begin VB.Image SplashImg 
      Height          =   9105
      Left            =   -120
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
