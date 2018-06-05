VERSION 5.00
Begin VB.Form frmSplash 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   ClientHeight    =   9105
   ClientLeft      =   2385
   ClientTop       =   630
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
      Interval        =   4500
      Left            =   480
      Top             =   360
   End
   Begin VB.Image SplashImg 
      Height          =   20745
      Left            =   0
      Picture         =   "frmSplash.frx":000C
      Stretch         =   -1  'True
      Top             =   0
      Width           =   33870
   End
End
Attribute VB_Name = "frmSplash"
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
    Load MainForm
    MainForm.Show
End Sub
