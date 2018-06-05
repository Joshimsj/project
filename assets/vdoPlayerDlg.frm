VERSION 5.00
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Begin VB.Form vdoPlayerDlg 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Viedeo Name"
   ClientHeight    =   7785
   ClientLeft      =   2115
   ClientTop       =   1485
   ClientWidth     =   16170
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7785
   ScaleWidth      =   16170
   ShowInTaskbar   =   0   'False
   Begin WMPLibCtl.WindowsMediaPlayer player 
      Height          =   7815
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   16215
      URL             =   ""
      rate            =   1
      balance         =   0
      currentPosition =   0
      defaultFrame    =   ""
      playCount       =   1
      autoStart       =   -1  'True
      currentMarker   =   0
      invokeURLs      =   -1  'True
      baseURL         =   ""
      volume          =   50
      mute            =   0   'False
      uiMode          =   "full"
      stretchToFit    =   0   'False
      windowlessVideo =   0   'False
      enabled         =   -1  'True
      enableContextMenu=   -1  'True
      fullScreen      =   0   'False
      SAMIStyle       =   ""
      SAMILang        =   ""
      SAMIFilename    =   ""
      captioningID    =   ""
      enableErrorDialogs=   0   'False
      _cx             =   28601
      _cy             =   13785
   End
End
Attribute VB_Name = "vdoPlayerDlg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    player.uiMode = "none"
    player.Top = 0
    player.Left = 0
    player.Height = Me.Height
    player.Width = Me.Width
End Sub

Public Sub Play_Video(ByVal path As String, Optional ByVal name As String = "VDO")
    player.url = path
    Me.Caption = name
End Sub
