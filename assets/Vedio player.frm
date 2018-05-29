VERSION 5.00
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmUtama 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "RR media player"
   ClientHeight    =   6930
   ClientLeft      =   150
   ClientTop       =   780
   ClientWidth     =   9660
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6930
   ScaleWidth      =   9660
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.StatusBar stbar 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   3
      Top             =   6555
      Width           =   9660
      _ExtentX        =   17039
      _ExtentY        =   661
      Style           =   1
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog cmdlg 
      Left            =   8640
      Top             =   4800
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Filter          =   "MP4 Video (*.mp4)|*.MP4"
   End
   Begin VB.ListBox lstSumber 
      Height          =   840
      Left            =   9720
      TabIndex        =   2
      Top             =   5760
      Width           =   1815
   End
   Begin VB.ListBox lstJudul 
      Height          =   5715
      Left            =   9720
      TabIndex        =   1
      Top             =   0
      Width           =   1815
   End
   Begin WMPLibCtl.WindowsMediaPlayer wmp 
      Height          =   6615
      Left            =   -120
      TabIndex        =   0
      Top             =   -120
      Width           =   9735
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
      _cx             =   17171
      _cy             =   11668
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuOpen 
         Caption         =   "&Open"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "&Exit"
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "&View"
      Begin VB.Menu mnuPlaylist 
         Caption         =   "&Playlist"
      End
   End
   Begin VB.Menu mnuAdd 
      Caption         =   "&Add to Playlist"
   End
End
Attribute VB_Name = "FrmUtama"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Me.Width = 9750
End Sub

Private Sub mnuAdd_Click()
lstSumber.AddItem cmdlg.FileName
lstJudul.AddItem cmdlg.FileTitle

End Sub

Private Sub mnuExit_Click()
End
End Sub

Private Sub mnuOpen_Click()
cmdlg.ShowOpen
If cmdlg.FileName = "" Then
' nothing happend
Else
wmp.URL = cmdlg.FileName
stbar.SimpleText = "Now Playing-" & cmdlg.FileTitle
End If
End Sub

Private Sub mnuPlaylist_Click()
If Me.Width = 9750 Then
Me.Width = 11640
Else
Me.Width = 9750
End If
End Sub
