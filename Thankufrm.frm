VERSION 5.00
Begin VB.Form Thankufrm 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   9465
   ClientLeft      =   2565
   ClientTop       =   705
   ClientWidth     =   15810
   LinkTopic       =   "Form1"
   ScaleHeight     =   9465
   ScaleWidth      =   15810
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer1 
      Interval        =   4500
      Left            =   0
      Top             =   9000
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Please Visit Again ! ! !"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   15.75
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C000&
      Height          =   615
      Left            =   600
      TabIndex        =   1
      Top             =   8760
      Width           =   7335
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Thank You"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   15.75
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C000&
      Height          =   495
      Left            =   3720
      TabIndex        =   0
      Top             =   1680
      Width           =   2775
   End
   Begin VB.Image Image1 
      Height          =   9480
      Left            =   0
      Picture         =   "Thankufrm.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   15840
   End
End
Attribute VB_Name = "Thankufrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Timer1_Timer()
End
End Sub


