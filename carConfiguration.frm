VERSION 5.00
Begin VB.Form carConfiguration 
   BackColor       =   &H8000000B&
   Caption         =   "Car Configuration"
   ClientHeight    =   10410
   ClientLeft      =   -240
   ClientTop       =   630
   ClientWidth     =   20250
   LinkTopic       =   "Form1"
   ScaleHeight     =   10410
   ScaleWidth      =   20250
   Begin VB.Timer tim 
      Interval        =   500
      Left            =   100
      Top             =   120
   End
   Begin VB.Shape c1 
      FillStyle       =   0  'Solid
      Height          =   495
      Left            =   3250
      Shape           =   3  'Circle
      Top             =   4750
      Width           =   495
   End
   Begin VB.Shape c3 
      BackColor       =   &H000000FF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H000000FF&
      Height          =   500
      Left            =   5750
      Shape           =   3  'Circle
      Top             =   4750
      Width           =   500
   End
   Begin VB.Label configOutput 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Configuration Output"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   5265
      TabIndex        =   0
      Top             =   840
      Width           =   1485
   End
   Begin VB.Shape cmain 
      BorderStyle     =   2  'Dash
      BorderWidth     =   2
      DrawMode        =   1  'Blackness
      Height          =   5000
      Left            =   3500
      Shape           =   3  'Circle
      Top             =   2500
      Width           =   5000
   End
End
Attribute VB_Name = "carConfiguration"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim centre(1) As Long
Dim radius As Long
Dim theta As Integer


Private Sub Form_Load()
    centre(1) = cmain.Top + (cmain.Height / 2)
    centre(0) = cmain.Left + (cmain.Width / 2)
    radius = cmain.Width / 2
    
    If tim.Enabled = False Then
    tim.Enabled = True
    End If
    
    theta = 0
    c1.Left = centre(0) + radius * Cos(theta) - c1.Height / 2
    c1.Top = centre(1) + radius * Sin(theta) - c1.Height / 2
    
    
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If X >= c1.Left And X <= c1.Left + c1.Width And Y >= c1.Top And Y <= c1.Top + c1.Height Then
        'xShift = c1.Left + 3000
        'Y = ((radius ^ 2) - (xShift - centre(0)) ^ 2) ^ (1 / 2) + centre(1)
        'c1.Left = xShift
        'c1.Top = Y
        'distance = 100
        'theta = distance / radius '2 pi (radian)
        
        theta = 0.6
        c1.Left = centre(0) + radius * Cos(theta) - c1.Height / 2
        c1.Top = centre(1) + radius * Sin(theta) - c1.Height / 2
    End If
End Sub

Private Sub tim_Timer()
    If theta > 2 * 22 / 7 Then
    theta = 0.6
    Else
    theta = theta + 0.6
    End If
    
    c1.Left = centre(0) + radius * Cos(theta) - c1.Height / 2
    c1.Top = centre(1) + radius * Sin(theta) - c1.Height / 2
End Sub
