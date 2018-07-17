VERSION 5.00
Begin VB.Form MainLogin 
   BorderStyle     =   0  'None
   Caption         =   "Login "
   ClientHeight    =   5970
   ClientLeft      =   5205
   ClientTop       =   2340
   ClientWidth     =   9135
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5970
   ScaleWidth      =   9135
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdLogin 
      Appearance      =   0  'Flat
      BackColor       =   &H000080FF&
      Caption         =   "Login"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   18
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1080
      MaskColor       =   &H000080FF&
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   4080
      UseMaskColor    =   -1  'True
      Width           =   2655
   End
   Begin VB.TextBox TxtPass 
      BackColor       =   &H00FF8080&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   14.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      IMEMode         =   3  'DISABLE
      Left            =   960
      TabIndex        =   3
      Text            =   "Password"
      Top             =   3240
      Width           =   2895
   End
   Begin VB.TextBox TxtUser 
      BackColor       =   &H00FF8080&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   14.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   960
      TabIndex        =   1
      Text            =   "Username"
      Top             =   2520
      Width           =   2895
   End
   Begin VB.Image Image2 
      Height          =   480
      Left            =   360
      Picture         =   "Login.frx":0000
      Stretch         =   -1  'True
      Top             =   3240
      Width           =   480
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Login "
      BeginProperty Font 
         Name            =   "Mistral"
         Size            =   72
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   1815
      Left            =   480
      TabIndex        =   0
      Top             =   0
      Width           =   4095
   End
   Begin VB.Image Image1 
      Height          =   7200
      Left            =   0
      Picture         =   "Login.frx":1A5B
      Top             =   -240
      Width           =   12780
   End
End
Attribute VB_Name = "MainLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public conn As ADODB.Connection


Private Sub cmdlogin_Click()
If TxtUser.Text = "" Then
MsgBox "Username is Empty.", vbInformation
TxtUser.SetFocus
Exit Sub
ElseIf TxtPass.Text = "" Then
MsgBox "Password is Empty"
TxtPass.SetFocus
Exit Sub
Else
Call login
End If
End Sub

Private Sub login()
Dim rs As New ADODB.Recordset
rs.Open "SELECT password FROM Logintab WHERE Username = '" & TxtUser.Text & "'", conn, adOpenStatic, adLockReadOnly
If rs.RecordCount < 1 Then
    MsgBox "Username is Invalid. Please try again.", vbInformation
    TxtUser.SetFocus
Exit Sub
Else
    If TxtPass.Text = rs!Password Then
        Unload Me
        Load MainForm
        MainForm.Show
    Exit Sub
    Else
        MsgBox "Password is Invalid. Please try again.", vbInformation
        TxtPass.SetFocus
    Exit Sub
    End If
End If
Set rs = Nothing
End Sub


Private Sub Form_Load()
    Set conn = New ADODB.Connection
    conn.Open "Provider=Microsoft.jet.OLEDB.4.0;Data Source =E:\project\assets\Logindb.mdb;persist security info=false"
End Sub

Private Sub TxtPass_Change()
    TxtPass.PasswordChar = "*"
End Sub

Private Sub TxtPass_Click()
    TxtPass.Text = ""
End Sub


Private Sub TxtPass_LostFocus()
    If TxtPass.Text = "" Then
        TxtPass.Text = "Password"
        TxtPass.PasswordChar = ""
    End If
End Sub
