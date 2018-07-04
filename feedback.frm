VERSION 5.00
Begin VB.Form feedbackfrm 
   BorderStyle     =   0  'None
   Caption         =   "Feedbackfrm"
   ClientHeight    =   10050
   ClientLeft      =   2160
   ClientTop       =   360
   ClientWidth     =   15930
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   10050
   ScaleWidth      =   15930
   ShowInTaskbar   =   0   'False
   Begin VB.ComboBox Combo2 
      Height          =   315
      Left            =   11760
      TabIndex        =   27
      Text            =   "Combo2"
      Top             =   3120
      Width           =   3255
   End
   Begin VB.ComboBox Agegroup 
      Height          =   315
      Left            =   6960
      TabIndex        =   26
      Text            =   "Combo1"
      Top             =   3000
      Width           =   2415
   End
   Begin VB.CommandButton CmdAdd 
      Caption         =   "Add"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   14.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   11880
      TabIndex        =   25
      Top             =   9480
      Width           =   1095
   End
   Begin VB.CommandButton CmdSubmit 
      Caption         =   "Submit"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   14.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   13200
      TabIndex        =   24
      Top             =   9480
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Close"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   14.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   14640
      TabIndex        =   23
      Top             =   9480
      Width           =   975
   End
   Begin VB.TextBox TxtQ7 
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
      Left            =   9360
      TabIndex        =   22
      Text            =   "Text7"
      Top             =   9360
      Width           =   2415
   End
   Begin VB.TextBox TxtQ6 
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
      Left            =   9360
      TabIndex        =   21
      Text            =   "Text6"
      Top             =   8760
      Width           =   6255
   End
   Begin VB.TextBox TxtQ5 
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
      Left            =   9360
      TabIndex        =   20
      Text            =   "Text5"
      Top             =   8160
      Width           =   6255
   End
   Begin VB.TextBox TxtQ4 
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
      Left            =   9360
      TabIndex        =   19
      Text            =   "Text4"
      Top             =   7560
      Width           =   6255
   End
   Begin VB.TextBox TxtQ3 
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
      Left            =   9360
      TabIndex        =   18
      Text            =   "Text3"
      Top             =   6960
      Width           =   6255
   End
   Begin VB.TextBox TxtQ2 
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
      Left            =   9360
      TabIndex        =   17
      Text            =   "Text2"
      Top             =   6360
      Width           =   6255
   End
   Begin VB.TextBox TxtQ1 
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
      Left            =   9360
      TabIndex        =   16
      Text            =   "Text1"
      Top             =   5760
      Width           =   6255
   End
   Begin VB.OptionButton Female 
      Caption         =   "Female"
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
      Left            =   2760
      TabIndex        =   5
      Top             =   3360
      Width           =   1695
   End
   Begin VB.OptionButton Male 
      Caption         =   "Male"
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
      Left            =   2760
      TabIndex        =   4
      Top             =   2760
      Width           =   1695
   End
   Begin VB.Label Label14 
      BackStyle       =   0  'Transparent
      Caption         =   "1-The manner Which your Greeted "
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   14.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   495
      Left            =   240
      TabIndex        =   15
      Top             =   5880
      Width           =   9015
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "5-Would you like to share our Experience with other fellows"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   14.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   495
      Left            =   240
      TabIndex        =   14
      Top             =   8160
      Width           =   9015
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "7-Did you like our Services throughtout the process"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   14.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   375
      Left            =   240
      TabIndex        =   13
      Top             =   9360
      Width           =   9015
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "6-Did our Sales Person  looked Well for you"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   14.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   495
      Left            =   240
      TabIndex        =   12
      Top             =   8760
      Width           =   9015
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "4-The Vehicle Price/Payment were Discussed Properly "
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   14.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   495
      Left            =   240
      TabIndex        =   11
      Top             =   7560
      Width           =   9015
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "3-Fullfield all Commitment  made by you"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   14.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   495
      Left            =   240
      TabIndex        =   10
      Top             =   6960
      Width           =   9015
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "2-Consideration of your time "
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   14.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   375
      Left            =   240
      TabIndex        =   9
      Top             =   6480
      Width           =   9015
   End
   Begin VB.Label Label7 
      BackColor       =   &H000080FF&
      Caption         =   "Please rate our Support Team....!!!!!!"
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
      Left            =   120
      TabIndex        =   8
      Top             =   5160
      Width           =   15615
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   " Level:-"
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
      Left            =   9600
      TabIndex        =   7
      Top             =   2760
      Width           =   2295
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Age Group:-"
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
      Left            =   5040
      TabIndex        =   6
      Top             =   2760
      Width           =   2055
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Gender:-"
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
      Left            =   720
      TabIndex        =   3
      Top             =   2760
      Width           =   1575
   End
   Begin VB.Label Label3 
      BackColor       =   &H000080FF&
      Caption         =   "Please tell a bit about  you......!!!!!"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   14.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   2040
      Width           =   15495
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   $"feedback.frx":0000
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   14.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1095
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   15495
   End
   Begin VB.Label Label1 
      BackColor       =   &H000080FF&
      Caption         =   "Auto Expo Satisfaction Survey"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   14.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4935
   End
   Begin VB.Image Image1 
      Height          =   10035
      Left            =   0
      Picture         =   "feedback.frx":0103
      Stretch         =   -1  'True
      Top             =   0
      Width           =   15885
   End
End
Attribute VB_Name = "feedbackfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim c As New ADODB.Connection
Dim r As New ADODB.Recordset
Dim s As String
Dim Gender As String



Private Sub CmdAdd_Click()
Option1.Value = False
Option2.Value = False
Option3.Value = False
Option4.Value = False
Option5.Value = False
Option6.Value = False
Option7.Value = False
Option8.Value = False
Option9.Value = False
Option10.Value = False
Option11.Value = False
TxtQ1.Text = ""
TxtQ2.Text = ""
TxtQ3.Text = ""
TxtQ4.Text = ""
TxtQ5.Text = ""
TxtQ6.Text = ""
TxtQ7.Text = ""
End Sub

Private Sub CmdSubmit_Click()
MsgBox (Gender)
End Sub

Private Sub Command1_Click()
MsgBox "Thank you for Visting"
End
End Sub

Private Sub Female_Click()
Gender = "Female"
End Sub

Private Sub Male_Click()
Gender = "Male"
End Sub

