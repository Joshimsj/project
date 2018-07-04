VERSION 5.00
Begin VB.Form feedbackfrm 
   BorderStyle     =   0  'None
   Caption         =   "Feedbackfrm"
   ClientHeight    =   10350
   ClientLeft      =   2160
   ClientTop       =   360
   ClientWidth     =   15915
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   10350
   ScaleWidth      =   15915
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox TxtQ8 
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
      Index           =   1
      Left            =   9360
      TabIndex        =   29
      Text            =   "Q8"
      Top             =   9000
      Width           =   6255
   End
   Begin VB.TextBox TxtSex 
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
      Left            =   6120
      TabIndex        =   27
      Text            =   "Enter Sex"
      Top             =   3360
      Width           =   1695
   End
   Begin VB.TextBox TxtName 
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   14.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   2040
      TabIndex        =   26
      Text            =   "Enter Name"
      Top             =   3360
      Width           =   1935
   End
   Begin VB.TextBox TxtIncome 
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
      Left            =   13320
      TabIndex        =   24
      Text            =   "Income"
      Top             =   3360
      Width           =   2175
   End
   Begin VB.TextBox TxtAge 
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
      TabIndex        =   23
      Text            =   "Enter Age"
      Top             =   3360
      Width           =   1695
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
      Left            =   11280
      TabIndex        =   22
      Top             =   9720
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
      Left            =   12720
      TabIndex        =   21
      Top             =   9720
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
      Left            =   14400
      TabIndex        =   20
      Top             =   9720
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
      Index           =   0
      Left            =   9360
      TabIndex        =   19
      Text            =   "Q7"
      Top             =   8400
      Width           =   6255
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
      TabIndex        =   18
      Text            =   "Q6"
      Top             =   7800
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
      TabIndex        =   17
      Text            =   "Q5"
      Top             =   7200
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
      TabIndex        =   16
      Text            =   "Q4"
      Top             =   6600
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
      TabIndex        =   15
      Text            =   "Q3"
      Top             =   6000
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
      TabIndex        =   14
      Text            =   "Q2"
      Top             =   5400
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
      TabIndex        =   13
      Text            =   "Q1"
      Top             =   4800
      Width           =   6255
   End
   Begin VB.Label Label14 
      BackStyle       =   0  'Transparent
      Caption         =   "8-Any Improvement needed in our service"
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
      Index           =   1
      Left            =   240
      TabIndex        =   30
      Top             =   9000
      Width           =   9015
   End
   Begin VB.Label Sex 
      BackStyle       =   0  'Transparent
      Caption         =   "Sex:- "
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
      Left            =   4440
      TabIndex        =   28
      Top             =   2760
      Width           =   1815
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Name:-"
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
      Left            =   480
      TabIndex        =   25
      Top             =   2760
      Width           =   1575
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
      Index           =   0
      Left            =   240
      TabIndex        =   12
      Top             =   4800
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
      TabIndex        =   11
      Top             =   7200
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
      Index           =   0
      Left            =   240
      TabIndex        =   10
      Top             =   8400
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
      TabIndex        =   9
      Top             =   7800
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
      TabIndex        =   8
      Top             =   6600
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
      TabIndex        =   7
      Top             =   6000
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
      TabIndex        =   6
      Top             =   5400
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
      TabIndex        =   5
      Top             =   4080
      Width           =   15615
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   " Annual Income:-"
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
      Left            =   11520
      TabIndex        =   4
      Top             =   2760
      Width           =   2535
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
      Left            =   8040
      TabIndex        =   3
      Top             =   2760
      Width           =   2055
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
      Height          =   10395
      Left            =   -120
      Picture         =   "feedback.frx":0103
      Stretch         =   -1  'True
      Top             =   0
      Width           =   16005
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
'Dim Gender As String

Private Sub CmdAdd_Click()
TxtName.Text = ""
TxtSex.Text = ""
TxtAge.Text = ""
TxtIncome.Text = ""
TxtQ1.Text = ""
TxtQ2.Text = ""
TxtQ3.Text = ""
TxtQ4.Text = ""
TxtQ5.Text = ""
TxtQ6.Text = ""
TxtQ7.Text = ""
TxtQ8.Text = ""
End Sub

Private Sub CmdSubmit_Click()
'MsgBox (Gender)
Dim fee As New Feedback
fee.Name = TxtName.Text
fee.Sex = TxtSex.Text
fee.Age = TxtAge.Text
fee.Income = TxtIncome.Text
fee.Q1 = TxtQ1.Text
fee.Q2 = TxtQ2.Text
fee.Q3 = TxtQ3.Text
fee.Q4 = TxtQ4.Text
fee.Q5 = TxtQ5.Text
fee.Q6 = TxtQ6.Text
fee.Q7 = TxtQ7.Text
fee.Q8 = TxtQ8.Text
Call fee.SaveData
End Sub

Private Sub Command1_Click()
MsgBox "Thank you for Visting"
End
End Sub

'Private Sub Female_Click()
'Gender = "Female"
'End Sub

'Private Sub Male_Click()
'Gender = "Male"
'End Sub

