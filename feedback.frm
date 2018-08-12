VERSION 5.00
Begin VB.Form feedbackfrm 
   BorderStyle     =   0  'None
   Caption         =   "Feedbackfrm"
   ClientHeight    =   8730
   ClientLeft      =   2160
   ClientTop       =   360
   ClientWidth     =   15780
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   8730
   ScaleWidth      =   15780
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox TxtQ7 
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
      Left            =   7680
      TabIndex        =   30
      Text            =   "Q7"
      Top             =   7560
      Width           =   7935
   End
   Begin VB.TextBox TxtQ8 
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
      Left            =   7680
      TabIndex        =   29
      Text            =   "Q8"
      Top             =   8160
      Width           =   4335
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
      Height          =   450
      Left            =   5400
      TabIndex        =   26
      Text            =   "Enter Sex"
      Top             =   2640
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
      Left            =   1800
      TabIndex        =   25
      Text            =   "Enter Name"
      Top             =   2640
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
      Height          =   450
      Left            =   12960
      TabIndex        =   23
      Text            =   "Income"
      Top             =   2640
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
      Height          =   450
      Left            =   9240
      MaxLength       =   3
      TabIndex        =   22
      Text            =   "Enter Age"
      Top             =   2640
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
      Left            =   12120
      TabIndex        =   21
      Top             =   8160
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
      Left            =   13320
      TabIndex        =   20
      Top             =   8160
      Width           =   1215
   End
   Begin VB.CommandButton CmdClose 
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
      TabIndex        =   19
      Top             =   8160
      Width           =   975
   End
   Begin VB.TextBox TxtQ6 
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
      Left            =   7680
      TabIndex        =   18
      Text            =   "Q6"
      Top             =   6960
      Width           =   7935
   End
   Begin VB.TextBox TxtQ5 
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
      Left            =   7680
      TabIndex        =   17
      Text            =   "Q5"
      Top             =   6360
      Width           =   7935
   End
   Begin VB.TextBox TxtQ4 
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
      Left            =   7680
      TabIndex        =   16
      Text            =   "Q4"
      Top             =   5760
      Width           =   7935
   End
   Begin VB.TextBox TxtQ3 
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
      Left            =   7680
      TabIndex        =   15
      Text            =   "Q3"
      Top             =   5160
      Width           =   7935
   End
   Begin VB.TextBox TxtQ2 
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
      Left            =   7680
      TabIndex        =   14
      Text            =   "Q2"
      Top             =   4560
      Width           =   7935
   End
   Begin VB.TextBox TxtQ1 
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
      Left            =   7680
      TabIndex        =   13
      Text            =   "Q1"
      Top             =   3960
      Width           =   7935
   End
   Begin VB.Label Label16 
      BackStyle       =   0  'Transparent
      Caption         =   "Feedback "
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
      Left            =   13920
      TabIndex        =   31
      Top             =   120
      Width           =   1695
   End
   Begin VB.Label Label15 
      BackStyle       =   0  'Transparent
      Caption         =   "8-Any room for improvements/Suggestions"
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
      TabIndex        =   28
      Top             =   8160
      Width           =   10215
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
      Left            =   4320
      TabIndex        =   27
      Top             =   2280
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
      Left            =   360
      TabIndex        =   24
      Top             =   2280
      Width           =   1575
   End
   Begin VB.Label Label14 
      BackStyle       =   0  'Transparent
      Caption         =   "1-The Reception"
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
      Top             =   3960
      Width           =   9015
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "5-Sharing your experience with others"
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
      Top             =   6360
      Width           =   9015
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "7-Our services throughtout the expo"
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
      Top             =   7560
      Width           =   9015
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "6-Attendance provided to you by our sales person"
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
      Top             =   6960
      Width           =   9015
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "4-Vehicle prices and payment policy "
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
      Top             =   5760
      Width           =   9015
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "3-Fullfillness of our commitments"
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
      Top             =   5160
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
      Top             =   4560
      Width           =   9015
   End
   Begin VB.Label Label7 
      BackColor       =   &H000080FF&
      Caption         =   "Please rate our Team on the basis of :- "
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
      Top             =   3240
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
      Left            =   10920
      TabIndex        =   4
      Top             =   2280
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
      Left            =   7440
      TabIndex        =   3
      Top             =   2280
      Width           =   2055
   End
   Begin VB.Label Label3 
      BackColor       =   &H000080FF&
      Caption         =   "Please tell us a bit about  you......!!!!!"
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
      Top             =   1800
      Width           =   15495
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   $"feedback.frx":0000
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   15.75
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
      Caption         =   "Auto Expo Satisfactory Survey"
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
      Height          =   8715
      Left            =   0
      Picture         =   "feedback.frx":00DF
      Stretch         =   -1  'True
      Top             =   0
      Width           =   15765
   End
End
Attribute VB_Name = "feedbackfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim C As New ADODB.Connection
Dim r As New ADODB.Recordset
Dim s As String

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

Private Sub CmdClose_Click()
Load Thankufrm
Thankufrm.Show
End Sub

Private Sub CmdSubmit_Click()
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


