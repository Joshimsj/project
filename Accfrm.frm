VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Accfrm 
   BorderStyle     =   0  'None
   Caption         =   "Account Details"
   ClientHeight    =   8415
   ClientLeft      =   6390
   ClientTop       =   1920
   ClientWidth     =   9750
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8415
   ScaleWidth      =   9750
   ShowInTaskbar   =   0   'False
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
      Height          =   495
      Left            =   7320
      TabIndex        =   23
      Top             =   7800
      Width           =   2055
   End
   Begin VB.CommandButton CmdAdd 
      Caption         =   "ADD"
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
      Left            =   7320
      TabIndex        =   22
      Top             =   7080
      Width           =   2055
   End
   Begin VB.CommandButton CmdSummit 
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
      Height          =   495
      Left            =   4320
      TabIndex        =   21
      Top             =   7800
      Width           =   2655
   End
   Begin MSComCtl2.DTPicker Dat 
      Height          =   495
      Left            =   4320
      TabIndex        =   20
      Top             =   7080
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   873
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Bookman Old Style"
         Size            =   14.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Format          =   97845249
      CurrentDate     =   43289
   End
   Begin VB.TextBox TxtDeMob 
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
      TabIndex        =   19
      Text            =   "Enter Depositor Mob No"
      Top             =   6360
      Width           =   5055
   End
   Begin VB.TextBox TxtDeName 
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
      TabIndex        =   18
      Text            =   "Enter Depositor Name"
      Top             =   5760
      Width           =   5055
   End
   Begin VB.TextBox TxtAmtWords 
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   14.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4320
      TabIndex        =   17
      Text            =   "Enter Amt in Words"
      Top             =   5040
      Width           =   5055
   End
   Begin VB.TextBox TxtAmtNo 
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   14.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4320
      TabIndex        =   16
      Text            =   "Enter Amt in No"
      Top             =   4320
      Width           =   5055
   End
   Begin VB.TextBox TxtBrName 
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   14.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   4320
      TabIndex        =   15
      Text            =   "Enter Branch Name"
      Top             =   3720
      Width           =   5055
   End
   Begin VB.TextBox TxtDName 
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   14.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4320
      TabIndex        =   14
      Text            =   "Enter Drawer Name"
      Top             =   3000
      Width           =   5055
   End
   Begin VB.TextBox TxtCode 
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
      TabIndex        =   13
      Text            =   "Enter Code"
      Top             =   2400
      Width           =   5055
   End
   Begin VB.TextBox TxtAccNo 
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
      TabIndex        =   12
      Text            =   "Enter Acc No"
      Top             =   1800
      Width           =   5055
   End
   Begin VB.TextBox TxtAccN 
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
      TabIndex        =   11
      Text            =   "Enter Acc Name"
      Top             =   1200
      Width           =   5055
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "Depositer Mob No:-"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   14.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Left            =   360
      TabIndex        =   10
      Top             =   6480
      Width           =   3135
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Depositer Name:-"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   14.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Left            =   360
      TabIndex        =   9
      Top             =   5760
      Width           =   3135
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Bank Acc Payment"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   24
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   735
      Left            =   2520
      TabIndex        =   8
      Top             =   240
      Width           =   5175
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Date:-"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   14.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Left            =   360
      TabIndex        =   7
      Top             =   7200
      Width           =   3135
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Amount in Words:- "
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   14.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Left            =   360
      TabIndex        =   6
      Top             =   5160
      Width           =   3135
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Branch Name:-"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   14.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Left            =   360
      TabIndex        =   5
      Top             =   3720
      Width           =   3135
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "IFSC Code:-"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   14.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Left            =   360
      TabIndex        =   4
      Top             =   2400
      Width           =   3135
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Amount in No:-"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   14.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Left            =   360
      TabIndex        =   3
      Top             =   4440
      Width           =   3135
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Bank Drawn Name:-"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   14.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Left            =   360
      TabIndex        =   2
      Top             =   3000
      Width           =   3135
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Acc No:-"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   14.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Left            =   360
      TabIndex        =   1
      Top             =   1800
      Width           =   3135
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Acc Holder Name:- "
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   14.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Left            =   360
      TabIndex        =   0
      Top             =   1200
      Width           =   3135
   End
   Begin VB.Image Image1 
      Height          =   8400
      Left            =   0
      Picture         =   "Accfrm.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   9720
   End
End
Attribute VB_Name = "Accfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Con As New ADODB.Connection
Dim r As New ADODB.Recordset
Dim s As String

Dim amount As String

Private Sub CmdAdd_Click()
TxtAccN.Text = ""
TxtAccNo.Text = ""
TxtCode.Text = ""
TxtDName.Text = ""
TxtBrName.Text = ""
TxtAmtNo.Text = ""
TxtAmtWords.Text = ""
TxtDeName.Text = ""
TxtDeMob.Text = ""
Dat.Value = "1/10/2018"
'Date.Value = "1/10/2018"
End Sub

Private Sub CmdClose_Click()
Load feedbackfrm
feedbackfrm.Show
End Sub

Private Sub CmdSummit_Click()
Dim An As New AccDetails
An.Acc_Holder_Name = TxtAccN.Text
An.Acc_No = TxtAccNo.Text
An.IFSC_Code = TxtCode.Text
An.Drawn_Name = TxtDName.Text
An.Branch_Name = TxtBrName.Text
An.Amt_no = TxtAmtNo.Text
An.Amt_Words = TxtAmtWords.Text
An.Depositor_Name = TxtDeName.Text
An.Depositor_Mob_No = TxtDeMob.Text
An.Acc_Dat = Dat.Value
Call An.Save
End Sub

Public Sub Load_Amount(price)
    amount = price
    TxtAmtNo.Text = amount
End Sub

