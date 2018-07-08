VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Chqfrm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cheque Option"
   ClientHeight    =   7845
   ClientLeft      =   4065
   ClientTop       =   1515
   ClientWidth     =   10335
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7845
   ScaleWidth      =   10335
   Begin MSComCtl2.DTPicker ChqD 
      Height          =   495
      Left            =   6360
      TabIndex        =   17
      Top             =   6480
      Width           =   3495
      _ExtentX        =   6165
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
      Format          =   97452033
      CurrentDate     =   43288
   End
   Begin VB.CommandButton C 
      Caption         =   "C"
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
      Left            =   6360
      TabIndex        =   16
      Top             =   7200
      Width           =   615
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
      Height          =   495
      Left            =   8520
      TabIndex        =   14
      Top             =   7200
      Width           =   1335
   End
   Begin VB.CommandButton Submit 
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
      Left            =   7080
      TabIndex        =   13
      Top             =   7200
      Width           =   1335
   End
   Begin VB.TextBox TxtCn 
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
      Left            =   6360
      TabIndex        =   12
      Text            =   "Chq No "
      Top             =   5640
      Width           =   3495
   End
   Begin VB.TextBox TxtAw 
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
      Left            =   6360
      TabIndex        =   11
      Text            =   "Chq Amt in Words"
      Top             =   4920
      Width           =   3495
   End
   Begin VB.TextBox TxtAno 
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
      Left            =   6360
      TabIndex        =   10
      Text            =   "Chq Amt in No"
      Top             =   4200
      Width           =   3495
   End
   Begin VB.TextBox TxtHol 
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
      Left            =   6360
      TabIndex        =   9
      Text            =   "Chq Holder Name"
      Top             =   3360
      Width           =   3495
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
      Height          =   615
      Left            =   6360
      TabIndex        =   8
      Text            =   "Enter bank Name"
      Top             =   2520
      Width           =   3495
   End
   Begin VB.TextBox Txtpay 
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
      Left            =   6360
      TabIndex        =   7
      Text            =   "Enter Pay to"
      Top             =   1680
      Width           =   3495
   End
   Begin VB.Label Label8 
      Caption         =   "Cheque Master"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   26.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   2880
      TabIndex        =   15
      Top             =   360
      Width           =   4815
   End
   Begin VB.Label Label7 
      Caption         =   "Cheque dated:- "
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
      Left            =   480
      TabIndex        =   6
      Top             =   6480
      Width           =   4455
   End
   Begin VB.Label Label6 
      Caption         =   "Cheque no:- "
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
      Left            =   480
      TabIndex        =   5
      Top             =   5760
      Width           =   4455
   End
   Begin VB.Label Label5 
      Caption         =   "Cheque Amount in Words:-"
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
      Left            =   480
      TabIndex        =   4
      Top             =   5040
      Width           =   4455
   End
   Begin VB.Label Label4 
      Caption         =   "Cheque Amount in No:-"
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
      Left            =   480
      TabIndex        =   3
      Top             =   4200
      Width           =   4455
   End
   Begin VB.Label Label3 
      Caption         =   "Cheque Holder Name:-"
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
      Left            =   480
      TabIndex        =   2
      Top             =   3360
      Width           =   4455
   End
   Begin VB.Label Label2 
      Caption         =   "Bank Name:-"
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
      Left            =   480
      TabIndex        =   1
      Top             =   2640
      Width           =   4455
   End
   Begin VB.Label Label1 
      Caption         =   "Pay To:- "
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   14.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   480
      TabIndex        =   0
      Top             =   1800
      Width           =   4455
   End
   Begin VB.Image Image1 
      Height          =   7815
      Left            =   0
      Top             =   0
      Width           =   10335
   End
End
Attribute VB_Name = "Chqfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Cn As New ADODB.Connection
Dim r As New ADODB.Recordset
Dim s As String

Private Sub C_Click()
Txtpay.Text = ""
TxtName.Text = ""
TxtHol.Text = ""
TxtAno.Text = ""
TxtAw.Text = ""
TxtCn.Text = ""
ChqD.Value = "1/10/2018"
End Sub

Private Sub CmdClose_Click()
End
End Sub

Private Sub Submit_Click()
Dim ch As New ChqDetails
ch.Pay_to = Txtpay.Text
ch.Bank_Name = TxtName.Text
ch.Cheq_Holder_Name = TxtHol.Text
ch.Cheq_Amt_No = TxtAno.Text
ch.Cheq_Amt_Words = TxtAw.Text
ch.Cheq_No = TxtCn.Text
ch.Cheq_Dated = ChqD.Value
Call ch.SaveD
End Sub
