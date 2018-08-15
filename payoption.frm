VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Chqfrm 
   BorderStyle     =   0  'None
   Caption         =   "Cheque Option"
   ClientHeight    =   7845
   ClientLeft      =   4020
   ClientTop       =   1140
   ClientWidth     =   10335
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7845
   ScaleWidth      =   10335
   ShowInTaskbar   =   0   'False
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
      CalendarBackColor=   0
      CalendarForeColor=   16777215
      CalendarTitleBackColor=   65535
      CalendarTitleForeColor=   255
      CalendarTrailingForeColor=   65280
      Format          =   109314049
      CurrentDate     =   43288
   End
   Begin VB.CommandButton C 
      BackColor       =   &H0080C0FF&
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
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   7200
      Width           =   615
   End
   Begin VB.CommandButton CmdClose 
      BackColor       =   &H0080C0FF&
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
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   7200
      Width           =   1335
   End
   Begin VB.CommandButton Submit 
      BackColor       =   &H0080C0FF&
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
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   7200
      Width           =   1335
   End
   Begin VB.TextBox TxtCn 
      BackColor       =   &H000080FF&
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   14.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   615
      Left            =   6360
      MaxLength       =   6
      TabIndex        =   12
      Top             =   5640
      Width           =   3495
   End
   Begin VB.TextBox TxtAw 
      BackColor       =   &H000080FF&
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   14.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   495
      Left            =   6360
      TabIndex        =   11
      Top             =   4920
      Width           =   3495
   End
   Begin VB.TextBox TxtAno 
      BackColor       =   &H000080FF&
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   14.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   495
      Left            =   6360
      TabIndex        =   10
      Top             =   4200
      Width           =   3495
   End
   Begin VB.TextBox TxtHol 
      BackColor       =   &H000080FF&
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   14.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   615
      Left            =   6360
      TabIndex        =   9
      Top             =   3360
      Width           =   3495
   End
   Begin VB.TextBox TxtName 
      BackColor       =   &H000080FF&
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   14.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   615
      Left            =   6360
      TabIndex        =   8
      Top             =   2520
      Width           =   3495
   End
   Begin VB.TextBox Txtpay 
      BackColor       =   &H000080FF&
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   14.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   615
      Left            =   6360
      TabIndex        =   7
      Top             =   1680
      Width           =   3495
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Cheque Master"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   26.25
         Charset         =   0
         Weight          =   600
         Underline       =   -1  'True
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   855
      Left            =   2880
      TabIndex        =   15
      Top             =   360
      Width           =   4815
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Cheque Dated :- "
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   14.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   480
      TabIndex        =   6
      Top             =   6480
      Width           =   4455
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Cheque no :- "
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
      Left            =   480
      TabIndex        =   5
      Top             =   5760
      Width           =   4455
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Cheque Amount( in Words) :-"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   14.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   480
      TabIndex        =   4
      Top             =   5040
      Width           =   4455
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Cheque Amount(in No) :-"
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
      Left            =   480
      TabIndex        =   3
      Top             =   4200
      Width           =   4455
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Cheque Holder Name :-"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   14.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   480
      TabIndex        =   2
      Top             =   3360
      Width           =   4455
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Bank Name :-"
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
      Left            =   480
      TabIndex        =   1
      Top             =   2640
      Width           =   4455
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Pay To :- "
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   14.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   480
      TabIndex        =   0
      Top             =   1800
      Width           =   4455
   End
   Begin VB.Image Image1 
      Height          =   7815
      Left            =   0
      Picture         =   "payoption.frx":0000
      Stretch         =   -1  'True
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
Dim amount As String

Private Sub C_Click()
ChqD.Value = "1/10/2018"
Txtpay.SetFocus
End Sub

Private Sub CmdClose_Click()
Invoice.Show
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

Invoice.Load_payment TxtAno.Text, "Cheque Pay"
End Sub

Public Sub Load_Amount(Price)
    amount = Price
    TxtAno.Text = amount
End Sub

Private Sub TxtCn_Change()
If IsNumeric(TxtCn.Text) = False Then
MsgBox ("Digits Only")
TxtCn.Text = ""
TxtCn.SetFocus
End If
End Sub

Private Sub TxtCn_GotFocus()
If (TxtHol.Text = "") Then
MsgBox ("Enter a valid Cheque Holder Name")
TxtHol.SetFocus
End If
End Sub

Private Sub TxtHol_Change()
If IsNumeric(TxtHol.Text) = True Then
MsgBox ("Text Only")
TxtHol.Text = ""
TxtHol.SetFocus
End If
End Sub

Private Sub TxtHol_GotFocus()
If (TxtName.Text = "") Then
MsgBox ("Enter a valid Bank Name")
TxtName.SetFocus
End If
End Sub

Private Sub TxtName_Change()
If IsNumeric(TxtName.Text) = True Then
MsgBox ("Text Only")
TxtName.Text = ""
TxtName.SetFocus
End If
End Sub

Private Sub TxtName_GotFocus()
If (Txtpay.Text = "") Then
Txtpay.Text = "Auto Expo Pvt Ltd"
Txtpay.SetFocus
End If
End Sub

Private Sub Txtpay_Change()
If IsNumeric(Txtpay.Text) = True Then
MsgBox ("Text Only")
Txtpay.Text = ""
Txtpay.SetFocus
End If
End Sub
