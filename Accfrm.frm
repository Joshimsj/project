VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Accfrm 
   BorderStyle     =   0  'None
   Caption         =   "Account Details"
   ClientHeight    =   8415
   ClientLeft      =   5640
   ClientTop       =   765
   ClientWidth     =   9750
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8415
   ScaleWidth      =   9750
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton CmdClose 
      BackColor       =   &H00C0C0C0&
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
      Left            =   7200
      MaskColor       =   &H00808080&
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   7800
      Width           =   2175
   End
   Begin VB.CommandButton CmdAdd 
      BackColor       =   &H00C0C0C0&
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
      Left            =   7200
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   7080
      Width           =   2175
   End
   Begin VB.CommandButton CmdSummit 
      BackColor       =   &H00C0C0C0&
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
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   7800
      Width           =   2535
   End
   Begin MSComCtl2.DTPicker Dat 
      Height          =   495
      Left            =   4320
      TabIndex        =   20
      Top             =   7080
      Width           =   2535
      _ExtentX        =   4471
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
      CalendarForeColor=   65535
      CalendarTitleBackColor=   255
      CalendarTitleForeColor=   65535
      CalendarTrailingForeColor=   65280
      Format          =   109510657
      CurrentDate     =   43289
   End
   Begin VB.TextBox TxtDeMob 
      BackColor       =   &H0000FFFF&
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   14.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   4320
      MaxLength       =   10
      TabIndex        =   19
      Top             =   6360
      Width           =   5055
   End
   Begin VB.TextBox TxtDeName 
      BackColor       =   &H0000FFFF&
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   14.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   4320
      TabIndex        =   18
      Top             =   5760
      Width           =   5055
   End
   Begin VB.TextBox TxtAmtWords 
      BackColor       =   &H0000FFFF&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   14.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   615
      Left            =   4320
      TabIndex        =   17
      Top             =   5040
      Width           =   5055
   End
   Begin VB.TextBox TxtAmtNo 
      BackColor       =   &H0000FFFF&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   14.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   615
      Left            =   4320
      TabIndex        =   16
      Top             =   4320
      Width           =   5055
   End
   Begin VB.TextBox TxtBrName 
      BackColor       =   &H0000FFFF&
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   14.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   525
      Left            =   4320
      TabIndex        =   15
      Top             =   3720
      Width           =   5055
   End
   Begin VB.TextBox TxtDName 
      BackColor       =   &H0000FFFF&
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   14.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   615
      Left            =   4320
      TabIndex        =   14
      Top             =   3000
      Width           =   5055
   End
   Begin VB.TextBox TxtCode 
      BackColor       =   &H0000FFFF&
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   14.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   4320
      MaxLength       =   11
      TabIndex        =   13
      Top             =   2400
      Width           =   5055
   End
   Begin VB.TextBox TxtAccNo 
      BackColor       =   &H0000FFFF&
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   14.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   4320
      MaxLength       =   15
      TabIndex        =   12
      Top             =   1800
      Width           =   5055
   End
   Begin VB.TextBox TxtAccN 
      BackColor       =   &H0000FFFF&
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   14.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   4320
      TabIndex        =   11
      Top             =   1200
      Width           =   5055
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "Depositer Mob No :-"
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
      Caption         =   "Depositer Name :-"
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
         Underline       =   -1  'True
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
      Caption         =   "Date :-"
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
      Caption         =   "Amount (in Words) :- "
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
      Caption         =   "Branch Name :-"
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
      Caption         =   "IFSC Code :-"
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
      Caption         =   "Amount (in Rs.) :-"
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
      Caption         =   "Bank Drawn Name :-"
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
      Caption         =   "Acc No :-"
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
      Caption         =   "Acc Holder Name :- "
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
Dat.Value = "1/10/2018"
TxtAccN.SetFocus
End Sub

Private Sub CmdClose_Click()
Invoice.Show
End Sub

Private Sub CmdSummit_Click()
Invoice.Load_payment TxtAmtNo.Text, "Account Pay"
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

Public Sub Load_Amount(Price)
    amount = Price
    TxtAmtNo.Text = amount
End Sub

Private Sub TxtAccN_Change()
If IsNumeric(TxtAccN.Text) = True Then
MsgBox ("Text Only")
TxtAccN.Text = ""
TxtAccN.SetFocus
End If
End Sub

Private Sub TxtAccNo_Change()
If IsNumeric(TxtAccNo.Text) = False Then
MsgBox ("Digits Only")
TxtAccNo.Text = ""
TxtAccNo.SetFocus
End If
End Sub

Private Sub TxtAccNo_GotFocus()
If (TxtAccN.Text = "") Then
MsgBox ("Enter a valid Bank Acc Name")
TxtAccN.SetFocus
End If
End Sub

Private Sub TxtBrName_Change()
If IsNumeric(TxtBrName.Text) = True Then
MsgBox ("Text Only")
TxtBrName.Text = ""
TxtBrName.SetFocus
End If
End Sub

Private Sub TxtBrName_GotFocus()
If (TxtDName.Text = "") Then
MsgBox ("Enter a valid Drawer Name")
TxtDName.SetFocus
End If
End Sub

Private Sub TxtCode_Change()
If IsNumeric(TxtCode.Text) = False Then
MsgBox ("Digits Only")
TxtCode.Text = ""
TxtCode.SetFocus
End If
End Sub

Private Sub TxtCode_GotFocus()
If (TxtAccNo.Text = "") Then
MsgBox ("Enter a valid Bank Acc No")
TxtAccNo.SetFocus
End If
End Sub

Private Sub TxtDeMob_Change()
If IsNumeric(TxtDeMob.Text) = False Then
MsgBox ("Digits Only")
TxtDeMob.Text = ""
TxtDeMob.SetFocus
End If
End Sub

Private Sub TxtDeMob_GotFocus()
If (TxtDeName.Text = "") Then
MsgBox ("Enter a valid Depositor Name")
TxtDeName.SetFocus
End If
End Sub

Private Sub TxtDeName_Change()
If IsNumeric(TxtDeName.Text) = True Then
MsgBox ("Text Only")
TxtDeName.Text = ""
TxtDeName.SetFocus
End If
End Sub

Private Sub TxtDeName_GotFocus()
If (TxtBrName.Text = "") Then
MsgBox ("Enter a valid Branch Name")
TxtBrName.SetFocus
End If
End Sub

Private Sub TxtDName_Change()
If IsNumeric(TxtDName.Text) = True Then
MsgBox ("Text Only")
TxtDName.Text = ""
TxtDName.SetFocus
End If
End Sub

Private Sub TxtDName_GotFocus()
If (TxtCode.Text = "") Then
MsgBox ("Enter a valid IFSC Code")
TxtCode.SetFocus
End If
End Sub
