VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form bookingfrm 
   BorderStyle     =   0  'None
   ClientHeight    =   6405
   ClientLeft      =   5070
   ClientTop       =   2115
   ClientWidth     =   11295
   LinkTopic       =   "Form1"
   Picture         =   "Booking .frx":0000
   ScaleHeight     =   6405
   ScaleWidth      =   11295
   ShowInTaskbar   =   0   'False
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   495
      Left            =   7440
      TabIndex        =   26
      Top             =   3960
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   873
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Bookman Old Style"
         Size            =   11.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Format          =   168165377
      CurrentDate     =   43294
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   11.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7440
      TabIndex        =   24
      Top             =   1200
      Visible         =   0   'False
      Width           =   3255
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   11.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2160
      TabIndex        =   22
      Top             =   1200
      Visible         =   0   'False
      Width           =   3135
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Next"
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
      TabIndex        =   20
      Top             =   5760
      Width           =   1455
   End
   Begin VB.CommandButton cmdadd 
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
      Height          =   495
      Left            =   3120
      TabIndex        =   10
      Top             =   5760
      Width           =   1815
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Close"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9360
      TabIndex        =   9
      Top             =   5760
      Width           =   1335
   End
   Begin VB.CommandButton CmdSave 
      Caption         =   "Save "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5280
      TabIndex        =   8
      Top             =   5760
      Width           =   1695
   End
   Begin VB.TextBox TxtZip 
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   11.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5640
      TabIndex        =   7
      Text            =   "Enter your Zipcode"
      Top             =   4800
      Width           =   2655
   End
   Begin VB.TextBox TxtCity 
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   11.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   7440
      TabIndex        =   6
      Text            =   "Enter your City"
      Top             =   3240
      Width           =   3255
   End
   Begin VB.TextBox TxtState 
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   11.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   7440
      TabIndex        =   5
      Text            =   "Enter your Sate"
      Top             =   2520
      Width           =   3255
   End
   Begin VB.TextBox TxtAdd 
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   11.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   7440
      TabIndex        =   4
      Text            =   "Enter your Address"
      Top             =   1800
      Width           =   3255
   End
   Begin MSComCtl2.DTPicker DOB 
      Height          =   495
      Left            =   2160
      TabIndex        =   3
      Top             =   3960
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   873
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Bookman Old Style"
         Size            =   11.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Format          =   168165377
      CurrentDate     =   43265
   End
   Begin VB.TextBox TxtMob 
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   11.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2160
      TabIndex        =   2
      Text            =   "Enter your Mobile No"
      Top             =   3240
      Width           =   3135
   End
   Begin VB.TextBox TxtCompany 
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   11.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2160
      TabIndex        =   1
      Text            =   "Enter Company name"
      Top             =   2520
      Width           =   3135
   End
   Begin VB.TextBox TxtName 
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   11.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2160
      TabIndex        =   0
      Text            =   "Enter Customer Name"
      Top             =   1800
      Width           =   3135
   End
   Begin VB.Label Label12 
      Caption         =   "DOP :-"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   11.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5640
      TabIndex        =   25
      Top             =   3960
      Width           =   1575
   End
   Begin VB.Label Label11 
      Caption         =   "Brand :-"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   11.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5640
      TabIndex        =   23
      Top             =   1200
      Width           =   1575
   End
   Begin VB.Label Label10 
      Caption         =   "Model_id :-"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   11.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   21
      Top             =   1200
      Width           =   1695
   End
   Begin VB.Label Label9 
      Caption         =   "Zipcode :-"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   11.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3840
      TabIndex        =   19
      Top             =   4800
      Width           =   1575
   End
   Begin VB.Label Label8 
      Caption         =   "City :-"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   11.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5640
      TabIndex        =   18
      Top             =   3240
      Width           =   1575
   End
   Begin VB.Label Label7 
      Caption         =   "Sate :-"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   11.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5640
      TabIndex        =   17
      Top             =   2520
      Width           =   1575
   End
   Begin VB.Label Label6 
      Caption         =   "Address :-"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   11.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5640
      TabIndex        =   16
      Top             =   1800
      Width           =   1575
   End
   Begin VB.Label Label5 
      Caption         =   "DOB :-"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   11.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   15
      Top             =   3960
      Width           =   1695
   End
   Begin VB.Label Label4 
      Caption         =   "Mobile No :-"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   11.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   14
      Top             =   3240
      Width           =   1695
   End
   Begin VB.Label Label3 
      Caption         =   "Com_Name :-"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   11.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   13
      Top             =   2520
      Width           =   1695
   End
   Begin VB.Label Label2 
      Caption         =   "Name :-"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   11.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   12
      Top             =   1800
      Width           =   1695
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Booking Detail"
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
      Left            =   4080
      TabIndex        =   11
      Top             =   240
      Width           =   3135
   End
   Begin VB.Image Image1 
      Height          =   6405
      Left            =   0
      Picture         =   "Booking .frx":B3F1D
      Stretch         =   -1  'True
      Top             =   0
      Width           =   11340
   End
End
Attribute VB_Name = "bookingfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim C As New ADODB.Connection
Dim r As New ADODB.Recordset
Dim s As String

Private Sub CmdAdd_Click()
TxtName.Text = ""
TxtCompany.Text = ""
TxtMob.Text = ""
DOB.Value = "1/01/1990"
TxtAdd.Text = ""
TxtState.Text = ""
TxtCity.Text = ""
TxtZip.Text = ""
'Com_ID1.Text = ""
'Txtdata.Text = ""
End Sub

Private Sub cmdSave_Click()
''r.Close
''s = "insert into cus_details(" & TxtName.Text & "," & TxtCompany.Text & "," & TxtMob.Text & "," & DOB.Value & "," & TxtAdd.Text & "," & TxtSate.Text & "," & TxtCity.Text & "," & TxtZip.Text & ")"
''r.Open s, c, adOpenDynamic, adLockOptimistic
''s = "select * from Customers"
''r.Open s, c, adOpenDynamic, adLockOptimistic
''If Not r.BOF And r.EOF Then
Dim cst As New customer
cst.Name = TxtName.Text
cst.Company = TxtCompany.Text
cst.MobileNumber = TxtMob.Text
cst.DOB = DOB.Value
cst.Address = TxtAdd.Text
cst.State = TxtState.Text
cst.City = TxtCity.Text
cst.Zip = TxtZip.Text
'cst.Identity = Com_ID1.Text
'cst.EnterDetails = com_ID2.Text
Call cst.SaveData
''Else
''MsgBox "Details are Sucessfully Added", vbInformation, "Customers"
''End If
End Sub

Private Sub Command1_Click()
Load Selfrom
Selfrom.Show
End Sub

Private Sub Command8_Click()
'Load feedbackfrm
'feedbackfrm.Show
End
End Sub

'Private Sub com_ID_Click()'
'Com_ID1.AddItem = "Aadhar Card"
'Com_ID1.AddItem = "Pancard"
'Com_ID1.AddItem = "Voter_ID"
'Com_ID1.AddItem = "Passport"
'End Sub


Private Sub Form_Load()
'c.Open "provider=microsoft.jet.oledb.4.0;data "
End Sub
