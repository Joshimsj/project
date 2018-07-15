VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form bookingfrm 
   BorderStyle     =   0  'None
   ClientHeight    =   6405
   ClientLeft      =   5070
   ClientTop       =   2115
   ClientWidth     =   11340
   LinkTopic       =   "Form1"
   Picture         =   "Booking .frx":0000
   ScaleHeight     =   6405
   ScaleWidth      =   11340
   ShowInTaskbar   =   0   'False
   Begin MSComCtl2.DTPicker DOP 
      Height          =   615
      Left            =   7440
      TabIndex        =   26
      Top             =   3240
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   1085
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
      Format          =   130023425
      CurrentDate     =   43294
   End
   Begin VB.TextBox Brand_Txt 
      Enabled         =   0   'False
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
      Width           =   3255
   End
   Begin VB.TextBox Model_Txt 
      Enabled         =   0   'False
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
      Left            =   9240
      TabIndex        =   20
      Top             =   4680
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
      Left            =   5640
      TabIndex        =   10
      Top             =   4680
      Width           =   1695
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
      Left            =   4920
      TabIndex        =   9
      Top             =   5400
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
      Left            =   7440
      TabIndex        =   8
      Top             =   4680
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
      Left            =   7440
      TabIndex        =   7
      Text            =   "Enter your Zipcode"
      Top             =   3960
      Width           =   3255
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
      Top             =   2520
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
      Top             =   1800
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
      Height          =   495
      Left            =   2160
      TabIndex        =   4
      Text            =   "Enter your Address"
      Top             =   3960
      Width           =   3135
   End
   Begin MSComCtl2.DTPicker DOB 
      Height          =   495
      Left            =   2160
      TabIndex        =   3
      Top             =   4680
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
      Format          =   130088961
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
      Height          =   495
      Left            =   5640
      TabIndex        =   25
      Top             =   3240
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
      Left            =   5640
      TabIndex        =   19
      Top             =   3960
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
      Top             =   2520
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
      Top             =   1800
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
      Height          =   375
      Left            =   240
      TabIndex        =   16
      Top             =   3960
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
      Top             =   4680
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
      Caption         =   """Customer Details"""
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

Dim model_id_selected As Integer
Dim brand_selected As String
Dim model_price As String

Dim car_category As String
Dim model_name As String

Dim cst As New customer

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

cst.Name = TxtName.Text
cst.Company = TxtCompany.Text
cst.MobileNumber = TxtMob.Text
cst.DOB = DOB.Value
cst.Address = TxtAdd.Text
cst.State = TxtState.Text
cst.City = TxtCity.Text
cst.Zip = TxtZip.Text
cst.model_id = Model_Txt.Text
cst.brand = Brand_Txt.Text
cst.DOP = DOP.Value
'cst.Identity = Com_ID1.Text
'cst.EnterDetails = com_ID2.Text
Call cst.SaveData
''Else
''MsgBox "Details are Sucessfully Added", vbInformation, "Customers"
''End If
End Sub

Private Sub Command1_Click()
Load Selfrom
Selfrom.Load_data model_price, car_category, brand_selected, model_id, model_name, cst
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

Public Sub Load_Selected_Data(model_id, brand, price, model, category)
    model_id_selected = model_id
    brand_selected = brand
    model_price = price
    model_name = model
    car_category = category
    
    Model_Txt.Text = model_id_selected
    Brand_Txt.Text = brand_selected
'c.Open "provider=microsoft.jet.oledb.4.0;data "
End Sub

