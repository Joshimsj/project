VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form bookingfrm 
   BorderStyle     =   0  'None
   ClientHeight    =   8850
   ClientLeft      =   2565
   ClientTop       =   945
   ClientWidth     =   15675
   LinkTopic       =   "Form1"
   ScaleHeight     =   8850
   ScaleWidth      =   15675
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox Txt_DelMid 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   11.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   11040
      TabIndex        =   32
      Top             =   4680
      Width           =   3255
   End
   Begin VB.TextBox Txt_DelMob 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   11.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   11040
      TabIndex        =   30
      Top             =   3960
      Width           =   3255
   End
   Begin VB.TextBox Txt_DelNam 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   11.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   11040
      TabIndex        =   28
      Top             =   3240
      Width           =   3255
   End
   Begin MSComCtl2.DTPicker DOP 
      Height          =   615
      Left            =   11040
      TabIndex        =   26
      Top             =   5400
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
      CalendarBackColor=   0
      CalendarForeColor=   16777215
      CalendarTitleBackColor=   65535
      CalendarTitleForeColor=   255
      CalendarTrailingForeColor=   65280
      Format          =   163774465
      CurrentDate     =   43294
   End
   Begin VB.TextBox Brand_Txt 
      BackColor       =   &H00000000&
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
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   11040
      TabIndex        =   24
      Top             =   1200
      Width           =   3255
   End
   Begin VB.TextBox Model_Txt 
      BackColor       =   &H00000000&
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
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   3840
      TabIndex        =   22
      Top             =   1200
      Width           =   3135
   End
   Begin VB.CommandButton CmdNext 
      BackColor       =   &H00C0C0C0&
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
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   7560
      Width           =   1695
   End
   Begin VB.CommandButton cmdadd 
      BackColor       =   &H00C0C0C0&
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
      Left            =   5400
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   7560
      Width           =   1695
   End
   Begin VB.CommandButton CmdClose 
      BackColor       =   &H00E0E0E0&
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
      Left            =   14280
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   8160
      Width           =   1335
   End
   Begin VB.CommandButton CmdSave 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Save"
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
      Left            =   7320
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   7560
      Width           =   1695
   End
   Begin VB.TextBox TxtZip 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   11.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   3840
      MaxLength       =   6
      TabIndex        =   7
      Top             =   4680
      Width           =   3135
   End
   Begin VB.TextBox TxtCity 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   11.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   11040
      TabIndex        =   6
      Top             =   2520
      Width           =   3255
   End
   Begin VB.TextBox TxtState 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   11.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   11040
      TabIndex        =   5
      Top             =   1800
      Width           =   3255
   End
   Begin VB.TextBox TxtAdd 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   11.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   3840
      TabIndex        =   4
      Top             =   3960
      Width           =   3135
   End
   Begin MSComCtl2.DTPicker DOB 
      Height          =   615
      Left            =   3840
      TabIndex        =   3
      Top             =   5400
      Width           =   3135
      _ExtentX        =   5530
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
      Format          =   163774465
      CurrentDate     =   43265
   End
   Begin VB.TextBox TxtMob 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   11.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   3840
      MaxLength       =   10
      TabIndex        =   2
      Top             =   3240
      Width           =   3135
   End
   Begin VB.TextBox TxtCompany 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   11.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   3840
      TabIndex        =   1
      Top             =   2520
      Width           =   3135
   End
   Begin VB.TextBox TxtName 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   11.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   3840
      TabIndex        =   0
      Top             =   1920
      Width           =   3135
   End
   Begin VB.Label Label15 
      BackStyle       =   0  'Transparent
      Caption         =   "Deal_Mail :-"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   11.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   9120
      TabIndex        =   31
      Top             =   4680
      Width           =   1695
   End
   Begin VB.Label Label14 
      BackStyle       =   0  'Transparent
      Caption         =   "Deal_Mobile :-"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   11.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   9120
      TabIndex        =   29
      Top             =   3960
      Width           =   1695
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "Deal_name :-"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   11.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   9120
      TabIndex        =   27
      Top             =   3240
      Width           =   1695
   End
   Begin VB.Label Label12 
      BackColor       =   &H8000000D&
      BackStyle       =   0  'Transparent
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
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   9120
      TabIndex        =   25
      Top             =   5400
      Width           =   1695
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
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
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   9120
      TabIndex        =   23
      Top             =   1200
      Width           =   1695
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
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
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   1920
      TabIndex        =   21
      Top             =   1200
      Width           =   1695
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
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
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   1920
      TabIndex        =   19
      Top             =   4800
      Width           =   1695
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
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
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   9120
      TabIndex        =   18
      Top             =   2520
      Width           =   1695
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "State :-"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   11.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   9120
      TabIndex        =   17
      Top             =   1800
      Width           =   1695
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
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
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   1920
      TabIndex        =   16
      Top             =   3960
      Width           =   1695
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Date :-"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   11.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   1920
      TabIndex        =   15
      Top             =   5520
      Width           =   1695
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
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
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   1920
      TabIndex        =   14
      Top             =   3240
      Width           =   1695
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
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
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   1920
      TabIndex        =   13
      Top             =   2520
      Width           =   1695
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
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
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   1920
      TabIndex        =   12
      Top             =   1800
      Width           =   1695
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   """Customer Details"""
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   14.25
         Charset         =   0
         Weight          =   600
         Underline       =   -1  'True
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   6720
      TabIndex        =   11
      Top             =   240
      Width           =   3135
   End
   Begin VB.Image Batman 
      Height          =   9285
      Left            =   -480
      Picture         =   "Booking .frx":0000
      Stretch         =   -1  'True
      Top             =   -480
      Width           =   16140
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
DOB.Value = "1/01/1990"
TxtName.SetFocus
End Sub

Private Sub CmdClose_Click()
End
End Sub

Private Sub CmdNext_Click()
Load Selfrom
Selfrom.Load_data model_price, car_category, brand_selected, model_id_selected, model_name, cst
Selfrom.Show
End Sub

Private Sub cmdSave_Click()
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
cst.Dealer_Name = Txt_DelNam.Text
cst.Dealer_Mob = Txt_DelMob.Text
cst.Dealer_Mid = Txt_DelMid.Text
Call cst.SaveData
End Sub

Public Sub Load_Selected_Data(car_model_id, brand, Price, model, category)
    model_id_selected = car_model_id
    brand_selected = brand
    model_price = Price
    model_name = model
    car_category = category
    
    Model_Txt.Text = model_id_selected
    Brand_Txt.Text = brand_selected
End Sub

Private Sub Txt_DelMid_GotFocus()
If (Txt_DelMob.Text = "") Then
MsgBox ("Enter valid Dealer Mobile No")
Txt_DelMob.SetFocus
End If
End Sub

Private Sub Txt_DelMob_Change()
If IsNumeric(Txt_DelMob.Text) = False Then
MsgBox ("Digits Only")
Txt_DelMob.Text = ""
Txt_DelMob.SetFocus
End If
End Sub

Private Sub Txt_DelMob_GotFocus()
If (Txt_DelNam.Text = "") Then
MsgBox ("Enter valid Dealer Name")
Txt_DelNam.SetFocus
End If
End Sub

Private Sub Txt_DelNam_Change()
If IsNumeric(Txt_DelNam.Text) = True Then
MsgBox ("Text Only")
Txt_DelNam.Text = ""
Txt_DelNam.SetFocus
End If
End Sub

Private Sub Txt_DelNam_GotFocus()
If (TxtCity.Text = "") Then
MsgBox ("Enter valid City Name")
TxtCity.SetFocus
End If
End Sub

Private Sub TxtAdd_GotFocus()
If (TxtMob.Text = "") Then
MsgBox ("Enter valid Mobile No")
TxtMob.SetFocus
End If
End Sub

Private Sub TxtCity_Change()
If IsNumeric(TxtCity.Text) = True Then
MsgBox ("Text Only")
TxtCity.Text = ""
TxtCity.SetFocus
End If
End Sub

Private Sub TxtCity_GotFocus()
If (TxtState.Text = "") Then
MsgBox ("Enter valid State Name")
TxtState.SetFocus
End If
End Sub

Private Sub TxtCompany_Change()
If IsNumeric(TxtCompany.Text) = True Then
MsgBox ("Text Only")
TxtCompany.Text = ""
TxtCompany.SetFocus
End If
End Sub

Private Sub TxtCompany_GotFocus()
If (TxtName.Text = "") Then
MsgBox ("Enter valid Name")
TxtName.SetFocus
End If
End Sub

Private Sub TxtMob_Change()
If IsNumeric(TxtMob.Text) = False Then
MsgBox ("Digits Only")
TxtMob.Text = ""
TxtMob.SetFocus
End If
End Sub

Private Sub TxtMob_GotFocus()
If (TxtCompany.Text = "") Then
MsgBox ("Enter valid Company Name")
TxtCompany.SetFocus
End If
End Sub

Private Sub TxtName_Change()
If IsNumeric(TxtName.Text) = True Then
MsgBox ("Text Only")
TxtName.Text = ""
TxtName.SetFocus
End If
End Sub

Private Sub TxtState_Change()
If IsNumeric(TxtState.Text) = True Then
MsgBox ("Text Only")
TxtState.Text = ""
TxtState.SetFocus
End If
End Sub

Private Sub TxtState_GotFocus()
If (TxtZip.Text = "") Then
MsgBox ("Enter valid Zipcode")
TxtZip.SetFocus
End If
End Sub

Private Sub TxtZip_Change()
If IsNumeric(TxtZip.Text) = False Then
MsgBox ("Digits Only")
TxtZip.Text = ""
TxtZip.SetFocus
End If
End Sub

Private Sub TxtZip_GotFocus()
If (TxtAdd.Text = "") Then
MsgBox ("Enter valid Address")
TxtAdd.SetFocus
End If
End Sub
