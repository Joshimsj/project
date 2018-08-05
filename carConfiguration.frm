VERSION 5.00
Begin VB.Form carConfiguration 
   BackColor       =   &H8000000B&
   BorderStyle     =   0  'None
   Caption         =   "Car Configuration"
   ClientHeight    =   11265
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   20400
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   11265
   ScaleWidth      =   20400
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.TextBox Txt_Des 
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
      Height          =   1455
      Left            =   360
      TabIndex        =   46
      Top             =   9480
      Width           =   9375
   End
   Begin VB.TextBox Txt_Pol 
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
      Left            =   12000
      TabIndex        =   45
      Top             =   10440
      Width           =   2175
   End
   Begin VB.TextBox Txt_Bui 
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
      Left            =   12000
      TabIndex        =   43
      Top             =   9480
      Width           =   2175
   End
   Begin VB.TextBox Txt_Whee 
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
      Height          =   405
      Left            =   17400
      TabIndex        =   40
      Top             =   10680
      Width           =   2535
   End
   Begin VB.TextBox Txt_Dri 
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
      Height          =   375
      Left            =   17400
      TabIndex        =   39
      Top             =   10080
      Width           =   2535
   End
   Begin VB.TextBox Txt_Stw 
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
      Height          =   375
      Left            =   17400
      TabIndex        =   38
      Top             =   9480
      Width           =   2535
   End
   Begin VB.TextBox Txt_Tank 
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
      Height          =   375
      Left            =   17400
      TabIndex        =   37
      Top             =   8880
      Width           =   2535
   End
   Begin VB.TextBox Txt_Size 
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
      Height          =   375
      Left            =   17400
      TabIndex        =   36
      Top             =   8280
      Width           =   2535
   End
   Begin VB.CommandButton CmdNex 
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
      Left            =   2520
      TabIndex        =   30
      Top             =   8160
      Width           =   1815
   End
   Begin VB.CommandButton CmdPrev 
      Caption         =   "Previous "
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
      TabIndex        =   29
      Top             =   8160
      Width           =   1575
   End
   Begin VB.CommandButton CmdExit 
      Caption         =   "Exit"
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
      Left            =   12240
      TabIndex        =   28
      Top             =   8160
      Width           =   1935
   End
   Begin VB.CommandButton CmdFeedb 
      Caption         =   "Feedback"
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
      TabIndex        =   27
      Top             =   8160
      Width           =   1695
   End
   Begin VB.CommandButton CmdBook 
      Caption         =   "Book"
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
      Left            =   9720
      TabIndex        =   26
      Top             =   8160
      Width           =   1935
   End
   Begin VB.CommandButton CmdMain 
      Caption         =   "Main"
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
      Left            =   4920
      TabIndex        =   25
      Top             =   8160
      Width           =   1815
   End
   Begin VB.TextBox Txt_Cost 
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
      Height          =   375
      Left            =   17400
      TabIndex        =   24
      Top             =   7680
      Width           =   2535
   End
   Begin VB.TextBox Txt_FuelType 
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
      Height          =   375
      Left            =   17400
      TabIndex        =   23
      Top             =   7080
      Width           =   2535
   End
   Begin VB.TextBox Txt_Torque 
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
      Height          =   375
      Left            =   17400
      TabIndex        =   22
      Top             =   6480
      Width           =   2535
   End
   Begin VB.TextBox Txt_Abs 
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
      Height          =   375
      Left            =   17400
      TabIndex        =   21
      Top             =   5880
      Width           =   2535
   End
   Begin VB.TextBox Txt_Airbags 
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
      Height          =   375
      Left            =   17400
      TabIndex        =   20
      Top             =   5280
      Width           =   2535
   End
   Begin VB.TextBox Txt_Speed 
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
      Height          =   375
      Left            =   17400
      TabIndex        =   19
      Top             =   4680
      Width           =   2535
   End
   Begin VB.TextBox Txt_Power 
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
      Height          =   375
      Left            =   17400
      TabIndex        =   18
      Top             =   4080
      Width           =   2535
   End
   Begin VB.TextBox Txt_Transmission 
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
      Height          =   375
      Left            =   17400
      TabIndex        =   17
      Top             =   3480
      Width           =   2535
   End
   Begin VB.TextBox Txt_Engine 
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
      Height          =   375
      Left            =   17400
      TabIndex        =   16
      Top             =   2880
      Width           =   2535
   End
   Begin VB.TextBox Txt_Model 
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
      Height          =   375
      Left            =   17400
      TabIndex        =   15
      Top             =   2280
      Width           =   2535
   End
   Begin VB.TextBox Txt_Brand 
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
      Height          =   375
      Left            =   17400
      TabIndex        =   14
      Top             =   1680
      Width           =   2535
   End
   Begin VB.TextBox Txt_Model_ID 
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
      Height          =   375
      Left            =   17400
      TabIndex        =   2
      Top             =   1080
      Width           =   2535
   End
   Begin VB.Label Label1 
      Caption         =   "Description"
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
      Left            =   360
      TabIndex        =   47
      Top             =   8880
      Width           =   2175
   End
   Begin VB.Label Label22 
      Caption         =   "Polution Check"
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
      Left            =   9960
      TabIndex        =   44
      Top             =   10440
      Width           =   1815
   End
   Begin VB.Label Label21 
      Caption         =   "Built Quality"
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
      Left            =   9960
      TabIndex        =   42
      Top             =   9480
      Width           =   1815
   End
   Begin VB.Label Label20 
      Alignment       =   2  'Center
      Caption         =   "Certification"
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
      Left            =   10920
      TabIndex        =   41
      Top             =   8880
      Width           =   2535
   End
   Begin VB.Label Label19 
      Caption         =   "Wheels Cover :-"
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
      Left            =   14640
      TabIndex        =   35
      Top             =   10680
      Width           =   2175
   End
   Begin VB.Label Label18 
      Caption         =   "Driving Mode :-"
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
      Left            =   14640
      TabIndex        =   34
      Top             =   10080
      Width           =   2175
   End
   Begin VB.Label Label17 
      Caption         =   "Steering Wheel :-"
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
      Left            =   14640
      TabIndex        =   33
      Top             =   9480
      Width           =   2175
   End
   Begin VB.Label Label16 
      Caption         =   "Fuel Tank (ltr) :-"
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
      Left            =   14640
      TabIndex        =   32
      Top             =   8880
      Width           =   2175
   End
   Begin VB.Label Label2 
      Caption         =   "Tyre Size :-"
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
      Left            =   14640
      TabIndex        =   31
      Top             =   8280
      Width           =   2175
   End
   Begin VB.Label Label15 
      Caption         =   "Cost (Rs) :-"
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
      Left            =   14640
      TabIndex        =   13
      Top             =   7680
      Width           =   2175
   End
   Begin VB.Label Label14 
      Caption         =   "Fuel Type(G/D) :-"
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
      Left            =   14640
      TabIndex        =   12
      Top             =   7080
      Width           =   2175
   End
   Begin VB.Label Label13 
      Caption         =   "Torque (NM) :-"
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
      Left            =   14640
      TabIndex        =   11
      Top             =   6480
      Width           =   2175
   End
   Begin VB.Label Label12 
      Caption         =   "ABS :-"
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
      Left            =   14640
      TabIndex        =   10
      Top             =   5880
      Width           =   2175
   End
   Begin VB.Label Label11 
      Caption         =   "Airbags :-"
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
      Left            =   14640
      TabIndex        =   9
      Top             =   5280
      Width           =   2175
   End
   Begin VB.Label Label10 
      Caption         =   "Speed (KPH) :-"
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
      Left            =   14640
      TabIndex        =   8
      Top             =   4680
      Width           =   2175
   End
   Begin VB.Label Label9 
      Caption         =   "Power (BHP) :-"
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
      Left            =   14640
      TabIndex        =   7
      Top             =   4080
      Width           =   2175
   End
   Begin VB.Label Label8 
      Caption         =   "Transmission :-"
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
      Left            =   14640
      TabIndex        =   6
      Top             =   3480
      Width           =   2175
   End
   Begin VB.Label Label7 
      Caption         =   "Engine (cc) :-"
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
      Left            =   14640
      TabIndex        =   5
      Top             =   2880
      Width           =   2175
   End
   Begin VB.Label Label6 
      Caption         =   "Model :-"
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
      Left            =   14640
      TabIndex        =   4
      Top             =   2280
      Width           =   2175
   End
   Begin VB.Label Label5 
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
      Height          =   375
      Left            =   14640
      TabIndex        =   3
      Top             =   1680
      Width           =   2175
   End
   Begin VB.Label Label4 
      Caption         =   "Model_Id :-"
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
      Left            =   14640
      TabIndex        =   1
      Top             =   1080
      Width           =   2175
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "Configuration"
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
      Left            =   15000
      TabIndex        =   0
      Top             =   120
      Width           =   4935
   End
   Begin VB.Image ImageDisplay 
      Height          =   8055
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   14415
   End
   Begin VB.Image Image1 
      Height          =   11295
      Left            =   0
      Top             =   0
      Width           =   20415
   End
End
Attribute VB_Name = "carConfiguration"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim model_id As Integer
Dim brand As String
Dim model_price As String
Dim car_category_list
Dim car_category As String
Dim model_name As String

'Connect with datbase

Dim conn As New ADODB.Connection
Dim car As New ADODB.Recordset
Dim pictures As New ADODB.Recordset

'set picture in Confi

Dim pic_context() As String
Dim display_pic As String
Dim current_loaded_pic_index As Integer

Private Sub CmdBook_Click()
    Load AccLogin
    bookingfrm.Load_Selected_Data model_id, brand, model_price, model_name, car_category
    AccLogin.Show
End Sub

Private Sub CmdExit_Click()
End
End Sub

Private Sub CmdFeedb_Click()
Load feedbackfrm
feedbackfrm.Show
End Sub

Private Sub CmdMain_Click()
Load MainForm
MainForm.Show
End Sub

Private Sub CmdNex_Click()
    If current_loaded_pic_index < UBound(pic_context) - 1 Then
        current_loaded_pic_index = current_loaded_pic_index + 1
        ImageDisplay.Picture = LoadPicture("E:\project\images\Extra\" & model_id & "\" & pic_context(current_loaded_pic_index))
    Else
        current_loaded_pic_index = -1
        ImageDisplay.Picture = LoadPicture("E:\project\images\" & display_pic)
    End If
End Sub

Private Sub CmdPrev_Click()
    If current_loaded_pic_index > 0 Then
        current_loaded_pic_index = current_loaded_pic_index - 1
        ImageDisplay.Picture = LoadPicture("E:\project\images\Extra\" & model_id & "\" & pic_context(current_loaded_pic_index))
    Else
        current_loaded_pic_index = -1
        ImageDisplay.Picture = LoadPicture("E:\project\images\" & display_pic)
    End If
End Sub

Private Sub Form_Load()
    conn.Open "Provider=Microsoft.jet.OLEDB.4.0;Data Source =E:\project\assets\cars3.mdb;persist security info=false"
End Sub

Public Sub Intilize_form()
    car_category_list = Array("Sports", "Vintage", "Luxury", "Hybrid", "Concept")
    
    If model_id > 0 Then
        query = "SELECT * FROM cars WHERE model_id = " & model_id
        car.Open query, conn, adUseClient, adLockOptimistic, adCmdText
        
        If car.RecordCount > 0 Then
            car_category = car_category_list(car!category)
            
            Txt_Model_ID.Text = car!model_id
            
            brand = car!brand
            Txt_Brand.Text = car!brand
            
            Txt_Model.Text = car!model
            model_name = car!model
            
            Txt_Engine.Text = car!engine
            Txt_Transmission.Text = car!transmission
            Txt_Power.Text = car!power
            Txt_Speed.Text = car!speed
            Txt_Airbags.Text = car!airbags
            Txt_Abs.Text = car!Abs
            Txt_Torque.Text = car!torque
            Txt_FuelType.Text = car!fuel_type
            
            model_price = car!cost
            Txt_Cost.Text = car!cost
            
            Txt_Size.Text = car!Tyre_Size
            Txt_Tank.Text = car!Fuel_tank
            Txt_Stw.Text = car!Steering_Wheel
            Txt_Dri.Text = car!Drive_mode
            Txt_Whee.Text = car!Wheels_Cover
            Txt_Bui.Text = car!Built
            Txt_Pol.Text = car!Polution
            Txt_Des.Text = car!Description
            
            'Load images from pictures database
            
            pic_query = "SELECT Pic FROM cars2 WHERE model_id = " & model_id
            pictures.Open pic_query, conn, adUseClient, adLockOptimistic, adCmdText
            
            If pictures.RecordCount > 0 Then
                ReDim pic_context(pictures.RecordCount) As String
                For i = 0 To pictures.RecordCount - 1
                    pic_context(i) = pictures!pic
                    pictures.MoveNext
                Next i
            End If
            
            If Not IsNull(car!display_pic) Then
                display_pic = car!display_pic
                ImageDisplay.Picture = LoadPicture("E:\project\images\" & display_pic)
                current_loaded_pic_index = -1
            End If
        Else
            MsgBox ("Model not found")
        End If
        
    Else
        MsgBox ("Invalid Model ID")
    End If
    car.Close
    pictures.Close
End Sub
Public Sub Add_model_id(ByVal model As Integer)
    model_id = model
End Sub
