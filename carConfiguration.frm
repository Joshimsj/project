VERSION 5.00
Begin VB.Form carConfiguration 
   BackColor       =   &H8000000B&
   BorderStyle     =   0  'None
   Caption         =   "Car Configuration"
   ClientHeight    =   11100
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   20400
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   11100
   ScaleWidth      =   20400
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command6 
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
      Height          =   615
      Left            =   18000
      TabIndex        =   28
      Top             =   9720
      Width           =   1935
   End
   Begin VB.CommandButton Command5 
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
      Height          =   615
      Left            =   15360
      TabIndex        =   27
      Top             =   9720
      Width           =   2055
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Book"
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
      Left            =   18000
      TabIndex        =   26
      Top             =   8760
      Width           =   1935
   End
   Begin VB.CommandButton Command3 
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
      Height          =   615
      Left            =   15360
      TabIndex        =   25
      Top             =   8760
      Width           =   2055
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
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2775
      Left            =   120
      TabIndex        =   29
      Top             =   8160
      Width           =   14295
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
      Height          =   11175
      Left            =   -120
      Top             =   0
      Width           =   20535
   End
End
Attribute VB_Name = "carConfiguration"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim model_id As Integer
Dim conn As New ADODB.Connection
Dim car As New ADODB.Recordset

Private Sub Command6_Click()
End
End Sub

Private Sub Form_Load()
    conn.Open "Provider=Microsoft.jet.OLEDB.4.0;Data Source =E:\project\assets\cars3.mdb;persist security info=false"
    
    'model_id = 5
    
    
    
End Sub

Public Sub Intilize_form()
    If model_id > 0 Then
        query = "SELECT * FROM cars WHERE model_id = " & model_id
        car.Open query, conn, adUseClient, adLockOptimistic, adCmdText
        
        If car.RecordCount > 0 Then
            Txt_Model_ID.Text = car!model_id
            Txt_Brand.Text = car!brand
            Txt_Model.Text = car!model
            Txt_Engine.Text = car!engine
            Txt_Transmission.Text = car!transmission
            Txt_Power.Text = car!power
            Txt_Speed.Text = car!speed
            Txt_Airbags.Text = car!airbags
            Txt_Abs.Text = car!Abs
            Txt_Torque.Text = car!torque
            Txt_FuelType.Text = car!fuel_type
            Txt_Cost.Text = car!cost
            
            If Not IsNull(car!display_pic) Then
                ImageDisplay.Picture = LoadPicture("E:\project\images\" & car!display_pic)
            End If
        Else
            MsgBox ("Model not found")
        End If
        
    Else
        MsgBox ("Invalid Model ID")
    End If
End Sub
Public Sub Add_model_id(ByVal model As Integer)
    model_id = model
   
End Sub
