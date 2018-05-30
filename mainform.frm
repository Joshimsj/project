VERSION 5.00
Begin VB.Form MainForm 
   Caption         =   "Form1"
   ClientHeight    =   6480
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   9900
   LinkTopic       =   "Form1"
   ScaleHeight     =   10950
   ScaleWidth      =   20250
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton loadcarBtn 
      Appearance      =   0  'Flat
      Caption         =   "View"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   12
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   100
      TabIndex        =   2
      Top             =   4410
      Width           =   1900
   End
   Begin VB.ListBox carList 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   12
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4230
      Left            =   100
      TabIndex        =   1
      Top             =   100
      Width           =   1900
   End
   Begin VB.CommandButton vTop 
      Caption         =   "Play Add"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   12
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   100
      TabIndex        =   0
      Top             =   5100
      Width           =   1900
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Top"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   14.25
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      TabIndex        =   5
      ToolTipText     =   "Click to view Top View"
      Top             =   0
      Width           =   1500
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Top"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   14.25
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      TabIndex        =   4
      ToolTipText     =   "Click to view Top View"
      Top             =   0
      Width           =   1500
   End
   Begin VB.Label TopView 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Top"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   14.25
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      TabIndex        =   3
      ToolTipText     =   "Click to view Top View"
      Top             =   0
      Width           =   1500
   End
   Begin VB.Image BGPic 
      Height          =   16200
      Left            =   -120
      Picture         =   "mainform.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   29400
   End
End
Attribute VB_Name = "MainForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim conn As New ADODB.Connection
Dim currentVdo As String
Dim currentCar As String

Private Sub carList_Click()
    loadcarBtn.Enabled = True
End Sub

Private Sub Form_Load()
    Me.WindowState = 2
    BGPic.Top = 0
    BGPic.Left = 0
    BGPic.Height = Me.Height
    BGPic.Width = Me.Width
    
    currentVdo = ""
    loadcarBtn.Enabled = False
    
    'Connect to database
     ConnectDatabase "C:\Users\MSJ\Desktop\VB_Smart\c_details.mdb"
    
    'Load Carlist
     Dim cars As New ADODB.Recordset
     query = "SELECT name FROM Cars"
     
     cars.Open query, conn, adUseClient, adLockOptimistic, adCmdText
     

    'Append carlist
    'Set carList.DataSource = cars
    'carList.DataField = "name"
    
    'Total cars in Database
    total_cars = cars.RecordCount
    
    For i = 0 To total_cars - 1
        carList.AddItem cars.Fields(0).Value, i
        cars.MoveNext
    Next i
    
    'carList.AddItem cars.Fields(1)
    cars.Close
    
    playVdoControl.Enabled = False
End Sub

Private Sub Form_Resize()
    BGPic.Height = Me.Height
    BGPic.Width = Me.Width
End Sub

Private Sub loadcarBtn_Click()
    playVdoControl.Enabled = True
    car_id = carList.ListIndex
    Dim car As New ADODB.Recordset
    q = "SELECT * FROM Cars WHERE ID = " & car_id + 1
    car.Open q, conn
    avatar = car!avatar
    BGPic.Picture = LoadPicture("Z:\MSJ\project\images\" & avatar)
    
    currentCar = car!name
    currentVdo = car!video
End Sub

'************************
' Database subrouteines (Functions)

Private Sub ConnectDatabase(ByVal database_path As String)
    conn.Open "Provider=Microsoft.jet.OLEDB.4.0;Data Source =" & database_path & ";persist security info=false"
End Sub



Private Sub playVdoControl_Click()
    Load vdoPlayerDlg
    vdoPlayerDlg.Show
    vdoPlayerDlg.Play_Video "Z:\MSJ\project\vdo\" & currentVdo, currentCar
End Sub
