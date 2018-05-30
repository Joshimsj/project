VERSION 5.00
Begin VB.Form MainForm 
   Caption         =   "Form1"
   ClientHeight    =   6480
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   9900
   LinkTopic       =   "Form1"
   ScaleHeight     =   6480
   ScaleWidth      =   9900
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox List1 
      Columns         =   2
      Height          =   2205
      Left            =   3720
      TabIndex        =   11
      Top             =   1080
      Width           =   1695
   End
   Begin VB.ListBox carModel 
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
      Height          =   2730
      Left            =   120
      TabIndex        =   10
      Top             =   3795
      Width           =   1905
   End
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
      Top             =   6690
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
      Height          =   2730
      Left            =   100
      TabIndex        =   1
      Top             =   550
      Width           =   1900
   End
   Begin VB.CommandButton playVdoControl 
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
      Left            =   120
      TabIndex        =   0
      Top             =   7440
      Width           =   1900
   End
   Begin VB.Label Model 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Car Model"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   9.75
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   345
      Index           =   1
      Left            =   120
      TabIndex        =   9
      Top             =   3360
      Width           =   1900
   End
   Begin VB.Label Brand 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Car Brand"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   9.75
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   345
      Index           =   0
      Left            =   120
      TabIndex        =   8
      Top             =   105
      Width           =   1905
   End
   Begin VB.Label topView 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
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
      ForeColor       =   &H80000008&
      Height          =   400
      Index           =   3
      Left            =   18500
      TabIndex        =   7
      ToolTipText     =   "Click to view Top View"
      Top             =   1400
      Width           =   1500
   End
   Begin VB.Label rearView 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Rear"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   14.25
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   400
      Index           =   2
      Left            =   18500
      TabIndex        =   6
      ToolTipText     =   "Click to view Top View"
      Top             =   900
      Width           =   1500
   End
   Begin VB.Label sideAView 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Side A"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   14.25
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   400
      Index           =   0
      Left            =   18500
      TabIndex        =   5
      ToolTipText     =   "Click to view Top View"
      Top             =   1900
      Width           =   1500
   End
   Begin VB.Label sideBView 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Side B"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   14.25
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   400
      Index           =   1
      Left            =   18500
      TabIndex        =   4
      ToolTipText     =   "Click to view Top View"
      Top             =   2400
      Width           =   1500
   End
   Begin VB.Label frontView 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Front"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   14.25
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   400
      Left            =   18500
      TabIndex        =   3
      ToolTipText     =   "Click to view Top View"
      Top             =   400
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

Private Sub Form_Load()
    Me.WindowState = 2
    BGPic.Top = 0
    BGPic.Left = 0
    BGPic.Height = Me.Height
    BGPic.Width = Me.Width
    
    currentVdo = ""
    loadcarBtn.Enabled = False
    
    'Connect to database
     ConnectDatabase "Z:\MSJ\project\assets\cars.mdb"
    
    'Load Carlist
     Dim cars As New ADODB.Recordset
     query = "SELECT name FROM cars"
     
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
    q = "SELECT * FROM car_details WHERE id = " & car_id + 1
    car.Open q, conn, adUseClient, adLockOptimistic, adCmdText
    
    avatar = car!avatar
    BGPic.Picture = LoadPicture("Z:\MSJ\project\images\" & avatar)
    
    currentCar = car!name
    currentVdo = car!video
    
    car.Close
End Sub

Private Sub carList_Click()
    loadcarBtn.Enabled = True
    
    carModel.Clear
    carModel.Refresh
    
    car_id = carList.ListIndex
    
    Dim car As New ADODB.Recordset
    q = "SELECT model FROM car_details WHERE car_id = " & car_id + 1
    car.Open q, conn, adUseClient, adLockOptimistic, adCmdText
    
    'avatar = car!avatar
    'BGPic.Picture = LoadPicture("Z:\MSJ\project\images\" & avatar)
    total_models = car.RecordCount
    
    For i = 0 To total_models - 1
        carModel.AddItem car.Fields(0).Value, i
        car.MoveNext
    Next i
    car.Close
End Sub

Private Sub carModel_Click()
    playVdoControl.Enabled = True
    car_id = carModel.ListIndex
    Dim car As New ADODB.Recordset
    q = "SELECT * FROM car_details WHERE id = " & car_id + 1
    car.Open q, conn, adUseClient, adLockOptimistic, adCmdText
    
    avatar = car!avatar
    BGPic.Picture = LoadPicture("Z:\MSJ\project\images\" & avatar)
    
    currentCar = car!name
    currentVdo = car!video
    
    car.Close
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
