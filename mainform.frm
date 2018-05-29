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
   Begin VB.CommandButton loadcarBtn 
      Caption         =   "View"
      Height          =   495
      Left            =   240
      TabIndex        =   2
      Top             =   5160
      Width           =   1815
   End
   Begin VB.ListBox CarList1 
      Height          =   4740
      Left            =   240
      TabIndex        =   1
      Top             =   120
      Width           =   1815
   End
   Begin VB.CommandButton playVdoControl 
      Caption         =   "play"
      Height          =   615
      Left            =   18720
      TabIndex        =   0
      Top             =   240
      Width           =   1335
   End
   Begin VB.Image BGPic 
      Height          =   16200
      Left            =   0
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
    
    'Connect to database
     ConnectDatabase "C:\Users\MSJ\Desktop\VB_Smart\c_details.mdb"
    
    'Load Carlist
     Dim cars As New ADODB.Recordset
     query = "SELECT name FROM Cars"
     
     cars.Open query, conn
     

    'Append carlist
    'CarList1.AddItem cars!name
    cars.Close

    CarList1.AddItem "Aston Martin", 0
    CarList1.AddItem "Aston Martin", 1
    CarList1.AddItem "Audi", 2
    playVdoControl.Enabled = False
End Sub

Private Sub Form_Resize()
    BGPic.Height = Me.Height
    BGPic.Width = Me.Width
End Sub

Private Sub loadcarBtn_Click()
    playVdoControl.Enabled = True
    car_id = CarList1.ListIndex
    Dim cars As New ADODB.Recordset
    q = "SELECT * FROM Cars WHERE ID = " & car_id + 1
    cars.Open q, conn
    avatar = cars!avatar
    BGPic.Picture = LoadPicture("Z:\MSJ\project\images\" & avatar)
    
    currentCar = cars!name
    currentVdo = cars!video
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