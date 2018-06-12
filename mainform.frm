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
   Begin VB.ComboBox brandType 
      Height          =   315
      Left            =   120
      TabIndex        =   10
      Text            =   "Combo1"
      Top             =   120
      Width           =   1900
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
      TabIndex        =   9
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
      TabIndex        =   8
      Top             =   3360
      Width           =   1900
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
Dim currentCategory As Integer

Dim currentModels() As Integer
Dim carBrands() As Integer

Private Sub brandType_Click()
    Index = brandType.ListIndex
    
    'Sets the current category on clipboard
    currentCategory = Index
    
    'Load Carlist
     Dim cars As New ADODB.Recordset
     query = "SELECT DISTINCT car_id, brand FROM cars WHERE category = " & Index
     
     cars.Open query, conn, adUseClient, adLockOptimistic, adCmdText
     

    'Append carlist
        
    'Total cars in Database
    total_cars = cars.RecordCount
    ReDim carBrands(total_cars) As Integer
    
    carList.Clear
    carList.Refresh
    
    If total_cars > 0 Then
    
        For i = 0 To total_cars - 1
            'Check for empty value
            carList.AddItem cars.Fields(1).Value, i
            carBrands(i) = cars.Fields(0).Value
            cars.MoveNext
        Next i
    End If
    
    'carList.AddItem cars.Fields(1)
    cars.Close
End Sub

Private Sub carModel_Click()
    loadcarBtn.Enabled = True
        
    Index = carModel.ListIndex
    model_id = currentModels(Index)
    
    Dim car As New ADODB.Recordset
    q = "SELECT * FROM cars WHERE model_id = " & model_id
    car.Open q, conn, adUseClient, adLockOptimistic, adCmdText
    
    If IsNull(car!display_pic) Then
        MsgBox ("No Pics found for this car")
    Else
        BGPic.Picture = LoadPicture("Z:\MSJ\project\images\" & car!display_pic)
    End If
    
    currentCar = car!brand & "(" & car!Model & ")"
    
    If IsNull(car!video) Then
    MsgBox ("No Vdo found for this car")
       playVdoControl.Enabled = False
    Else
        playVdoControl.Enabled = True
        currentVdo = car!video
    End If
    
    car.Close
End Sub

Private Sub Form_Load()
    Me.WindowState = 2
    BGPic.Top = 0
    BGPic.Left = 0
    BGPic.Height = Me.Height
    BGPic.Width = Me.Width
    
    currentVdo = ""
    loadcarBtn.Enabled = False
    
    'carTypeTables = Array("sport", "vintage", "luxury", "hybrid", "evision")
    
    brandType.AddItem "Sports", 0
    brandType.AddItem "Vintage", 1
    brandType.AddItem "Luxury", 2
    brandType.AddItem "Hybrid", 3
    brandType.AddItem "Evision", 4
    
    
    'Connect to database
    ConnectDatabase "Z:\MSJ\project\assets\cars3.mdb"
    
    playVdoControl.Enabled = False
End Sub

Private Sub Form_Resize()
    BGPic.Height = Me.Height
    BGPic.Width = Me.Width
End Sub

Private Sub frontView_Click()
    Dim pic As New ADODB.Recordset
        
    Index = carModel.ListIndex
    model_id = currentModels(Index)
    
    Dim car As New ADODB.Recordset
    q = "SELECT pic FROM pictures WHERE model_id = " & model_id & " AND view = 'front'"
    car.Open q, conn, adUseClient, adLockOptimistic, adCmdText
    
    If IsNull(car!pic) Then
        MsgBox ("No Pictures found for this car")
    Else
        BGPic.Picture = LoadPicture("Z:\MSJ\project\images\" & car!pic)
    End If
    
End Sub

Private Sub loadcarBtn_Click()
    Load carConfiguration
    carConfiguration.Show
End Sub


Private Sub carList_Click()

    Index = carList.ListIndex
    
    Dim car As New ADODB.Recordset
    
    q = "SELECT model_id, model FROM cars WHERE car_id = " & carBrands(Index) & " AND category = " & currentCategory
    car.Open q, conn, adUseClient, adLockOptimistic, adCmdText
    
    'avatar = car!avatar
    'BGPic.Picture = LoadPicture("Z:\MSJ\project\images\" & avatar)
    total_models = car.RecordCount
    
    ReDim currentModels(total_models) As Integer
    carModel.Clear
    carModel.Refresh
    
    For i = 0 To total_models - 1
        currentModels(i) = car.Fields(0).Value
        carModel.AddItem car.Fields(1).Value, i
        car.MoveNext
    Next i
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
