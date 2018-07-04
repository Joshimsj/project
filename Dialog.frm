VERSION 5.00
Begin VB.Form Dialog 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Selection "
   ClientHeight    =   6000
   ClientLeft      =   5085
   ClientTop       =   2355
   ClientWidth     =   8895
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6000
   ScaleWidth      =   8895
   Begin VB.CommandButton Search 
      Caption         =   "Search"
      Height          =   495
      Left            =   240
      TabIndex        =   9
      Top             =   5160
      Width           =   1335
   End
   Begin VB.ComboBox PriceRange 
      Height          =   315
      Left            =   4800
      TabIndex        =   8
      Text            =   "Price Range"
      Top             =   2040
      Width           =   3495
   End
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   4800
      TabIndex        =   7
      Text            =   "Text2"
      Top             =   3840
      Width           =   3615
   End
   Begin VB.TextBox Speed 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   16393
         SubFormatType   =   1
      EndProperty
      Height          =   495
      Left            =   4800
      TabIndex        =   6
      Top             =   2880
      Width           =   3615
   End
   Begin VB.ComboBox CarCategory 
      Height          =   315
      Left            =   4800
      TabIndex        =   5
      Text            =   "Car Category"
      Top             =   960
      Width           =   3495
   End
   Begin VB.Label Label5 
      Caption         =   "Brand name"
      Height          =   615
      Left            =   360
      TabIndex        =   4
      Top             =   3840
      Width           =   3375
   End
   Begin VB.Label Label4 
      Caption         =   "Speed"
      Height          =   615
      Left            =   360
      TabIndex        =   3
      Top             =   2880
      Width           =   3375
   End
   Begin VB.Label Label3 
      Caption         =   "Price Range"
      Height          =   615
      Left            =   360
      TabIndex        =   2
      Top             =   1920
      Width           =   3375
   End
   Begin VB.Label Label2 
      Caption         =   "Select Category"
      Height          =   495
      Left            =   360
      TabIndex        =   1
      Top             =   1080
      Width           =   3375
   End
   Begin VB.Label Label1 
      Caption         =   "Selection of your Choice...!!!"
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
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   4455
   End
   Begin VB.Image Image1 
      Height          =   6015
      Left            =   0
      Top             =   0
      Width           =   8895
   End
End
Attribute VB_Name = "Dialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim price_range(2) As Double

Dim conn As New ADODB.Connection
Dim carBrands() As Integer

Private Sub Form_Load()
    CarCategory.AddItem "Sports", 0
    CarCategory.AddItem "Vintage", 1
    CarCategory.AddItem "Luxury", 2
    CarCategory.AddItem "Hybrid", 3
    CarCategory.AddItem "Evision", 4
    
    PriceRange.AddItem "Bellow 20 lac", 0
    price_range(0) = 2000000
    
    PriceRange.AddItem "Between 30 lac - 70 lac", 1
    price_range(1) = 7000000
    
    PriceRange.AddItem "Above 3 cr", 2
    price_range(2) = 30000000
    
    'Connect to database
    ConnectDatabase "E:\project\assets\cars3.mdb"
       
End Sub

Private Sub Search_Click()
    Dim category As Integer
    Dim range As Integer
    Dim total_cars As Integer
    Dim i As Integer
    Dim filter_query As String
    
    category = CarCategory.ListIndex
    range = PriceRange.ListIndex
    
    '' debug code for empty input
    
    filter_query = " cost > " & price_range(range)
    
    Select Case range
        Case 0
            filter_query = " cost < " & price_range(range)
        Case 1
            filter_query = " cost > " & price_range(0) & " AND cost < " & price_range(1)
        Case 2
            filter_query = " cost > " & price_range(range)
    End Select
    
    'Load Carlist
     Dim cars As New ADODB.Recordset
     
     cars.Open "SELECT DISTINCT car_id, brand FROM cars WHERE category = " & category & " AND " & filter_query, conn, adUseClient, adLockOptimistic, adCmdText
     

    'Append carlist
        
    'Total cars in Database
    total_cars = cars.RecordCount
    ReDim carBrands(total_cars) As Integer
    
    Load MainForm
    
    If total_cars > 0 Then
        For i = 0 To total_cars - 1
            'Check for empty value
            MainForm.carList.AddItem cars.Fields(1).Value, i
            carBrands(i) = cars.Fields(0).Value
            cars.MoveNext
        Next i
    End If
    
    MainForm.loadBrands carBrands, total_cars
    MainForm.Show
    
    
    'carList.AddItem cars.Fields(1)
    cars.Close
End Sub

Private Sub ConnectDatabase(ByVal database_path As String)
    conn.Open "Provider=Microsoft.jet.OLEDB.4.0;Data Source =" & database_path & ";persist security info=false"
End Sub
