VERSION 5.00
Begin VB.Form Selection 
   BorderStyle     =   0  'None
   Caption         =   "Selection "
   ClientHeight    =   3300
   ClientLeft      =   5040
   ClientTop       =   1980
   ClientWidth     =   8670
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3300
   ScaleWidth      =   8670
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Search 
      Caption         =   "Search"
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
      Left            =   3720
      TabIndex        =   5
      Top             =   2640
      Width           =   1335
   End
   Begin VB.ComboBox PriceRange 
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   14.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   4800
      TabIndex        =   4
      Text            =   "Price Range"
      Top             =   1920
      Width           =   3495
   End
   Begin VB.ComboBox CarCategory 
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   14.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   4800
      TabIndex        =   3
      Text            =   "Car Category"
      Top             =   1200
      Width           =   3495
   End
   Begin VB.Label Label3 
      Caption         =   "Price Range"
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
      TabIndex        =   2
      Top             =   1920
      Width           =   3375
   End
   Begin VB.Label Label2 
      Caption         =   "Select Category"
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
      TabIndex        =   1
      Top             =   1200
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
      Height          =   3375
      Left            =   0
      Top             =   0
      Width           =   8655
   End
End
Attribute VB_Name = "Selection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim price_range(2) As Double

Dim conn As New ADODB.Connection
Dim carBrands() As Integer
Dim brandNames() As String

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
    ReDim brandNames(total_cars) As String
    
    If total_cars > 0 Then
        Load MainForm
        For i = 0 To total_cars - 1
            'Check for empty value
            MainForm.carList.AddItem cars.Fields(1).Value, i
            carBrands(i) = cars.Fields(0).Value
            brandNames(i) = cars.Fields(1).Value
            cars.MoveNext
        Next i
        
        MainForm.loadFilters filter_query, category
        MainForm.loadBrands carBrands, brandNames, total_cars
        MainForm.Show
    Else
        MsgBox ("No Cars Found")
    End If
    
    
    cars.Close
End Sub

Private Sub ConnectDatabase(ByVal database_path As String)
    conn.Open "Provider=Microsoft.jet.OLEDB.4.0;Data Source =" & database_path & ";persist security info=false"
End Sub
