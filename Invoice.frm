VERSION 5.00
Begin VB.Form Invoice 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   9450
   ClientLeft      =   2325
   ClientTop       =   465
   ClientWidth     =   15930
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9450
   ScaleWidth      =   15930
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox TxtDelMid 
      BackColor       =   &H00C0C000&
      Enabled         =   0   'False
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
      Left            =   4320
      TabIndex        =   34
      Top             =   8520
      Width           =   2535
   End
   Begin VB.TextBox TxtDelN 
      BackColor       =   &H00C0C000&
      Enabled         =   0   'False
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
      Left            =   4320
      TabIndex        =   33
      Top             =   6840
      Width           =   2535
   End
   Begin VB.TextBox TxtDelMno 
      BackColor       =   &H00C0C000&
      Enabled         =   0   'False
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
      Left            =   4320
      MaxLength       =   10
      TabIndex        =   32
      Top             =   7680
      Width           =   2535
   End
   Begin VB.TextBox TxtCategory 
      BackColor       =   &H00C0C000&
      Enabled         =   0   'False
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
      Left            =   12120
      TabIndex        =   27
      Top             =   2640
      Width           =   2415
   End
   Begin VB.CommandButton CmdExit 
      BackColor       =   &H0000FFFF&
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
      Left            =   12960
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   8880
      Width           =   1455
   End
   Begin VB.TextBox TxtPur 
      BackColor       =   &H00C0C000&
      Enabled         =   0   'False
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
      Left            =   13320
      TabIndex        =   23
      Top             =   240
      Width           =   2055
   End
   Begin VB.TextBox TxtCos 
      BackColor       =   &H00C0C000&
      Enabled         =   0   'False
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
      Left            =   12120
      TabIndex        =   21
      Top             =   7680
      Width           =   2415
   End
   Begin VB.TextBox TxtPay 
      BackColor       =   &H00C0C000&
      Enabled         =   0   'False
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
      Left            =   12120
      TabIndex        =   19
      Top             =   6840
      Width           =   2415
   End
   Begin VB.TextBox TxtMod 
      BackColor       =   &H00C0C000&
      Enabled         =   0   'False
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
      Left            =   12120
      TabIndex        =   16
      Top             =   5160
      Width           =   2415
   End
   Begin VB.TextBox TxtBra 
      BackColor       =   &H00C0C000&
      Enabled         =   0   'False
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
      Left            =   12120
      TabIndex        =   15
      Top             =   4320
      Width           =   2415
   End
   Begin VB.TextBox Txtmodid 
      BackColor       =   &H00C0C000&
      Enabled         =   0   'False
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
      Left            =   12120
      TabIndex        =   14
      Top             =   3480
      Width           =   2415
   End
   Begin VB.TextBox TxtCit 
      BackColor       =   &H00C0C000&
      Enabled         =   0   'False
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
      Left            =   4320
      TabIndex        =   13
      Top             =   5160
      Width           =   2535
   End
   Begin VB.TextBox TxtMob 
      BackColor       =   &H00C0C000&
      Enabled         =   0   'False
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
      Left            =   4320
      MaxLength       =   10
      TabIndex        =   12
      Top             =   4320
      Width           =   2535
   End
   Begin VB.TextBox TxtAdd 
      BackColor       =   &H00C0C000&
      Enabled         =   0   'False
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
      Left            =   4320
      TabIndex        =   11
      Top             =   3480
      Width           =   2535
   End
   Begin VB.TextBox TxtName 
      BackColor       =   &H00C0C000&
      Enabled         =   0   'False
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
      Left            =   4320
      TabIndex        =   10
      Top             =   2640
      Width           =   2535
   End
   Begin VB.Label Label20 
      BackStyle       =   0  'Transparent
      Caption         =   "Dealer Details :-"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   14.25
         Charset         =   0
         Weight          =   600
         Underline       =   -1  'True
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   495
      Left            =   240
      TabIndex        =   31
      Top             =   6000
      Width           =   2415
   End
   Begin VB.Label Label19 
      BackStyle       =   0  'Transparent
      Caption         =   "Dealer Mail Id :-"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   14.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Left            =   840
      TabIndex        =   30
      Top             =   8400
      Width           =   2775
   End
   Begin VB.Label Label18 
      BackStyle       =   0  'Transparent
      Caption         =   "Dealer Mobile No :-"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   14.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Left            =   840
      TabIndex        =   29
      Top             =   7680
      Width           =   2775
   End
   Begin VB.Label Label17 
      BackStyle       =   0  'Transparent
      Caption         =   "Dealer Name :-"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   14.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   495
      Left            =   840
      TabIndex        =   28
      Top             =   6840
      Width           =   2775
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "Category :- "
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   14.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   495
      Left            =   9120
      TabIndex        =   26
      Top             =   2640
      Width           =   2295
   End
   Begin VB.Label Label16 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Invoice"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   18
         Charset         =   0
         Weight          =   600
         Underline       =   -1  'True
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   495
      Left            =   6720
      TabIndex        =   25
      Top             =   120
      Width           =   1935
   End
   Begin VB.Label Label15 
      BackStyle       =   0  'Transparent
      Caption         =   "Date of Purchase :-"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   14.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   495
      Left            =   9960
      TabIndex        =   22
      Top             =   240
      Width           =   3015
   End
   Begin VB.Label Label14 
      BackStyle       =   0  'Transparent
      Caption         =   "Cost Paid :-"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   14.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   495
      Left            =   9000
      TabIndex        =   20
      Top             =   7680
      Width           =   2535
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "Payment Type :- "
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   14.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   495
      Left            =   9000
      TabIndex        =   18
      Top             =   6840
      Width           =   2535
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "Payment Details :-"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   14.25
         Charset         =   0
         Weight          =   600
         Underline       =   -1  'True
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   495
      Left            =   7080
      TabIndex        =   17
      Top             =   6000
      Width           =   2895
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Model :-"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   14.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   495
      Left            =   9120
      TabIndex        =   9
      Top             =   5160
      Width           =   2295
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Brand :-"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   14.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   495
      Left            =   9120
      TabIndex        =   8
      Top             =   4320
      Width           =   2295
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Model Id :- "
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   14.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   495
      Left            =   9120
      TabIndex        =   7
      Top             =   3480
      Width           =   2295
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Car Details :-"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   14.25
         Charset         =   0
         Weight          =   600
         Underline       =   -1  'True
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   495
      Left            =   7080
      TabIndex        =   6
      Top             =   1800
      Width           =   2415
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Cus City :- "
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   14.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   495
      Left            =   1320
      TabIndex        =   5
      Top             =   5160
      Width           =   2535
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Cus Mobile :-"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   14.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   495
      Left            =   1320
      TabIndex        =   4
      Top             =   4320
      Width           =   2535
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Cus Address :-"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   14.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   495
      Left            =   1320
      TabIndex        =   3
      Top             =   3480
      Width           =   2535
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Cus Name :- "
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   14.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   495
      Left            =   1320
      TabIndex        =   2
      Top             =   2640
      Width           =   2535
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Sold To :- "
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   14.25
         Charset         =   0
         Weight          =   600
         Underline       =   -1  'True
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   495
      Left            =   240
      TabIndex        =   1
      Top             =   1800
      Width           =   2175
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00808080&
      Caption         =   "The Best of The Best Auto Expo"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   18
         Charset         =   0
         Weight          =   600
         Underline       =   -1  'True
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   1095
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4575
   End
   Begin VB.Image Image1 
      Height          =   9480
      Left            =   0
      Picture         =   "Invoice.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   15960
   End
End
Attribute VB_Name = "Invoice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Sub Load_data(car_model_id, m_name, category, brand, ByRef customer_data As customer)
    TxtName.Text = customer_data.Name
    TxtCategory.Text = category
    Txtmodid.Text = car_model_id
    TxtBra.Text = brand
    TxtMod.Text = m_name
    TxtAdd.Text = customer_data.Address
    TxtMob.Text = customer_data.MobileNumber
    TxtCit.Text = customer_data.City
    TxtPur.Text = customer_data.DOP
    TxtDelN.Text = customer_data.Dealer_Name
    TxtDelMno.Text = customer_data.Dealer_Mob
    TxtDelMid.Text = customer_data.Dealer_Mid
End Sub

Public Sub Load_payment(amount, method)
    TxtPay.Text = method
    TxtCos.Text = amount
End Sub


Private Sub CmdExit_Click()
 Load feedbackfrm
 feedbackfrm.Show
End Sub
