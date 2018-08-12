VERSION 5.00
Begin VB.Form Selfrom 
   BorderStyle     =   0  'None
   Caption         =   "Payment Option"
   ClientHeight    =   4605
   ClientLeft      =   6210
   ClientTop       =   2880
   ClientWidth     =   8775
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4605
   ScaleWidth      =   8775
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command2 
      Caption         =   "X"
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
      Left            =   7560
      TabIndex        =   5
      Top             =   3840
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Submit"
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
      Left            =   5760
      TabIndex        =   4
      Top             =   3840
      Width           =   1575
   End
   Begin VB.OptionButton Option2 
      BackColor       =   &H000040C0&
      Caption         =   "Cheque "
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
      TabIndex        =   3
      Top             =   2280
      Width           =   2895
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H000040C0&
      Caption         =   "Bank Acc"
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
      Left            =   2400
      TabIndex        =   2
      Top             =   2280
      Width           =   2655
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Select your Choice"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   14.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   375
      Left            =   360
      TabIndex        =   1
      Top             =   1320
      Width           =   3135
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Payment Option"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   14.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   615
      Left            =   3600
      TabIndex        =   0
      Top             =   120
      Width           =   2415
   End
   Begin VB.Image Image1 
      Height          =   4575
      Left            =   0
      Picture         =   "payment.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   8775
   End
End
Attribute VB_Name = "Selfrom"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim amount As String
Dim model_id As Integer
Dim model_name As String
Dim category As String
Dim brand As String


Private Sub Command1_Click()
If Option1.Value = True Then
Selfrom.Visible = False
Accfrm.Visible = True
Chqfrm.Visible = False
Else
Selfrom.Visible = False
Accfrm.Visible = False
Chqfrm.Visible = True
End If
End Sub

Private Sub Command2_Click()
Invoice.Show
End Sub

Public Sub Load_data(Price, car_category, car_brand, car_model_id, m_name, ByRef customer_object As customer)
    amount = Price
    category = car_category
    brand = car_brand
    model_id = car_model_id
    model_name = m_name
    
    Load Invoice
    Invoice.Load_data car_model_id, model_name, category, brand, customer_object
       
    
    Load Accfrm
    Accfrm.Load_Amount Price
    
    Load Chqfrm
    Chqfrm.Load_Amount Price
End Sub

