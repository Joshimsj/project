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
      BackColor       =   &H8000000D&
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
      Left            =   7800
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   120
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H8000000D&
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
      Left            =   3600
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3840
      Width           =   1815
   End
   Begin VB.OptionButton Option2 
      BackColor       =   &H000080FF&
      Caption         =   "CHEQUE"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   14.25
         Charset         =   0
         Weight          =   600
         Underline       =   -1  'True
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5760
      TabIndex        =   3
      Top             =   3120
      Width           =   1815
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H000080FF&
      Caption         =   "Bank ACC"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   14.25
         Charset         =   0
         Weight          =   600
         Underline       =   -1  'True
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1200
      TabIndex        =   2
      Top             =   3120
      Width           =   1935
   End
   Begin VB.Label Label2 
      BackColor       =   &H000000FF&
      Caption         =   "Select Payment Mode :-"
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
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   1200
      Width           =   3495
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Payment"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   21.75
         Charset         =   0
         Weight          =   600
         Underline       =   -1  'True
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   735
      Left            =   3960
      TabIndex        =   0
      Top             =   120
      Width           =   2775
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

