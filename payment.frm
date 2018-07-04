VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5100
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   8100
   LinkTopic       =   "Form1"
   ScaleHeight     =   5100
   ScaleWidth      =   8100
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   495
      Left            =   6120
      TabIndex        =   2
      Top             =   3240
      Width           =   975
   End
   Begin VB.ComboBox Payment_type 
      Height          =   315
      Left            =   3120
      TabIndex        =   0
      Text            =   "Combo1"
      Top             =   1440
      Width           =   3495
   End
   Begin VB.Label Label1 
      Caption         =   "Payment Type"
      Height          =   735
      Left            =   360
      TabIndex        =   1
      Top             =   1440
      Width           =   2295
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdSave_Click()
MsgBox (Payment_type.ItemData(0))
End Sub

Private Sub Form_Load()
Payment_type.AddItem "NetBanking", Val("Net Banking")
'Payment_type.ItemData(0) = Val("NetBanking")
Payment_type.AddItem "Debit card", Val("Net Banking")
'Payment_type.ItemData(1) = Val("Debit Card")
Payment_type.AddItem "Credit card", Val("Net Banking")
'Payment_type.ItemData(2) = Val("Credit Card")
End Sub
