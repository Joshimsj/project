VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "customer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private cName As String
Private cMobileNumber As String
Private cAddress As String
Private cCompany As String
'Private cIdentity As String
'Private cEnterDetails As String
Private cState As String
Private cCity As String
Private cZip As String
Private cDOB As String
Public Property Let Name(customerName As String)
    cName = customerName
End Property

Public Property Let MobileNumber(customerContact As String)
    cMobileNumber = customerContact
End Property

Public Property Let Address(customerAddress As String)
    cAddress = customerAddress
End Property

Public Property Let Company(customerCompany As String)
    cCompany = customerCompany
End Property

'Public Property Let Identity(customerIdentity As String)
    'cIdentity = customerIdentity
'End Property

'Public Property Let EnterDetails(customerEnterDetails As String)
    'cEnterDetails = customerEnterDetails
'End Property

Public Property Let State(customerState As String)
    cState = customerState
End Property

Public Property Let City(customerCity As String)
    cCity = customerCity
End Property

Public Property Let Zip(customerZip As String)
    cZip = customerZip
End Property

Public Property Let DOB(customerDOB As String)
    cDOB = customerDOB
End Property

Public Property Get Name() As String
    Name = cName
End Property

Public Property Get MobileNumber() As String
    MobileNumber = cMobileNumber
End Property

Public Property Get Address() As String
    Address = cAddress
End Property

Public Property Get Company() As String
    Company = cCompany
End Property

'Public Property Get Identity() As String
    'Identity = cIdentity
'End Property

'Public Property Get EnterDetails() As String
    'EnterDetails = cEnterDetails
'End Property

Public Property Get State() As String
    State = cState
End Property

Public Property Get City() As String
    City = cCity
End Property

Public Property Get Zip() As String
    Zip = cZip
End Property

Public Property Get DOB() As String
    DOB = cDOB
End Property

Public Sub SaveData()
'create an adodb conn, adodb recordset
Dim conn As New ADODB.Connection
Dim rec As New ADODB.Recordset
Dim s As String
'connect
 conn.Open "Provider=Microsoft.jet.OLEDB.4.0;Data Source =E:\project\assets\Customers.mdb;persist security info=false"
'query
  's = "INSERT INTO cus_details (name, company, contact, dob, identity, address, state, city, pincode) VALUES (" & Me.Name & "," & Me.Company & "," & Me.MobileNumber & "," & Me.DOB & "," & Me.Identity & ", " & Me.Address & ", " & Me.State & "," & Me.City & "," & Me.Zip & ")"
   s = "INSERT INTO cus_details (name, company, contact, DOB, address, state, city, pincode) VALUES('" & Me.Name & "','" & Me.Company & "','" & Me.MobileNumber & "','" & Me.DOB & "','" & Me.Address & "','" & Me.State & "','" & Me.City & "','" & Me.Zip & "')"

   rec.Open s, conn, adOpenDynamic, adLockOptimistic
   conn.Close
   MsgBox "Customer Saved", vbInformation
End Sub
