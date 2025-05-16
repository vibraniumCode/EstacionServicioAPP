VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClaseCliente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mCliente As String
Private mCUIT As String
Private mDireccion As String

Public Property Get Cliente() As String
    Cliente = mCliente
End Property

Public Property Get Cuit() As String
    Cuit = mCUIT
End Property

Public Property Get Direccion() As String
    Direccion = mDireccion
End Property

Public Property Let Cliente(ByVal Valor As String)
    mCliente = Valor
End Property

Public Property Let Cuit(ByVal Valor As String)
    mCUIT = Valor
End Property

Public Property Let Direccion(ByVal Valor As String)
    mDireccion = Valor
End Property

