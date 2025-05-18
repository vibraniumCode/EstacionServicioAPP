VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClaseCombustible"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mCombustible As String
Private mPrecio As Double

' Propiedad
Public Property Get Combustible() As String
    Combustible = mCombustible
End Property

Public Property Let Combustible(ByVal Valor As String)
    mCombustible = Valor
End Property

Public Property Get Precio() As Double
    Precio = mPrecio
End Property

Public Property Let Precio(ByVal Valor As Double)
    mPrecio = Valor
End Property
