VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClaseProducto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mPrecioNeto As Double
Private mPrecioUnitario As Double
Private mCantidad As Integer
Private mDescripcion As String


' Propiedad para el precio unitario
Public Property Get precioNeto() As Double
    precioNeto = mPrecioNeto
End Property

Public Property Let precioNeto(ByVal Valor As Double)
    mPrecioNeto = Valor
End Property

' Propiedad para el precio unitario
Public Property Get PrecioUnitario() As Double
    PrecioUnitario = mPrecioUnitario
End Property

Public Property Let PrecioUnitario(ByVal Valor As Double)
    mPrecioUnitario = Valor
End Property

' Propiedad para la cantidad
Public Property Get Cantidad() As Integer
    Cantidad = mCantidad
End Property

Public Property Let Cantidad(ByVal Valor As Integer)
    If Valor < 0 Then
        mCantidad = 0
    Else
        mCantidad = Valor
    End If
End Property

'Propiedad para la descripcion
Public Property Get Descripcion() As String
    Descripcion = mDescripcion
End Property

Public Property Let Descripcion(ByVal Valor As String)
    mDescripcion = Valor
End Property

' Calcular el precio neto
Public Function CalcularPrecioNeto() As Double
    If mCantidad = 0 Then
        MostrarAlerta "La cantidad no puede ser cero. Ingrese un valor válido."
        CalcularPrecioNeto = 0
        Exit Function
    End If
    
    CalcularPrecioNeto = mPrecioUnitario * mCantidad
End Function
