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

Private pLitros As Double
Private pPrecioUnitario As Double
Private pPrecioNeto As Double
Private pMontoITC As Double


' Propiedades públicas
Public Property Let PrecioUnitario(ByVal value As Double)
    pPrecioUnitario = value
End Property

Public Property Get PrecioUnitario() As Double
    PrecioUnitario = pPrecioUnitario
End Property

Public Property Let litros(ByVal value As Double)
    pLitros = value
End Property

Public Property Get litros() As Double
    litros = pLitros
End Property

Public Property Let precioNeto(ByVal value As Double)
    pPrecioNeto = value
End Property

Public Property Get precioNeto() As Double
    precioNeto = pPrecioNeto
End Property

Public Property Let MontoITC(ByVal value As Double)
    pMontoITC = value
End Property

Public Property Get MontoITC() As Double
    MontoITC = pMontoITC
End Property

' Métodos de cálculo
Public Function CalcularPrecioNeto() As Double
    CalcularPrecioNeto = Me.litros * Me.PrecioUnitario
End Function

Public Function CalcularIVA() As Double
    CalcularIVA = pPrecioNeto * 0.21
End Function

Public Function CalcularITC() As Double
    CalcularITC = pLitros * pMontoITC
End Function

Public Function CalcularTotal() As Double
    CalcularTotal = pPrecioNeto + CalcularIVA() + CalcularITC()
End Function

