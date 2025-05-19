VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClaseImpuesto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mImpuesto As String
Private mMonto As Double
Private mFechaOperacion As Date

' Propiedad
Public Property Get Impuesto() As String
    Impuesto = mImpuesto
End Property

Public Property Let Impuesto(ByVal Valor As String)
    mImpuesto = Valor
End Property

Public Property Get Monto() As Double
    Monto = mMonto
End Property

Public Property Let Monto(ByVal Valor As Double)
    mMonto = Valor
End Property

Public Property Get FechaOperacion() As Date
    FechaOperacion = mFechaOperacion
End Property

Public Property Let FechaOperacion(ByVal Valor As Date)
    mFechaOperacion = Valor
End Property


