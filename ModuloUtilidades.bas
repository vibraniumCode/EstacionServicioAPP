Attribute VB_Name = "ModuloUtilidades"
Option Explicit

Global nroFactura As Long

'Public Function ValidadorMsg() As Boolean
'If Facturacion.txtNombre.Text = "" Then
'    MsgBox "Ingrese el nombre del cliente", vbInformation, "Doom"
'    ValidadorMsg = True
'ElseIf Facturacion.txtCuit.Text = "" Then
'    MsgBox "Ingrese el C.U.I.T. del cliente", vbInformation, "Doom"
'    ValidadorMsg = True
'ElseIf Facturacion.txtDireccion.Text = "" Then
'    MsgBox "Ingrese la dirección del cliente", vbInformation, "Doom"
'    ValidadorMsg = True
'Else
'    ValidadorMsg = False
'End If
'End Function


' Convierte un valor de texto en número limpio
Public Function LimpiarValor(ByVal texto As String) As Double
    Dim valorLimpio As String
    valorLimpio = Replace(texto, "$", "")
    valorLimpio = Replace(valorLimpio, ",", "")
    valorLimpio = Trim(valorLimpio)
    
    If IsNumeric(valorLimpio) Then
        LimpiarValor = Val(valorLimpio)
    Else
        LimpiarValor = 0
    End If
End Function

' Formatea un número como precio
Public Function FormatoPrecio(ByVal Valor As Double) As String
    FormatoPrecio = "$" & Format$(Valor, "#,##0.00")
End Function

' Muestra una alerta estándar
Public Sub MostrarAlerta(ByVal mensaje As String)
    MsgBox mensaje, vbExclamation, "Advertencia"
End Sub

Public Function FormatearNumeroFactura(ByVal numero As Long) As String
    FormatearNumeroFactura = Format(numero, "00000000")
End Function
