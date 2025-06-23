Attribute VB_Name = "ModuloUtilidades"
Option Explicit

Global nroFactura As String
Global codFactura As String
Global idSeleccionado As String
Global executeSQL As String
Global idEstacionEmpresa As Integer

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

Public Sub PintarFilasAlternadasFlex(ByRef grilla As MSFlexGrid)
    Dim i As Integer, j As Integer
    With grilla
        .GridLines = flexGridFlat
        .GridColor = RGB(200, 200, 200)
        .GridColorFixed = RGB(160, 160, 160)
        .BackColorFixed = RGB(220, 220, 220)
        .ForeColorFixed = vbBlack
        .CellAlignment = flexAlignLeftCenter
        
        For i = .FixedRows To .Rows - 1
            For j = 0 To .Cols - 1
                .Row = i
                .col = j
                .CellBackColor = IIf(i Mod 2 = 0, RGB(240, 240, 240), vbWhite)
                .CellAlignment = flexAlignLeftCenter
            Next j
        Next i
    End With
End Sub

