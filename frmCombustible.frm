VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmCombustible 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Combustible"
   ClientHeight    =   3135
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   9135
   Icon            =   "frmCombustible.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3135
   ScaleWidth      =   9135
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Estacion de servicio"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   2895
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   8895
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
         Height          =   2535
         Left            =   4680
         TabIndex        =   8
         Top             =   300
         Width           =   4095
         _ExtentX        =   7223
         _ExtentY        =   4471
         _Version        =   393216
         SelectionMode   =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.CommandButton btnActualizar 
         Caption         =   "&Actualizar"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   1560
         Picture         =   "frmCombustible.frx":08CA
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   2160
         Width           =   1335
      End
      Begin VB.CommandButton btnSalir 
         Caption         =   "&Volver"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   3000
         Picture         =   "frmCombustible.frx":0E54
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   2160
         Width           =   1335
      End
      Begin VB.Frame Frame3 
         Caption         =   "Precio Actual"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   120
         TabIndex        =   2
         Top             =   1200
         Width           =   4455
         Begin VB.TextBox txtPrecioActual 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            TabIndex        =   4
            Text            =   "$00.00"
            Top             =   360
            Width           =   4215
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Tipo"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   4455
         Begin VB.TextBox txtTipo 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            TabIndex        =   3
            Top             =   360
            Width           =   4215
         End
      End
      Begin VB.CommandButton btnIngresar 
         Caption         =   "&Ingresar"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   120
         Picture         =   "frmCombustible.frx":13DE
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   2160
         Width           =   1335
      End
   End
   Begin VB.Menu mnuListView 
      Caption         =   "Menu"
      Visible         =   0   'False
      Begin VB.Menu mEliminar 
         Caption         =   "Eliminar"
      End
      Begin VB.Menu mActualizar 
         Caption         =   "Actualizar"
      End
   End
End
Attribute VB_Name = "frmCombustible"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Combustibles As New ClaseCombustible
Dim idCombustible As Integer

Private Sub btnActualizar_Click()
    Dim rs As New ADODB.Recordset
    
    If DatosValidador Then Exit Sub
    
    ' Conectar a la base de datos utilizando el módulo de conexión
    Call ConectarBD

    On Error GoTo ErrHandler
    
    rs.Open "exec sp_OperacionCombustible 'MODIFICAR'," & idCombustible & ",'" & txtTipo.Text & "','" & txtPrecioActual.Text & "'", conn, adOpenStatic, adLockReadOnly
    MsgBox rs(0), vbInformation, "ESAPP"
    
    If rs.State = adStateOpen Then rs.Close
    If conn.State = adStateOpen Then Call DesconectarBD
    Call CargarlvCombustible
    Call LimpiarCampos
    
    btnActualizar.Enabled = False
    btnIngresar.Enabled = True
    
    Exit Sub

ErrHandler:
    MsgBox "Error al cargar datos: " & Err.Description, vbCritical, "Error"
    If rs.State = adStateOpen Then rs.Close
    If conn.State = adStateOpen Then Call DesconectarBD
End Sub

Private Sub btnIngresar_Click()
Dim rs As New ADODB.Recordset
    
    If DatosValidador Then Exit Sub
    
    ' Conectar a la base de datos utilizando el módulo de conexión
    Call ConectarBD

    On Error GoTo ErrHandler
    
    With Combustibles
    rs.Open "exec sp_OperacionCombustible 'INSERTAR',NULL,'" & .Combustible & "'," & .Precio, conn, adOpenStatic, adLockReadOnly
    End With
    
    MsgBox rs(0), vbInformation, "ESAPP"
    
    ' Cerrar el recordset
    If rs.State = adStateOpen Then rs.Close
    If conn.State = adStateOpen Then Call DesconectarBD
    Call CargarlvCombustible
    Call LimpiarCampos
    Exit Sub

ErrHandler:
    MsgBox "Error al cargar datos: " & Err.Description, vbCritical, "Error"
    If rs.State = adStateOpen Then rs.Close
    If conn.State = adStateOpen Then Call DesconectarBD
End Sub

Public Function DatosValidador() As Boolean
If txtTipo.Text = "" Then
    MsgBox "Ingrese el tipo del combustible", vbInformation, "ESAPP"
    DatosValidador = True
ElseIf txtPrecioActual.Text = "" Then
    MsgBox "Ingrese el precio actual del combustible", vbInformation, "ESAPP"
    DatosValidador = True
Else
    DatosValidador = False
End If
End Function

Private Sub LimpiarCampos()
    txtTipo.Text = ""
    txtPrecioActual.Text = ""
End Sub

Private Sub btnSalir_Click()
Unload Me
End Sub

Private Sub Form_Load()
With MSFlexGrid1
    .Rows = 3
    .Cols = 3
    .FixedRows = 2
    .FixedCols = 0
    .TextMatrix(0, 0) = "id"
    .TextMatrix(0, 1) = "Tipo"
    .TextMatrix(0, 2) = "Precio"
    .ColWidth(0) = 800
    .ColWidth(1) = 4500
    .ColWidth(1) = 3500
End With

' Cargar datos en el ListView
Call CargarlvCombustible
End Sub

Private Sub CargarlvCombustible()
    Dim rs As New ADODB.Recordset
    Dim fila As Integer

    ' Conectar a la base de datos utilizando el módulo de conexión
    Call ConectarBD

    ' Obtener datos de la base
    On Error GoTo ErrHandler

    rs.Open "SELECT id, tipo, precio FROM Combustible", conn, adOpenStatic, adLockReadOnly

    ' Configurar columnas del MSFlexGrid
    With MSFlexGrid1
        .Clear
        .Rows = 1 ' Solo cabecera
        .Cols = 3
        .TextMatrix(0, 0) = "Id"
        .TextMatrix(0, 1) = "Tipo"
        .TextMatrix(0, 2) = "Precio"

        ' Agregar datos fila por fila
        Do While Not rs.EOF
            .Rows = .Rows + 1
            fila = .Rows - 1
            .TextMatrix(fila, 0) = rs("id")
            .TextMatrix(fila, 1) = UCase(rs("tipo"))
            .TextMatrix(fila, 2) = UCase(rs("precio"))
            rs.MoveNext
        Loop
    End With
    With MSFlexGrid1
        .ColWidth(0) = 700 ' Ancho del Código
        .ColWidth(1) = 2280
        .ColWidth(2) = 1000
    End With
    
    If rs.State = adStateOpen Then rs.Close
    If conn.State = adStateOpen Then Call DesconectarBD
    
     ' Pintar filas alternadas
    Call PintarFilasAlternadasFlex(MSFlexGrid1)
    Exit Sub

ErrHandler:
    MsgBox "Error al cargar datos: " & Err.Description, vbCritical, "Error"
    If rs.State = adStateOpen Then rs.Close
    If conn.State = adStateOpen Then Call DesconectarBD
End Sub

Private Sub lvCombustible_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbRightButton Then
    PopupMenu mnuListView
End If
End Sub

Private Sub mActualizar_Click()

On Error Resume Next

If Err.Number <> 0 Then
    MsgBox "Por favor, selecciione un tipo de combustible para actualizar", vbExclamation, "ESAPP"
    Exit Sub
End If
On Error GoTo 0

Call CargarDatosBox

btnActualizar.Enabled = True
btnIngresar.Enabled = False

End Sub

Private Sub CargarDatosBox()

MSFlexGrid1.col = 0
idCombustible = MSFlexGrid1.Text
MSFlexGrid1.col = 1
txtTipo.Text = MSFlexGrid1.Text
MSFlexGrid1.col = 2
txtPrecioActual.Text = MSFlexGrid1.Text
End Sub

Private Sub mEliminar_Click()
Dim rs As New ADODB.Recordset

Call ConectarBD

On Error GoTo ErrHandler
    rs.Open "EXEC sp_OperacionCombustible 'ELIMINAR', " & idSeleccionado, conn, adOpenStatic, adLockReadOnly
    MsgBox rs(0), vbInformation, "ESAPP"
    
    If rs.State = adStateOpen Then rs.Close
    If conn.State = adStateOpen Then Call DesconectarBD
    Call CargarlvCombustible
    Exit Sub

ErrHandler:
    MsgBox "Error al cargar datos: " & Err.Description, vbCritical, "Error"
    If rs.State = adStateOpen Then rs.Close
    If conn.State = adStateOpen Then Call DesconectarBD
End Sub

Private Sub MSFlexGrid1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbRightButton Then
    Dim fila As Long
    fila = MSFlexGrid1.FixedRows + (Y \ MSFlexGrid1.RowHeight(0)) - 1
    
    ' Validar que la fila clickeada sea válida (dentro del rango del grid)
    If fila >= MSFlexGrid1.FixedRows And fila < MSFlexGrid1.Rows Then
        MSFlexGrid1.Row = fila
        MSFlexGrid1.col = 0

        idSeleccionado = MSFlexGrid1.Text

        PopupMenu mnuListView
    End If
End If
End Sub

Private Sub txtPrecioActual_KeyPress(KeyAscii As Integer)
If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 And KeyAscii <> 44 Then
        KeyAscii = 0 ' Cancelar el caracter si no es número, backspace o coma
    End If
    
    ' Evitar múltiples comas
    If KeyAscii = 44 And InStr(txtPrecioActual.Text, ",") > 0 Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtPrecioActual_LostFocus()
With Combustibles
    .Precio = Replace(txtPrecioActual.Text, ",", ".")
    txtPrecioActual.Text = FormatoPrecio(.Precio)
End With
End Sub

Private Sub txtTipo_LostFocus()
Combustibles.Combustible = txtTipo.Text
End Sub
