VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FRMCombustible 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Combustible"
   ClientHeight    =   2910
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   9135
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2910
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
      Height          =   2655
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   8895
      Begin MSComctlLib.ListView lvCombustible 
         Height          =   2295
         Left            =   4680
         TabIndex        =   8
         Top             =   240
         Width           =   4095
         _ExtentX        =   7223
         _ExtentY        =   4048
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   0
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
         Height          =   375
         Left            =   1560
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
         Height          =   375
         Left            =   3000
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
         Height          =   375
         Left            =   120
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
    
    ' Cerrar el recordset
    rs.Close
    Call DesconectarBD
    Call CargarlvCombustible
    Call LimpiarCampos
    
    btnActualizar.Enabled = False
    btnIngresar.Enabled = True
    
    Exit Sub

ErrHandler:
    MsgBox "Error al cargar datos: " & Err.Description, vbCritical, "Error"
    Call DesconectarBD
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
    rs.Close
    Call DesconectarBD
    Call CargarlvCombustible
    Call LimpiarCampos
    Exit Sub

ErrHandler:
    MsgBox "Error al cargar datos: " & Err.Description, vbCritical, "Error"
    Call DesconectarBD
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
' Configuración del ListView
With lvCombustible
    .View = lvwReport
    .ColumnHeaders.Add , , "", 0
    .ColumnHeaders.Add , , "id", 500
    .ColumnHeaders.Add , , "Tipo", 2000
    .ColumnHeaders.Add , , "Precio", 1500
End With

' Cargar datos en el ListView
Call CargarlvCombustible
End Sub

Private Sub CargarlvCombustible()
    Dim rs As New ADODB.Recordset

    ' Conectar a la base de datos utilizando el módulo de conexión
    Call ConectarBD

    ' Limpiar el ListView antes de agregar los nuevos datos
    lvCombustible.ListItems.Clear

    ' Obtener datos de la base
    On Error GoTo ErrHandler

    rs.Open "SELECT '' as valor, id, tipo, precio FROM Combustible", conn, adOpenStatic, adLockReadOnly

    ' Cargar datos en el ListView
    If Not rs.EOF Then
        Do While Not rs.EOF
            With lvCombustible.ListItems.Add(, , rs("valor"))
                .SubItems(1) = rs("id")
                .SubItems(2) = rs("tipo")
                .SubItems(3) = FormatoPrecio(rs("precio"))
            End With
            rs.MoveNext
        Loop
    Else
        MsgBox "No hay combustible registrados.", vbExclamation, "Aviso"
    End If
    lvCombustible.ColumnHeaders(1).Width = 0
    ' Cerrar el recordset
    rs.Close
    
    ' Desconectar
    Call DesconectarBD
    Exit Sub

ErrHandler:
    MsgBox "Error al cargar datos: " & Err.Description, vbCritical, "Error"
    Call DesconectarBD
End Sub

Private Sub lvCombustible_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbRightButton Then
    PopupMenu mnuListView
End If
End Sub

Private Sub mActualizar_Click()
If lvCombustible.ListItems.Count = 0 Then
    MsgBox "No hay combustible para actualizar", vbInformation, "ESAPP"
    Exit Sub
End If

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

idCombustible = lvCombustible.SelectedItem.SubItems(1)
txtTipo.Text = lvCombustible.SelectedItem.SubItems(2)
txtPrecioActual.Text = lvCombustible.SelectedItem.SubItems(3)

End Sub

Private Sub mEliminar_Click()
Dim rs As New ADODB.Recordset

'Verificamos si tenemos elemento seleccionado
If lvCombustible.ListItems.Count <> 0 Then
    If lvCombustible.SelectedItem Is Nothing Then
        MsgBox "Por favor, seleccione un elemento para eliminar", vbExclamation
        Exit Sub
    End If
Else
    MsgBox "No hay tipo de combustible para eliminar", vbInformation, "ESAPP"
    Exit Sub
End If

idCombustible = lvCombustible.SelectedItem.SubItems(1)

Call ConectarBD

On Error GoTo ErrHandler
    rs.Open "EXEC sp_OperacionCombustible 'ELIMINAR', " & idCombustible, conn, adOpenStatic, adLockReadOnly
    MsgBox rs(0), vbInformation, "ESAPP"
    
    rs.Close
    Call DesconectarBD
    Call CargarlvCombustible
    Exit Sub

ErrHandler:
    MsgBox "Error al cargar datos: " & Err.Description, vbCritical, "Error"
    Call DesconectarBD
End Sub

Private Sub txtPrecioActual_KeyPress(KeyAscii As Integer)
If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 Then
        KeyAscii = 0 ' Cancelar el caracter si no es un numero
    End If
End Sub

Private Sub txtPrecioActual_LostFocus()
With Combustibles
    .Precio = txtPrecioActual.Text
    txtPrecioActual.Text = FormatoPrecio(.Precio)
End With
End Sub

Private Sub txtTipo_LostFocus()
Combustibles.Combustible = txtTipo.Text
End Sub
