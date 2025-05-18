VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form FRMCliente 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Administacion de Clientes - (LOCAL - DEFAULT)"
   ClientHeight    =   8310
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   18735
   Icon            =   "FRMCliente.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8310
   ScaleWidth      =   18735
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "Perpetua"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   8055
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   18495
      Begin VB.Frame Frame5 
         Height          =   735
         Left            =   360
         TabIndex        =   12
         Top             =   7080
         Width           =   17775
         Begin VB.CommandButton btnSalir 
            Caption         =   "&Salir"
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
            Left            =   15600
            TabIndex        =   13
            Top             =   240
            Width           =   2055
         End
         Begin VB.Line Line1 
            BorderColor     =   &H8000000D&
            BorderWidth     =   2
            X1              =   2280
            X2              =   15600
            Y1              =   420
            Y2              =   420
         End
         Begin VB.Label lbTotalClientes 
            AutoSize        =   -1  'True
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   195
            Left            =   120
            TabIndex        =   14
            Top             =   285
            Width           =   60
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Clientes"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   4695
         Left            =   360
         TabIndex        =   10
         Top             =   2280
         Width           =   17775
         Begin MSComctlLib.ListView lvClientes 
            Height          =   4215
            Left            =   120
            TabIndex        =   11
            Top             =   360
            Width           =   17535
            _ExtentX        =   30930
            _ExtentY        =   7435
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
      End
      Begin VB.Frame Frame3 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   735
         Left            =   360
         TabIndex        =   8
         Top             =   1440
         Width           =   17775
         Begin VB.CommandButton btnNuevoCliente 
            Caption         =   "&Nuevo Cliente"
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
            TabIndex        =   9
            Top             =   240
            Width           =   2055
         End
         Begin VB.CommandButton bntActualizar 
            Caption         =   "&Actualizar Cliente"
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
            Left            =   2400
            TabIndex        =   15
            Top             =   240
            Width           =   2055
         End
         Begin VB.Label lbSeleccion 
            AutoSize        =   -1  'True
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   240
            Left            =   4680
            TabIndex        =   16
            Top             =   300
            Width           =   75
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Datos Generales"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   1095
         Left            =   360
         TabIndex        =   1
         Top             =   240
         Width           =   17775
         Begin VB.TextBox txtDir 
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
            Left            =   7680
            TabIndex        =   5
            Top             =   600
            Width           =   9975
         End
         Begin VB.TextBox txtCuit 
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
            Left            =   4320
            TabIndex        =   4
            Top             =   600
            Width           =   3255
         End
         Begin VB.TextBox txtNombre 
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
            TabIndex        =   2
            Top             =   600
            Width           =   4095
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Dirección"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   7680
            TabIndex        =   7
            Top             =   360
            Width           =   795
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "C.U.I.T."
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   4320
            TabIndex        =   6
            Top             =   360
            Width           =   675
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Nombre"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   120
            TabIndex        =   3
            Top             =   360
            Width           =   675
         End
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
Attribute VB_Name = "FRMCliente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Clientes As New ClaseCliente

Dim Cantidad As Integer
Dim IdCliente As String

Private Sub bntActualizar_Click()
    Dim rs As New ADODB.Recordset
    
    If DatosValidador Then Exit Sub
    
    ' Conectar a la base de datos utilizando el módulo de conexión
    Call ConectarBD

    On Error GoTo ErrHandler
    
    If ValidarCUITCompleto Then Exit Sub
    
    rs.Open "exec sp_OperacionCliente 'MODIFICAR'," & IdCliente & ",'" & txtNombre.Text & "','" & txtDir.Text & "','" & txtCuit.Text & "'", conn, adOpenStatic, adLockReadOnly
    MsgBox rs(0), vbInformation, "ESAPP"
    
    ' Cerrar el recordset
    rs.Close
    Call DesconectarBD
    Call CargarlvClientes
    Call LimpiarCampos
    
    bntActualizar.Enabled = False
    btnNuevoCliente.Enabled = True
    
    Exit Sub

ErrHandler:
    MsgBox "Error al cargar datos: " & Err.Description, vbCritical, "Error"
    Call DesconectarBD
End Sub

Private Sub btnNuevoCliente_Click()
    Dim rs As New ADODB.Recordset
    
    If DatosValidador Then Exit Sub
    
    ' Conectar a la base de datos utilizando el módulo de conexión
    Call ConectarBD

    On Error GoTo ErrHandler
    
    If ValidarCUITCompleto Then Exit Sub
    
    With Clientes
    
    rs.Open "exec sp_OperacionCliente 'INSERTAR',NULL,'" & .Cliente & "','" & .Direccion & "','" & .Cuit & "'", conn, adOpenStatic, adLockReadOnly
    MsgBox rs(0), vbInformation, "ESAPP"
    
    End With
    
    ' Cerrar el recordset
    rs.Close
    Call DesconectarBD
    Call CargarlvClientes
    Call LimpiarCampos
    Exit Sub

ErrHandler:
    MsgBox "Error al cargar datos: " & Err.Description, vbCritical, "Error"
    Call DesconectarBD
End Sub

Private Sub LimpiarCampos()
    txtNombre.Text = ""
    txtDir.Text = ""
    txtCuit.Text = ""
    lbSeleccion.Caption = ""
End Sub

Public Function DatosValidador() As Boolean
If txtNombre.Text = "" Then
    MsgBox "Ingrese el nombre del cliente", vbInformation, "ESAPP"
    DatosValidador = True
ElseIf txtCuit.Text = "" Then
    MsgBox "Ingrese el C.U.I.T. del cliente", vbInformation, "ESAPP"
    DatosValidador = True
ElseIf txtDir.Text = "" Then
    MsgBox "Ingrese la dirección del cliente", vbInformation, "ESAPP"
    DatosValidador = True
Else
    DatosValidador = False
End If
End Function

Private Sub btnSalir_Click()
Unload Me
End Sub



Private Sub Form_Load()

'' Conectar a la base de datos
'Call ConectarBD

' Configuración del ListView
With lvClientes
    .View = lvwReport
    .ColumnHeaders.Add , , "", 0
    .ColumnHeaders.Add , , "Nro Cliente", 2000
    .ColumnHeaders.Add , , "Nombre", 4000
    .ColumnHeaders.Add , , "Direccion", 5000
    .ColumnHeaders.Add , , "C.U.I.T.", 3000
End With

' Cargar datos en el ListView
Call CargarlvClientes
End Sub

Private Sub CargarlvClientes()
    Dim rs As New ADODB.Recordset

    ' Conectar a la base de datos utilizando el módulo de conexión
    Call ConectarBD

    ' Limpiar el ListView antes de agregar los nuevos datos
    lvClientes.ListItems.Clear

    ' Obtener datos de la base
    On Error GoTo ErrHandler

    rs.Open "SELECT '' as valor, FORMAT(id, '00000000') AS id, nombre, direccion, cuit FROM clientes", conn, adOpenStatic, adLockReadOnly

    ' Cargar datos en el ListView
    If Not rs.EOF Then
        Do While Not rs.EOF
            With lvClientes.ListItems.Add(, , rs("valor"))
                .SubItems(1) = rs("id")
                .SubItems(2) = rs("nombre")
                .SubItems(3) = rs("direccion")
                .SubItems(4) = rs("cuit")
            End With
            rs.MoveNext
        Loop
    Else
        MsgBox "No hay productos registrados.", vbExclamation, "Aviso"
    End If
    lvClientes.ColumnHeaders(1).Width = 0
    ' Cerrar el recordset
    rs.Close
    
    Cantidad = lvClientes.ListItems.Count
    lbTotalClientes.Caption = "Total de clientes: " & Cantidad
    
    ' Desconectar
    Call DesconectarBD
    Exit Sub

ErrHandler:
    MsgBox "Error al cargar datos: " & Err.Description, vbCritical, "Error"
    Call DesconectarBD
End Sub

Private Sub lvClientes_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbRightButton Then
    PopupMenu mnuListView
End If
End Sub

Private Sub mActualizar_Click()

If lvClientes.ListItems.Count = 0 Then
    MsgBox "No hay clientes para actualizar", vbInformation, "ESAPP"
    Exit Sub
End If

On Error Resume Next

If Err.Number <> 0 Then
    MsgBox "Por favor, selecciione un cliente para actualizar", vbExclamation, "ESAPP"
    Exit Sub
End If
On Error GoTo 0

Call CargarDatosBox

bntActualizar.Enabled = True
btnNuevoCliente.Enabled = False

End Sub

Private Sub CargarDatosBox()

IdCliente = lvClientes.SelectedItem.SubItems(1)
txtNombre.Text = lvClientes.SelectedItem.SubItems(2)
txtDir.Text = lvClientes.SelectedItem.SubItems(3)
txtCuit.Text = lvClientes.SelectedItem.SubItems(4)
    
lbSeleccion.Caption = "Cliente seleccionado: " + Format$(CLng(IdCliente), "0000000000")

'Format$(Valor, "#,##0.00")
End Sub

Private Sub mEliminar_Click()

Dim rs As New ADODB.Recordset

'Verificamos si tenemos elemento seleccionado
If lvClientes.ListItems.Count <> 0 Then
    If lvClientes.SelectedItem Is Nothing Then
        MsgBox "Por favor, seleccione un elemento para eliminar", vbExclamation
        Exit Sub
    End If
Else
    MsgBox "No hay clientes para eliminar", vbInformation, "ESAPP"
    Exit Sub
End If

IdCliente = lvClientes.SelectedItem.SubItems(1)

Call ConectarBD

On Error GoTo ErrHandler
    rs.Open "EXEC sp_OperacionCliente 'ELIMINAR', " & IdCliente, conn, adOpenStatic, adLockReadOnly
    MsgBox rs(0), vbInformation, "ESAPP"
    
    rs.Close
    Call DesconectarBD
    Call CargarlvClientes
    Exit Sub

ErrHandler:
    MsgBox "Error al cargar datos: " & Err.Description, vbCritical, "Error"
    Call DesconectarBD

End Sub

Private Sub txtCuit_Change()
Dim strCUIT As String
    Dim strSoloNumeros As String
    Dim i As Integer
    
    ' Guardar la posición del cursor
    Dim cursorPos As Integer
    cursorPos = txtCuit.SelStart
    
    ' Obtener solo los números del texto actual
    strCUIT = txtCuit.Text
    strSoloNumeros = ""
    
    For i = 1 To Len(strCUIT)
        If IsNumeric(Mid(strCUIT, i, 1)) Then
            strSoloNumeros = strSoloNumeros & Mid(strCUIT, i, 1)
        End If
    Next i
    
    ' Limitar a 11 dígitos (formato CUIT)
    If Len(strSoloNumeros) > 11 Then
        strSoloNumeros = Left(strSoloNumeros, 11)
    End If
    
    ' Formatear con guiones según XX-XXXXXXXX-X
    Dim strFormateado As String
    strFormateado = ""
    
    For i = 1 To Len(strSoloNumeros)
        strFormateado = strFormateado & Mid(strSoloNumeros, i, 1)
        ' Agregar guiones después del segundo y décimo dígito
        If i = 2 Or i = 10 Then
            If i < Len(strSoloNumeros) Then
                strFormateado = strFormateado & "-"
            End If
        End If
    Next i
    
    ' Contar cuántos guiones hay antes de la posición actual del cursor
    Dim guionesPrevios As Integer
    guionesPrevios = 0
    
    For i = 1 To cursorPos
        If i <= Len(strCUIT) Then
            If Mid(strCUIT, i, 1) = "-" Then
                guionesPrevios = guionesPrevios + 1
            End If
        End If
    Next i
    
    ' Contar cuántos dígitos hay antes de la posición actual del cursor
    Dim digitosPrevios As Integer
    digitosPrevios = 0
    
    For i = 1 To cursorPos
        If i <= Len(strCUIT) Then
            If IsNumeric(Mid(strCUIT, i, 1)) Then
                digitosPrevios = digitosPrevios + 1
            End If
        End If
    Next i
    
    ' Evitar llamadas recursivas al evento Change
    If txtCuit.Text <> strFormateado Then
        txtCuit.Text = strFormateado
        
        ' Calcular nueva posición del cursor basada en los dígitos ingresados
        Dim nuevaPos As Integer
        nuevaPos = digitosPrevios
        
        ' Ajustar por los guiones en el nuevo formato
        If nuevaPos > 2 Then
            nuevaPos = nuevaPos + 1 ' Añadir el primer guión
        End If
        If nuevaPos > 10 Then
            nuevaPos = nuevaPos + 1 ' Añadir el segundo guión
        End If
        
        ' Asegurarse de que la posición no se salga del rango
        If nuevaPos > Len(strFormateado) Then
            nuevaPos = Len(strFormateado)
        End If
        
        txtCuit.SelStart = nuevaPos
    End If
End Sub

Private Sub txtCuit_KeyPress(KeyAscii As Integer)
If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 Then
        KeyAscii = 0 ' Cancelar el caracter si no es un numero
    End If
End Sub

Private Function ValidarCUITCompleto() As Boolean
    Dim mensaje As String
    
    If Len(txtCuit.Text) < 13 Then
        mensaje = "El CUIT está incompleto. Debe tener el formato XX-XXXXXXXX-X"
        MsgBox mensaje, vbExclamation, "CUIT Incompleto"
        ValidarCUITCompleto = True
        Exit Function
    End If
    
    If ValidarCUIT(txtCuit.Text) Then
        ValidarCUITCompleto = False
    Else
        mensaje = "El CUIT no es válido. El dígito verificador no corresponde."
        MsgBox mensaje, vbExclamation, "CUIT Inválido"
        ValidarCUITCompleto = True
    End If
End Function

Public Function ValidarCUIT(ByVal strCUIT As String) As Boolean
    ' Validar longitud y formato XX-XXXXXXXX-X
    If Len(strCUIT) <> 13 Then
        ValidarCUIT = False
        Exit Function
    End If
    
    If Mid(strCUIT, 3, 1) <> "-" Or Mid(strCUIT, 12, 1) <> "-" Then
        ValidarCUIT = False
        Exit Function
    End If
    
    ' Validar que las partes solo contengan números
    Dim parte1 As String, parte2 As String, parte3 As String
    parte1 = Left(strCUIT, 2)
    parte2 = Mid(strCUIT, 4, 8)
    parte3 = Right(strCUIT, 1)
    
    If Not IsNumeric(parte1) Or Not IsNumeric(parte2) Or Not IsNumeric(parte3) Then
        ValidarCUIT = False
        Exit Function
    End If
    
    ' Algoritmo para validar el dígito verificador del CUIT
    Dim cuitSinGuiones As String
    Dim multiplicadores As Variant
    Dim suma As Integer
    Dim digitoVerificador As Integer
    Dim i As Integer
    
    ' Quitar los guiones para trabajar solo con los números
    cuitSinGuiones = parte1 & parte2 & parte3
    
    ' La serie de multiplicadores es [5,4,3,2,7,6,5,4,3,2]
    multiplicadores = Array(5, 4, 3, 2, 7, 6, 5, 4, 3, 2)
    suma = 0
    
    ' Multiplicar cada dígito por su correspondiente multiplicador y sumar
    For i = 0 To 9
        suma = suma + (Val(Mid(cuitSinGuiones, i + 1, 1)) * multiplicadores(i))
    Next i
    
    ' El dígito verificador es 11 menos el resto de la división por 11
    Dim resto As Integer
    resto = suma Mod 11
    digitoVerificador = 11 - resto
    
    ' Si el resultado es 11, el dígito verificador es 0
    ' Si el resultado es 10, el dígito verificador es 9
    If digitoVerificador = 11 Then
        digitoVerificador = 0
    ElseIf digitoVerificador = 10 Then
        digitoVerificador = 9
    End If
    
    ' Verificar si el dígito calculado coincide con el ingresado
    ValidarCUIT = (digitoVerificador = Val(parte3))
End Function

Private Sub txtCuit_LostFocus()
Clientes.Cuit = txtCuit.Text
End Sub

Private Sub txtDir_LostFocus()
Clientes.Direccion = txtDir.Text
End Sub

Private Sub txtNombre_LostFocus()
Clientes.Cliente = txtNombre.Text
End Sub
