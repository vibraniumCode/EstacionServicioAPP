VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form FRMFacturacion 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   6735
   ClientLeft      =   150
   ClientTop       =   -1590
   ClientWidth     =   13815
   Icon            =   "FRMFacturacion.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6735
   ScaleWidth      =   13815
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   6495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   13575
      Begin VB.Frame Frame2 
         Caption         =   "Estación"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   855
         Left            =   120
         TabIndex        =   27
         Top             =   480
         Width           =   9255
         Begin VB.ComboBox cboEstaciones 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   120
            TabIndex        =   28
            Text            =   "Combo1"
            Top             =   360
            Width           =   9015
         End
      End
      Begin VB.Frame Frame12 
         Height          =   855
         Left            =   120
         TabIndex        =   21
         Top             =   5520
         Width           =   13335
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            Caption         =   "@vibraniumcode - mlopez developer"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   120
            TabIndex        =   24
            Top             =   360
            Width           =   3285
         End
      End
      Begin VB.Frame Frame9 
         Caption         =   "Cliente"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   855
         Left            =   120
         TabIndex        =   14
         Top             =   1320
         Width           =   13335
         Begin VB.ComboBox cboClientes 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   120
            TabIndex        =   15
            Text            =   "Combo1"
            Top             =   360
            Width           =   13095
         End
      End
      Begin MSComCtl2.DTPicker fecEmision 
         Height          =   345
         Left            =   11280
         TabIndex        =   12
         Top             =   720
         Width           =   2100
         _ExtentX        =   3704
         _ExtentY        =   609
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   143785985
         CurrentDate     =   45777
      End
      Begin VB.Frame Frame7 
         Height          =   1335
         Left            =   120
         TabIndex        =   6
         Top             =   4200
         Width           =   13335
         Begin VB.TextBox txtImpuesto 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   3360
            TabIndex        =   22
            Text            =   "$00.00"
            Top             =   600
            Width           =   3135
         End
         Begin VB.TextBox txtTotal 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   6600
            TabIndex        =   11
            Text            =   "$00.00"
            Top             =   600
            Width           =   3375
         End
         Begin VB.TextBox txtIva 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            TabIndex        =   10
            Text            =   "$00.00"
            Top             =   600
            Width           =   3135
         End
         Begin VB.CommandButton btnFinalizar 
            Caption         =   "&Finalizar venta"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Left            =   11640
            Picture         =   "FRMFacturacion.frx":08CA
            Style           =   1  'Graphical
            TabIndex        =   26
            Top             =   360
            Width           =   1575
         End
         Begin VB.Line Line4 
            BorderWidth     =   2
            Index           =   2
            X1              =   13310
            X2              =   0
            Y1              =   750
            Y2              =   800
         End
         Begin VB.Line Line2 
            BorderColor     =   &H000040C0&
            BorderWidth     =   3
            DrawMode        =   14  'Copy Pen
            Index           =   2
            X1              =   3360
            X2              =   5640
            Y1              =   480
            Y2              =   480
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "IMPUESTOS"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   3360
            TabIndex        =   23
            Top             =   240
            Width           =   1065
         End
         Begin VB.Line Line2 
            BorderColor     =   &H00000080&
            BorderWidth     =   3
            DrawMode        =   14  'Copy Pen
            Index           =   1
            X1              =   6600
            X2              =   8880
            Y1              =   480
            Y2              =   480
         End
         Begin VB.Line Line3 
            BorderColor     =   &H00008000&
            BorderWidth     =   3
            DrawMode        =   14  'Copy Pen
            X1              =   120
            X2              =   1800
            Y1              =   480
            Y2              =   480
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            Caption         =   "TOTAL"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   6600
            TabIndex        =   9
            Top             =   240
            Width           =   585
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "ALICUOTA IVA"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   120
            TabIndex        =   7
            Top             =   240
            Width           =   1320
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Carga"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   2055
         Left            =   120
         TabIndex        =   1
         Top             =   2160
         Width           =   13335
         Begin MSComctlLib.ProgressBar ProgressBar1 
            Height          =   255
            Left            =   6120
            TabIndex        =   25
            Top             =   1560
            Width           =   7095
            _ExtentX        =   12515
            _ExtentY        =   450
            _Version        =   393216
            Appearance      =   1
            OLEDropMode     =   1
            Scrolling       =   1
         End
         Begin VB.Frame Frame11 
            Caption         =   "PRECIO NETO"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Left            =   8880
            TabIndex        =   19
            Top             =   360
            Width           =   2175
            Begin VB.TextBox precioNeto 
               Alignment       =   1  'Right Justify
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00004000&
               Height          =   375
               Left            =   120
               Locked          =   -1  'True
               TabIndex        =   20
               Text            =   "$00.00"
               Top             =   300
               Width           =   1935
            End
         End
         Begin VB.Frame Frame10 
            Caption         =   "COMBUSTIBLE + PRECIO $$$"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Left            =   120
            TabIndex        =   16
            Top             =   360
            Width           =   5655
            Begin VB.TextBox txtMontoCombustible 
               Alignment       =   1  'Right Justify
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000C0&
               Height          =   330
               Left            =   3000
               TabIndex        =   18
               Text            =   "$00.00"
               Top             =   300
               Width           =   2535
            End
            Begin VB.ComboBox cboCombustible 
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Left            =   120
               Style           =   2  'Dropdown List
               TabIndex        =   17
               Top             =   300
               Width           =   2775
            End
         End
         Begin VB.CommandButton btnProcesarDatos 
            Caption         =   "&Procesar Datos"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   120
            Picture         =   "FRMFacturacion.frx":1194
            Style           =   1  'Graphical
            TabIndex        =   8
            Top             =   1320
            Width           =   1695
         End
         Begin VB.Frame Frame4 
            Caption         =   "LITROS"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Left            =   5880
            TabIndex        =   2
            Top             =   360
            Width           =   2895
            Begin VB.CommandButton restar 
               Caption         =   "-"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Left            =   600
               TabIndex        =   5
               Top             =   300
               Width           =   375
            End
            Begin VB.CommandButton sumar 
               Caption         =   "+"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Left            =   120
               TabIndex        =   4
               Top             =   300
               Width           =   375
            End
            Begin VB.TextBox txtCantidadLt 
               Alignment       =   1  'Right Justify
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000009&
               Height          =   360
               Left            =   1080
               TabIndex        =   3
               Text            =   "0"
               Top             =   300
               Width           =   1695
            End
         End
         Begin VB.Label lblCarga 
            Alignment       =   2  'Center
            Caption         =   "Calculando el importe final con impuestos..."
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00008000&
            Height          =   255
            Left            =   1920
            TabIndex        =   29
            Top             =   1560
            Visible         =   0   'False
            Width           =   4095
         End
         Begin VB.Line Line5 
            BorderWidth     =   2
            X1              =   0
            X2              =   13305
            Y1              =   760
            Y2              =   760
         End
         Begin VB.Line Line4 
            BorderWidth     =   2
            Index           =   0
            X1              =   13310
            X2              =   0
            Y1              =   1680
            Y2              =   1680
         End
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "FECHA DE EMISION"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   9600
         TabIndex        =   13
         Top             =   840
         Width           =   1545
      End
      Begin VB.Line Line1 
         BorderWidth     =   2
         X1              =   120
         X2              =   13440
         Y1              =   360
         Y2              =   360
      End
   End
   Begin VB.Menu mnuListView 
      Caption         =   "&mnuListView"
      Visible         =   0   'False
      Begin VB.Menu Eliminar 
         Caption         =   "Eliminar"
      End
      Begin VB.Menu Actualizar 
         Caption         =   "Actualizar"
      End
   End
End
Attribute VB_Name = "FRMFacturacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Clientes As New ClaseCliente
Dim producto As New ClaseProducto

Dim ClientesArray() As New ClaseCliente
Dim ClientesIDs() As New ClaseCliente

Dim fecImpITC As Date


Dim alertaMostrada As Boolean
Dim idProducto As Long
'Dim nroFactura As Long
' Método 1: Bloquear botones de la ventana usando API de Windows
Private Declare Function GetSystemMenu Lib "user32" (ByVal hWnd As Long, ByVal bRevert As Long) As Long
Private Declare Function DeleteMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long

Private Const SC_MAXIMIZE = &HF030
Private Const MF_BYCOMMAND = &H0

Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Private Sub btnFinalizar_Click()
If MsgBox("¿Desea finalizar la venta?", vbYesNo + vbQuestion, "Confirmación") = vbYes Then
    Dim rs As New ADODB.Recordset
    Dim litros As Long
    Dim idEmpresaSeleccionado As Integer
    
        
    Call ConectarBD
    On Error GoTo ErrHandler
    
    Dim idClienteSeleccionado As Long
    
    If cboEstaciones.ListIndex <> -1 Then
        idEmpresaSeleccionado = cboEstaciones.ItemData(cboEstaciones.ListIndex)
    Else
        MsgBox "No hay estacion seleccionado", vbInformation, "ESAPP"
    End If
    
    ' Verificamos que haya un cliente seleccionado
    If cboClientes.ListIndex <> -1 Then
        idClienteSeleccionado = cboClientes.ItemData(cboClientes.ListIndex)
    Else
        MsgBox "No hay cliente seleccionado", vbInformation, "ESAPP"
        Exit Sub
    End If
    
    Dim idCombustibleSeleccionado As Long
    
    If cboCombustible.ListIndex <> -1 Then
        idCombustibleSeleccionado = cboCombustible.ItemData(cboCombustible.ListIndex)
    Else
        MsgBox "No hay combustible seleccionado", vbInformation, "ESAPP"
        Exit Sub
    End If
    
    If cboEstaciones.ListIndex <> -1 Then
        idEmpresaSeleccionado = cboEstaciones.ItemData(cboEstaciones.ListIndex)
    Else
        MsgBox "No hay empresa seleccionado", vbInformation, "ESAPP"
        Exit Sub
    End If
    
    If Not ValidarEntradas(producto.litros) Then
        MsgBox "Datos inválidos, revise los valores.", vbExclamation
        Exit Sub
    End If
    
'    producto.precioNeto = CDbl(LimpiarValor(precioNeto.Text))
'    producto.PrecioUnitario = CDbl(LimpiarValor(txtMontoCombustible.Text))
'    producto.MontoITC = CDbl(LimpiarValor(txtMontoITC.Text))
    
'    Dim ivaCalc As Double
'    Dim itcCalc As Double
'    Dim totalCalc As Double
'
'    ivaCalc = producto.CalcularIVA()
'    itcCalc = producto.CalcularITC()
'    totalCalc = producto.CalcularTotal()
    
    Dim fecha As String
    fecha = Format(fecEmision.value, "yyyymmdd")
    Dim sql As String
    sql = "exec sp_facturacion " & idEmpresaSeleccionado & "," & idClienteSeleccionado & "," & idCombustibleSeleccionado & "," & producto.litros & ",'" & fecha & "',0"
    
    rs.Open "exec sp_facturacion " & _
    idEmpresaSeleccionado & "," & _
    idClienteSeleccionado & "," & _
    idCombustibleSeleccionado & "," & _
    producto.litros & ",'" & _
    fecha & "',0" _
    , conn, adOpenStatic, adLockReadOnly
    
    codFactura = Format(rs(0), "0000")
    nroFactura = Format(rs(1), "00000000")
    
    MsgBox "Factura generada con el Nro: " & codFactura & "-" & nroFactura, vbInformation, "ESAPP"
    
    If rs.State = adStateOpen Then rs.Close
    If conn.State = adStateOpen Then Call DesconectarBD
    
    frmComprobantes.Show vbModal
    Exit Sub

ErrHandler:
    MsgBox "Error al cargar datos: " & Err.Description, vbCritical, "Error"
    If rs.State = adStateOpen Then rs.Close
    If conn.State = adStateOpen Then Call DesconectarBD
End If

End Sub

Private Sub btnProcesarDatos_Click()
    Dim litros As Double
    
    If Not ValidarEntradas(litros) Then Exit Sub

    producto.litros = litros
    producto.precioNeto = CDbl(precioNeto.Text)
    
    ' Verificamos que haya un cliente seleccionado
    If cboClientes.ListIndex <> -1 Then
        If cboCombustible.ListIndex <> -1 Then
            lblCarga.Visible = True
            CargarBarraProgreso
            CargarControl
        Else
            MsgBox "No hay combustible seleccionado", vbInformation, "ESAPP"
            Exit Sub
        End If
    Else
        MsgBox "No hay cliente seleccionado", vbInformation, "ESAPP"
        Exit Sub
    End If
   
'    txtIva.Text = FormatoPrecio(producto.CalcularIVA())
'    txtITC.Text = FormatoPrecio(producto.CalcularITC())
'    txtTotal.Text = FormatoPrecio(producto.CalcularTotal())

    MsgBox "Carga completada", vbInformation, "ESAPP"
    frmControl.Show vbModal
    btnFinalizar.Enabled = True
End Sub
Private Sub CargarControl()
idEstacionEmpresa = cboEstaciones.ItemData(cboEstaciones.ListIndex)

executeSQL = "exec sp_facturacion " & cboEstaciones.ItemData(cboEstaciones.ListIndex) & ", " & cboClientes.ItemData(cboClientes.ListIndex)
executeSQL = executeSQL + ", " & cboCombustible.ItemData(cboCombustible.ListIndex) & ", " & producto.litros & ",'" & Format(fecEmision.value, "yyyymmdd") & "',1"
End Sub
Private Sub CargarBarraProgreso()
    Dim i As Integer
    ProgressBar1.Min = 0
    ProgressBar1.Max = 100
    For i = 0 To 100
        ProgressBar1.value = i
        DoEvents
        Sleep 20
    Next i
End Sub

Private Function ValidarEntradas(ByRef litros As Double) As Boolean
    Dim texto As String
    texto = Trim(Replace(UCase(txtCantidadLt.Text), "LT", ""))

    If Not IsNumeric(texto) Or Val(texto) <= 0 Then
        MsgBox "Ingrese la cantidad de litros válida", vbInformation, "ESAPP"
        ValidarEntradas = False
        Exit Function
    End If

   If Not IsNumeric(LimpiarValor(precioNeto.Text)) Or Val(LimpiarValor(precioNeto.Text)) = 0 Then
        MsgBox "El precio neto es inválido", vbInformation, "ESAPP"
        ValidarEntradas = False
        Exit Function
    End If

    litros = CDbl(texto)
    ValidarEntradas = True
End Function

Private Sub btnTickets_Click()

End Sub

Private Sub cboCombustible_Click()
    Dim rs As New ADODB.Recordset
    Dim idSeleccionado As Integer
 
    ' Obtener el ID seleccionado desde el ComboBox
    If cboCombustible.ListIndex <> -1 Then
        idSeleccionado = cboCombustible.ItemData(cboCombustible.ListIndex)
        
        ' Conectar a la base de datos
        Call ConectarBD
        
        On Error GoTo ErrHandler
        
        ' Buscar el precio del combustible seleccionado
        rs.Open "SELECT precio FROM Combustible WHERE id = " & idSeleccionado, conn, adOpenStatic, adLockReadOnly
        
        If Not rs.EOF Then
            txtMontoCombustible.Text = Format(rs("precio"), "0.00")
        Else
            txtMontoCombustible.Text = ""
        End If
        
        Call calcularNeto
        
        txtMontoCombustible.Text = FormatoPrecio(txtMontoCombustible.Text)
        
        If rs.State = adStateOpen Then rs.Close
        If conn.State = adStateOpen Then Call DesconectarBD
    End If
    Exit Sub

ErrHandler:
    MsgBox "Error al obtener el precio del combustible: " & Err.Description, vbCritical, "Error"
    If rs.State = adStateOpen Then rs.Close
    If conn.State = adStateOpen Then Call DesconectarBD
End Sub

Private Sub Form_Load()
Dim Cantidad As Integer

  Dim hMenu As Long

' Obtener el menú del sistema
hMenu = GetSystemMenu(Me.hWnd, False)

' Eliminar solo el botón de maximizar
DeleteMenu hMenu, SC_MAXIMIZE, MF_BYCOMMAND

'CargarNumeroFactura

' Cargar los datos de clientes ANTES de cargar el combo
ModuloClientes.CargarDatosClientes

Call CargarEstaciones
Call CargarClientesCombo
Call CargarComboCombustible
'Call CargarImpuestoITC
End Sub

Private Sub Form_Resize()
    ' Restaurar el tamaño original si se intenta maximizar
    If Me.WindowState = vbMaximized Then
        Me.WindowState = vbNormal
    End If
End Sub

Private Sub Frame5_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Frame5.Caption = "ITC $"
End Sub

Private Sub Image1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Frame5.Caption = "ITC $ - " & fecImpITC
End Sub

Private Sub precioNeto_LostFocus()
    producto.precioNeto = LimpiarValor(precioNeto.Text)
    precioNeto.Text = FormatoPrecio(producto.precioNeto)
End Sub

Private Sub CargarPrecio()
    precioNeto.Text = FormatoPrecio(producto.CalcularPrecioNeto())
End Sub

Private Sub sumar_Click()
    Dim texto As String
    Dim litros As Long

    alertaMostrada = False

    texto = Trim(Replace(UCase(txtCantidadLt.Text), "LT", ""))

    If IsNumeric(texto) Then
        litros = CLng(texto)
        producto.litros = litros + 1
        txtCantidadLt.Text = CStr(producto.litros) & " LT"
    Else
        txtCantidadLt.Text = ""
        producto.litros = 0
    End If

    Call calcularNeto
End Sub

Private Sub restar_Click()
    Dim texto As String
    Dim litros As Long

    alertaMostrada = False

    texto = Trim(Replace(UCase(txtCantidadLt.Text), "LT", ""))

    If IsNumeric(texto) Then
        litros = CLng(texto)
        producto.litros = IIf(litros > 0, litros - 1, 0)
        txtCantidadLt.Text = CStr(producto.litros) & " LT"
    Else
        txtCantidadLt.Text = ""
        producto.litros = 0
    End If

    Call calcularNeto
End Sub


Private Sub ActualizarPrecio()
    precioNeto.Text = FormatoPrecio(producto.CalcularPrecioNeto())
    precioNeto_LostFocus
End Sub

' Procedimiento para limpiar los campos
Private Sub LimpiarCampos()
    txtDescripcion.Text = ""
    btnCantidad.Text = 1
    Preciouni.Text = "$" & Format(0, "#,##0.00")
    precioNeto.Text = "$" & Format(0, "#,##0.00")
    Me.Tag = ""  ' Limpiar el ID guardado
End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub CargarClientesCombo()
    Dim i As Integer
    
    ' Limpiar el ComboBox
    cboClientes.Clear
    
    ' Verificar si hay elementos en el array antes de recorrerlo
    If IsArrayInitialized(ModuloClientes.ClientesArray) Then
        ' Cargar todos los clientes en el combo
        For i = 0 To UBound(ModuloClientes.ClientesArray)
            cboClientes.AddItem ModuloClientes.ClientesArray(i)
            cboClientes.ItemData(cboClientes.NewIndex) = ModuloClientes.ClientesIDs(i)
        Next i
    End If
End Sub

' Función para verificar si un array está inicializado
Private Function IsArrayInitialized(arr As Variant) As Boolean
    On Error Resume Next
    IsArrayInitialized = False
    
    ' Intentar obtener el límite superior del array
    Dim temp As Integer
    temp = UBound(arr)
    
    ' Si no hay error y el límite superior es al menos 0, el array está inicializado
    If Err.Number = 0 And temp >= 0 Then
        IsArrayInitialized = True
    End If
    
    On Error GoTo 0
End Function

Private Sub cboClientes_KeyUp(KeyCode As Integer, Shift As Integer)
    Dim filtro As String
    Dim i As Integer
    
    ' Obtener el texto actual del ComboBox
    filtro = LCase(cboClientes.Text)
    
    ' No hacer nada para teclas especiales
    If KeyCode = vbKeyDown Or KeyCode = vbKeyUp Or KeyCode = vbKeyReturn Then
        Exit Sub
    End If
    
    ' Guardar la posición actual del cursor
    Dim cursorPos As Integer
    cursorPos = cboClientes.SelStart
    
    ' Limpiar y recargar los elementos que coincidan con el filtro
    cboClientes.Clear
    
    ' Verificar si el array está inicializado
    If IsArrayInitialized(ModuloClientes.ClientesArray) Then
        For i = 0 To UBound(ModuloClientes.ClientesArray)
            If InStr(1, LCase(ModuloClientes.ClientesArray(i)), filtro) > 0 Then
                cboClientes.AddItem ModuloClientes.ClientesArray(i)
                cboClientes.ItemData(cboClientes.NewIndex) = ModuloClientes.ClientesIDs(i)
            End If
        Next i
    End If
    
    ' Restaurar el texto y la posición del cursor
    cboClientes.Text = filtro
    cboClientes.SelStart = cursorPos
End Sub

Private Sub cboClientes_LostFocus()
    ' Si el ComboBox está vacío, recargar todos los elementos
    If Trim(cboClientes.Text) = "" Then
        CargarClientesCombo
    End If
End Sub
Private Sub CargarImpuestoITC()
     Dim rs As New ADODB.Recordset
    
    ' Conectar a la base de datos
    Call ConectarBD
    
    On Error GoTo ErrHandler
    
    rs.Open "select top 1 monto, fechaOperacion from Impuestos order by fechaOperacion desc", conn, adOpenStatic, adLockReadOnly
    
    txtMontoITC.Text = FormatoPrecio(rs(0))
    fecImpITC = rs(1)
    
    If rs.State = adStateOpen Then rs.Close
    If conn.State = adStateOpen Then Call DesconectarBD
    Exit Sub

ErrHandler:
    MsgBox "Error al cargar el impuesto ITC: " & Err.Description, vbCritical, "Error"
    If rs.State = adStateOpen Then rs.Close
    If conn.State = adStateOpen Then Call DesconectarBD
End Sub

Private Sub CargarComboCombustible()
    Dim rs As New ADODB.Recordset
    
    ' Conectar a la base de datos
    Call ConectarBD
    
    On Error GoTo ErrHandler
    
    cboCombustible.Clear
    
    rs.Open "SELECT id, tipo, precio FROM Combustible ORDER BY id", conn, adOpenStatic, adLockReadOnly
    
    ' Cargar los meses desde la base de datos al ComboBox
    Do While Not rs.EOF
        ' Puedes guardar el ID en ItemData si querés usarlo después
        cboCombustible.AddItem rs("tipo")
        cboCombustible.ItemData(cboCombustible.NewIndex) = rs("id")
        rs.MoveNext
    Loop
    
    If cboCombustible.ListCount > 0 Then
        cboCombustible.ListIndex = 0
    End If
    
    If rs.State = adStateOpen Then rs.Close
    If conn.State = adStateOpen Then Call DesconectarBD
    Exit Sub
    
ErrHandler:
    MsgBox "Error al cargar el listado de combustible: " & Err.Description, vbCritical, "Error"
    If rs.State = adStateOpen Then rs.Close
    If conn.State = adStateOpen Then Call DesconectarBD
End Sub

Private Sub CargarEstaciones()
    Dim rs As New ADODB.Recordset
    
    ' Conectar a la base de datos
    Call ConectarBD
    
    On Error GoTo ErrHandler
    
    cboEstaciones.Clear
    
    rs.Open "SELECT id, nombre FROM Empresas ORDER BY id", conn, adOpenStatic, adLockReadOnly
    
    ' Cargar los meses desde la base de datos al ComboBox
    Do While Not rs.EOF
        ' Puedes guardar el ID en ItemData si querés usarlo después
        cboEstaciones.AddItem rs("nombre")
        cboEstaciones.ItemData(cboEstaciones.NewIndex) = rs("id")
        rs.MoveNext
    Loop
    
    If cboEstaciones.ListCount > 0 Then
        cboEstaciones.ListIndex = 0
    End If
    
    If rs.State = adStateOpen Then rs.Close
    If conn.State = adStateOpen Then Call DesconectarBD
    Exit Sub
    
ErrHandler:
    MsgBox "Error al cargar el listado de estaciones: " & Err.Description, vbCritical, "Error"
    If rs.State = adStateOpen Then rs.Close
    If conn.State = adStateOpen Then Call DesconectarBD
End Sub

Private Sub txtCantidadLt_KeyPress(KeyAscii As Integer)
If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 And KeyAscii <> 44 Then
        KeyAscii = 0 ' Cancelar el caracter si no es número, backspace o coma
    End If
    
    ' Evitar múltiples comas
    If KeyAscii = 44 And InStr(txtCantidadLt.Text, ",") > 0 Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtCantidadLt_LostFocus()
Call calcularNeto
End Sub

Private Sub calcularNeto()
    Dim numero As Long
    Dim texto As String
    Dim litros As Double

    ' Limpiar y preparar el texto
    texto = Trim(Replace(UCase(Replace(txtCantidadLt.Text, ",", ".")), "LT", ""))
    
    ' Validar si es numérico antes de convertir
    If IsNumeric(texto) Then
        litros = CDbl(texto)
        txtCantidadLt.Text = CStr(litros) & " LT"
        
        ' Actualizar datos del producto y recalcular
        producto.litros = litros
        producto.PrecioUnitario = CDbl(txtMontoCombustible.Text)
        precioNeto.Text = FormatoPrecio(producto.CalcularPrecioNeto())
    Else
        txtCantidadLt.Text = ""
        producto.litros = 0
        precioNeto.Text = FormatoPrecio(0)
    End If
End Sub

Private Sub C_Change()

End Sub

Private Sub txtMontoCombustible_LostFocus()
txtMontoCombustible.Text = FormatoPrecio(txtMontoCombustible.Text)
End Sub

Private Sub CerrarRSyConexion(rs As ADODB.Recordset)
    If Not rs Is Nothing Then If rs.State = adStateOpen Then rs.Close
    If conn.State = adStateOpen Then Call DesconectarBD
End Sub
