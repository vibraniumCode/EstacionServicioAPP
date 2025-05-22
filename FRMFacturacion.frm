VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FRMFacturacion 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Carga masiva"
   ClientHeight    =   10935
   ClientLeft      =   150
   ClientTop       =   -1590
   ClientWidth     =   13845
   Icon            =   "FRMFacturacion.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10935
   ScaleWidth      =   13845
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   10695
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   13575
      Begin VB.Frame Frame12 
         Height          =   855
         Left            =   120
         TabIndex        =   35
         Top             =   9720
         Width           =   13335
         Begin VB.CommandButton btnTickets 
            Caption         =   "&Listado de Tickets"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   8520
            TabIndex        =   40
            Top             =   240
            Width           =   2295
         End
         Begin VB.CommandButton btnFinalizar 
            Caption         =   "&Finalizar Venta"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   11040
            TabIndex        =   39
            Top             =   240
            Width           =   2055
         End
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
            TabIndex        =   38
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
         TabIndex        =   26
         Top             =   1080
         Width           =   13335
         Begin VB.ComboBox cboClientes 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   120
            TabIndex        =   27
            Text            =   "Combo1"
            Top             =   360
            Width           =   13095
         End
      End
      Begin MSComCtl2.DTPicker fecEmision 
         Height          =   345
         Left            =   11280
         TabIndex        =   24
         Top             =   600
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
         Format          =   138149889
         CurrentDate     =   45777
      End
      Begin VB.Frame Frame7 
         Height          =   1335
         Left            =   120
         TabIndex        =   14
         Top             =   8280
         Width           =   13335
         Begin VB.TextBox txtITC 
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
            TabIndex        =   36
            Text            =   "$00.00"
            Top             =   600
            Width           =   3135
         End
         Begin VB.TextBox txtTotal 
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
            Left            =   9840
            TabIndex        =   21
            Text            =   "$00.00"
            Top             =   600
            Width           =   3375
         End
         Begin VB.TextBox txtIva 
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
            TabIndex        =   20
            Text            =   "$00.00"
            Top             =   600
            Width           =   3135
         End
         Begin VB.TextBox txtSubtotal 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   120
            TabIndex        =   19
            Text            =   "$00.00"
            Top             =   600
            Width           =   3135
         End
         Begin VB.Line Line2 
            BorderColor     =   &H000040C0&
            BorderWidth     =   3
            DrawMode        =   14  'Copy Pen
            Index           =   2
            X1              =   6600
            X2              =   8880
            Y1              =   480
            Y2              =   480
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "ITC"
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
            TabIndex        =   37
            Top             =   240
            Width           =   300
         End
         Begin VB.Line Line2 
            BorderColor     =   &H00000080&
            BorderWidth     =   3
            DrawMode        =   14  'Copy Pen
            Index           =   1
            X1              =   9840
            X2              =   12120
            Y1              =   480
            Y2              =   480
         End
         Begin VB.Line Line3 
            BorderColor     =   &H00008000&
            BorderWidth     =   3
            DrawMode        =   14  'Copy Pen
            X1              =   3360
            X2              =   5040
            Y1              =   480
            Y2              =   480
         End
         Begin VB.Line Line2 
            BorderColor     =   &H8000000D&
            BorderWidth     =   3
            DrawMode        =   14  'Copy Pen
            Index           =   0
            X1              =   120
            X2              =   2160
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
            Left            =   9840
            TabIndex        =   18
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
            Left            =   3360
            TabIndex        =   16
            Top             =   240
            Width           =   1320
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "SUBTOTAL"
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
            TabIndex        =   15
            Top             =   240
            Width           =   945
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
         TabIndex        =   6
         Top             =   2040
         Width           =   13335
         Begin VB.Frame Frame5 
            Caption         =   "ITC $"
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
            Left            =   1800
            TabIndex        =   33
            Top             =   1200
            Width           =   2415
            Begin VB.TextBox txtMontoITC 
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
               Height          =   360
               Left            =   120
               TabIndex        =   34
               Top             =   300
               Width           =   1935
            End
            Begin VB.Image Image1 
               Height          =   255
               Left            =   2080
               Picture         =   "FRMFacturacion.frx":67DA
               Stretch         =   -1  'True
               Top             =   360
               Width           =   255
            End
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
            TabIndex        =   31
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
               TabIndex        =   32
               Text            =   "$00.00"
               Top             =   240
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
            TabIndex        =   28
            Top             =   360
            Width           =   5655
            Begin VB.TextBox txtImpuestoITC 
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
               TabIndex        =   30
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
               TabIndex        =   29
               Top             =   300
               Width           =   2775
            End
         End
         Begin VB.CommandButton btnActualizarproducto 
            Caption         =   "&Actualizar producto"
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
            Height          =   495
            Left            =   11160
            TabIndex        =   23
            Top             =   1440
            Width           =   2055
         End
         Begin VB.CommandButton btnIngresarproducto 
            Caption         =   "&Ingresar producto"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   9000
            TabIndex        =   17
            Top             =   1440
            Width           =   2055
         End
         Begin VB.Frame Frame6 
            Caption         =   "IVA %"
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
            TabIndex        =   11
            Top             =   1200
            Width           =   1575
            Begin VB.TextBox iva 
               Alignment       =   1  'Right Justify
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
               Height          =   360
               Left            =   120
               TabIndex        =   13
               Text            =   "21.00"
               Top             =   300
               Width           =   1335
            End
            Begin VB.Label Label8 
               Caption         =   "IVA"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   -840
               TabIndex        =   12
               Top             =   1080
               Width           =   1335
            End
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
            TabIndex        =   7
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
               TabIndex        =   10
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
               TabIndex        =   9
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
               Locked          =   -1  'True
               TabIndex        =   8
               Text            =   "0"
               Top             =   300
               Width           =   1695
            End
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
            X1              =   13310
            X2              =   0
            Y1              =   1680
            Y2              =   1680
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Registros"
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
         Height          =   3975
         Left            =   120
         TabIndex        =   5
         Top             =   4200
         Width           =   13335
         Begin MSComctlLib.ListView Grilla 
            Height          =   3615
            Left            =   120
            TabIndex        =   22
            Top             =   240
            Width           =   13095
            _ExtentX        =   23098
            _ExtentY        =   6376
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
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            NumItems        =   0
         End
      End
      Begin VB.TextBox factura 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1800
         TabIndex        =   2
         Top             =   240
         Width           =   2895
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
         TabIndex        =   25
         Top             =   675
         Width           =   1545
      End
      Begin VB.Label Label4 
         Caption         =   "A"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1800
         TabIndex        =   4
         Top             =   600
         Width           =   255
      End
      Begin VB.Label Label3 
         Caption         =   "Tipo de factura:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   720
         Width           =   1695
      End
      Begin VB.Line Line1 
         BorderWidth     =   2
         X1              =   4680
         X2              =   13440
         Y1              =   405
         Y2              =   405
      End
      Begin VB.Label Label2 
         Caption         =   "FACTURA:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   1575
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
Dim Producto As New ClaseProducto
Dim Clientes As New ClaseCliente

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

' Declaraciones (poner en el formulario o en un módulo .bas)


Private Sub Actualizar_Click()
    Dim test As String
    
    test = Grilla.SelectedItem.Text
    
    ' Verificar si hay elementos en la grilla
    If Grilla.ListItems.Count = 0 Then
        MsgBox "No hay productos para actualizar", vbExclamation
        Exit Sub
    End If
    
    ' Verificar si hay un elemento seleccionado
    On Error Resume Next
    
    If Err.Number <> 0 Then
        MsgBox "Por favor, seleccione un producto para actualizar", vbExclamation
        Exit Sub
    End If
    On Error GoTo 0
    
    ' Una vez confirmado que hay un elemento seleccionado, cargar sus datos
    Call CargarDatosParaActualizar
    
    ' Cambiar visibilidad de los botones (si es necesario)
    btnIngresarproducto.Visible = False
    btnActualizarproducto.Visible = True
End Sub

Private Sub CargarDatosParaActualizar()
    ' Obtener el ID del producto seleccionado
    idProducto = CLng(Grilla.SelectedItem.Text)
    
    ' Guardar el ID en el Tag del formulario
    Me.Tag = CStr(idProducto)
    
    ' Cargar datos en los TextBox
    txtDescripcion.Text = Grilla.SelectedItem.SubItems(1)
    btnCantidad.Text = Grilla.SelectedItem.SubItems(2)
    Preciouni.Text = Grilla.SelectedItem.SubItems(3)
    precioNeto.Text = Grilla.SelectedItem.SubItems(4)
    ' Agregar más campos según sea necesario
End Sub

Private Sub btnActualizarproducto_Click()
' Validar los datos
    If Trim(txtDescripcion.Text) = "" Then
        MsgBox "La descripción no puede estar vacía", vbExclamation
        txtDescripcion.SetFocus
        Exit Sub
    End If
    
    ' Conectar a la base de datos
    Call ConectarBD
    
    If txtDescripcion.Text = "" Then
        MostrarAlerta "Ingrese una descripción del producto."
        Exit Sub
    ElseIf btnCantidad = 0 Then
        MostrarAlerta "La cantidad no puede ser cero. Ingrese un valor válido."
        Exit Sub
    ElseIf Preciouni.Text = 0 Then
        MostrarAlerta "El precio unitario no puede ser cero. Ingrese un valor válido."
        Exit Sub
    End If
    
    ' Actualizar el registro en la base de datos
    On Error GoTo ErrHandler
    conn.Execute "UPDATE PRODUCTOS_VENTAS SET " & _
                "DESCRIPCION = '" & Replace(txtDescripcion.Text, "'", "''") & "', " & _
                "CANTIDAD = " & Replace(btnCantidad.Text, ",", ".") & ", " & _
                "PRECIO_UNITARIO = " & Producto.PrecioUnitario & ", " & _
                "PRECIO_NETO = " & Producto.precioNeto & _
                " WHERE ID = " & idProducto
    
    ' Desconectar de la base de datos
    Call DesconectarBD
    
    ' Actualizar la grilla
    Call CargarGrilla
    Call CalculoGral
    ' Limpiar los campos y restablecer botones
    LimpiarCampos
    btnIngresarproducto.Visible = True
    btnActualizarproducto.Visible = False
    
    MsgBox "Producto actualizado correctamente", vbInformation
    
    Exit Sub
    
ErrHandler:
    MsgBox "Error al actualizar el producto: " & Err.Description, vbCritical, "Error"
    Call DesconectarBD
End Sub

Private Sub btnFinalizar_Click()

If Grilla.ListItems.Count > 0 Then
    Call ConectarBD
    On Error GoTo ErrHandler
    conn.Execute "INSERT INTO FACTURAS SELECT 'A'"
    MsgBox "Proceso finalizado", vbInformation
    Call DesconectarBD
    CargarNumeroFactura
    Exit Sub
Else
    MostrarAlerta "No se puede finalizar si aun no ingresaste productos"
    Exit Sub
End If

ErrHandler:
    MsgBox "Error al eliminar el producto: " & Err.Description, vbCritical, "Error"
    Call DesconectarBD
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
            txtImpuestoITC.Text = Format(rs("precio"), "0.00")
        Else
            txtImpuestoITC.Text = ""
        End If
        
        txtImpuestoITC.Text = FormatoPrecio(txtImpuestoITC.Text)
        
        rs.Close
        Call DesconectarBD
    End If
    Exit Sub

ErrHandler:
    MsgBox "Error al obtener el precio del combustible: " & Err.Description, vbCritical, "Error"
    Call DesconectarBD
End Sub




Private Sub Eliminar_Click()
    Dim idProducto As Long
    
    ' Verificar si hay un elemento seleccionado
    If Grilla.SelectedItem Is Nothing Then
        MsgBox "Por favor, seleccione un elemento para eliminar", vbExclamation
        Exit Sub
    End If
    
    ' Obtener el ID del producto desde el ListView
    idProducto = CLng(Grilla.SelectedItem.Text)
    
    ' Conectar a la base de datos
    Call ConectarBD
    
    ' Eliminar el producto de la base de datos usando Execute
    On Error GoTo ErrHandler
    conn.Execute "DELETE FROM PRODUCTOS_VENTAS WHERE id = " & idProducto
    
    ' Desconectar de la base de datos
    Call DesconectarBD
    
    ' Eliminar el item seleccionado del ListView
    Grilla.ListItems.Remove Grilla.SelectedItem.Index
    
    MsgBox "Registro eliminado correctamente", vbInformation
    Call CalculoGral
    Exit Sub
    
ErrHandler:
    MsgBox "Error al eliminar el producto: " & Err.Description, vbCritical, "Error"
    Call DesconectarBD
End Sub



Private Sub Form_Load()
Dim Cantidad As Integer

  Dim hMenu As Long

' Obtener el menú del sistema
hMenu = GetSystemMenu(Me.hWnd, False)

' Eliminar solo el botón de maximizar
DeleteMenu hMenu, SC_MAXIMIZE, MF_BYCOMMAND

CargarNumeroFactura

' Cargar los datos de clientes ANTES de cargar el combo
ModuloClientes.CargarDatosClientes
Call CargarClientesCombo
Call CargarComboCombustible
Call CargarImpuestoITC

'Producto.Cantidad = 0
'Producto.PrecioUnitario = 0
'Producto.precioNeto = 0
'Producto.Descripcion = ""
'
'' Conectar a la base de datos
'Call ConectarBD
'
'' Configuración del ListView
'With Grilla
'    .View = lvwReport
'    .ColumnHeaders.Add , , "Id", 1000
'    .ColumnHeaders.Add , , "Descripción", 2000
'    .ColumnHeaders.Add , , "Cantidad", 1000
'    .ColumnHeaders.Add , , "Precio Unitario", 1500
'    .ColumnHeaders.Add , , "Precio Neto", 1500
'End With
'
'' Cargar datos en el ListView
'Call CargarGrilla
'
'
'Cantidad = Grilla.ListItems.Count
'If Cantidad > 0 Then
'
'    Call CalculoGral
'End If



End Sub
    
Private Sub btnIngresarproducto_Click()
Dim cmd As New ADODB.Command
'Dim facturaNumero As Double

'facturaNumero = Val(factura.Text)
If Not IsNumeric(Producto.Cantidad) Or Not IsNumeric(Producto.PrecioUnitario) Then
    MsgBox "La cantidad y el precio deben ser números válidos.", vbCritical, "Error"
    Exit Sub
End If

' Conectar a la base de datos utilizando el módulo de conexión
    Call ConectarBD

    If Producto.Descripcion = "" Then
        MostrarAlerta "Ingrese una descripción del producto."
        Exit Sub
    ElseIf Producto.Cantidad = 0 Then
        MostrarAlerta "La cantidad no puede ser cero. Ingrese un valor válido."
        Exit Sub
    ElseIf Producto.PrecioUnitario = 0 Then
        MostrarAlerta "El precio unitario no puede ser cero. Ingrese un valor válido."
        Exit Sub
    End If

 'Preparar comando SQL para insertar datos
    On Error GoTo ErrHandler
    With cmd
        .ActiveConnection = conn
        .CommandType = adCmdText
        .CommandText = "INSERT INTO PRODUCTOS_VENTAS (DESCRIPCION, CANTIDAD, PRECIO_UNITARIO, PRECIO_NETO, FACTURA) VALUES (?, ?, ?, ?, ?)"
        .Parameters.Append .CreateParameter("DESCRIPCION", adVarChar, adParamInput, 255, Producto.Descripcion)
        .Parameters.Append .CreateParameter("CANTIDAD", adInteger, adParamInput, , Producto.Cantidad)
        .Parameters.Append .CreateParameter("PRECIO_UNITARIO", adDouble, adParamInput, , Producto.PrecioUnitario)
        .Parameters.Append .CreateParameter("PRECIO_NETO", adDouble, adParamInput, , Producto.precioNeto)
        .Parameters.Append .CreateParameter("FACTURA", adDouble, adParamInput, , nroFactura)
        .Execute
    End With

    ' Actualizar el ListView después de la inserción
    Call CargarGrilla
    Call CalculoGral
    LimpiarCampos
    
    ' Desconectar
    Call DesconectarBD
    Exit Sub

ErrHandler:
    MsgBox "Error al insertar: " & Err.Description, vbCritical, "Error"
    Call DesconectarBD
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

Private Sub Grilla_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
 ' Mostrar el menú contextual solo si se hace clic derecho
    If Button = vbRightButton Then
        ' Mostrar el menú emergente
        PopupMenu mnuListView
    End If
End Sub





Private Sub Image1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Frame5.Caption = "ITC $ - " & fecImpITC
End Sub

Private Sub precioNeto_LostFocus()
    Producto.precioNeto = LimpiarValor(precioNeto.Text)
    precioNeto.Text = FormatoPrecio(Producto.precioNeto)
End Sub





Private Sub CargarPrecio()
    precioNeto.Text = FormatoPrecio(Producto.CalcularPrecioNeto())
End Sub

Private Sub sumar_Click()
    Dim valor As String
    Dim numero As Long
    alertaMostrada = False
    
    valor = Trim(Replace(UCase(txtCantidadLt.Text), "LT", ""))
    
    If IsNumeric(valor) Then
        Producto.Cantidad = CLng(valor) + 1
        txtCantidadLt.Text = CStr(Producto.Cantidad) & " LT"
    Else
        txtCantidadLt.Text = ""
    End If
'    Producto.Cantidad = Producto.Cantidad + 1
'    btnCantidad.Text = Producto.Cantidad
'    Call ActualizarPrecio
End Sub

Private Sub restar_Click()
    Dim valor As String
    Dim numero As Long
    alertaMostrada = False
    
    valor = Trim(Replace(UCase(txtCantidadLt.Text), "LT", ""))
    
    If IsNumeric(valor) Then
        Producto.Cantidad = CLng(valor) - 1
        txtCantidadLt.Text = CStr(Producto.Cantidad) & " LT"
    Else
        txtCantidadLt.Text = ""
    End If
'    Producto.Cantidad = Producto.Cantidad - 1
'    btnCantidad.Text = Producto.Cantidad
'    Call ActualizarPrecio
End Sub

Private Sub ActualizarPrecio()
    precioNeto.Text = FormatoPrecio(Producto.CalcularPrecioNeto())
    precioNeto_LostFocus
End Sub










Private Sub CargarGrilla()
    Dim rs As New ADODB.Recordset

    ' Conectar a la base de datos utilizando el módulo de conexión
    Call ConectarBD

    ' Limpiar el ListView antes de agregar los nuevos datos
    Grilla.ListItems.Clear

    ' Obtener datos de la base
    On Error GoTo ErrHandler

    rs.Open "SELECT id, descripcion, cantidad, precio_unitario, precio_neto FROM PRODUCTOS_VENTAS WHERE FACTURA = " & nroFactura, conn, adOpenStatic, adLockReadOnly

    ' Cargar datos en el ListView
    If Not rs.EOF Then
        Do While Not rs.EOF
            With Grilla.ListItems.Add(, , rs("id"))
                .SubItems(1) = rs("descripcion")
                .SubItems(2) = rs("cantidad")
                .SubItems(3) = "$" & Format(rs("precio_unitario"), "#,##0.00")
                .SubItems(4) = "$" & Format(rs("precio_neto"), "#,##0.00")
            End With
            rs.MoveNext
        Loop
    Else
        MsgBox "No hay productos registrados.", vbExclamation, "Aviso"
    End If
    Grilla.ColumnHeaders(1).Width = 0
    ' Cerrar el recordset
    rs.Close
    
    ' Desconectar
    Call DesconectarBD
    Exit Sub

ErrHandler:
    MsgBox "Error al cargar datos: " & Err.Description, vbCritical, "Error"
    Call DesconectarBD
End Sub
Private Sub CalculoGral()
    Dim rs As New ADODB.Recordset

    ' Conectar a la base de datos utilizando el módulo de conexión
    Call ConectarBD

    ' Obtener datos de la base
    On Error GoTo ErrHandler
    rs.Open "SELECT " & _
                "SUM(PRECIO_NETO) SUBTOTAL, " & _
                "SUM(PRECIO_NETO * 0.21) AS IVA, " & _
                "SUM(PRECIO_NETO) + SUM(PRECIO_NETO * 0.21) AS TOTAL " & _
                "FROM PRODUCTOS_VENTAS Where factura = " & nroFactura, conn, adOpenStatic, adLockReadOnly
    
    txtSubtotal.Text = Format(rs(0), "$0.00")
    txtIva.Text = Format(rs(1), "$0.00")
    txtTotal.Text = Format(rs(2), "$0.00")
    
    ' Cerrar el recordset
    rs.Close
    
    ' Desconectar
    Call DesconectarBD
    Exit Sub

ErrHandler:
    MsgBox "Error al cargar datos: " & Err.Description, vbCritical, "Error"
    Call DesconectarBD
End Sub

' Procedimiento para limpiar los campos
Private Sub LimpiarCampos()
    txtDescripcion.Text = ""
    btnCantidad.Text = 1
    Preciouni.Text = "$" & Format(0, "#,##0.00")
    precioNeto.Text = "$" & Format(0, "#,##0.00")
    Me.Tag = ""  ' Limpiar el ID guardado
End Sub

Private Sub CargarNumeroFactura()
    Dim rs As New ADODB.Recordset

    ' Conectar a la base de datos
    Call ConectarBD

    ' Obtener el último número de factura
    On Error GoTo ErrHandler
    rs.Open "SELECT MAX(factura) AS UltimoNro FROM FACTURAS", conn, adOpenStatic, adLockReadOnly

    ' Verificar si hay datos
    If Not rs.EOF Then
        nroFactura = rs("UltimoNro")
        factura.Text = "N°0001-" & FormatearNumeroFactura(nroFactura)
    Else
        MsgBox "No hay facturas registradas.", vbExclamation, "Aviso"
    End If

    ' Cerrar el recordset y desconectar
    rs.Close
    Call DesconectarBD
    Exit Sub

ErrHandler:
    MsgBox "Error al obtener el número de factura: " & Err.Description, vbCritical, "Error"
    Call DesconectarBD
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
    
    ' Cerrar el recordset
    rs.Close
    
    ' Desconectar
    Call DesconectarBD
    Exit Sub

ErrHandler:
    MsgBox "Error al cargar el impuesto ITC: " & Err.Description, vbCritical, "Error"
    Call DesconectarBD
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
    
    cboCombustible.ListIndex = 0
    
    If rs.State = adStateOpen Then
        ' Cerrar el recordset
        rs.Close
    End If
    
    If conn.State = adStateOpen Then
        ' Desconectar
        Call DesconectarBD
    End If
    
    Exit Sub
    
ErrHandler:
    MsgBox "Error al cargar el listado de combustible: " & Err.Description, vbCritical, "Error"
    Call DesconectarBD
End Sub

Private Sub txtCantidadLt_LostFocus()
Dim valor As String
Dim numero As Long
    
valor = Trim(Replace(UCase(txtCantidadLt.Text), "LT", ""))
    
If IsNumeric(valor) Then
    numero = CLng(valor)
    txtCantidadLt.Text = CStr(numero) & " LT"
Else
    txtCantidadLt.Text = ""
End If


End Sub

Private Sub txtImpuestoITC_LostFocus()
txtImpuestoITC.Text = FormatoPrecio(txtImpuestoITC.Text)
End Sub

