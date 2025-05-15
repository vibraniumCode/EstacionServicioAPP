VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form Facturacion 
   Caption         =   "Carga masiva"
   ClientHeight    =   9885
   ClientLeft      =   120
   ClientTop       =   765
   ClientWidth     =   13575
   Icon            =   "Facturacion.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   9885
   ScaleWidth      =   13575
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame8 
      Height          =   615
      Left            =   120
      TabIndex        =   32
      Top             =   9240
      Width           =   13335
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
         Height          =   375
         Left            =   11160
         TabIndex        =   36
         Top             =   160
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
         TabIndex        =   33
         Top             =   240
         Width           =   3285
      End
   End
   Begin VB.Frame Frame1 
      Height          =   9135
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   13335
      Begin MSComCtl2.DTPicker fecEmision 
         Height          =   345
         Left            =   11040
         TabIndex        =   44
         Top             =   600
         Width           =   2100
         _ExtentX        =   3704
         _ExtentY        =   609
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   188678145
         CurrentDate     =   45777
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
         Height          =   1215
         Left            =   120
         TabIndex        =   37
         Top             =   2520
         Width           =   13095
         Begin VB.TextBox txtDireccion 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   1080
            TabIndex        =   43
            Top             =   770
            Width           =   11895
         End
         Begin VB.TextBox txtCuit 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   9240
            TabIndex        =   41
            Top             =   300
            Width           =   3735
         End
         Begin VB.TextBox txtNombre 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   1080
            TabIndex        =   39
            Top             =   300
            Width           =   7575
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            Caption         =   "DIRECCIÓN"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   120
            TabIndex        =   42
            Top             =   840
            Width           =   855
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            Caption         =   "CUIT"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   8760
            TabIndex        =   40
            Top             =   360
            Width           =   360
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "NOMBRE"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   120
            TabIndex        =   38
            Top             =   360
            Width           =   630
         End
      End
      Begin VB.Frame Frame7 
         Height          =   1335
         Left            =   120
         TabIndex        =   21
         Top             =   7680
         Width           =   13095
         Begin VB.CommandButton ImpArturo 
            Caption         =   "&Comprobante neumaticos arturo"
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
            Left            =   9600
            MaskColor       =   &H00E0E0E0&
            Style           =   1  'Graphical
            TabIndex        =   30
            Top             =   240
            Width           =   3375
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
            Left            =   6000
            TabIndex        =   29
            Text            =   "$00.00"
            Top             =   600
            Width           =   3495
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
            Left            =   3240
            TabIndex        =   28
            Text            =   "$00.00"
            Top             =   600
            Width           =   2655
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
            TabIndex        =   27
            Text            =   "$00.00"
            Top             =   600
            Width           =   3015
         End
         Begin VB.CommandButton ImpFederal 
            Caption         =   "&Comprobante federal parts"
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
            Left            =   9600
            TabIndex        =   31
            Top             =   840
            Width           =   3375
         End
         Begin VB.Line Line2 
            BorderColor     =   &H00000080&
            BorderWidth     =   3
            DrawMode        =   14  'Copy Pen
            Index           =   1
            X1              =   6000
            X2              =   8280
            Y1              =   480
            Y2              =   480
         End
         Begin VB.Line Line3 
            BorderColor     =   &H00008000&
            BorderWidth     =   3
            DrawMode        =   14  'Copy Pen
            X1              =   3240
            X2              =   4920
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
            Left            =   6000
            TabIndex        =   26
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
            Left            =   3240
            TabIndex        =   23
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
            TabIndex        =   22
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
         Height          =   1575
         Left            =   120
         TabIndex        =   6
         Top             =   960
         Width           =   13095
         Begin VB.CommandButton btnActualizarproducto 
            Caption         =   "&Actualizar producto"
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
            Left            =   10920
            TabIndex        =   35
            Top             =   1080
            Visible         =   0   'False
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
            Left            =   10920
            TabIndex        =   25
            Top             =   840
            Width           =   2055
         End
         Begin VB.Frame Frame6 
            Height          =   735
            Left            =   9240
            TabIndex        =   18
            Top             =   720
            Width           =   1575
            Begin VB.TextBox iva 
               Alignment       =   1  'Right Justify
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   405
               Left            =   600
               TabIndex        =   20
               Text            =   "21.00"
               Top             =   240
               Width           =   855
            End
            Begin VB.Label Label11 
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
               Left            =   120
               TabIndex        =   24
               Top             =   360
               Width           =   375
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
               TabIndex        =   19
               Top             =   1080
               Width           =   1335
            End
         End
         Begin VB.Frame Frame5 
            Caption         =   "PRECIOS"
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
            Left            =   2520
            TabIndex        =   13
            Top             =   720
            Width           =   6615
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
               Left            =   4680
               Locked          =   -1  'True
               TabIndex        =   17
               Text            =   "$00.00"
               Top             =   240
               Width           =   1815
            End
            Begin VB.TextBox Preciouni 
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
               Height          =   375
               Left            =   1680
               TabIndex        =   15
               Text            =   "$00.00"
               Top             =   240
               Width           =   1815
            End
            Begin VB.Label Label6 
               Caption         =   "PRECIO NETO"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   3600
               TabIndex        =   16
               Top             =   360
               Width           =   1335
            End
            Begin VB.Label Label7 
               Caption         =   "PRECIO UNITARIO"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   120
               TabIndex        =   14
               Top             =   360
               Width           =   1455
            End
         End
         Begin VB.Frame Frame4 
            Caption         =   "CANTIDAD"
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
            Left            =   120
            TabIndex        =   9
            Top             =   720
            Width           =   2295
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
               TabIndex        =   12
               Top             =   240
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
               TabIndex        =   11
               Top             =   240
               Width           =   375
            End
            Begin VB.TextBox btnCantidad 
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
               Height          =   330
               Left            =   1080
               Locked          =   -1  'True
               TabIndex        =   10
               Text            =   "1"
               Top             =   240
               Width           =   1095
            End
         End
         Begin VB.TextBox txtDescripcion 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   1320
            TabIndex        =   8
            Top             =   240
            Width           =   11655
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "DESCRIPCIÓN:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   120
            TabIndex        =   7
            Top             =   360
            Width           =   1095
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
         Top             =   3720
         Width           =   13095
         Begin MSComctlLib.ListView Grilla 
            Height          =   3615
            Left            =   120
            TabIndex        =   34
            Top             =   240
            Width           =   12855
            _ExtentX        =   22675
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
         Left            =   9240
         TabIndex        =   45
         Top             =   670
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
         Top             =   680
         Width           =   1695
      End
      Begin VB.Line Line1 
         X1              =   4680
         X2              =   13200
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
      Begin VB.Menu Eliminar 
         Caption         =   "Eliminar"
      End
      Begin VB.Menu Actualizar 
         Caption         =   "Actualizar"
      End
   End
End
Attribute VB_Name = "Facturacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Producto As New ClaseProducto
Dim Clientes As New ClaseCliente

Dim alertaMostrada As Boolean
Dim idProducto As Long
'Dim nroFactura As Long
' Método 1: Bloquear botones de la ventana usando API de Windows
Private Declare Function GetSystemMenu Lib "user32" (ByVal hwnd As Long, ByVal bRevert As Long) As Long
Private Declare Function DeleteMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long

Private Const SC_MAXIMIZE = &HF030
Private Const MF_BYCOMMAND = &H0

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

Private Sub Command1_Click()

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

Private Sub fecha_Click()

End Sub

Private Sub Form_Load()
    Dim Cantidad As Integer
    
      Dim hMenu As Long
    
    ' Obtener el menú del sistema
    hMenu = GetSystemMenu(Me.hwnd, False)
    
    ' Eliminar solo el botón de maximizar
    DeleteMenu hMenu, SC_MAXIMIZE, MF_BYCOMMAND
    
    CargarNumeroFactura

    Producto.Cantidad = 0
    Producto.PrecioUnitario = 0
    Producto.precioNeto = 0
    Producto.Descripcion = ""
    
    ' Conectar a la base de datos
    Call ConectarBD

    ' Configuración del ListView
    With Grilla
        .View = lvwReport
        .ColumnHeaders.Add , , "Id", 1000
        .ColumnHeaders.Add , , "Descripción", 2000
        .ColumnHeaders.Add , , "Cantidad", 1000
        .ColumnHeaders.Add , , "Precio Unitario", 1500
        .ColumnHeaders.Add , , "Precio Neto", 1500
    End With

    ' Cargar datos en el ListView
    Call CargarGrilla
    

    Cantidad = Grilla.ListItems.Count
    If Cantidad > 0 Then

        Call CalculoGral
    End If
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

Private Sub Grilla_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
 ' Mostrar el menú contextual solo si se hace clic derecho
    If Button = vbRightButton Then
        ' Mostrar el menú emergente
        PopupMenu mnuListView
    End If
End Sub

Private Sub ImpArturo_Click()
    If ValidadorMsg Then Exit Sub
    GenerarComprobante txtSubtotal.Text, txtIva.Text, txtTotal.Text, Facturacion
End Sub

Private Sub ImpFederal_Click()
    If ValidadorMsg Then Exit Sub
    GenerarComprobante2 txtSubtotal.Text, txtIva.Text, txtTotal.Text, Facturacion
End Sub

Private Sub precioNeto_LostFocus()
    Producto.precioNeto = LimpiarValor(precioNeto.Text)
    precioNeto.Text = FormatoPrecio(Producto.precioNeto)
End Sub

Private Sub Preciouni_KeyPress(KeyAscii As Integer)
    ' Permitir solo números, el signo de dólar, el punto decimal, la retroceso y la barra espaciadora
    If Not ((KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 8 Or KeyAscii = 46 Or KeyAscii = 32) Then
        KeyAscii = 0
    End If
End Sub

Private Sub Preciouni_LostFocus()
    Producto.PrecioUnitario = LimpiarValor(Preciouni.Text)
    Preciouni.Text = FormatoPrecio(Producto.PrecioUnitario)
    Producto.Cantidad = Val(btnCantidad.Text)

    Call ActualizarPrecio
End Sub

Private Sub btnCantidad_Change()
    alertaMostrada = False
    Producto.Cantidad = Val(btnCantidad.Text)
    Call ActualizarPrecio
End Sub

Private Sub CargarPrecio()
    precioNeto.Text = FormatoPrecio(Producto.CalcularPrecioNeto())
End Sub

Private Sub sumar_Click()
    alertaMostrada = False
    Producto.Cantidad = Producto.Cantidad + 1
    btnCantidad.Text = Producto.Cantidad
    Call ActualizarPrecio
End Sub

Private Sub restar_Click()
    Producto.Cantidad = Producto.Cantidad - 1
    btnCantidad.Text = Producto.Cantidad
    Call ActualizarPrecio
End Sub

Private Sub ActualizarPrecio()
    precioNeto.Text = FormatoPrecio(Producto.CalcularPrecioNeto())
    precioNeto_LostFocus
End Sub


Private Sub txtCuit_LostFocus()
    Clientes.Cuit = txtCuit.Text
End Sub

Private Sub txtDescripcion_LostFocus()
    Producto.Descripcion = txtDescripcion.Text
End Sub

Private Sub txtDireccion_LostFocus()
    Clientes.Direccion = txtDireccion.Text
End Sub

Private Sub txtNombre_LostFocus()
    Clientes.Cliente = txtNombre.Text
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

