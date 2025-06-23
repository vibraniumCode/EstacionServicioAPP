VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form FRMControl 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Control"
   ClientHeight    =   5175
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4815
   Icon            =   "FRMControl.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5175
   ScaleWidth      =   4815
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      Height          =   4935
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4575
      Begin VB.CommandButton btnSalir 
         Caption         =   "&Aceptar"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   3360
         Picture         =   "FRMControl.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   4200
         Width           =   1095
      End
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
         Height          =   3135
         Left            =   120
         TabIndex        =   4
         Top             =   960
         Width           =   4335
         _ExtentX        =   7646
         _ExtentY        =   5530
         _Version        =   393216
         BackColorFixed  =   -2147483643
         BackColorBkg    =   -2147483643
         FillStyle       =   1
         GridLines       =   0
         GridLinesFixed  =   0
         FormatString    =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Frame Frame2 
         Height          =   615
         Left            =   120
         TabIndex        =   2
         Top             =   120
         Width           =   4335
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            Caption         =   "10/07/2025"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   3
            Top             =   240
            Width           =   4095
         End
      End
      Begin VB.Line Line1 
         BorderWidth     =   2
         X1              =   3480
         X2              =   120
         Y1              =   4500
         Y2              =   4500
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00800000&
         Caption         =   "TOTAL CON IMPUESTO"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Left            =   105
         TabIndex        =   1
         Top             =   720
         Width           =   4335
      End
   End
End
Attribute VB_Name = "frmControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

End Sub

Private Sub btnSalir_Click()
Unload Me
End Sub

Private Sub Form_Load()
    Dim rs As New ADODB.Recordset
    Dim campos As Variant
    Dim i As Integer

    Call ConectarBD

    On Error GoTo ErrHandler

    rs.Open executeSQL, conn, adOpenStatic, adLockReadOnly
    
    Select Case idEstacionEmpresa
        Case 1
            campos = Array("Litros cargados", "Precio por litro", "Impuesto ITC por litro", "Impuesto detallados", "Subtotal sin impuestos", _
                   "IVA", "Impuesto interno", "Impuesto IDC", "Total de impuestos", "Total a pagar")
            FRMFacturacion.txtIva.Text = rs(5)
            FRMFacturacion.txtImpuesto.Text = rs(8)
            FRMFacturacion.txtTotal.Text = rs(9)
        Case 2
            campos = Array("Litros cargados", "Precio por litro", "Neto Gravado", "IVA", "Impuesto interno", _
                   "Impuesto ITC por litro", "Total a pagar")
            FRMFacturacion.txtIva.Text = rs(3)
            FRMFacturacion.txtImpuesto.Text = rs(4)
            FRMFacturacion.txtTotal.Text = rs(6)
        Case 3
            campos = Array("Litros cargados", "Precio por litro", "Neto Gravado", "IVA", "Impuesto interno", _
                   "Impuesto ITC por litro", "Total a pagar")
            FRMFacturacion.txtIva.Text = rs(3)
            FRMFacturacion.txtImpuesto.Text = rs(4)
            FRMFacturacion.txtTotal.Text = rs(6)
        Case 4
            campos = Array("Litros cargados", "Precio por litro", "Neto Gravado", "IVA", "Impuesto interno", _
                   "Impuesto ITC por litro", "Total a pagar")
            FRMFacturacion.txtIva.Text = rs(3)
            FRMFacturacion.txtImpuesto.Text = rs(4)
            FRMFacturacion.txtTotal.Text = rs(6)
        Case 5
            campos = Array("Litros cargados", "Precio por litro", "Neto Gravado", "IVA", "Impuesto interno", _
                   "Impuesto ITC por litro", "Imp. Hidr. de Carbono", "Imp. Combustibles Líq.", "Imp. a Mat. Variables", "Total a pagar")
            FRMFacturacion.txtIva.Text = rs(3)
            FRMFacturacion.txtImpuesto.Text = rs(4)
            FRMFacturacion.txtTotal.Text = rs(9)
    End Select

    If Not rs.EOF Then
        With MSFlexGrid1
            .Clear
            .FixedRows = 0 ' quitar encabezado real
            .Cols = 2
            .Rows = UBound(campos) + 1 ' solo filas para datos
            
            ' Cargar datos desde fila 0
            For i = 0 To UBound(campos)
                .TextMatrix(i, 0) = campos(i)
                .TextMatrix(i, 1) = rs(campos(i))
            Next i

            ' Ajustar anchos
            .ColWidth(0) = 2000
            .ColWidth(1) = 2200

            ' Sin líneas de grilla
            .GridLines = flexGridNone

            ' Pintar la última fila
            Dim fila As Integer, col As Integer
            fila = .Rows - 1
            For col = 0 To .Cols - 1
                .Row = fila
                .col = col
                .CellBackColor = vbBlue
                .CellForeColor = vbWhite
            Next col
        End With
    End If
    
    ' Ajustar altura del grid según cantidad de filas cargadas
    Dim alturaFila As Long
    alturaFila = 270 ' Tamaño típico de una fila en twips
    MSFlexGrid1.Height = alturaFila * MSFlexGrid1.Rows
    
    MSFlexGrid1.Row = 0

    If rs.State = adStateOpen Then rs.Close
    If conn.State = adStateOpen Then Call DesconectarBD
    Exit Sub

ErrHandler:
    MsgBox "Error al cargar datos: " & Err.Description, vbCritical, "Error"
    If rs.State = adStateOpen Then rs.Close
    If conn.State = adStateOpen Then Call DesconectarBD
End Sub

