VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FRMListadoITC 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Listado de ITC Mensual"
   ClientHeight    =   7950
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   5535
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7950
   ScaleWidth      =   5535
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame2 
      Height          =   735
      Left            =   120
      TabIndex        =   7
      Top             =   7080
      Width           =   5295
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
         Left            =   3960
         TabIndex        =   8
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Impuesto ITC"
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
      Height          =   6855
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5295
      Begin MSComctlLib.ListView lvlITC 
         Height          =   4935
         Left            =   600
         TabIndex        =   6
         Top             =   1800
         Width           =   3975
         _ExtentX        =   7011
         _ExtentY        =   8705
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
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
      Begin VB.CommandButton btnCargar 
         Caption         =   "&Cargar"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   600
         TabIndex        =   5
         Top             =   1320
         Width           =   3975
      End
      Begin VB.ComboBox cboMeses 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   2640
         TabIndex        =   2
         Text            =   "-- TODO --"
         Top             =   840
         Width           =   1935
      End
      Begin VB.TextBox txtAño 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   600
         TabIndex        =   1
         Text            =   "2025"
         Top             =   840
         Width           =   1935
      End
      Begin VB.Label Label2 
         Caption         =   "Mes Proceso"
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
         Height          =   255
         Left            =   2640
         TabIndex        =   4
         Top             =   600
         Width           =   1935
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Año Proceso"
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
         Height          =   195
         Left            =   600
         TabIndex        =   3
         Top             =   600
         Width           =   1215
      End
   End
End
Attribute VB_Name = "FRMListadoITC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub enero_Change()

End Sub

Public Sub CargarMeses()
    Dim rs As New ADODB.Recordset
    Dim i As Integer
    
    ' Conectar a la base de datos
    Call ConectarBD
    
    On Error GoTo ErrHandler
    
    cboMeses.Clear
    
        ' Agregar la opción "TODO" con id = 999
    cboMeses.AddItem "-- TODO --"
    cboMeses.ItemData(cboMeses.NewIndex) = 999
    
    rs.Open "SELECT id, mes FROM mes ORDER BY id", conn, adOpenStatic, adLockReadOnly
    
    ' Cargar los meses desde la base de datos al ComboBox
    Do While Not rs.EOF
        ' Puedes guardar el ID en ItemData si querés usarlo después
        cboMeses.AddItem rs("mes")
        cboMeses.ItemData(cboMeses.NewIndex) = rs("id")
        rs.MoveNext
    Loop
    
    ' Seleccionar "TODO" por defecto
    cboMeses.ListIndex = 0
    
    ' Cerrar el recordset
    rs.Close
    
    ' Desconectar
    Call DesconectarBD
    Exit Sub
    
ErrHandler:
    MsgBox "Error al cargar los meses: " & Err.Description, vbCritical, "Error"
    Call DesconectarBD
End Sub

Public Sub CargarGrilla()
    Dim rs As New ADODB.Recordset
    
    Call ConectarBD
    
    On Error GoTo ErrHandler
    
    ' Limpiar el ListView antes de agregar los nuevos datos
    lvlITC.ListItems.Clear
    
    If cboMeses.ItemData(cboMeses.ListIndex) = 999 Then
        rs.Open "SELECT '' AS valor,fechaOperacion,tipo,monto FROM Impuestos WHERE Year(fechaOperacion) = " & txtAño & " ORDER BY fechaOperacion ASC", conn, adOpenStatic, adLockReadOnly
    Else
        rs.Open "SELECT '' AS valor,fechaOperacion,tipo,monto FROM Impuestos WHERE Year(fechaOperacion) = " & txtAño & " AND Month(fechaOperacion) = " & cboMeses.ItemData(cboMeses.ListIndex) & " ORDER BY fechaOperacion ASC", conn, adOpenStatic, adLockReadOnly
    End If
     ' Cargar datos en el ListView
    If Not rs.EOF Then
        Do While Not rs.EOF
            With lvlITC.ListItems.Add(, , rs("valor"))
                .SubItems(1) = rs("fechaOperacion")
                .SubItems(2) = rs("tipo")
                .SubItems(3) = FormatoPrecio(rs("monto"))
            End With
            rs.MoveNext
        Loop
    Else
        MsgBox "No hay impuestos registrados.", vbExclamation, "Aviso"
    End If
    
    ' Cerrar el recordset
    rs.Close
    
    ' Desconectar
    Call DesconectarBD
    Exit Sub
    
ErrHandler:
    MsgBox "Error al cargar los impuesto: " & Err.Description, vbCritical, "Error"
    Call DesconectarBD
End Sub

Private Sub btnCargar_Click()
Dim año As String
año = Trim(txtAño.Text)
If Len(año) = 4 Then
    Call CargarGrilla
Else
    MsgBox "Ingrese correctamente el año", vbCritical, "ESAPP"
End If
End Sub

Private Sub btnSalir_Click()
Unload Me
End Sub

Private Sub Form_Load()
Call CargarMeses

' Configuración del ListView
With lvlITC
    .View = lvwReport
    .ColumnHeaders.Add , , "", 0
    .ColumnHeaders.Add , , "Fecha", 1500
    .ColumnHeaders.Add , , "Tipo", 1100
    .ColumnHeaders.Add , , "Monto", 1300
End With
End Sub

Private Sub txtAño_KeyPress(KeyAscii As Integer)
If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 Then
        KeyAscii = 0 ' Cancelar el caracter si no es un numero
    End If
End Sub
