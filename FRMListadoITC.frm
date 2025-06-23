VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form FRMListadoImpuestos 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Listado de impuestos"
   ClientHeight    =   8535
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6390
   Icon            =   "FRMListadoITC.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8535
   ScaleWidth      =   6390
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame2 
      Height          =   975
      Left            =   120
      TabIndex        =   6
      Top             =   7440
      Width           =   6135
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
         Height          =   615
         Left            =   4800
         Picture         =   "FRMListadoITC.frx":058A
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Impuesto "
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
      Height          =   7335
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6135
      Begin VB.ComboBox cboEstaciones 
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
         Left            =   600
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   1200
         Width           =   4935
      End
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
         Height          =   4335
         Left            =   120
         TabIndex        =   9
         Top             =   2880
         Width           =   5895
         _ExtentX        =   10398
         _ExtentY        =   7646
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.ComboBox cboImpuestos 
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
         Left            =   600
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   1680
         Width           =   4935
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
         Height          =   615
         Left            =   600
         Picture         =   "FRMListadoITC.frx":0B14
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   2160
         Width           =   1215
      End
      Begin VB.ComboBox cboMeses 
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
         Left            =   2640
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   720
         Width           =   2895
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
         Top             =   720
         Width           =   1935
      End
      Begin VB.Line Line1 
         BorderWidth     =   2
         X1              =   1800
         X2              =   5520
         Y1              =   2470
         Y2              =   2470
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
         ForeColor       =   &H80000011&
         Height          =   255
         Left            =   2640
         TabIndex        =   4
         Top             =   480
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
         ForeColor       =   &H80000011&
         Height          =   195
         Left            =   600
         TabIndex        =   3
         Top             =   480
         Width           =   1215
      End
   End
End
Attribute VB_Name = "FRMListadoImpuestos"
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
    
    If rs.State = adStateOpen Then rs.Close
    If conn.State = adStateOpen Then Call DesconectarBD
    Exit Sub
    
ErrHandler:
    MsgBox "Error al cargar los meses: " & Err.Description, vbCritical, "Error"
    If rs.State = adStateOpen Then rs.Close
    If conn.State = adStateOpen Then Call DesconectarBD
End Sub

Public Sub CargarGrilla()
    Dim rs As New ADODB.Recordset
    Dim sql As String
    Dim fila As Integer
    Dim idImpuestoSeleccionado As Integer
    Dim idEmpresaSeleccionado As Integer
    Call ConectarBD
    
    On Error GoTo ErrHandler
    
    sql = "exec sp_impuestos " & cboImpuestos.ItemData(cboImpuestos.ListIndex) & ", NULL, NULL, NULL, "
    sql = sql + "" & cboEstaciones.ItemData(cboEstaciones.ListIndex) & ", " & txtAño.Text
    sql = sql + ", " & cboMeses.ItemData(cboMeses.ListIndex) & ", GRL"
    
    rs.Open sql, conn, adOpenStatic, adLockReadOnly
   
    ' Configurar columnas del MSFlexGrid
    With MSFlexGrid1
        .Clear
        .Rows = 1 ' Solo cabecera
        .Cols = 3
        .TextMatrix(0, 0) = "Fecha"
        .TextMatrix(0, 1) = "Impuesto"
        .TextMatrix(0, 2) = "Monto"

        ' Agregar datos fila por fila
        Do While Not rs.EOF
            .Rows = .Rows + 1
            fila = .Rows - 1
            .TextMatrix(fila, 0) = rs("fechaOperacion")
            .TextMatrix(fila, 1) = UCase(rs("tipo"))
            .TextMatrix(fila, 2) = UCase(rs("monto"))
            rs.MoveNext
        Loop
    End With
    
    If rs.State = adStateOpen Then rs.Close
    If conn.State = adStateOpen Then Call DesconectarBD
    ' Pintar filas alternadas
    Call PintarFilasAlternadasFlex(MSFlexGrid1)
    Exit Sub
    
ErrHandler:
    MsgBox "Error al cargar los impuesto: " & Err.Description, vbCritical, "Error"
    If rs.State = adStateOpen Then rs.Close
    If conn.State = adStateOpen Then Call DesconectarBD
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

Public Sub CargarImpuesto() 'new
    Dim rs As New ADODB.Recordset

    ' Conectar a la base de datos
    Call ConectarBD
    
    On Error GoTo ErrHandler
    
    cboImpuestos.Clear
    
    ' Agregar la opción "TODO" con id = 999
    cboImpuestos.AddItem "-- TODO --"
    cboImpuestos.ItemData(cboImpuestos.NewIndex) = 999
    
    rs.Open "select distinct t.id, t.tipo from Empresa_Impuesto ei join Timpuestos t on t.id = ei.idTipo where ei.idEmpresa = " & cboEstaciones.ItemData(cboEstaciones.ListIndex) & " ORDER BY id", conn, adOpenStatic, adLockReadOnly
    
    ' Cargar los meses desde la base de datos al ComboBox
    Do While Not rs.EOF
        ' Puedes guardar el ID en ItemData si querés usarlo después
        cboImpuestos.AddItem rs("id") & " - " & rs("tipo")
        cboImpuestos.ItemData(cboImpuestos.NewIndex) = rs("id")
        rs.MoveNext
    Loop
    
    ' Seleccionar "TODO" por defecto
    cboImpuestos.ListIndex = 0
    
    If rs.State = adStateOpen Then rs.Close
    If conn.State = adStateOpen Then Call DesconectarBD
    Exit Sub
    
ErrHandler:
    MsgBox "Error al cargar el listado de impuestos: " & Err.Description, vbCritical, "Error"
    If rs.State = adStateOpen Then rs.Close
    If conn.State = adStateOpen Then Call DesconectarBD
End Sub

Private Sub cboEstaciones_Click()
Call CargarImpuesto
End Sub

Private Sub Form_Load()
Call CargarMeses
'Call CargarImpuesto
Call CargarEstaciones

With MSFlexGrid1
    .Rows = 3
    .Cols = 3
    .FixedRows = 2
    .FixedCols = 0
    .TextMatrix(0, 0) = "Fecha"
    .TextMatrix(0, 1) = "Tipo"
    .TextMatrix(0, 2) = "Monto"
        .ColWidth(0) = 1300 ' Ancho del Código
        .ColWidth(1) = 3500
        .ColWidth(2) = 1000
End With
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
Private Sub txtAño_KeyPress(KeyAscii As Integer)
If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 Then
        KeyAscii = 0 ' Cancelar el caracter si no es un numero
    End If
End Sub
