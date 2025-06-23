VERSION 5.00
Begin VB.Form FRMImpuestos 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Impuesto"
   ClientHeight    =   2820
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   7935
   Icon            =   "FRMImpuestos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2820
   ScaleWidth      =   7935
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2535
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7695
      Begin VB.ComboBox cboEstaciones 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   120
         TabIndex        =   9
         Text            =   "Combo1"
         Top             =   480
         Width           =   7455
      End
      Begin VB.ComboBox cboImpuestos 
         BeginProperty Font 
            Name            =   "Tahoma"
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
         TabIndex        =   8
         Top             =   1080
         Width           =   4935
      End
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
         Left            =   6480
         Picture         =   "FRMImpuestos.frx":058A
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   1800
         Width           =   1095
      End
      Begin VB.CommandButton btnGrabar 
         Caption         =   "&Grabar"
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
         Left            =   5160
         Picture         =   "FRMImpuestos.frx":0B14
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   1800
         Width           =   1095
      End
      Begin VB.TextBox txtMonto 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5160
         TabIndex        =   4
         Text            =   "$00.00"
         Top             =   1080
         Width           =   2415
      End
      Begin VB.TextBox txtFecOperacion 
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
         TabIndex        =   1
         Top             =   1920
         Width           =   2415
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Empresa"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   855
      End
      Begin VB.Line Line1 
         BorderWidth     =   2
         X1              =   2640
         X2              =   6480
         Y1              =   2160
         Y2              =   2160
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Monto"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   5160
         TabIndex        =   5
         Top             =   840
         Width           =   585
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Impuesto"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   3
         Top             =   840
         Width           =   930
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Operación"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   2
         Top             =   1680
         Width           =   1635
      End
   End
   Begin VB.Menu Operaciones 
      Caption         =   "&Operaciones"
      Visible         =   0   'False
      Begin VB.Menu timpuesto 
         Caption         =   "&Impuestos"
      End
   End
End
Attribute VB_Name = "FRMImpuestos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Impuestos As New ClaseImpuesto

Private Sub btnGrabar_Click()
Dim rs As New ADODB.Recordset
Dim idImpuestoSeleccionado As Integer
Dim idEmpresaSeleccionado As Integer

    If DatosValidador Then Exit Sub
    
    ' Conectar a la base de datos utilizando el módulo de conexión
    Call ConectarBD

    On Error GoTo ErrHandler
    
    If cboEstaciones.ListIndex <> -1 Then
        idEmpresaSeleccionado = cboEstaciones.ItemData(cboEstaciones.ListIndex)
    Else
        MsgBox "No hay empresa seleccionado", vbInformation, "ESAPP"
        Exit Sub
    End If
    
    If cboImpuestos.ListIndex <> -1 Then
        idImpuestoSeleccionado = cboImpuestos.ItemData(cboImpuestos.ListIndex)
    Else
        MsgBox "No hay impuestos seleccionado", vbInformation, "ESAPP"
        Exit Sub
    End If

    rs.Open "exec sp_impuestos " & idImpuestoSeleccionado & ",null," & Replace(Impuestos.Monto, ",", ".") & ",'" & Format(txtFecOperacion.Text, "yyyymmdd") & "'," & idEmpresaSeleccionado & ",NULL,NULL,'MOD'", conn, adOpenStatic, adLockReadOnly
    
    MsgBox rs(1), vbInformation, "ESAPP"
    If rs(0) = 1 Then
        If MsgBox("¿Desea modificar el monto de este mes?", vbYesNo + vbQuestion, "Confirmación") = vbYes Then
            rs.Close
            rs.Open "exec sp_impuestos " & idImpuestoSeleccionado & ",null," & Replace(Impuestos.Monto, ",", ".") & ",'" & Format(txtFecOperacion.Text, "yyyymmdd") & "'," & idEmpresaSeleccionado & ",NULL,NULL,'UPD'", conn, adOpenStatic, adLockReadOnly
            
            MsgBox rs(0), vbInformation, "ESAPP"
        End If
    End If
    
    ' Cerrar el recordset
    If rs.State = adStateOpen Then rs.Close
    If conn.State = adStateOpen Then Call DesconectarBD
'    Call CargarUltImpu
    Call LimpiarCampos
    Exit Sub

ErrHandler:
    MsgBox "Error al cargar datos: " & Err.Description, vbCritical, "Error"
    If rs.State = adStateOpen Then rs.Close
    If conn.State = adStateOpen Then Call DesconectarBD
End Sub

Private Sub btnSalir_Click()
Unload Me
End Sub

Private Sub cboEstaciones_Click()
Call CargarImpuesto
End Sub

Private Sub Form_Load()
txtFecOperacion.Text = Date
Call CargarEstaciones
Call CargarImpuesto
'Call CargarUltImpu
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

Public Sub CargarImpuesto()
    Dim rs As New ADODB.Recordset

    ' Conectar a la base de datos
    Call ConectarBD
    
    On Error GoTo ErrHandler
    
    cboImpuestos.Clear
    Dim SQL As String
    
    rs.Open "select distinct t.id, t.tipo from Empresa_Impuesto ei join Timpuestos t on t.id = ei.idTipo where ei.idEmpresa = " & cboEstaciones.ItemData(cboEstaciones.ListIndex) & " ORDER BY id", conn, adOpenStatic, adLockReadOnly
    
    ' Cargar los meses desde la base de datos al ComboBox
    Do While Not rs.EOF
        ' Puedes guardar el ID en ItemData si querés usarlo después
        cboImpuestos.AddItem rs("id") & " - " & rs("tipo")
        cboImpuestos.ItemData(cboImpuestos.NewIndex) = rs("id")
        rs.MoveNext
    Loop
    
    If cboImpuestos.ListCount > 0 Then
        cboImpuestos.ListIndex = 0
    End If
    
    If rs.State = adStateOpen Then rs.Close
    If conn.State = adStateOpen Then Call DesconectarBD
    Exit Sub
    
ErrHandler:
    MsgBox "Error al cargar el listado de impuestos: " & Err.Description, vbCritical, "Error"
    If rs.State = adStateOpen Then rs.Close
    If conn.State = adStateOpen Then Call DesconectarBD
End Sub

Public Function DatosValidador() As Boolean
If txtMonto.Text = "$0.00" Then
    MsgBox "Ingrese el monto del impuesto", vbInformation, "ESAPP"
    DatosValidador = True
Else
    DatosValidador = False
End If
End Function

Private Sub timpuesto_Click()
frmTimpuesto.Show vbModal
End Sub

Private Sub txtMonto_KeyPress(KeyAscii As Integer)
If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 And KeyAscii <> 44 Then
        KeyAscii = 0 ' Cancelar el caracter si no es número, backspace o coma
    End If
    
    ' Evitar múltiples comas
    If KeyAscii = 44 And InStr(txtMonto.Text, ",") > 0 Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtMonto_LostFocus()
With Impuestos
    ' Convertir el texto ingresado a formato numérico usando Replace
    .Monto = Replace(txtMonto.Text, ",", ".")
    txtMonto.Text = FormatoPrecio(Val(.Monto))
End With
End Sub

Public Function FormatoPrecioCorto(ByVal Valor As Double) As String
    FormatoPrecioCorto = Format$(Valor, "#.##0,00")  ' Sin símbolo $
End Function

Private Sub LimpiarCampos()
    txtMonto.Text = FormatoPrecio("$00.00")
End Sub
