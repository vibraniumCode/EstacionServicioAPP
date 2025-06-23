VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form FRMTimpuesto 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Tipos de impuesto"
   ClientHeight    =   4980
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   7470
   Icon            =   "frmTimpuesto.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4980
   ScaleWidth      =   7470
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      Height          =   4695
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7215
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
         Height          =   2895
         Left            =   120
         TabIndex        =   4
         Top             =   1680
         Width           =   6975
         _ExtentX        =   12303
         _ExtentY        =   5106
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
      Begin VB.CommandButton btnIngresar 
         Caption         =   "&Ingresar Impuesto"
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
         Picture         =   "frmTimpuesto.frx":08CA
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   960
         Width           =   1815
      End
      Begin VB.TextBox txtTimpuesto 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   2
         Top             =   480
         Width           =   6975
      End
      Begin VB.Line Line1 
         BorderWidth     =   2
         X1              =   1920
         X2              =   7080
         Y1              =   1320
         Y2              =   1320
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Impuesto"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   240
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   915
      End
   End
   Begin VB.Menu gestion 
      Caption         =   "&Gestion"
      Visible         =   0   'False
      Begin VB.Menu EliminarImp 
         Caption         =   "Eliminar Impuesto"
      End
   End
End
Attribute VB_Name = "frmTimpuesto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub btnIngresar_Click()
Dim rs As New ADODB.Recordset
    
    ' Conectar a la base de datos utilizando el módulo de conexión
    Call ConectarBD

    On Error GoTo ErrHandler
    
    If txtTimpuesto.Text = "" Then
        MsgBox "Ingrese la descripcion del impuesto", vbInformation, "ESAPP"
        Exit Sub
    End If

    rs.Open "exec sp_impuestos null,'" & txtTimpuesto.Text & "',null,null" & ",'IMP'", conn, adOpenStatic, adLockReadOnly
    
    MsgBox rs(0), vbInformation, "ESAPP"
    
    Call CargarlvImpuesto
    txtTimpuesto.Text = ""
    
    ' Cerrar el recordset
    If rs.State = adStateOpen Then rs.Close
    If conn.State = adStateOpen Then Call DesconectarBD
    Exit Sub

ErrHandler:
    MsgBox "Error al cargar datos: " & Err.Description, vbCritical, "Error"
    If rs.State = adStateOpen Then rs.Close
    If conn.State = adStateOpen Then Call DesconectarBD
End Sub

Private Sub EliminarImp_Click()
    Dim idImpuesto As Integer
    Dim filasAfectadas As Long
    On Error GoTo ErrHandler

    Call ConectarBD

    ' Ejecutamos el DELETE y contamos filas afectadas
    conn.Execute "DELETE FROM timpuestos WHERE id = " & idSeleccionado, filasAfectadas

    If filasAfectadas > 0 Then
        MsgBox "Impuesto eliminado correctamente.", vbInformation, "Éxito"

        Call CargarlvImpuesto
    Else
        MsgBox "No se pudo eliminar el impuesto. Verifique el ID.", vbExclamation, "Aviso"
    End If

    Call DesconectarBD
    Exit Sub
ErrHandler:
    MsgBox "Error al eliminar datos: " & Err.Description, vbCritical, "Error"
    Call DesconectarBD
End Sub


Private Sub Form_Load()
With MSFlexGrid1
    .Rows = 2
    .Cols = 2
    .FixedRows = 1
    .FixedCols = 0
    .TextMatrix(0, 0) = "Código"
    .TextMatrix(0, 1) = "Tipo de impuesto"
    .ColWidth(0) = 800
    .ColWidth(1) = 4500
End With

MSFlexGrid1.GridLines = flexGridSolid

' Cargar datos en el ListView
Call CargarlvImpuesto
Call PintarFilasAlternadasFlex(MSFlexGrid1)

'Call PintarFilasAlternadasFlex
End Sub

Public Sub CargarlvImpuesto()
     Dim rs As New ADODB.Recordset
    Dim fila As Integer

    ' Conectar a la base de datos
    Call ConectarBD

    On Error GoTo ErrHandler

    ' Consulta a la base de datos
    rs.Open "SELECT id, tipo FROM timpuestos", conn, adOpenStatic, adLockReadOnly

    ' Configurar columnas del MSFlexGrid
    With MSFlexGrid1
        .Clear
        .Rows = 1 ' Solo cabecera
        .Cols = 2
        .TextMatrix(0, 0) = "Código"
        .TextMatrix(0, 1) = "Tipo de impuesto"

        ' Agregar datos fila por fila
        Do While Not rs.EOF
            .Rows = .Rows + 1
            fila = .Rows - 1
            .TextMatrix(fila, 0) = rs("id")
            .TextMatrix(fila, 1) = UCase(rs("tipo"))
            rs.MoveNext
        Loop
    End With
    
        With MSFlexGrid1
        .ColWidth(0) = 1000 ' Ancho del Código
        .ColWidth(1) = .Width - .ColWidth(0) - 100 ' El resto para el Tipo de impuesto
    End With

    
    rs.Close
    Call DesconectarBD

    ' Pintar filas alternadas
    Call PintarFilasAlternadasFlex(MSFlexGrid1)
    Exit Sub

ErrHandler:
    MsgBox "Error al cargar datos: " & Err.Description, vbCritical, "Error"
    If rs.State = adStateOpen Then rs.Close
    Call DesconectarBD
End Sub

Private Sub Form_Unload(Cancel As Integer)
Call FRMImpuestos.CargarImpuesto
End Sub

Private Sub MSFlexGrid1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

If Button = vbRightButton Then
    Dim fila As Long
    fila = MSFlexGrid1.FixedRows + (Y \ MSFlexGrid1.RowHeight(0)) - 1
    
    ' Validar que la fila clickeada sea válida (dentro del rango del grid)
    If fila >= MSFlexGrid1.FixedRows And fila < MSFlexGrid1.Rows Then
        MSFlexGrid1.Row = fila
        MSFlexGrid1.Col = 0

        idSeleccionado = MSFlexGrid1.Text

        PopupMenu gestion
    End If
End If
End Sub

