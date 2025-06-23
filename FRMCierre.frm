VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FRMCierre 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Reporte"
   ClientHeight    =   2430
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   13710
   Icon            =   "FRMCierre.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2430
   ScaleWidth      =   13710
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   2175
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   13455
      Begin VB.CommandButton btnPlanilla 
         Caption         =   "&Planilla"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   120
         Picture         =   "FRMCierre.frx":058A
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   1200
         Width           =   1095
      End
      Begin VB.Frame Frame3 
         Caption         =   "Empresa"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   855
         Left            =   5160
         TabIndex        =   4
         Top             =   240
         Width           =   8175
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
            TabIndex        =   5
            Text            =   "Combo1"
            Top             =   360
            Width           =   7935
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Fecha Desde/Hasta"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   855
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   4935
         Begin MSComCtl2.DTPicker DTPicker2 
            Height          =   375
            Left            =   2520
            TabIndex        =   3
            Top             =   360
            Width           =   2295
            _ExtentX        =   4048
            _ExtentY        =   661
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
            Format          =   143785985
            CurrentDate     =   45828
         End
         Begin MSComCtl2.DTPicker DTPicker1 
            Height          =   375
            Left            =   120
            TabIndex        =   2
            Top             =   360
            Width           =   2295
            _ExtentX        =   4048
            _ExtentY        =   661
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
            Format          =   143785985
            CurrentDate     =   45828
         End
      End
   End
End
Attribute VB_Name = "FRMCierre"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub btnPlanilla_Click()
    Dim rs As New ADODB.Recordset
    Dim sql As String

    ' Conectar a la base de datos
    Call ConectarBD

    ' Armás el SQL completo
    sql = "SELECT F.fecEmision AS FECHA, " & _
          "SUBSTRING(CAST(F.horaEmision AS VARCHAR), 1, 8) AS HORA, " & _
          "CONCAT('N° ', RIGHT('0000' + CAST(F.codFactura AS VARCHAR), 4), '-', RIGHT('00000000' + CAST(F.nroFactura AS VARCHAR), 8)) AS TICKET, " & _
          "TCC.litros AS LITROS, " & _
          "CONCAT('$ ', CAST(FI.imp_neto AS VARCHAR)) AS NETO, " & _
          "CONCAT('$ ', CAST(FI.imp_iva AS VARCHAR)) AS IVA, " & _
          "CONCAT('$ ', CAST(FI.impuesto_total AS VARCHAR)) AS TRIBUTOS, " & _
          "CONCAT('$ ', CAST(FI.imp_total AS VARCHAR)) AS TOTAL " & _
          "FROM Facturacion F " & _
          "JOIN tcarga_combustible TCC ON TCC.codFactura = F.codFactura AND TCC.nroFactura = F.nroFactura " & _
          "JOIN Facturacion_importe FI ON FI.codFactura = F.codFactura AND FI.nroFactura = F.nroFactura " & _
          "WHERE F.fecEmision BETWEEN ' " & Format(DTPicker1.value, "yyyymmdd") & "' AND '" & Format(DTPicker2.value, "yyyymmdd") & "' " & _
          "AND FI.empresa =" & cboEstaciones.ItemData(cboEstaciones.ListIndex) & _
          " ORDER BY F.fecEmision"

    ' Usás el SQL correcto
    rs.Open sql, conn, adOpenStatic, adLockReadOnly

    ' Mostrar el DataReport si hay datos
    If Not rs.EOF Then
        Set DataReport1.DataSource = rs
        DataReport1.Show vbModal
    Else
        MsgBox "No hay datos para mostrar.", vbInformation
    End If
End Sub

' Necesitas instalar PDFCreator en el sistema
Private Sub GenerarPDF()
    On Error GoTo ErrorHandler
    
    ' Verificar si existe impresora PDF
    Dim i As Integer
    Dim pdfPrinter As String
    pdfPrinter = ""
    
    For i = 0 To Printers.Count - 1
        If InStr(UCase(Printers(i).DeviceName), "PDF") > 0 Then
            pdfPrinter = Printers(i).DeviceName
            Exit For
        End If
    Next i
    
    If pdfPrinter = "" Then
        MsgBox "No se encontró impresora PDF en el sistema"
        Exit Sub
    End If
    
    ' Guardar impresora actual
    Dim currentPrinter As String
    currentPrinter = Printer.DeviceName
    
    ' Cambiar a impresora PDF
    Set Printer = Printers(i)
    
    ' Imprimir
    txtTicket2.SelPrint True
    
    ' Restaurar impresora original
    For i = 0 To Printers.Count - 1
        If Printers(i).DeviceName = currentPrinter Then
            Set Printer = Printers(i)
            Exit For
        End If
    Next i
    
    MsgBox "PDF generado exitosamente"
    Exit Sub
    
ErrorHandler:
    MsgBox "Error al generar PDF: " & Err.Description
End Sub

Private Sub Form_Load()
Call CargarEstaciones
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


