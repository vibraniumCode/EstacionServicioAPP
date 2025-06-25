VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmComprobantes 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Comprobante"
   ClientHeight    =   9375
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6615
   Icon            =   "frmComprobantes.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9375
   ScaleWidth      =   6615
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame frmComprobante 
      Caption         =   "Diseño de comprobante"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   9135
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6375
      Begin VB.CommandButton btnD 
         Caption         =   "&Desabilitar Ticket"
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
         Left            =   5040
         Picture         =   "frmComprobantes.frx":08CA
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   8160
         Width           =   1215
      End
      Begin VB.CommandButton btnH 
         Caption         =   "&Habilitar ticket"
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
         Left            =   5040
         Picture         =   "frmComprobantes.frx":0E54
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   7200
         Width           =   1215
      End
      Begin RichTextLib.RichTextBox txtTicket2 
         Height          =   8655
         Left            =   120
         TabIndex        =   3
         Top             =   360
         Width           =   4575
         _ExtentX        =   8070
         _ExtentY        =   15266
         _Version        =   393217
         Enabled         =   0   'False
         ScrollBars      =   2
         TextRTF         =   $"frmComprobantes.frx":13DE
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "SimSun"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.CommandButton btnPrint 
         Caption         =   "&Imprimir Ticket"
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
         Left            =   5040
         Picture         =   "frmComprobantes.frx":145A
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   6240
         Width           =   1215
      End
      Begin VB.TextBox txtTicket 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "SimSun"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   7635
         HelpContextID   =   1
         HideSelection   =   0   'False
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   1
         Top             =   360
         Width           =   4215
      End
   End
End
Attribute VB_Name = "frmComprobantes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btnD_Click()
txtTicket2.Enabled = False
End Sub

Private Sub btnH_Click()
txtTicket2.Enabled = True
End Sub

Private Sub btnPrint_Click()
If MsgBox("¿Desea imprimir el ticket?", vbYesNo + vbQuestion, "Confirmación") = vbYes Then
'    MsgBox "Verificando impresora...", vbInformation, "ESAPP"
    Dim x As Printer
    For Each x In Printers
    If x.DeviceName = "POS-80-Series" Then
        MsgBox "Imprimiendo comprobante"
        Printer.FontName = "SinSum"
        Printer.FontSize = 9
        Printer.FontBold = True
        Printer.Print txtTicket.Text
        Set Printer = x
    Exit For
    End If
    Next
End If
End Sub

Private Sub Form_Load()
Dim rs As New ADODB.Recordset
Dim sql As String
        
Call ConectarBD
On Error GoTo ErrHandler

    sql = "sp_comprobante " & codFactura & "," & nroFactura & "," & idEstacionEmpresa
    
    Select Case idEstacionEmpresa
        Case 1
            Call Ticket_1(sql, rs)
        Case 2
            Call Ticket_2(sql, rs)
        Case 3, 4
            Call Ticket_3(sql, rs)
        Case 5
            txtTicket2.Width = 4815
            Call Ticket_5(sql, rs)
    End Select
    
    If rs.State = adStateOpen Then rs.Close
    If conn.State = adStateOpen Then Call DesconectarBD
    Exit Sub

ErrHandler:
    MsgBox "Error al cargar datos: " & Err.Description, vbCritical, "Error"
    If rs.State = adStateOpen Then rs.Close
    If conn.State = adStateOpen Then Call DesconectarBD
End Sub
Private Sub Ticket_1(sql As String, rs As Recordset)
rs.Open sql, conn, adOpenStatic, adLockReadOnly

    txtTicket2.SelText = rs(2) + vbCrLf
    txtTicket2.SelText = txtTicket2.SelText + "C.U.I.T. Nro.:" + rs(3) + vbCrLf
    txtTicket2.SelText = txtTicket2.SelText + "Ing. Brutos: " + rs(4) + vbCrLf
    txtTicket2.SelText = txtTicket2.SelText + "Domicilio: " + rs(5) + vbCrLf
    txtTicket2.SelText = txtTicket2.SelText + rs(6) + vbCrLf
    txtTicket2.SelText = txtTicket2.SelText + "TEL.: " + rs(7) + vbCrLf
    txtTicket2.SelText = txtTicket2.SelText + "Inicio de Actividades: " + rs(8) + vbCrLf
    txtTicket2.SelText = txtTicket2.SelText + "IVA RESPONSABLE INSCRIPTO" + vbCrLf
    txtTicket2.SelBold = True
    txtTicket2.SelText = txtTicket2.SelText + "            PETRORAFAELA  S R L" + vbCrLf
    txtTicket2.SelBold = False
    txtTicket2.SelText = txtTicket2.SelText + "-------------------------------------------------" + vbCrLf
    txtTicket2.SelBold = True
    txtTicket2.SelText = txtTicket2.SelText + "TIQUET FACTURA A (Cód.081) N° "
    txtTicket2.SelBold = False
    txtTicket2.SelText = txtTicket2.SelText + CStr(Format(rs(0), "0000")) & "-" + CStr(Format(rs(1), "00000000")) + vbCrLf
    txtTicket2.SelText = txtTicket2.SelText + "                                 Fecha " + rs(9) + vbCrLf
    txtTicket2.SelText = txtTicket2.SelText + "                                    Hora " + rs(10) + vbCrLf
    txtTicket2.SelText = txtTicket2.SelText + "-------------------------------------------------" + vbCrLf
    txtTicket2.SelText = txtTicket2.SelText + UCase(rs(11)) + vbCrLf
    txtTicket2.SelText = txtTicket2.SelText + "C.U.I.T. Nro.: " + rs(12) + vbCrLf
    txtTicket2.SelText = txtTicket2.SelText + "IVA RESPONSABLE INSCRIPTO" + vbCrLf
    txtTicket2.SelText = txtTicket2.SelText + UCase(rs(13)) + vbCrLf
    txtTicket2.SelText = txtTicket2.SelText + "CONTADO" + vbCrLf
    txtTicket2.SelText = txtTicket2.SelText + "COMPROBANTES ASOCIADO:" + vbCrLf
    txtTicket2.SelText = txtTicket2.SelText + "Cód. 001                           00001-00000001" + vbCrLf
    txtTicket2.SelBold = True
    txtTicket2.SelText = txtTicket2.SelText + "            PETRORAFAELA  S R L" + vbCrLf
    txtTicket2.SelBold = False
    txtTicket2.SelText = txtTicket2.SelText + "-------------------------------------------------"
    txtTicket2.SelText = txtTicket2.SelText + CStr(rs(16)) + " x " + CStr(rs(17)) + vbCrLf
    txtTicket2.SelText = txtTicket2.SelText + "(" + CStr(rs(14)) + ")" + UCase(rs(15)) + vbCrLf
    txtTicket2.SelText = txtTicket2.SelText + "ITC Unit.x Lt: " + CStr(rs(20)) + vbCrLf
    txtTicket2.SelText = txtTicket2.SelText + CStr(rs(16)) + " x " + CStr(rs(18)) + "   (21)[83,88]         " + CStr(rs(19)) + vbCrLf
    txtTicket2.SelText = txtTicket2.SelText + "Ajuste por redondeo       (21)               0,00" + vbCrLf
    txtTicket2.SelText = txtTicket2.SelText + vbCrLf
    txtTicket2.SelText = txtTicket2.SelText + "SUBTOT. IMP. NETO GRAVADO               " + CStr(rs(19)) + vbCrLf
    txtTicket2.SelText = txtTicket2.SelText + vbCrLf
    txtTicket2.SelText = txtTicket2.SelText + "ALICUOTA 21,00%                        " + CStr(rs(23)) + vbCrLf
    txtTicket2.SelText = txtTicket2.SelText + vbCrLf
    txtTicket2.SelText = txtTicket2.SelText + "10-Impuesto interno a nivel item         " + CStr(rs(22)) + vbCrLf
    txtTicket2.SelText = txtTicket2.SelText + "07 - IDC                                   " + CStr(rs(21)) + vbCrLf
    txtTicket2.SelText = txtTicket2.SelText + "IMPORTE TOTAL OTROS TRIBUTOS             " + CStr(rs(24)) + vbCrLf
    txtTicket2.SelText = txtTicket2.SelText + vbCrLf
    txtTicket2.SelBold = True
    txtTicket2.SelText = txtTicket2.SelText + "TOTAL                            " + CStr(rs(25)) + vbCrLf
    txtTicket2.SelBold = False
    txtTicket2.SelText = txtTicket2.SelText + vbCrLf
    txtTicket2.SelText = txtTicket2.SelText + "RECIBI/MOS" + vbCrLf
    txtTicket2.SelText = txtTicket2.SelText + "Cant. Cuota: 1" + vbCrLf
    txtTicket2.SelText = txtTicket2.SelText + "Efectivo" + vbCrLf
    txtTicket2.SelItalic = True
    txtTicket2.SelText = txtTicket2.SelText + "   CF"
    txtTicket2.SelItalic = False
    txtTicket2.SelText = txtTicket2.SelText + vbCrLf
End Sub

'''''''''''''''''''''

Private Sub Ticket_2(sql As String, rs As Recordset)
rs.Open sql, conn, adOpenStatic, adLockReadOnly

    txtTicket2.SelText = rs(2) + vbCrLf
    txtTicket2.SelText = txtTicket2.SelText + "C.U.I.T. Nro.:" + CStr(rs(3)) + vbCrLf
    txtTicket2.SelText = txtTicket2.SelText + "Ing. Brutos: " + CStr(rs(4)) + vbCrLf
    txtTicket2.SelText = txtTicket2.SelText + CStr(rs(5)) + vbCrLf
    txtTicket2.SelText = txtTicket2.SelText + CStr(rs(6)) + vbCrLf
    txtTicket2.SelText = txtTicket2.SelText + "Inicio de Actividades: " + CStr(rs(8)) + vbCrLf
    txtTicket2.SelText = txtTicket2.SelText + "IVA Responsable Inscripto" + vbCrLf
    txtTicket2.SelText = txtTicket2.SelText + "-------------------------------------------------" + vbCrLf
    txtTicket2.SelBold = True
    txtTicket2.SelText = txtTicket2.SelText + "          Cód. 081 - TIQUET FACTURA A" + vbCrLf
    txtTicket2.SelBold = False
    txtTicket2.SelText = txtTicket2.SelText + "                                N° " + CStr(Format(rs(0), "0000")) & "-" + CStr(Format(rs(1), "00000000")) + vbCrLf
    txtTicket2.SelText = txtTicket2.SelText + "Fecha " + CStr(rs(9)) + "                    Hora " + CStr(rs(10)) + vbCrLf
    txtTicket2.SelText = txtTicket2.SelText + "-------------------------------------------------" + vbCrLf
    txtTicket2.SelText = txtTicket2.SelText + UCase(rs(11)) + vbCrLf
    txtTicket2.SelText = txtTicket2.SelText + "C.U.I.T. Nro.: " + rs(12) + vbCrLf
    txtTicket2.SelText = txtTicket2.SelText + "IVA RESPONSABLE INSCRIPTO" + vbCrLf
    txtTicket2.SelText = txtTicket2.SelText + UCase(rs(13)) + vbCrLf
    txtTicket2.SelText = txtTicket2.SelText + "CONTADO" + vbCrLf
    txtTicket2.SelText = txtTicket2.SelText + "SD" + vbCrLf
    txtTicket2.SelText = txtTicket2.SelText + "-------------------------------------------------"
    txtTicket2.SelBold = True
    txtTicket2.SelText = txtTicket2.SelText + "Cantidad/Precio unit            código" + vbCrLf
    txtTicket2.SelBold = True
    txtTicket2.SelText = txtTicket2.SelText + "Descripción    (%IVA) [%B.I.]  Precio Neto" + vbCrLf
    txtTicket2.SelBold = False
    txtTicket2.SelText = txtTicket2.SelText + CStr(rs(15)) + vbCrLf
    txtTicket2.SelText = txtTicket2.SelText + CStr(rs(16)) + " LS A " + CStr(rs(17)) + "(21)[90,95]           " + CStr(rs(19)) + vbCrLf
    txtTicket2.SelText = txtTicket2.SelText + vbCrLf
    txtTicket2.SelBold = True
    txtTicket2.SelText = txtTicket2.SelText + "SUBTOTAL IMPORTE NETO NO GRAVADO      0,00" + vbCrLf
    txtTicket2.SelBold = True
    txtTicket2.SelText = txtTicket2.SelText + "SUBTOTAL IMPORTE IMPORTE EXENTO       0,00" + vbCrLf
    txtTicket2.SelBold = True
    txtTicket2.SelText = txtTicket2.SelText + "SUBTOTAL IMPORTE NETO FRAVADO     " + CStr(rs(19)) + vbCrLf
    txtTicket2.SelBold = False
    txtTicket2.SelText = txtTicket2.SelText + vbCrLf
    txtTicket2.SelText = txtTicket2.SelText + "Concepto" + vbCrLf
    txtTicket2.SelText = txtTicket2.SelText + "Alicuota 21,00%                        " + CStr(rs(23)) + vbCrLf
    txtTicket2.SelText = txtTicket2.SelText + "IMPORTE TOTAL IVA                      " + CStr(rs(23)) + vbCrLf
    txtTicket2.SelText = txtTicket2.SelText + vbCrLf
    txtTicket2.SelText = txtTicket2.SelText + "10-Impuesto interno a nivel item         " + CStr(rs(22)) + vbCrLf
    txtTicket2.SelBold = True
    txtTicket2.SelText = txtTicket2.SelText + "IMPORTE TOTAL OTROS TRIBUTOS      " + CStr(rs(22)) + vbCrLf
    txtTicket2.SelText = txtTicket2.SelText + vbCrLf
    txtTicket2.SelBold = True
    txtTicket2.SelText = txtTicket2.SelText + "TOTAL                            " + CStr(rs(25)) + vbCrLf
    txtTicket2.SelBold = False
    txtTicket2.SelText = txtTicket2.SelText + "RECIBI(MOS)" + vbCrLf
    txtTicket2.SelText = txtTicket2.SelText + "Contado-------------------------------" + vbCrLf
    txtTicket2.SelText = txtTicket2.SelText + "Pesos Argentinos----------------------  " + CStr(rs(25)) + vbCrLf
    txtTicket2.SelText = txtTicket2.SelText + "CAMBIO                                       0,00" + vbCrLf
    txtTicket2.SelText = txtTicket2.SelText + "RECIBI" + vbCrLf
    txtTicket2.SelText = txtTicket2.SelText + "ITC Unit x Lt: $" + CStr(rs(20)) + vbCrLf
    txtTicket2.SelItalic = True
    txtTicket2.SelText = txtTicket2.SelText + "   CF" + vbCrLf
    txtTicket2.SelItalic = False
    txtTicket2.SelText = txtTicket2.SelText + "           HSHSAB0000045122              V: 01.00"
    txtTicket2.SelText = txtTicket2.SelText + vbCrLf
End Sub

'''''''''''''''''''''

Private Sub Ticket_3(sql As String, rs As Recordset)
rs.Open sql, conn, adOpenStatic, adLockReadOnly
    txtTicket2.SelBold = True
    txtTicket2.SelText = CStr(rs(2)) + vbCrLf
    txtTicket2.SelBold = False
    txtTicket2.SelText = txtTicket2.SelText + "C.U.I.T. Nro.:" + CStr(rs(3)) + vbCrLf
    txtTicket2.SelText = txtTicket2.SelText + "Ing. Brutos: " + CStr(rs(4)) + vbCrLf
    txtTicket2.SelText = txtTicket2.SelText + "Domicilio: " + CStr(rs(5)) + vbCrLf
    txtTicket2.SelText = txtTicket2.SelText + CStr(rs(6)) + vbCrLf
    txtTicket2.SelText = txtTicket2.SelText + "Inicio de Actividades: " + CStr(rs(8)) + vbCrLf
    txtTicket2.SelText = txtTicket2.SelText + "IVA RESPONSABLE INSCRIPTO" + vbCrLf
    txtTicket2.SelText = txtTicket2.SelText + "-------------------------------------------------" + vbCrLf
    txtTicket2.SelBold = True
    txtTicket2.SelText = txtTicket2.SelText + "TIQUET FACTURA A (Cód.081) N° "
    txtTicket2.SelBold = False
    txtTicket2.SelText = txtTicket2.SelText + CStr(Format(rs(0), "0000")) & "-" + CStr(Format(rs(1), "00000000")) + vbCrLf
    txtTicket2.SelText = txtTicket2.SelText + "                                 Fecha " + rs(9) + vbCrLf
    txtTicket2.SelText = txtTicket2.SelText + "                                    Hora " + rs(10) + vbCrLf
    txtTicket2.SelText = txtTicket2.SelText + "-------------------------------------------------" + vbCrLf
    txtTicket2.SelText = txtTicket2.SelText + UCase(rs(11)) + vbCrLf
    txtTicket2.SelText = txtTicket2.SelText + "C.U.I.T. Nro.: " + rs(12) + vbCrLf
    txtTicket2.SelText = txtTicket2.SelText + "IVA RESPONSABLE INSCRIPTO" + vbCrLf
    txtTicket2.SelText = txtTicket2.SelText + UCase(rs(13)) + vbCrLf
    txtTicket2.SelText = txtTicket2.SelText + "Cond. Vta: CONTADO" + vbCrLf
    txtTicket2.SelText = txtTicket2.SelText + "COMPROBANTES ASOCIADOS:" + vbCrLf
    txtTicket2.SelText = txtTicket2.SelText + "Cód. 001                           00001-00000001" + vbCrLf
    txtTicket2.SelText = txtTicket2.SelText + "-------------------------------------------------"
    
    txtTicket2.SelText = txtTicket2.SelText + CStr(rs(16)) + " (00) x " + CStr(rs(17)) + vbCrLf
    txtTicket2.SelText = txtTicket2.SelText + UCase(rs(15)) + vbCrLf
    txtTicket2.SelText = txtTicket2.SelText + "Valor ITC $" + CStr(rs(20)) + "  (21)[90,95]          " + CStr(rs(19)) + vbCrLf
    txtTicket2.SelText = txtTicket2.SelText + vbCrLf
    txtTicket2.SelText = txtTicket2.SelText + "SUBTOT. IMP. NETO GRAVADO               " + CStr(rs(19)) + vbCrLf
    txtTicket2.SelText = txtTicket2.SelText + vbCrLf
    txtTicket2.SelText = txtTicket2.SelText + "ALICUOTA 21,00%                        " + CStr(rs(23)) + vbCrLf
    txtTicket2.SelText = txtTicket2.SelText + vbCrLf
    txtTicket2.SelText = txtTicket2.SelText + "10-Impuesto interno a nivel item         " + CStr(rs(22)) + vbCrLf
    txtTicket2.SelText = txtTicket2.SelText + "IMPORTE TOTAL OTROS TRIBUTOS             " + CStr(rs(22)) + vbCrLf
    txtTicket2.SelText = txtTicket2.SelText + vbCrLf
    txtTicket2.SelBold = True
    txtTicket2.SelText = txtTicket2.SelText + "TOTAL                            " + CStr(rs(25)) + vbCrLf
    txtTicket2.SelBold = False
    txtTicket2.SelText = txtTicket2.SelText + vbCrLf
    txtTicket2.SelText = txtTicket2.SelText + "RECIBI/MOS" + vbCrLf
    txtTicket2.SelText = txtTicket2.SelText + "Efectivo                                " + CStr(rs(25)) + vbCrLf
    txtTicket2.SelText = txtTicket2.SelText + "Suma de sus pagos                       " + CStr(rs(25)) + vbCrLf
    txtTicket2.SelText = txtTicket2.SelText + "Su vuelto                                    0,00" + vbCrLf
    txtTicket2.SelText = txtTicket2.SelText + vbCrLf
    txtTicket2.SelItalic = True
    txtTicket2.SelText = txtTicket2.SelText + "CF" + vbCrLf
    txtTicket2.SelItalic = False
    txtTicket2.SelText = txtTicket2.SelText + "       REGISTRO:   EPEPAA0000032653" + vbCrLf
    txtTicket2.SelText = txtTicket2.SelText + "       V: 1.02 Jano"
    txtTicket2.SelText = txtTicket2.SelText + vbCrLf
End Sub


'''''''''''''''''''''

Private Sub Ticket_5(sql As String, rs As Recordset)
rs.Open sql, conn, adOpenStatic, adLockReadOnly
    txtTicket2.SelBold = True
    txtTicket2.SelText = CStr(rs(2)) + vbCrLf
    txtTicket2.SelBold = False
    txtTicket2.SelText = txtTicket2.SelText + "CUIT:" + CStr(rs(3)) + vbCrLf
    txtTicket2.SelText = txtTicket2.SelText + "Domicilio: " + CStr(rs(5)) + vbCrLf
    txtTicket2.SelText = txtTicket2.SelText + "IIBB: " + CStr(rs(4)) + vbCrLf
    txtTicket2.SelText = txtTicket2.SelText + CStr(rs(6)) + vbCrLf
    txtTicket2.SelText = txtTicket2.SelText + "INICIO ACTIVIDADES: " + CStr(rs(8)) + vbCrLf
    txtTicket2.SelText = txtTicket2.SelText + "IVA Responsable Inscripto" + vbCrLf
    txtTicket2.SelText = txtTicket2.SelText + "-------------------------------------------------" + vbCrLf
    txtTicket2.SelBold = True
    txtTicket2.SelFontSize = 12
    txtTicket2.SelText = txtTicket2.SelText + "  FACTURA A ORIGINAL (COD. 001)          " + vbCrLf
    txtTicket2.SelFontSize = 8
    txtTicket2.SelBold = False
    txtTicket2.SelText = txtTicket2.SelText + "-------------------------------------------------" + vbCrLf
    txtTicket2.SelText = txtTicket2.SelText + UCase(rs(11)) + vbCrLf
    txtTicket2.SelText = txtTicket2.SelText + "CUIT:" + rs(12) + " IVA RESPONSABLE INSCRIPTO" + vbCrLf
    txtTicket2.SelText = txtTicket2.SelText + UCase(rs(13)) + vbCrLf
    txtTicket2.SelText = txtTicket2.SelText + "-------------------------------------------------" + vbCrLf
    txtTicket2.SelText = txtTicket2.SelText + "Fecha " + CStr(rs(9)) + "                    Hora " + CStr(rs(10)) + vbCrLf
    txtTicket2.SelText = txtTicket2.SelText + "FC: 9352  N°: 2      TR: 1092544  SUC.NRO: 02563" + vbCrLf
    txtTicket2.SelText = txtTicket2.SelText + "OPERADOR: Jaime" + vbCrLf
    txtTicket2.SelText = txtTicket2.SelText + "NRO.COMP: " + CStr(Format(rs(0), "0000")) & "-" + CStr(Format(rs(1), "00000000")) + " TERMINAL: 07" + vbCrLf
    txtTicket2.SelText = txtTicket2.SelText + "-------------------------------------------------" + vbCrLf
    txtTicket2.SelText = txtTicket2.SelText + "Cant. Precio Unit." + vbCrLf
    txtTicket2.SelText = txtTicket2.SelText + "Descripción (s/IVA)                       IMPORTE" + vbCrLf
    txtTicket2.SelText = txtTicket2.SelText + "-------------------------------------------------" + vbCrLf
    txtTicket2.SelBold = True
    txtTicket2.SelText = txtTicket2.SelText + UCase(rs(15)) + vbCrLf
    txtTicket2.SelBold = True
    txtTicket2.SelText = txtTicket2.SelText + CStr(rs(16)) + " x " + CStr(rs(17)) + "                 " + CStr(rs(19)) + vbCrLf
    txtTicket2.SelBold = False
    txtTicket2.SelText = txtTicket2.SelText + "ITC: " + CStr(rs(20)) + vbCrLf
    txtTicket2.SelBold = True
    txtTicket2.SelText = txtTicket2.SelText + "SUBTOTAL SIN DESCUENTO           " + CStr(rs(19)) + vbCrLf
    txtTicket2.SelBold = False
    txtTicket2.SelText = txtTicket2.SelText + "TOTAL DESCUENTOS                             0,00" + vbCrLf
    txtTicket2.SelText = txtTicket2.SelText + "TOTAL NETO SIN IVA                     " + CStr(rs(19)) + vbCrLf
    txtTicket2.SelText = txtTicket2.SelText + "SUBTOTAL 21.00 %                       " + CStr(rs(19)) + vbCrLf
    txtTicket2.SelText = txtTicket2.SelText + "IVA 21,00 %                            " + CStr(rs(23)) + vbCrLf
    txtTicket2.SelText = txtTicket2.SelText + "CONCEPTOS NO GRAVADOS                  " + CStr(rs(24)) + vbCrLf
    txtTicket2.SelBold = True
    txtTicket2.SelText = txtTicket2.SelText + "TOTAL                             " + CStr(rs(25)) + vbCrLf
    txtTicket2.SelBold = False
    txtTicket2.SelText = txtTicket2.SelText + vbCrLf
    
    txtTicket2.SelText = txtTicket2.SelText + "OTROS IMPUESTOS NACIONALES" + vbCrLf
    txtTicket2.SelText = txtTicket2.SelText + "Imp. Líq. Carb.                         " + CStr(rs(26)) + vbCrLf
    txtTicket2.SelText = txtTicket2.SelText + "Imp. Comb. Liq.                          " + CStr(rs(27)) + vbCrLf
    txtTicket2.SelText = txtTicket2.SelText + "OTROS IMPUESTOS" + vbCrLf
    txtTicket2.SelText = txtTicket2.SelText + "IVA Matanza Variable 1.5%               " + CStr(rs(28)) + vbCrLf
    
    txtTicket2.SelText = txtTicket2.SelText + vbCrLf
    txtTicket2.SelText = txtTicket2.SelText + vbCrLf
    txtTicket2.SelBold = True
    txtTicket2.SelText = txtTicket2.SelText + "RECIBI/MOS" + vbCrLf
    txtTicket2.SelBold = False
    txtTicket2.SelText = txtTicket2.SelText + "Efectivo                             " + CStr(rs(25)) + vbCrLf
    txtTicket2.SelText = txtTicket2.SelText + "Suma de sus pagos                     " + CStr(rs(25)) + vbCrLf
    txtTicket2.SelBold = True
    txtTicket2.SelText = txtTicket2.SelText + "Su vuelto                          0,00" + vbCrLf
    txtTicket2.SelBold = False
    txtTicket2.SelText = txtTicket2.SelText + "Cantidad Unidades                               1" + vbCrLf
    txtTicket2.SelText = txtTicket2.SelText + vbCrLf
    txtTicket2.SelText = txtTicket2.SelText + "-------------------------------------------------" + vbCrLf
    txtTicket2.SelText = txtTicket2.SelText + "#### REFERENCIA ELECTRONICA DEL COMPROBANTE ####" + vbCrLf
    txtTicket2.SelText = txtTicket2.SelText + "      C.A.E.: 75185950369313 Vto.: 20250516" + vbCrLf
    txtTicket2.SelText = txtTicket2.SelText + "-------------------------------------------------" + vbCrLf
    txtTicket2.SelText = txtTicket2.SelText + vbCrLf
    txtTicket2.SelText = txtTicket2.SelText + "    Venta combustible por cuenta y orden YPF" + vbCrLf
    txtTicket2.SelText = txtTicket2.SelText + "     Print.Al.Consumidor Bs As 0800-222-9042" + vbCrLf
    txtTicket2.SelText = txtTicket2.SelText + vbCrLf
End Sub



