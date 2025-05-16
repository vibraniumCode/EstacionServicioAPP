VERSION 5.00
Begin VB.Form FRMMenu 
   BackColor       =   &H8000000A&
   Caption         =   "Sistema de Carga y Facturación    ---    Base de Datos: ""Local""    -    Usuario: 'Default'"
   ClientHeight    =   3135
   ClientLeft      =   165
   ClientTop       =   510
   ClientWidth     =   4680
   Icon            =   "Menu.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3135
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Menu Clientes 
      Caption         =   "Clientes"
      Begin VB.Menu CClientes 
         Caption         =   "&Carga"
      End
   End
   Begin VB.Menu Facturacion 
      Caption         =   "Facturación"
      Begin VB.Menu Conbustible 
         Caption         =   "&Conbustible"
      End
      Begin VB.Menu pventa 
         Caption         =   "&Punto de venta"
      End
   End
   Begin VB.Menu PEsp 
      Caption         =   "Procesos Especiales"
      Begin VB.Menu Cierre 
         Caption         =   "&Cierre"
      End
   End
End
Attribute VB_Name = "FRMMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub mnuNavbar_Click()

End Sub

Private Sub CClientes_Click()
FRMCliente.Show vbModal
End Sub

Private Sub Conbustible_Click()
frmCombustible.Show vbModal
End Sub

Private Sub pventa_Click()
FRMFacturacion.Show vbModal
End Sub
