Attribute VB_Name = "modConexion"
Option Explicit

' Módulo para manejar la conexión global
Public conn As New ADODB.Connection

Public Sub ConectarBD()
    
    Dim rutaIni As String
    Dim strConn As String
    
    ' Obtener ruta del archivo INI (en el mismo directorio que la aplicación)
    rutaIni = App.Path & "\login.ini"
    
    ' Leer configuración
    Dim Provider As String
    Dim DataSource As String
    Dim InitialCatalog As String
    Dim UserID As String
    Dim Password As String
    
    Provider = LeerIni("DatabaseConfig", "Provider", rutaIni)
    DataSource = LeerIni("DatabaseConfig", "DataSource", rutaIni)
    InitialCatalog = LeerIni("DatabaseConfig", "InitialCatalog", rutaIni)
    UserID = LeerIni("DatabaseConfig", "UserID", rutaIni)
    Password = LeerIni("DatabaseConfig", "Password", rutaIni)
    
    ' Armar cadena de conexión
    strConn = "Provider=SQLOLEDB;" & _
          "Data Source=" & DataSource & ";" & _
          "Initial Catalog=" & InitialCatalog & ";" & _
          "Integrated Security=SSPI;" & _
          "TrustServerCertificate=True"
    

    ' Abrir conexión
    On Error GoTo ErrHandler
    If conn.State = 0 Then
        conn.Open strConn
    End If
    Exit Sub

ErrHandler:
    MsgBox "Error al conectar: " & Err.Description, vbCritical, "Error"
End Sub

Public Sub DesconectarBD()
    ' Cierra la conexión si está abierta
    If conn.State = 1 Then
        conn.Close
    End If
End Sub


