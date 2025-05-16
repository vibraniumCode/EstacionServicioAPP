Attribute VB_Name = "ModuloClientes"
' En un módulo estándar llamado "ModuloClientes"
Option Explicit

' Variables públicas para almacenar los datos de clientes
Public ClientesArray() As String  ' Para almacenar las descripciones
Public ClientesIDs() As Long      ' Para almacenar los IDs correspondientes

Public Sub CargarDatosClientes()
    Dim rs As New ADODB.Recordset
    Dim i As Integer
    
    ' Inicializar con arrays vacios por defecto
    ReDim ClientesArray(0 To 0)
    ReDim ClientesIDs(0 To 0)
    
    ' Conectar a la base de datos
    Call ConectarBD
    
    On Error GoTo ErrHandler
    rs.Open "SELECT id, concat(FORMAT(id, '00000000'),'-',nombre,'-',direccion,'-',cuit) AS descripcion FROM clientes ORDER BY nombre", conn, adOpenStatic, adLockReadOnly
    
    ' Determinar el numero de registros
    If Not rs.EOF Then
        rs.MoveLast
        ReDim ClientesArray(0 To rs.RecordCount - 1)
        ReDim ClientesIDs(0 To rs.RecordCount - 1)
        rs.MoveFirst
        
        i = 0
        Do While Not rs.EOF
            ' Guardar datos en arrays
            ClientesArray(i) = rs("descripcion")
            ClientesIDs(i) = rs("id")
            
            i = i + 1
            rs.MoveNext
        Loop
    End If
    
    ' Cerrar el recordset
    rs.Close
    
    ' Desconectar
    Call DesconectarBD
    Exit Sub
    
ErrHandler:
    MsgBox "Error al cargar datos de clientes: " & Err.Description, vbCritical, "Error"
    Call DesconectarBD
    ' Asegurar que los arrays estén inicializados incluso en caso de error
    ReDim ClientesArray(0 To 0)
    ReDim ClientesIDs(0 To 0)
End Sub
