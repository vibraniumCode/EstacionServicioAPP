Attribute VB_Name = "modINI"
' Declaraciones de API para manejar archivos INI
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" ( _
    ByVal lpApplicationName As String, _
    ByVal lpKeyName As String, _
    ByVal lpDefault As String, _
    ByVal lpReturnedString As String, _
    ByVal nSize As Long, _
    ByVal lpFileName As String) As Long

Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" ( _
    ByVal lpApplicationName As String, _
    ByVal lpKeyName As String, _
    ByVal lpString As String, _
    ByVal lpFileName As String) As Long
    
    
Public Function LeerIni(ByVal Seccion As String, ByVal Clave As String, ByVal Archivo As String) As String
    Dim Buffer As String * 255
    Dim Resultado As Long
    
    Resultado = GetPrivateProfileString(Seccion, Clave, "", Buffer, 255, Archivo)
    LeerIni = Left$(Buffer, Resultado)
End Function

Public Sub EscribirIni(ByVal Seccion As String, ByVal Clave As String, ByVal Valor As String, ByVal Archivo As String)
    WritePrivateProfileString Seccion, Clave, Valor, Archivo
End Sub

