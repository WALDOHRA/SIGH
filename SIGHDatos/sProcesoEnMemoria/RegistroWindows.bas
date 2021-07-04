Attribute VB_Name = "RegistroWindows"
'------------------------------------------------------------------------------------
'        Organizaci�n: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: M�dulo para Registro en WIndows
'        Programado por: Barrantes D
'        Fecha: Enero 2009
'
'------------------------------------------------------------------------------------
Option Explicit
  
'Declaraci�n de constantes
'****************************
  
  
Private Const REG_SZ As Long = 1
Private Const REG_DWORD As Long = 4
  
Private Const ERROR_NONE = 0
Private Const ERROR_BADDB = 1
Private Const ERROR_BADKEY = 2
Private Const ERROR_CANTOPEN = 3
Private Const ERROR_CANTREAD = 4
Private Const ERROR_CANTWRITE = 5
Private Const ERROR_OUTOFMEMORY = 6
Private Const ERROR_INVALID_PARAMETER = 7
Private Const ERROR_ACCESS_DENIED = 8
Private Const ERROR_INVALID_PARAMETERS = 87
Private Const ERROR_NO_MORE_ITEMS = 259
  
Private Const KEY_ALL_ACCESS = &H3F
  
Private Const REG_OPTION_NON_VOLATILE = 0
  
  
'Declaraci�n de las funciones api para el registro
'*************************************************
  
  
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function RegCreateKeyEx Lib "advapi32.dll" Alias "RegCreateKeyExA" _
       (ByVal hKey As Long, _
        ByVal lpSubKey As String, _
        ByVal Reserved As Long, _
        ByVal lpClass As String, _
        ByVal dwOptions As Long, _
        ByVal samDesired As Long, _
        ByVal lpSecurityAttributes As Long, _
        phkResult As Long, _
        lpdwDisposition As Long) As Long
  
Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" _
       (ByVal hKey As Long, _
        ByVal lpSubKey As String, _
        ByVal ulOptions As Long, _
        ByVal samDesired As Long, _
        phkResult As Long) As Long
  
Private Declare Function RegQueryValueExString Lib "advapi32.dll" Alias "RegQueryValueExA" _
       (ByVal hKey As Long, _
        ByVal lpValueName As String, _
        ByVal lpReserved As Long, _
        lpType As Long, _
        ByVal lpData As String, _
        lpcbData As Long) As Long
  
Private Declare Function RegQueryValueExLong Lib "advapi32.dll" Alias "RegQueryValueExA" _
       (ByVal hKey As Long, _
        ByVal lpValueName As String, _
        ByVal lpReserved As Long, _
        lpType As Long, _
        lpData As Long, _
        lpcbData As Long) As Long
           
Private Declare Function RegQueryValueExNULL Lib "advapi32.dll" Alias "RegQueryValueExA" _
       (ByVal hKey As Long, _
        ByVal lpValueName As String, _
        ByVal lpReserved As Long, _
        lpType As Long, _
        ByVal lpData As Long, _
        lpcbData As Long) As Long
  
Private Declare Function RegSetValueExString Lib "advapi32.dll" Alias "RegSetValueExA" _
        (ByVal hKey As Long, _
         ByVal lpValueName As String, _
         ByVal Reserved As Long, _
         ByVal dwType As Long, _
         ByVal lpValue As String, _
         ByVal cbData As Long) As Long
Private Declare Function RegSetValueExLong Lib "advapi32.dll" Alias _
         "RegSetValueExA" (ByVal hKey As Long, _
         ByVal lpValueName As String, _
         ByVal Reserved As Long, _
         ByVal dwType As Long, _
         lpValue As Long, _
         ByVal cbData As Long) As Long
  
Private Declare Function RegDeleteKey& Lib "advapi32.dll" Alias "RegDeleteKeyA" _
        (ByVal hKey As Long, _
         ByVal lpSubKey As String)
  
Private Declare Function RegDeleteValue& Lib "advapi32.dll" Alias "RegDeleteValueA" _
        (ByVal hKey As Long, _
         ByVal lpValueName As String)
  
  
'Funciones p�blicas para crear, eliminar, consultar los datos
  
'****************************************************************
  
  
' Funci�n que elimina una clave especifica utilizando el Api RegDeleteKey
  
Public Function EliminarClave(clave As Long, Nombre_clave As String)
       
    Dim ret As Long
       
    ret = RegDeleteKey(clave, Nombre_clave)
       
End Function
  
' Funci�n que elimina un dato utilizando el Api RegDeleteValue
  
Public Function EliminarValor(clave As Long, _
                              Nombre_clave As String, _
                              Nombre_valor As String)
  
  
       Dim ret As Long
       Dim Handle_clave As Long
          
       ' Abre la clave del registro
       ret = RegOpenKeyEx(clave, Nombre_clave, 0, KEY_ALL_ACCESS, Handle_clave)
          
       'Elimina el valor del registro
       ret = RegDeleteValue(Handle_clave, Nombre_valor)
          
       'Cierra la vlave del registro abierta
       RegCloseKey (Handle_clave)
  
End Function
  
' Funci�n que crea una nueva Clave utilizando el Api RegCreateKeyEx
  
Public Function CrearNuevaClave(clave As Long, Nombre_clave As String)
  
    Dim Handle_clave As Long
    Dim ret As Long
       
    ret = RegCreateKeyEx(clave, _
                         Nombre_clave, 0&, vbNullString, _
                         REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, 0&, _
                         Handle_clave, ret)
       
    RegCloseKey (Handle_clave)
       
End Function
  
' Funci�n que establece un nuevo valor mediante el Api SetValueEx
  
Public Function EstablecerValor(clave As Long, _
                                Nombre_clave As String, _
                                Nombre_valor As String, _
                                el_Valor As Variant, _
                                Tipo_Valor As Long)
  
  
       Dim ret As Long
       Dim Handle_clave As Long
  
       ret = RegOpenKeyEx(clave, Nombre_clave, 0, KEY_ALL_ACCESS, Handle_clave)
       ret = SetValueEx(Handle_clave, Nombre_valor, Tipo_Valor, el_Valor)
          
       RegCloseKey (Handle_clave)
  
End Function
  
' Funci�n que consulta un dato del registro usando QueryValueEx
  
Public Function ConsultarValor(clave As Long, Nombre_clave As String, Nombre_valor As String)
  
       Dim Handle_clave As Long
       Dim valor As Variant
  
       Dim ret As Long
  
       ret = RegOpenKeyEx(clave, Nombre_clave, 0, KEY_ALL_ACCESS, Handle_clave)
          
       ret = QueryValueEx(Handle_clave, Nombre_valor, valor)
       ' REtorna el valor del registro a la funci�n
       ConsultarValor = valor
       'Cierra la clave abierta del registro
       RegCloseKey (Handle_clave)
End Function
  
  
  
' Funciones privadas del m�dulo
  
Private Function SetValueEx(ByVal Handle_clave As Long, _
                            Nombre_valor As String, _
                            Tipo As Long, _
                            el_Valor As Variant) As Long
       
    Dim ret As Long
    Dim sValue As String
  
    Select Case Tipo
           
        ' Valor de tipo cadena
        Case REG_SZ
               
            sValue = el_Valor
            SetValueEx = RegSetValueExString(Handle_clave, _
                                             Nombre_valor, 0&, _
                                             Tipo, sValue, Len(sValue))
           
        'Valor Entero
        Case REG_DWORD
            ret = el_Valor
            SetValueEx = RegSetValueExLong(Handle_clave, Nombre_valor, 0&, Tipo, ret, 4)
        End Select
  
End Function
  
Private Function QueryValueEx(ByVal lhKey As Long, _
                              ByVal Name_Valor As String, _
                              el_Valor As Variant) As Long
       
    Dim cch As Long
    Dim lrc As Long
    Dim Tipo As Long
    Dim ret_Valor As Long
    Dim dato As String
  
    On Error GoTo QueryValueExError
  
    lrc = RegQueryValueExNULL(lhKey, Name_Valor, 0&, Tipo, 0&, cch)
    If lrc <> ERROR_NONE Then Error 5
  
    Select Case Tipo
           
        Case REG_SZ:
               
            dato = String(cch, 0)
            lrc = RegQueryValueExString(lhKey, Name_Valor, 0&, Tipo, dato, cch)
            If lrc = ERROR_NONE Then
                el_Valor = Left$(dato, cch)
            Else
                el_Valor = Empty
            End If
  
        Case REG_DWORD:
            lrc = RegQueryValueExLong(lhKey, Name_Valor, 0&, Tipo, ret_Valor, cch)
            If lrc = ERROR_NONE Then el_Valor = ret_Valor
        Case Else
            lrc = -1
    End Select
  
QueryValueExExit:
  
    QueryValueEx = lrc
    Exit Function
  
QueryValueExError:
  
    Resume QueryValueExExit
  
End Function


