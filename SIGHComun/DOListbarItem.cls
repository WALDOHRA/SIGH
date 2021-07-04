VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DOListbarItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organizaci�n: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa:  Clase para capa de estructura de la tabla ListBarItems
'        Programado por: Barrantes D.
'        Fecha: Mayo 2007
'
'------------------------------------------------------------------------------------
Option Explicit

Dim ml_Auditoria As Long
Dim ms_KeyIcon As String
Dim ml_Indice As Long
Dim ms_Clave As String
Dim ms_Texto As String
Dim ml_IdListGrupo As Long
Dim ml_IdListItem As Long
Property Let IdUsuarioAuditoria(lValue As Long)
   ml_Auditoria = lValue
End Property
Property Get IdUsuarioAuditoria() As Long
   IdUsuarioAuditoria = ml_Auditoria
End Property
Property Let KeyIcon(sValue As String)
   ms_KeyIcon = sValue
End Property
Property Get KeyIcon() As String
   KeyIcon = ms_KeyIcon
End Property
Property Let Indice(lValue As Long)
   ml_Indice = lValue
End Property
Property Get Indice() As Long
   Indice = ml_Indice
End Property
Property Let Clave(sValue As String)
   ms_Clave = sValue
End Property
Property Get Clave() As String
   Clave = ms_Clave
End Property
Property Let Texto(sValue As String)
   ms_Texto = sValue
End Property
Property Get Texto() As String
   Texto = ms_Texto
End Property
Property Let IdListGrupo(lValue As Long)
   ml_IdListGrupo = lValue
End Property
Property Get IdListGrupo() As Long
   IdListGrupo = ml_IdListGrupo
End Property
Property Let IdListItem(lValue As Long)
   ml_IdListItem = lValue
End Property
Property Get IdListItem() As Long
   IdListItem = ml_IdListItem
End Property

