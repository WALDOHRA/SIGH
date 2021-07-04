VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DORolPermiso"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organizaci�n: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa:  Clase para capa de estructura de la tabla RolesPermisos
'        Programado por: Castro W
'        Fecha: Marzo 2006
'
'------------------------------------------------------------------------------------
Option Explicit

Dim ml_Auditoria As Long
Dim ml_IdRol As Long
Dim ml_IdPermiso As Long
Dim ml_IdRolPermiso As Long
Property Let IdUsuarioAuditoria(lValue As Long)
   ml_Auditoria = lValue
End Property
Property Get IdUsuarioAuditoria() As Long
   IdUsuarioAuditoria = ml_Auditoria
End Property
Property Let IdRol(lValue As Long)
   ml_IdRol = lValue
End Property
Property Get IdRol() As Long
   IdRol = ml_IdRol
End Property
Property Let IdPermiso(lValue As Long)
   ml_IdPermiso = lValue
End Property
Property Get IdPermiso() As Long
   IdPermiso = ml_IdPermiso
End Property
Property Let IdRolPermiso(lValue As Long)
   ml_IdRolPermiso = lValue
End Property
Property Get IdRolPermiso() As Long
   IdRolPermiso = ml_IdRolPermiso
End Property

