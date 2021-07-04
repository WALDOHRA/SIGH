VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DOSubclasificacionDiagnostico"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organizaci�n: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa:  Clase para capa de estructura de la tabla SubclasificacionDiagnosticos
'        Programado por: Castro W
'        Fecha: Julio 2004
'
'------------------------------------------------------------------------------------
Option Explicit

Dim ml_Auditoria As Long
Dim ml_IdTipoServicio As Long
Dim ml_IdClasificacionDx As Long
Dim ms_Descripcion As String
Dim ms_Codigo As String
Dim ml_IdSubClasificacionDX As Long
Property Let IdUsuarioAuditoria(lValue As Long)
   ml_Auditoria = lValue
End Property
Property Get IdUsuarioAuditoria() As Long
   IdUsuarioAuditoria = ml_Auditoria
End Property
Property Let IdTipoServicio(lValue As Long)
   ml_IdTipoServicio = lValue
End Property
Property Get IdTipoServicio() As Long
   IdTipoServicio = ml_IdTipoServicio
End Property
Property Let IdClasificacionDx(lValue As Long)
   ml_IdClasificacionDx = lValue
End Property
Property Get IdClasificacionDx() As Long
   IdClasificacionDx = ml_IdClasificacionDx
End Property
Property Let Descripcion(sValue As String)
   ms_Descripcion = sValue
End Property
Property Get Descripcion() As String
   Descripcion = ms_Descripcion
End Property
Property Let Codigo(sValue As String)
   ms_Codigo = sValue
End Property
Property Get Codigo() As String
   Codigo = ms_Codigo
End Property
Property Let IdSubClasificacionDX(lValue As Long)
   ml_IdSubClasificacionDX = lValue
End Property
Property Get IdSubClasificacionDX() As Long
   IdSubClasificacionDX = ml_IdSubClasificacionDX
End Property

