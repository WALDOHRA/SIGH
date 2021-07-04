VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DoImagMovimientoIngresos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organizaci�n: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa:  Clase para capa de estructura de la tabla ImagMovimientoIngresos
'        Programado por: Barrantes D
'        Fecha: Abril 2009
'
'------------------------------------------------------------------------------------
Option Explicit

Dim ml_IdUsuarioAuditoria As Long
Dim ml_IdMovimiento As Long
Dim ms_NroDocumento As String
Dim ml_IdPersonaRecepciona As Long

Property Let IdUsuarioAuditoria(lValue As Long)
   ml_IdUsuarioAuditoria = lValue
End Property
Property Get IdUsuarioAuditoria() As Long
   IdUsuarioAuditoria = ml_IdUsuarioAuditoria
End Property

Property Let IdMovimiento(iValue As Long)
   ml_IdMovimiento = iValue
End Property
Property Get IdMovimiento() As Long
   IdMovimiento = ml_IdMovimiento
End Property
Property Let NroDocumento(sValue As String)
   ms_NroDocumento = sValue
End Property
Property Get NroDocumento() As String
   NroDocumento = ms_NroDocumento
End Property
Property Let IdPersonaRecepciona(iValue As Long)
   ml_IdPersonaRecepciona = iValue
End Property
Property Get IdPersonaRecepciona() As Long
   IdPersonaRecepciona = ml_IdPersonaRecepciona
End Property

