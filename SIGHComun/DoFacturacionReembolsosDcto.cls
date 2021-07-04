VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DoFacturacionReembolsosDcto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'------------------------------------------------------------------------------------
'        Inicio de c�digo autogenerado para la clase: FactReembolsosDocumentos
'        Autor: Daniel Enrique Barrantes
'        Fecha: 08/04/2010 11:46:48 a.m.
'        Auspiciado por: Hospital Regional Ayacucho
'        Control De Cambios:
'------------------------------------------------------------------------------------
'        Autor                      Fecha                      Cambio
'------------------------------------------------------------------------------------


Dim ml_IdUsuarioAuditoria As Long
Dim ml_idFactReembolso As Long
Dim ms_NroSerie As String
Dim ms_NroDocumento As String
Dim ms_MotivoAnulacion As String

Property Let IdUsuarioAuditoria(lValue As Long)
   ml_IdUsuarioAuditoria = lValue
End Property
Property Get IdUsuarioAuditoria() As Long
   IdUsuarioAuditoria = ml_IdUsuarioAuditoria
End Property

Property Let idFactReembolso(iValue As Long)
   ml_idFactReembolso = iValue
End Property
Property Get idFactReembolso() As Long
   idFactReembolso = ml_idFactReembolso
End Property
Property Let NroSerie(sValue As String)
   ms_NroSerie = sValue
End Property
Property Get NroSerie() As String
   NroSerie = ms_NroSerie
End Property
Property Let NroDocumento(sValue As String)
   ms_NroDocumento = sValue
End Property
Property Get NroDocumento() As String
   NroDocumento = ms_NroDocumento
End Property
Property Let MotivoAnulacion(sValue As String)
   ms_MotivoAnulacion = sValue
End Property
Property Get MotivoAnulacion() As String
   MotivoAnulacion = ms_MotivoAnulacion
End Property