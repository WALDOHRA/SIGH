VERSION 5.00
Begin VB.Form PacientesBusqueda 
   Caption         =   "B?squeda de Pacientes"
   ClientHeight    =   6150
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11505
   Icon            =   "PacientesBusqueda.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6150
   ScaleWidth      =   11505
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Height          =   1065
      Left            =   45
      TabIndex        =   1
      Top             =   5025
      Width           =   11400
      Begin VB.CommandButton btnCancelar 
         Caption         =   "Cancelar (ESC)"
         DisabledPicture =   "PacientesBusqueda.frx":0CCA
         DownPicture     =   "PacientesBusqueda.frx":118E
         Height          =   700
         Left            =   5715
         Picture         =   "PacientesBusqueda.frx":167A
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   240
         Width           =   1365
      End
      Begin VB.CommandButton btnAceptar 
         Caption         =   "Aceptar (F2)"
         DisabledPicture =   "PacientesBusqueda.frx":1B66
         DownPicture     =   "PacientesBusqueda.frx":1FC6
         Height          =   700
         Left            =   4170
         Picture         =   "PacientesBusqueda.frx":243B
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   240
         Width           =   1365
      End
   End
   Begin Galenhos.ucPacientesLista ucPacientesLista1 
      Height          =   5010
      Left            =   15
      TabIndex        =   0
      Top             =   15
      Width           =   11430
      _extentx        =   20161
      _extenty        =   8837
   End
End
Attribute VB_Name = "PacientesBusqueda"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mi_BotonPresionado As sghBotonDetallePresionado

Property Set DataSource(oValue As ADODB.Recordset)
    Set ucPacientesLista1.DataSource = oValue
End Property
Property Get DataSource() As ADODB.Recordset
    Set DataSource = ucPacientesLista1.DataSource
End Property
Property Let IdRegistroSeleccionado(lValue As Long)
    ucPacientesLista1.IdRegistroSeleccionado = lValue
End Property
Property Get IdRegistroSeleccionado() As Long
    IdRegistroSeleccionado = ucPacientesLista1.IdRegistroSeleccionado
End Property
Property Get BotonPresionado() As sghBotonDetallePresionado
    BotonPresionado = mi_BotonPresionado
End Property
Property Let TipoFiltro(lValue As sghTipoFiltroPacientes)
    ucPacientesLista1.TipoFiltro = lValue
End Property
Property Get TipoFiltro() As sghTipoFiltroPacientes
    TipoFiltro = ucPacientesLista1.TipoFiltro
End Property


Private Sub btnAceptar_Click()
    mi_BotonPresionado = sghAceptar
    Me.Visible = False
End Sub

Private Sub btnCancelar_Click()
    mi_BotonPresionado = sghCancelar
    Me.Visible = False
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    AdministrarKeyPreview KeyCode
End Sub

Private Sub Form_Load()
    Me.ucPacientesLista1.Titulo = "B?squeda de Pacientes"
End Sub
    
Sub AdministrarKeyPreview(KeyCode As Integer)
   Select Case KeyCode
        Case vbKeyF6
            ucPacientesLista1.RealizarBusqueda
        Case vbKeyEscape
           btnCancelar_Click
        Case vbKeyF2
           btnAceptar_Click
       End Select
End Sub

