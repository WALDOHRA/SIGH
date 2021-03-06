VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form SiProgramacionDetalle 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   3450
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5250
   Icon            =   "SiProgramacionDetalle.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3450
   ScaleWidth      =   5250
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Height          =   1080
      Left            =   75
      TabIndex        =   11
      Top             =   2280
      Width           =   5085
      Begin VB.CommandButton btnCancelar 
         Cancel          =   -1  'True
         Caption         =   "Cancelar (ESC)"
         DisabledPicture =   "SiProgramacionDetalle.frx":08CA
         DownPicture     =   "SiProgramacionDetalle.frx":0D8E
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   700
         Left            =   2580
         Picture         =   "SiProgramacionDetalle.frx":127A
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   225
         Width           =   1365
      End
      Begin VB.CommandButton btnAceptar 
         Caption         =   "Aceptar (F2)"
         DisabledPicture =   "SiProgramacionDetalle.frx":1766
         DownPicture     =   "SiProgramacionDetalle.frx":1BC6
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   700
         Left            =   1020
         Picture         =   "SiProgramacionDetalle.frx":203B
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   225
         Width           =   1365
      End
   End
   Begin VB.Frame fraProg 
      Height          =   2220
      Left            =   75
      TabIndex        =   9
      Top             =   15
      Width           =   5100
      Begin VB.ComboBox cmbIdServicio 
         DataField       =   "IdEspecialidad"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         ItemData        =   "SiProgramacionDetalle.frx":24B0
         Left            =   1410
         List            =   "SiProgramacionDetalle.frx":24B2
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   585
         Width           =   3500
      End
      Begin VB.ComboBox cmbIdTurno 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1410
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   1365
         Width           =   3495
      End
      Begin VB.TextBox txtFechaFin 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   3690
         TabIndex        =   3
         Top             =   990
         Visible         =   0   'False
         Width           =   1200
      End
      Begin VB.TextBox txtMedico 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1425
         TabIndex        =   0
         Top             =   195
         Width           =   3500
      End
      Begin VB.TextBox txtFechaIni 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1410
         TabIndex        =   2
         Top             =   990
         Width           =   1200
      End
      Begin MSMask.MaskEdBox txtHoraInicio 
         Height          =   315
         Left            =   1410
         TabIndex        =   5
         Top             =   1755
         Width           =   780
         _ExtentX        =   1376
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   5
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##:##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtHoraFin 
         Height          =   315
         Left            =   4110
         TabIndex        =   6
         Top             =   1740
         Width           =   750
         _ExtentX        =   1323
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   5
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##:##"
         PromptChar      =   "_"
      End
      Begin VB.Label lblServicio 
         Caption         =   "Punto de Carga"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   615
         Width           =   1290
      End
      Begin VB.Label lblHoraFin 
         Caption         =   "Hora final"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   3165
         TabIndex        =   15
         Top             =   1785
         Width           =   1005
      End
      Begin VB.Label lblHoraInicio 
         Caption         =   "Hora inicio"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   120
         TabIndex        =   14
         Top             =   1800
         Width           =   1005
      End
      Begin VB.Label Label44 
         Caption         =   "Responsable"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   135
         TabIndex        =   13
         Top             =   225
         Width           =   1365
      End
      Begin VB.Label Label2 
         Caption         =   "Turno"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   120
         TabIndex        =   12
         Top             =   1410
         Width           =   1665
      End
      Begin VB.Label lblFecha 
         Caption         =   "Fecha"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   120
         TabIndex        =   10
         Top             =   1020
         Width           =   1005
      End
   End
End
Attribute VB_Name = "SiProgramacionDetalle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'------------------------------------------------------------------------------------
'        Organizaci?n: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Programar a M?dicos
'        Programado por: Barrantes D
'        Fecha: Enero 2009
'
'------------------------------------------------------------------------------------
Option Explicit

Dim mo_Teclado As New SIGHEntidades.Teclado
Dim mo_Formulario As New SIGHEntidades.Formulario
Dim ml_idUsuario As Long
Dim ms_MensajeError As String
Dim mi_Opcion As sghOpciones
Dim mb_ExistenDatos As Boolean
Dim ml_IdProgramacion As Long
Dim ml_IdMedico As Long
Dim ms_NombreMedico  As String
Dim ml_IdDepartamento As Long
Dim ml_IdEspecialidad As Long
Dim ml_IdTipoServicio As Long
Dim mo_ReglasImagenes As New SIGHNegocios.ReglasImagenes
Dim mo_AdminServiciosComunes As New SIGHNegocios.ReglasComunes
Dim mo_AdminServiciosHosp As New SIGHNegocios.ReglasServiciosHosp
Dim mo_AdminProgMedica As New SIGHNegocios.ReglasDeProgMedica
Dim mo_ReglasFacturacion As New SIGHNegocios.ReglasFacturacion
Dim mo_ReglasAdmision As New SIGHNegocios.ReglasAdmision
Dim mo_ReglasDeSeguridad As New SIGHNegocios.ReglasDeSeguridad
Dim ml_NroCuposCE As Long
Dim mo_Diario As PVDayView.PVDayView
Dim mo_Calendario As PVCalendar
Dim mb_SeHaModificadoProgramacion As Boolean
Dim mo_FechaHora As New SIGHEntidades.FechaHora
Dim ml_IdPuntoCarga As Long
Dim mo_cmbIdTurno As New SIGHEntidades.ListaDespleglable
Dim mo_cmbIdServicio As New SIGHEntidades.ListaDespleglable
Dim lcBuscaParametro As New SIGHDatos.Parametros
Dim mo_lnIdTablaLISTBARITEMS As Long
Dim mo_lcNombrePc As String
Dim lbSoloCambiaHoraFinal As Boolean
Dim ldFechaRegistro As Date
Property Let lcNombrePc(lValue As String)
   mo_lcNombrePc = lValue
End Property
Property Let lnIdTablaLISTBARITEMS(lValue As Long)
   mo_lnIdTablaLISTBARITEMS = lValue
End Property
Property Set Diario(oValue As PVDayView.PVDayView)
   Set mo_Diario = oValue
End Property
Property Get Diario() As PVDayView.PVDayView
   Set Diario = mo_Diario
End Property
Property Set Calendario(oValue As PVCalendar)
   Set mo_Calendario = oValue
End Property
Property Get Calendario() As PVCalendar
   Set Calendario = mo_Calendario
End Property

Property Let ExistenDatos(bValue As Boolean)
   mb_ExistenDatos = bValue
End Property
Property Get ExistenDatos() As Boolean
   ExistenDatos = mb_ExistenDatos
End Property
Property Let Opcion(iValue As sghOpciones)
   mi_Opcion = iValue
End Property
Property Get Opcion() As sghOpciones
   Opcion = mi_Opcion
End Property
Property Let MensajeError(sValue As String)
   ms_MensajeError = sValue
End Property
Property Get MensajeError() As String
   MensajeError = ms_MensajeError
End Property
Property Let idUsuario(lValue As Long)
   ml_idUsuario = lValue
End Property
Property Get idUsuario() As Long
   idUsuario = ml_idUsuario
End Property
Property Let IdProgramacion(lValue As Long)
   ml_IdProgramacion = lValue
End Property
Property Get IdProgramacion() As Long
   IdProgramacion = ml_IdProgramacion
End Property
Property Let idMedico(lValue As Long)
   ml_IdMedico = lValue
End Property
Property Get idMedico() As Long
   idMedico = ml_IdMedico
End Property
Property Let NombreMedico(sValue As String)
   ms_NombreMedico = sValue
End Property
Property Get NombreMedico() As String
   NombreMedico = ms_NombreMedico
End Property

Property Let NroCuposCE(lValue As Long)
   ml_NroCuposCE = lValue
End Property
Property Get NroCuposCE() As Long
   NroCuposCE = ml_NroCuposCE
End Property
Property Get SeHaModificadoProgramacion() As Long
   SeHaModificadoProgramacion = mb_SeHaModificadoProgramacion
End Property

Property Let idPuntoCarga(sValue As Long)
    ml_IdPuntoCarga = sValue
End Property







Private Sub cmbIdServicio_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, cmbIdServicio
    AdministrarKeyPreview KeyCode

End Sub




Sub CargarComboBoxes()
       
        mo_cmbIdTurno.BoundColumn = "IdTurno"
        mo_cmbIdTurno.ListField = "DescripcionLarga"
        Set mo_cmbIdTurno.RowSource = mo_AdminProgMedica.TurnosSeleccionarPorIdTipoServicio(1)
       
        Dim oRsSalas As New Recordset
        Set oRsSalas = mo_ReglasImagenes.SiCitasSalasSeleccionarTodas
        oRsSalas.Filter = "idPuntoCarga=" & ml_IdPuntoCarga
        mo_cmbIdServicio.BoundColumn = "idSala"
        mo_cmbIdServicio.ListField = "sala"
        Set mo_cmbIdServicio.RowSource = oRsSalas
        If oRsSalas.RecordCount = 1 Then
           oRsSalas.MoveFirst
           mo_cmbIdServicio.BoundText = oRsSalas!idSala
        End If
        Set oRsSalas = Nothing
        
        
End Sub









Private Sub cmbIdTurno_Click()
Dim oDOTurno As New doTurno
Dim IdTurno As Long

    IdTurno = mo_cmbIdTurno.ObtenerItemDataDeComboxBox(cmbIdTurno)
    mo_cmbIdTurno.BoundText = IdTurno

    Set oDOTurno = mo_AdminProgMedica.TurnosSeleccionarPorId(Val(mo_cmbIdTurno.BoundText))
    If Not oDOTurno Is Nothing Then
        Me.txtHoraInicio.Text = oDOTurno.HoraInicio
        Me.txtHoraFin.Text = oDOTurno.HoraFin
    End If
End Sub

Private Sub cmbIdTurno_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, cmbIdTurno
    AdministrarKeyPreview KeyCode
End Sub

Private Sub cmbIdTurno_LostFocus()
Dim oDOTurno As New doTurno

    If cmbIdTurno.Text = "" Then
        Exit Sub
    End If
    
    'Set oDOTurno = mo_AdminProgMedica.TurnosSeleccionarPorCodigo(Trim(Split(cmbIdTurno.Text, "=")(0)))
    Set oDOTurno = mo_AdminProgMedica.TurnosSeleccionarPorId(mo_cmbIdTurno.BoundText)
    If oDOTurno.IdTurno <> 0 Then
        mo_cmbIdTurno.BoundText = oDOTurno.IdTurno
   End If
   
End Sub

Private Sub Form_Initialize()
    

    Set mo_cmbIdServicio.MiComboBox = Me.cmbIdServicio
    Set mo_cmbIdTurno.MiComboBox = Me.cmbIdTurno
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    LimpiarVariablesDeMemoria
End Sub








Private Sub txtHoraFin_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtHoraFin
    AdministrarKeyPreview KeyCode
End Sub

Private Sub txtHoraFin_LostFocus()
  If Not SIGHEntidades.ValidaHora(txtHoraFin.Text) Then
            'MsgBox "La hora ingresada no es correcta", vbInformation, Me.Caption
            Dim oMensaje As New SIGHNegocios.clMensaje
            oMensaje.MostrarFormulario "La hora ingresada no es correcta", Me.Caption
            Set oMensaje = Nothing
            
            txtHoraFin.Text = SIGHEntidades.HORA_VACIA_HM
        End If
End Sub

Private Sub txtHoraFin_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsParaHora(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
End Sub


Private Sub txtHoraInicio_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtHoraInicio
    AdministrarKeyPreview KeyCode
End Sub

Private Sub txtHoraInicio_LostFocus()
    
       If Not SIGHEntidades.ValidaHora(txtHoraInicio.Text) Then
            'MsgBox "La hora ingresada no es correcta", vbInformation, Me.Caption
            Dim oMensaje As New SIGHNegocios.clMensaje
            oMensaje.MostrarFormulario "La hora ingresada no es correcta", Me.Caption
            Set oMensaje = Nothing
            
             txtHoraInicio.Text = SIGHEntidades.HORA_VACIA_HM
        End If
    
    
End Sub

Private Sub txtHoraInicio_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsParaHora(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
End Sub


'------------------------------------------------------------------------------------
'   CargarDatosAlFormulario
'   Descripci?n:    Seleccionar un registro unico de la tabla ProgramacionMedica
'   Par?metros:     Ninguno
'------------------------------------------------------------------------------------

Sub CargarDatosAlFormulario()
        '
        Dim oRsPermisos As New Recordset
        Set oRsPermisos = mo_ReglasDeSeguridad.UsuariosRolesSeleccionarPermisosTodos(ml_idUsuario)
        oRsPermisos.Filter = "idPermiso=500"
        If oRsPermisos.RecordCount = 1 Then
           mo_Formulario.HabilitarDeshabilitar txtHoraInicio, False
           mo_Formulario.HabilitarDeshabilitar txtHoraFin, False
        End If
        '
        lbSoloCambiaHoraFinal = False
        If mi_Opcion = sghModificar Then
           oRsPermisos.Filter = "idPermiso=501"
           If oRsPermisos.RecordCount = 1 Then
              lbSoloCambiaHoraFinal = True
           End If
        End If
        oRsPermisos.Close
        Set oRsPermisos = Nothing
        '
        Select Case mi_Opcion
            Case sghAgregar
               ValoresPorDefecto
            Case sghModificar
                CargarDatosAlosControles
            Case sghConsultar
                CargarDatosAlosControles
            Case sghEliminar
                CargarDatosAlosControles
        End Select
 
        Select Case mi_Opcion
       Case sghAgregar
       Case sghModificar
       Case sghConsultar
            fraProg.Enabled = False
            Me.btnAceptar.Enabled = False
       Case sghEliminar
            fraProg.Enabled = False
       End Select
 
End Sub
Sub ValoresPorDefecto()

    Me.txtMedico = Me.NombreMedico
    
    If mo_Calendario.SelectedDateCount = 1 Then
        Me.txtFechaIni = Format(mo_Calendario.Value, SIGHEntidades.DevuelveFechaSoloFormato_DMY)
        Me.txtFechaFin.Visible = False
    ElseIf mo_Calendario.SelectedDateCount > 1 Then
        Dim DiaSeleccionado As Date
        DiaSeleccionado = mo_Calendario.Value
        Me.txtFechaIni = Format(mo_Calendario.Value, SIGHEntidades.DevuelveFechaSoloFormato_DMY)
        Do While DiaSeleccionado <> 0
            Me.txtFechaFin.Text = Format(DiaSeleccionado, SIGHEntidades.DevuelveFechaSoloFormato_DMY)
            DiaSeleccionado = mo_Calendario.NextSelectedDate(DiaSeleccionado)
        Loop
        Me.txtFechaFin.Visible = True
    End If
    
    Dim rsEspecialidad As Recordset

End Sub


'------------------------------------------------------------------------------------
'   CargarDatosAlFormulario
'   Descripci?n:    Seleccionar un registro unico de la tabla ProgramacionMedica
'   Par?metros:     Ninguno
'------------------------------------------------------------------------------------

Sub Form_Load()
       Select Case mi_Opcion
       Case sghAgregar
           Me.Caption = "Agregar programaci?n"
       Case sghModificar
           Me.Caption = "Modificar programaci?n"
       Case sghConsultar
           Me.Caption = "Consultar programaci?n"
       Case sghEliminar
           Me.Caption = "Eliminar programaci?n"
       End Select

       CargarComboBoxes
       CargarDatosAlFormulario
       mo_Formulario.ConfigurarTipoLetra "Tahoma", "9", Me
       If lbSoloCambiaHoraFinal = True And mi_Opcion = sghModificar Then
            mo_Formulario.HabilitarDeshabilitar cmbIdServicio, False
            mo_Formulario.HabilitarDeshabilitar cmbIdTurno, False
            mo_Formulario.HabilitarDeshabilitar txtHoraInicio, False
            mo_Formulario.HabilitarDeshabilitar txtHoraFin, True
        End If
End Sub

'------------------------------------------------------------------------------------
'   CargarDatosAlFormulario
'   Descripci?n:    Seleccionar un registro unico de la tabla ProgramacionMedica
'   Par?metros:     Ninguno
'------------------------------------------------------------------------------------

Sub Form_Activate()
   If mi_Opcion <> sghAgregar Then
       If Not mb_ExistenDatos Then
           Me.Visible = False
           LimpiarVariablesDeMemoria
       End If
   Else
        If CDate(txtFechaIni.Text) < lcBuscaParametro.RetornaFechaServidorSQL Then
          ' MsgBox "S?lo puede programar Fechas mayores a " & lcBuscaParametro.RetornaFechaServidorSQL, vbInformation, Me.Caption
           Dim oMensaje As New SIGHNegocios.clMensaje
           oMensaje.MostrarFormulario "S?lo puede programar Fechas mayores a " & lcBuscaParametro.RetornaFechaServidorSQL, Me.Caption
           Set oMensaje = Nothing
           
           
           Me.Visible = False
           LimpiarVariablesDeMemoria
        End If
   End If
End Sub
Sub AdministrarKeyPreview(KeyCode As Integer)
   Select Case KeyCode
       Case vbKeyEscape
       Case vbKeyF2
           btnAceptar_Click
       End Select
End Sub
Private Sub btnAceptar_Click()
   If btnAceptar.Enabled = False Then
      Exit Sub
   End If
   Select Case mi_Opcion
   Case sghAgregar
       If ValidarDatosObligatorios() Then
           If ValidarReglas() Then
               If AgregarDatos() Then
               End If
           End If
       End If
   Case sghModificar
       If ValidarDatosObligatorios() Then
           If ValidarReglas() Then
               If ModificarDatos() Then
               End If
           End If
       End If
   Case sghEliminar
            Dim oMensaje2 As New SIGHNegocios.clMensaje
            oMensaje2.MostrarFormulario Chr(13) & "Esta seguro ?", Me.Caption, 20, , , True
            If oMensaje2.BotonPresionado = sghAceptar Then
                If EliminarDatos() Then
                End If
            End If
            Set oMensaje2 = Nothing
   End Select
End Sub

Private Sub btnCancelar_Click()
    mb_SeHaModificadoProgramacion = False
    Me.Visible = False
    LimpiarVariablesDeMemoria
End Sub

Function ValidarDatosObligatorios() As Boolean
Dim sMensaje As String

   ValidarDatosObligatorios = False
   
    
    If Me.txtHoraInicio.Text = SIGHEntidades.HORA_VACIA_HM Then
       sMensaje = sMensaje + "Ingrese la hora inicio" + Chr(13)
    Else
        If Not SIGHEntidades.EsHora(txtHoraInicio) Then
            sMensaje = sMensaje + "La hora de inicio no es correcta" + Chr(13)
        End If
    End If
    If Me.txtHoraFin.Text = SIGHEntidades.HORA_VACIA_HM Then
       sMensaje = sMensaje + "Ingrese la hora fin" + Chr(13)
    Else
        If Not SIGHEntidades.EsHora(txtHoraFin) Then
            sMensaje = sMensaje + "La hora final no es correcta" + Chr(13)
        End If
   End If
   If Me.cmbIdServicio.Text = "" Then
      sMensaje = sMensaje + "Elija el PUNTO DE CARGA" + Chr(13)
   End If
   
   If mo_cmbIdTurno.BoundText = "" Then
       sMensaje = sMensaje + "Ingrese el turno" + Chr(13)
   End If
   If sMensaje <> "" Then
       Dim oMensaje As New SIGHNegocios.clMensaje
       oMensaje.MostrarFormulario sMensaje, Me.Caption
       Set oMensaje = Nothing
       
       Exit Function
   End If
   ValidarDatosObligatorios = True
End Function
Function ValidarReglas() As Boolean
   Dim sMensaje As String
   ValidarReglas = False
   If Not SIGHEntidades.ValidaHora(Me.txtHoraInicio) Then
        Dim oMensaje As New SIGHNegocios.clMensaje
        oMensaje.MostrarFormulario "La hora inicial ingresada no es v?lida", Me.Caption
        Set oMensaje = Nothing
        
        Exit Function
   End If
   
   If Not SIGHEntidades.ValidaHora(Me.txtHoraFin) Then
        Dim oMensaje0 As New SIGHNegocios.clMensaje
        oMensaje0.MostrarFormulario "La hora final ingresada no es v?lida", Me.Caption
        Set oMensaje0 = Nothing
        
        Exit Function
   End If
   
   If mi_Opcion <> sghAgregar Then
       If -DateDiff("h", CDate(Date & " " & Me.txtHoraInicio), CDate(Date & " " & txtHoraFin)) > 0 Then
            Dim oMensaje1 As New SIGHNegocios.clMensaje
            oMensaje1.MostrarFormulario "La hora de salida no puede ser menor a la hora de ingreso", Me.Caption
            Set oMensaje1 = Nothing
            
            Exit Function
        End If
    End If
    
    If validarProgramacionMedicoEnOtrosServicios(Val(mo_cmbIdServicio.BoundText), Me.idMedico, _
                    txtFechaIni.Text, txtHoraInicio.Text, txtHoraFin.Text) = False Then
        Exit Function
    End If
    '
    Dim oRsTmp As New Recordset
    Dim oMensaje2 As New SIGHNegocios.clMensaje
    
    Dim dFechaInicio As Date, dFechaFin As Date
    Dim dFechaInicioProgramacion As Date, dFechaFinProgracion As Date
    
    dFechaInicioProgramacion = getFechaInicio(CDate(txtFechaIni.Text), Me.txtHoraInicio.Text)
    dFechaFinProgracion = getFechaFin(CDate(txtFechaIni.Text), Me.txtHoraInicio.Text, Me.txtHoraFin.Text)
    'debb-27/05/2015
    If mi_Opcion = sghAgregar Then
        Set oRsTmp = mo_ReglasAdmision.ServiciosAtenSimultaneaSeleccionarXServicioAtencionSimultanea(Val(mo_cmbIdServicio.BoundText))
        If oRsTmp.RecordCount > 0 Then
            ms_MensajeError = "El CONSULTORIO elegido NO SE PUEDE PROGRAMAR, porque es un CONSULTORIO QUE SE ATIENDE EN FORMA SIMULTANEA en: " & _
                   Chr(13) & oRsTmp!nombre & Chr(13) & _
                   "verifique en opci?n GENERAL->SERVICIOS->CONSULTORIOS EXTERNOS->" & oRsTmp!nombre
            oMensaje2.MostrarFormulario ms_MensajeError, Me.Caption
            Set oMensaje2 = Nothing
            Exit Function
        End If
        oRsTmp.Close
    End If
    '
    If ExisteProgramacionParaEseServicioFecha(Val(mo_cmbIdServicio.BoundText), CDate(txtFechaIni.Text), _
                                              Me.txtHoraInicio.Text, Me.txtHoraFin.Text) = True Then
       Exit Function
    End If
    '
    sMensaje = HayCruceDeProgramacionesDeConsultaExterna()
    If sMensaje <> "" Then
         Dim oMensaje22 As New SIGHNegocios.clMensaje
         oMensaje22.MostrarFormulario "Hay un cruces con la(s) siguiente(s) programacion(es) de consulta externa " & Chr(13) & sMensaje, Me.Caption
         Set oMensaje22 = Nothing
         
         Exit Function
    End If

     ValidarReglas = True
End Function
Function HayCruceDeProgramacionesDeConsultaExterna() As String
Dim daDiaSeleccionado As Date
Dim programacion As PVAppointment
Dim sTitulo As String
Dim sHoras() As String
Dim iHoraIni As Integer
Dim iHoraFin As Integer
Dim bTurnoProgramado As Boolean
Dim oDOProgramacion As New DOSiProgramacion
Dim sMensaje  As String

    sMensaje = ""
    daDiaSeleccionado = Me.Calendario.Value
    Do While daDiaSeleccionado <> 0
        Set programacion = Me.Diario.AppointmentSet.Get(daDiaSeleccionado)
        bTurnoProgramado = False
        Do While Not programacion Is Nothing
            'Verifica que la programacion sea del mismo dia
             Set oDOProgramacion = programacion.DataVariant
            'Verifica si la nueva porg es de CE
                If Me.IdProgramacion <> oDOProgramacion.IdProgramacion Then
                    'Verifica si la prog existente tambien es de CE
                        If oDOProgramacion.fecha = daDiaSeleccionado Then
                            'Las validaciones se hacen solo si la hora inicio nueva es menor que la hora de fin existente
                            If ConvertirAMinutos(Me.txtHoraInicio) < ConvertirAMinutos(oDOProgramacion.HoraFin) Then
                                If (ConvertirAMinutos(Me.txtHoraInicio) <= ConvertirAMinutos(oDOProgramacion.HoraInicio) And ConvertirAMinutos(Me.txtHoraFin) >= ConvertirAMinutos(oDOProgramacion.HoraInicio)) Or _
                                    (ConvertirAMinutos(Me.txtHoraInicio) >= ConvertirAMinutos(oDOProgramacion.HoraInicio) And ConvertirAMinutos(Me.txtHoraFin) <= ConvertirAMinutos(oDOProgramacion.HoraFin)) Or _
                                    (ConvertirAMinutos(Me.txtHoraInicio) <= ConvertirAMinutos(oDOProgramacion.HoraFin) And ConvertirAMinutos(Me.txtHoraFin) >= ConvertirAMinutos(oDOProgramacion.HoraFin)) Then
                                    sMensaje = sMensaje & programacion.Description & " " & oDOProgramacion.fecha & " Hora Inicio " & oDOProgramacion.HoraInicio & " Hora Fin " & oDOProgramacion.HoraFin & Chr(13)
                                End If
                            End If
                        Else
                            Exit Do
                        End If
                End If
                
            Set programacion = Diario.AppointmentSet.GetNext(programacion)
        Loop
        daDiaSeleccionado = Me.Calendario.NextSelectedDate(daDiaSeleccionado)
    Loop

    HayCruceDeProgramacionesDeConsultaExterna = sMensaje

End Function
Function ConvertirAMinutos(sHora As String) As Integer
Dim sHoras() As String

    sHoras = Split(sHora, ":")
    ConvertirAMinutos = Val(sHoras(0)) * 60 + Val(sHoras(1))

End Function
'------------------------------------------------------------------------------------
'   Cargar datos al objetos de datos
'   Descripci?n:    Seleccionar un registro unico de la tabla ProgramacionMedica
'   Par?metros:     Ninguno
'------------------------------------------------------------------------------------

Sub CargaDatosAlObjetosDeDatos()

   
End Sub

'------------------------------------------------------------------------------------
'        Agregar Datos
'------------------------------------------------------------------------------------

Function AgregarDatos() As Boolean
Dim daDiaSeleccionado As Date
Dim programacion As PVAppointment
Dim sTitulo As String
Dim sHoras() As String
Dim iHoraIni As Integer
Dim iHoraFin As Integer
Dim iHoraIni1 As Double
Dim iHoraFin1 As Double
Dim bTurnoProgramado As Boolean
Dim doProgramacion As DOSiProgramacion
Dim programacionAdicional As PVAppointment
Dim daDiaSeleccionadoFantasma As Date
Dim bAgregarProgFicticio As Boolean
Dim sHoraFin As String
Dim oCollProgramaciones As New Collection
Dim ldFechaHoraServidor As Date
Dim oRsTmp As New Recordset
Dim lnTiempoPromedioAtencion As Long, ldFechaSeleccionada As Date
    ldFechaHoraServidor = lcBuscaParametro.RetornaFechaHoraServidorSQL
    mb_SeHaModificadoProgramacion = False
    daDiaSeleccionado = mo_Calendario.Value
    '
    lnTiempoPromedioAtencion = 0
    Set oRsTmp = mo_AdminServiciosComunes.FactPuntosCargaSeleccionarPorId(ml_IdPuntoCarga)
    If oRsTmp.RecordCount > 0 Then
        If Not IsNull(oRsTmp.Fields!nroCuposMinutos) Then
            lnTiempoPromedioAtencion = oRsTmp.Fields!nroCuposMinutos
        End If
    End If
    oRsTmp.Close
    '
    Do While daDiaSeleccionado <> 0
        ldFechaSeleccionada = daDiaSeleccionado
        If ExisteProgramacionParaEseServicioFecha(Val(mo_cmbIdServicio.BoundText), ldFechaSeleccionada, _
                                                  Me.txtHoraInicio.Text, Me.txtHoraFin.Text) = False Then
    
            sTitulo = ""
            Set programacion = mo_Diario.AppointmentSet.Get(daDiaSeleccionado)
            bTurnoProgramado = False
            
            'Agrega programacion
            sHoras = Split(Me.txtHoraInicio, ":")
            iHoraIni = Val(sHoras(0)) + Val(sHoras(1)) / 60
            iHoraIni1 = Val(sHoras(0)) + Val(sHoras(1)) / 60
            
            sHoras = Split(Me.txtHoraFin, ":")
            iHoraFin = Val(sHoras(0)) + Val(sHoras(1)) / 60
            iHoraFin1 = Val(sHoras(0)) + Val(sHoras(1)) / 60
            
            If iHoraIni1 < iHoraFin1 Then
                Set programacion = mo_Diario.AppointmentSet.Add(Me.cmbIdTurno.Text, daDiaSeleccionado + iHoraIni / 24, daDiaSeleccionado + iHoraFin / 24)
                On Error Resume Next
                sHoraFin = Me.txtHoraFin
                bAgregarProgFicticio = False
            Else
                iHoraFin = 23 + 59 / 60
                sHoraFin = "23:59"
                Set programacion = mo_Diario.AppointmentSet.Add(Me.cmbIdTurno.Text, daDiaSeleccionado + iHoraIni / 24, daDiaSeleccionado + iHoraFin / 24)
                bAgregarProgFicticio = True
            End If
            mo_Calendario.DATEText(daDiaSeleccionado) = IIf(mo_Calendario.DATEText(daDiaSeleccionado) = "", SIGHEntidades.ExtraerCadena(Me.cmbIdTurno.Text, 1, "="), mo_Calendario.DATEText(daDiaSeleccionado) + "," + SIGHEntidades.ExtraerCadena(Me.cmbIdTurno.Text, 1, "="))
            Set doProgramacion = New DOSiProgramacion
            'Datos de la programacion
            
            doProgramacion.fecha = Format(daDiaSeleccionado, SIGHEntidades.DevuelveFechaSoloFormato_DMY)
            doProgramacion.HoraInicio = Me.txtHoraInicio
            doProgramacion.HoraFin = sHoraFin
            doProgramacion.IdTurno = Val(mo_cmbIdTurno.BoundText)
            doProgramacion.IdUsuarioAuditoria = ml_idUsuario
            doProgramacion.idResponsable = Me.idMedico
            doProgramacion.IdProgramacion = 0
            doProgramacion.idSala = Val(mo_cmbIdServicio.BoundText)
            doProgramacion.FechaReg = ldFechaHoraServidor
            doProgramacion.TiempoPromedioAtencion = lnTiempoPromedioAtencion
            
            oCollProgramaciones.Add doProgramacion
            
            programacion.DataVariant = doProgramacion
        
            'Para los turnos que continuan ma?ana
            If bAgregarProgFicticio Then
                iHoraIni = 0 + 0 / 60
                iHoraFin = Val(sHoras(0)) + Val(sHoras(1)) / 60
                daDiaSeleccionadoFantasma = DateAdd("d", 1, daDiaSeleccionado)
                
                Set doProgramacion = New DOProgramacionMedica
                doProgramacion.fecha = Format(daDiaSeleccionadoFantasma, SIGHEntidades.DevuelveFechaSoloFormato_DMY)
                doProgramacion.HoraInicio = "00:00"
                doProgramacion.HoraFin = Me.txtHoraFin
                doProgramacion.IdTurno = Val(mo_cmbIdTurno.BoundText)
                doProgramacion.IdUsuarioAuditoria = ml_idUsuario
                doProgramacion.idResponsable = Me.idMedico
                doProgramacion.IdProgramacion = 0
                doProgramacion.FechaReg = ldFechaHoraServidor
                doProgramacion.idSala = Val(mo_cmbIdServicio.BoundText)
                
                oCollProgramaciones.Add doProgramacion
                
                Set programacionAdicional = mo_Diario.AppointmentSet.Add(Me.cmbIdTurno.Text, daDiaSeleccionadoFantasma + iHoraIni / 24, daDiaSeleccionadoFantasma + iHoraFin / 24)
                
                programacionAdicional.DataVariant = doProgramacion
                mo_Calendario.DATEText(daDiaSeleccionadoFantasma) = IIf(mo_Calendario.DATEText(daDiaSeleccionadoFantasma) = "", SIGHEntidades.ExtraerCadena(Me.cmbIdTurno.Text, 1, "="), mo_Calendario.DATEText(daDiaSeleccionadoFantasma) + "," + SIGHEntidades.ExtraerCadena(Me.cmbIdTurno.Text, 1, "="))
            End If
        End If
        daDiaSeleccionado = mo_Calendario.NextSelectedDate(daDiaSeleccionado)
        mb_SeHaModificadoProgramacion = True
    Loop


    
    AgregarDatos = mo_ReglasImagenes.SiProgramacionAgregar(oCollProgramaciones, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, Trim(txtMedico.Text) & " " & Trim(cmbIdServicio.Text) & txtFechaIni.Text)

    Me.Visible = False
    LimpiarVariablesDeMemoria
    Set oRsTmp = Nothing
End Function

'------------------------------------------------------------------------------------
'        Modificar Datos
'------------------------------------------------------------------------------------

Function ModificarDatos() As Boolean
Dim iHoraIni As Integer
Dim iHoraFin As Integer
Dim iHoras() As Integer
Dim sHoras() As String
Dim programacion As PVAppointment
Dim lnTiempoPromedioAtencion As Long
Dim oRsTmp As New Recordset
    '
    lnTiempoPromedioAtencion = 0
    Set oRsTmp = mo_AdminServiciosComunes.FactPuntosCargaSeleccionarPorId(ml_IdPuntoCarga)
    If oRsTmp.RecordCount > 0 Then
        If Not IsNull(oRsTmp.Fields!nroCuposMinutos) Then
            lnTiempoPromedioAtencion = oRsTmp.Fields!nroCuposMinutos
        End If
    End If
    oRsTmp.Close
    '

    mb_SeHaModificadoProgramacion = True
    
    sHoras = Split(Me.txtHoraInicio.Text, ":")
    iHoraIni = CInt(Val(sHoras(0)) + Val(sHoras(1)) / 60)
    
    sHoras = Split(Me.txtHoraFin, ":")
    iHoraFin = CInt(Val(sHoras(0)) + Val(sHoras(1)) / 60)
    
    Set programacion = mo_Diario.AppointmentSet.GetSelectedAppointment
    On Error Resume Next
    mo_Diario.AppointmentSet.Remove programacion.Key
    
    Set programacion = mo_Diario.AppointmentSet.Add(Me.cmbIdTurno.Text, mo_Diario.CurrentDate + iHoraIni / 24, mo_Diario.CurrentDate + iHoraFin / 24)
    
    Dim doProgramacion As New DOSiProgramacion
    
    'Datos de la programacion
    doProgramacion.IdProgramacion = Me.IdProgramacion
    If Me.txtFechaIni.Enabled = True Then
       doProgramacion.fecha = Format(programacion.StartDateTime, SIGHEntidades.DevuelveFechaSoloFormato_DMY)
    Else
       doProgramacion.fecha = Me.txtFechaIni.Text
    End If
    doProgramacion.HoraInicio = Me.txtHoraInicio
    doProgramacion.HoraFin = Me.txtHoraFin
    doProgramacion.IdTurno = mo_cmbIdTurno.BoundText
    doProgramacion.IdUsuarioAuditoria = ml_idUsuario
    doProgramacion.idResponsable = Me.idMedico
    doProgramacion.idSala = Val(mo_cmbIdServicio.BoundText)
    doProgramacion.TiempoPromedioAtencion = lnTiempoPromedioAtencion
    doProgramacion.FechaReg = CDate(lcBuscaParametro.RetornaFechaHoraServidorSQL)
        
    programacion.DataVariant = doProgramacion

    mo_Calendario.DATEText(mo_Diario.CurrentDate) = ObtieneCodigoDeProgramacionPorDia(mo_Diario.CurrentDate)
    Me.Visible = False
    LimpiarVariablesDeMemoria

    ModificarDatos = mo_ReglasImagenes.SiProgramacionModificar(doProgramacion, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, Trim(txtMedico.Text) & " " & Trim(cmbIdServicio.Text) & txtFechaIni.Text)
    Set oRsTmp = Nothing
End Function

'------------------------------------------------------------------------------------
'        Eliminar Datos
'------------------------------------------------------------------------------------

Function EliminarDatos() As Boolean
Dim programacion As PVAppointment
Dim oDOProgramacion As New DOSiProgramacion
Dim oProgMedicas As New Collection
    
    mb_SeHaModificadoProgramacion = True
    
    Set programacion = mo_Diario.AppointmentSet.GetSelectedAppointment
    mo_Diario.AppointmentSet.Remove programacion.Key

    mo_Calendario.DATEText(mo_Diario.CurrentDate) = ObtieneCodigoDeProgramacionPorDia(mo_Diario.CurrentDate)

    'Agrega la unica programacion medica a eliminar
    
    EliminarDatos = mo_ReglasImagenes.SiProgramacionEliminar(programacion.DataVariant, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, Trim(txtMedico.Text) & " " & Trim(cmbIdServicio.Text) & txtFechaIni.Text)
    
    Me.Visible = False
    LimpiarVariablesDeMemoria

End Function


'------------------------------------------------------------------------------------
'   Llenar Datos Al Formulario
'   Descripci?n:    Seleccionar un registro unico de la tabla ProgramacionMedica
'   Par?metros:     Ninguno
'------------------------------------------------------------------------------------

Sub CargarDatosAlosControles()
    Dim doProgramacion As DOSiProgramacion
    Dim oRsTmp1 As New Recordset
    Dim programacion As PVAppointment
    Set programacion = mo_Diario.AppointmentSet.GetSelectedAppointment
    
    If programacion Is Nothing Then
        'MsgBox "Por favor seleccione la programaci?n que desea modificar o eliminar", vbInformation, Me.Caption
        Dim oMensaje2 As New SIGHNegocios.clMensaje
        oMensaje2.MostrarFormulario "Por favor seleccione la programaci?n que desea modificar o eliminar", Me.Caption
        Set oMensaje2 = Nothing
        
        Exit Sub
    End If
    
    Set doProgramacion = programacion.DataVariant
    Me.IdProgramacion = doProgramacion.IdProgramacion
    
    If lbSoloCambiaHoraFinal = False And (mi_Opcion = sghModificar Or mi_Opcion = sghEliminar) Then
        Set oRsTmp1 = mo_ReglasImagenes.siCitasXidprogramacion(doProgramacion.IdProgramacion)
'        Set oRsTmp1 = mo_AdminProgMedica.CitasSeleccionarPorServicioYfecha(doProgramacion.idSala, doProgramacion.fecha)
        If oRsTmp1.RecordCount > 0 Then
           oRsTmp1.MoveFirst
           Do While Not oRsTmp1.EOF
              If oRsTmp1.Fields!IdProgramacion = doProgramacion.IdProgramacion Then
                 Dim oMensaje As New SIGHNegocios.clMensaje
                 oMensaje.MostrarFormulario "No se puede modificar o eliminar porque ya existe un " & _
                                            "Paciente con CITA", Me.Caption
                 Set oMensaje = Nothing
                 '
                 Set oRsTmp1 = Nothing
                 btnCancelar_Click
                 Exit Sub
              End If
              oRsTmp1.MoveNext
           Loop
        End If
    End If
    
    Me.txtMedico = Me.NombreMedico

    Me.txtFechaIni = Format(doProgramacion.fecha, SIGHEntidades.DevuelveFechaSoloFormato_DMY)
    Me.txtFechaFin.Visible = False
    
    mo_cmbIdServicio.BoundText = doProgramacion.idSala
    
    mo_cmbIdTurno.BoundText = doProgramacion.IdTurno
    Me.txtHoraInicio = doProgramacion.HoraInicio
    Me.txtHoraFin = doProgramacion.HoraFin
    ldFechaRegistro = doProgramacion.FechaReg
    mb_ExistenDatos = True
   
    Set oRsTmp1 = Nothing
End Sub

'------------------------------------------------------------------------------------
'   Llenar Datos Al Formulario
'   Descripci?n:    Seleccionar un registro unico de la tabla ProgramacionMedica
'   Par?metros:     Ninguno
'------------------------------------------------------------------------------------

Sub LimpiarFormulario()

           Me.IdProgramacion = 0
           Me.idMedico = 0
           Me.NroCuposCE = 0
   
End Sub

Function ObtieneCodigoDeProgramacionPorDia(daDiaSeleccionado As Date)
Dim sTitulo As String
Dim programacion As PVAppointment
Dim oDOProgramacion As New DOSiProgramacion
Dim oDOTurno As New doTurno


        sTitulo = ""
        Set programacion = Diario.AppointmentSet.Get(daDiaSeleccionado)
        Do While Not programacion Is Nothing
            If Format(programacion.StartDateTime, SIGHEntidades.DevuelveFechaSoloFormato_DMY) <> daDiaSeleccionado Then
                Exit Do
            End If
            
            Set oDOProgramacion = programacion.DataVariant
            Set oDOTurno = mo_AdminProgMedica.TurnosSeleccionarPorId(oDOProgramacion.IdTurno)
            
            sTitulo = sTitulo + oDOTurno.Codigo + Chr(13)
            Set programacion = Diario.AppointmentSet.GetNext(programacion)
        Loop
        If sTitulo <> "" Then
            ObtieneCodigoDeProgramacionPorDia = Left(sTitulo, Len(sTitulo) - 1)
        End If


End Function



Sub LimpiarVariablesDeMemoria()
    On Error Resume Next
    Set mo_Teclado = Nothing
    Set mo_Formulario = Nothing
    Set mo_AdminServiciosComunes = Nothing
    Set mo_AdminServiciosHosp = Nothing
    Set mo_AdminProgMedica = Nothing
    Set mo_FechaHora = Nothing
    
    Set mo_cmbIdTurno = Nothing
    Set mo_cmbIdServicio = Nothing
    Set lcBuscaParametro = Nothing
End Sub


'actualizado 26102014 yamill
Private Function getFechaInicio(sFecha As Date, sHoraInicio As String) As Date
    getFechaInicio = Format(sFecha, SIGHEntidades.DevuelveFechaSoloFormato_DMY) & " " & sHoraInicio
End Function


Private Function getFechaFin(sFecha As Date, sHoraInicio As String, sHoraFin As String) As Date
    Dim ldFechaFinal As Date
    sHoraInicio = Format(sHoraInicio, "hh:mm")
    sHoraFin = Format(sHoraFin, "hh:mm")
    If sHoraFin > sHoraInicio Then
        getFechaFin = Format(sFecha, SIGHEntidades.DevuelveFechaSoloFormato_DMY) & " " & sHoraFin
    Else
'        sFecha = DateAdd("d", 1, sFecha)
'        getFechaFin = Format(sFecha, sighentidades.DevuelveFechaSoloFormato_DMY) & " " & sHoraFin
        ldFechaFinal = DateAdd("d", 1, sFecha)
        getFechaFin = Format(ldFechaFinal, SIGHEntidades.DevuelveFechaSoloFormato_DMY) & " " & sHoraFin
        
    End If
End Function

'Actualizado 30102014 yamill palomino
Private Function validarProgramacionMedicoEnOtrosServicios(lIdServicio As Long, lIDMedico As Long, _
                    sFecha As String, sHoraInicio As String, sHoraFin As String) As Boolean
                    
    Dim oRsTmp As ADODB.Recordset
    Dim dFechaInicio As Date, dFechaFin As Date
    Dim dFechaInicioProgramacion As Date, dFechaFinProgracion As Date
    Dim oDOProgramacionMedica As New DOSiProgramacion
    Dim returnValue As Boolean
   
    
    oDOProgramacionMedica.idSala = lIdServicio
    oDOProgramacionMedica.idResponsable = lIDMedico
    oDOProgramacionMedica.fecha = sFecha
    Dim oMensaje2 As New SIGHNegocios.clMensaje
    
    validarProgramacionMedicoEnOtrosServicios = True
    
    dFechaInicioProgramacion = getFechaInicio(CDate(sFecha), sHoraInicio)
    dFechaFinProgracion = getFechaFin(CDate(sFecha), sHoraInicio, sHoraFin)
    
    Set oRsTmp = mo_ReglasImagenes.SiProgramacionXresponsableYfechaYsala(oDOProgramacionMedica)
    If oRsTmp.RecordCount > 0 Then
       oRsTmp.MoveFirst
       Do While Not oRsTmp.EOF
          If oRsTmp.Fields!fecha = CDate(sFecha) Then
                
                dFechaInicio = getFechaInicio(oRsTmp.Fields!fecha, oRsTmp.Fields!HoraInicio)
                dFechaFin = getFechaFin(oRsTmp.Fields!fecha, oRsTmp.Fields!HoraInicio, oRsTmp.Fields!HoraFin)

                If (dFechaInicio <= dFechaInicioProgramacion And dFechaFin >= dFechaFinProgracion) _
                    Or (dFechaInicioProgramacion <= dFechaInicio And dFechaFinProgracion >= dFechaFin) _
                    Or (dFechaInicioProgramacion > dFechaInicio And dFechaInicioProgramacion < dFechaFin) _
                    Or (dFechaFinProgracion > dFechaInicio And dFechaFinProgracion < dFechaFin) Then
                    If mi_Opcion = sghAgregar Then
                        oMensaje2.MostrarFormulario "RESPONSABLE fu? programado en otra SALA para esa FECHA/HORA", Me.Caption
                        Set oMensaje2 = Nothing
                        validarProgramacionMedicoEnOtrosServicios = False
                        Exit Function
                    ElseIf mi_Opcion = sghModificar And oRsTmp!IdProgramacion <> ml_IdProgramacion Then
                        oMensaje2.MostrarFormulario "RESPONSABLE fu? programado en otra SALA para esa FECHA/HORA", Me.Caption
                        Set oMensaje2 = Nothing
                        validarProgramacionMedicoEnOtrosServicios = False
                        Exit Function
                    End If
                End If
          End If
          oRsTmp.MoveNext
       Loop
    End If
    oRsTmp.Close
    Set oRsTmp = Nothing
End Function

Function ExisteProgramacionParaEseServicioFecha(lnIdSala As Long, lcFecha As Date, _
                                                lcHoraInicio As String, lcHoraFin As String) As Boolean
    Dim oRsTmp As New Recordset
    Dim oMensaje2 As New SIGHNegocios.clMensaje
    Dim dFechaInicio As Date, dFechaFin As Date
    Dim dFechaInicioProgramacion As Date, dFechaFinProgracion As Date
    
    dFechaInicioProgramacion = getFechaInicio(lcFecha, lcHoraInicio)
    dFechaFinProgracion = getFechaFin(lcFecha, lcHoraInicio, lcHoraFin)
                                                
                                                
    Set oRsTmp = mo_ReglasImagenes.siProgramacionXsalaYfecha(lnIdSala, lcFecha)
    If oRsTmp.RecordCount > 0 Then
       oRsTmp.MoveFirst
       Do While Not oRsTmp.EOF
          If oRsTmp.Fields!fecha = CDate(lcFecha) Then
                'Actualizado 27102014 yamill palomino
                dFechaInicio = getFechaInicio(oRsTmp.Fields!fecha, oRsTmp.Fields!HoraInicio)
                dFechaFin = getFechaFin(oRsTmp.Fields!fecha, oRsTmp.Fields!HoraInicio, oRsTmp.Fields!HoraFin)
                

                If (dFechaInicio <= dFechaInicioProgramacion And dFechaFin >= dFechaFinProgracion) _
                    Or (dFechaInicioProgramacion <= dFechaInicio And dFechaFinProgracion >= dFechaFin) _
                    Or (dFechaInicioProgramacion > dFechaInicio And dFechaInicioProgramacion < dFechaFin) _
                    Or (dFechaFinProgracion > dFechaInicio And dFechaFinProgracion < dFechaFin) Then ' Actualizado 16/10/2014 Yamill palomino
                    If mi_Opcion = sghAgregar Then
                        oMensaje2.MostrarFormulario "El PUNTO DE CARGA elegido ya fu? programado para esa FECHA/HORA, al RESPONSABLE: " & Trim(oRsTmp.Fields!ApellidoPaterno) & " " & Trim(oRsTmp.Fields!ApellidoMaterno) & " " & Trim(oRsTmp.Fields!Nombres), Me.Caption
                        Set oMensaje2 = Nothing
                        ExisteProgramacionParaEseServicioFecha = True
                        Exit Function
                    ElseIf mi_Opcion = sghModificar And oRsTmp!IdProgramacion <> ml_IdProgramacion Then
                        oMensaje2.MostrarFormulario "El PUNTO DE CARGA elegido ya fu? programado para esa FECHA/HORA, al RESPONSABLEo: " & Trim(oRsTmp.Fields!ApellidoPaterno) & " " & Trim(oRsTmp.Fields!ApellidoMaterno) & " " & Trim(oRsTmp.Fields!Nombres), Me.Caption
                        ExisteProgramacionParaEseServicioFecha = True
                        Set oMensaje2 = Nothing
                        Exit Function
                    End If
                End If
          End If
          oRsTmp.MoveNext
       Loop
    End If
    oRsTmp.Close
    Set oRsTmp = Nothing
    Set oMensaje2 = Nothing


End Function


