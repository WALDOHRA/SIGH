VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.UserControl UcPacienteDatos 
   ClientHeight    =   2895
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4425
   ScaleHeight     =   2895
   ScaleWidth      =   4425
   Begin VB.Frame fraDatosPaciente 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2850
      Left            =   0
      TabIndex        =   17
      Top             =   0
      Width           =   4395
      Begin VB.TextBox txtEdad 
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
         Left            =   3780
         MaxLength       =   20
         TabIndex        =   16
         Top             =   2400
         Width           =   465
      End
      Begin VB.ComboBox cmbIdTipoGenHistoriaClinica 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2640
         TabIndex        =   2
         Top             =   210
         Width           =   1635
      End
      Begin VB.CommandButton cmdBuscaPacientePorApellidos 
         Caption         =   "..."
         Height          =   300
         Left            =   3930
         TabIndex        =   11
         ToolTipText     =   "Busca Datos personales por Apellidos y Nombres"
         Top             =   1695
         Width           =   315
      End
      Begin VB.TextBox txtIdNroHistoria 
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
         Left            =   1650
         MaxLength       =   35
         TabIndex        =   1
         Top             =   210
         Width           =   945
      End
      Begin VB.TextBox txtPrimerNombre 
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
         Left            =   1650
         MaxLength       =   20
         TabIndex        =   8
         Top             =   1320
         Width           =   2595
      End
      Begin VB.TextBox txtApellidoPaterno 
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
         Left            =   1650
         MaxLength       =   20
         TabIndex        =   4
         Top             =   580
         Width           =   2595
      End
      Begin VB.TextBox txtApellidoMaterno 
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
         Left            =   1650
         MaxLength       =   20
         TabIndex        =   6
         Top             =   950
         Width           =   2595
      End
      Begin VB.ComboBox cmbIdTipoSexo 
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
         Left            =   1650
         TabIndex        =   13
         Top             =   2055
         Width           =   2625
      End
      Begin MSMask.MaskEdBox txtFechaNacimiento 
         Height          =   315
         Left            =   1650
         TabIndex        =   15
         Top             =   2415
         Width           =   1350
         _ExtentX        =   2381
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.TextBox txtSegundoNombre 
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
         Left            =   1650
         MaxLength       =   20
         TabIndex        =   10
         Top             =   1680
         Width           =   2595
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Segundo Nombre"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   120
         TabIndex        =   9
         Top             =   1710
         Width           =   1440
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Nro Historia "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   150
         TabIndex        =   0
         Top             =   240
         Width           =   990
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Edad"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   3300
         TabIndex        =   19
         Top             =   2445
         Width           =   405
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Apellido &Paterno"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   150
         TabIndex        =   3
         Top             =   610
         Width           =   1335
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "&Primer Nombre"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   150
         TabIndex        =   7
         Top             =   1350
         Width           =   1215
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Apellido &Materno"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   150
         TabIndex        =   5
         Top             =   960
         Width           =   1365
      End
      Begin VB.Label Label17 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "&Fecha Nacimiento"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   150
         TabIndex        =   14
         Top             =   2445
         Width           =   1440
      End
      Begin VB.Label Label30 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Sexo"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   150
         TabIndex        =   12
         Top             =   2085
         Width           =   405
      End
      Begin VB.Label Label7 
         Caption         =   "Datos del paciente"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   135
         TabIndex        =   18
         Top             =   0
         Width           =   1545
      End
   End
End
Attribute VB_Name = "UcPacienteDatos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Dim mo_Teclado As New SIGHComun.Teclado
Dim mo_Formulario As New SIGHComun.Formulario
Dim ms_MensajeError As String
Dim mi_Opcion As sghOpciones
Dim ml_idUsuario As Long
Dim mb_ExistenDatos As Boolean
Dim mo_AdminAdmision As New SIGHNegocios.ReglasAdmision
Dim mo_AdminServiciosComunes As New SIGHNegocios.ReglasComunes
Dim mo_AdminServiciosGeograficos As New SIGHNegocios.ReglasServGeograf
Dim mo_AdminArchivoClinico As New SIGHNegocios.ReglasArchivoClinico
Dim mo_AdminFacturacion As New ReglasFacturacion
Dim mo_AdminServiciosHosp As New SIGHNegocios.ReglasServiciosHosp
Dim mo_AdminProgramacion As New SIGHNegocios.ReglasDeProgMedica
Dim ml_TipoServicio As sghTipoServicio
Dim mo_AdminReportes As New SIGHNegocios.ReglasReportes
Dim mo_AdminHoteleria As New SIGHNegocios.ReglasHoteleria
Dim mrs_Diagnosticos As New ADODB.Recordset
Dim mo_Apariencia As New SIGHComun.GridInfragistic
Dim ml_TipoVistaForm As sghTipoVistaFormAtenciones
Dim mb_PacienteNoIdentificado As Boolean
Public Event SeModificoFechaNacimiento(sFechaNacimiento As String)
Public Event SePresionoTeclaEspecial(KeyCode As Integer)
Public Event SeModificoSexo(lIdTipoSexo As Long)

Dim mo_cmbIdTipoGenHistoriaClinica As New SIGHComun.ListaDespleglable
Dim mo_CmbIdTipoSexo As New SIGHComun.ListaDespleglable

'------------------------------------------------------------------------------------
'                               VARIABLE PARA LA FILIACION
'------------------------------------------------------------------------------------
Dim ml_idPaciente As Long
Dim ms_Autogenerado As String
Dim lcBuscaParametro As New SIGHDatos.Parametros
Dim ml_FechaRegistro As Date

Property Let FechaRegistro(lValue As Date)
   ml_FechaRegistro = lValue
End Property
Property Let idPaciente(lValue As Long)
   ml_idPaciente = lValue
End Property
Property Get idPaciente() As Long
   idPaciente = ml_idPaciente
End Property

Property Let Autogenerado(sValue As String)
   ms_Autogenerado = sValue
End Property
Property Get Autogenerado() As String
   Autogenerado = ms_Autogenerado
End Property
Property Let FechaNacimiento(sValue As String)
   txtFechaNacimiento.Text = sValue
End Property
Property Get FechaNacimiento() As String
   FechaNacimiento = txtFechaNacimiento.Text
End Property

Property Let NroHistoriaClinica(lValue As Long)
   txtIdNroHistoria.Text = CStr(lValue)
End Property
Property Get NroHistoriaClinica() As Long
   NroHistoriaClinica = Val(txtIdNroHistoria.Text)
End Property

Property Get ExistePaciente() As Boolean
   ExistePaciente = mb_ExistenDatos
End Property


Private Sub cmbIdTipoSexo_Change()
    RaiseEvent SeModificoSexo(Val(mo_CmbIdTipoSexo.BoundText))
End Sub


Private Sub cmdBuscaPacientePorApellidos_Click()
    Dim oDOPacienteTmp As New doPaciente
    Dim rsPacientesTmp As New Recordset
    oDOPacienteTmp.idPaciente = 0
    oDOPacienteTmp.ApellidoPaterno = txtApellidoPaterno.Text
    oDOPacienteTmp.ApellidoMaterno = txtApellidoMaterno.Text
    oDOPacienteTmp.PrimerNombre = txtPrimerNombre.Text
    oDOPacienteTmp.SegundoNombre = txtSegundoNombre.Text
    Set rsPacientesTmp = mo_AdminAdmision.PacientesObtenerConElMismoNombre(oDOPacienteTmp)
    If rsPacientesTmp.RecordCount > 0 Then
       ml_idPaciente = rsPacientesTmp.Fields!idPaciente
       CargarDatosDePacienteALosControles
    Else
       ml_idPaciente = 0
       txtIdNroHistoria.Text = ""
       mo_cmbIdTipoGenHistoriaClinica.BoundText = sghHistoriaTemporalCOnsultaExterna
    End If
    Set oDOPacienteTmp = Nothing
    Set rsPacientesTmp = Nothing
End Sub



Private Sub txtFechaNacimiento_Change()
    RaiseEvent SeModificoFechaNacimiento(txtFechaNacimiento.Text)
End Sub


Private Sub cmbIdTipoSexo_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, cmbIdTipoSexo
RaiseEvent SePresionoTeclaEspecial(KeyCode)
End Sub


Private Sub cmbIdTipoSexo_LostFocus()
   If cmbIdTipoSexo.Text <> "" Then
        On Error Resume Next
       mo_CmbIdTipoSexo.BoundText = Val(Split(cmbIdTipoSexo.Text, " = ")(0))
       
       If Err.Number <> 0 Then
        cmbIdTipoSexo.Text = ""
       End If
       
   End If
   mo_Formulario.MarcarComoVacio cmbIdTipoSexo
End Sub

Private Sub cmbIdTipoSexo_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsNumero(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
End Sub


Private Sub txtFechaNacimiento_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, txtFechaNacimiento
    RaiseEvent SePresionoTeclaEspecial(KeyCode)
End Sub


Private Sub txtFechaNacimiento_LostFocus()
       If txtFechaNacimiento <> SIGHComun.FECHA_VACIA_DMY Then
            If Not esfecha(txtFechaNacimiento, "DD/MM/AAAA") Then
                MsgBox "La fecha ingresada no es v�lida", vbInformation, "Datos de paciente"
                 txtFechaNacimiento = SIGHComun.FECHA_VACIA_DMY
                 txtFechaNacimiento.SetFocus
            Else
                txtEdad.Text = Trim(Str(EdadActual(CDate(txtFechaNacimiento.Text), ml_FechaRegistro)))
            End If
        End If
   mo_Formulario.MarcarComoVacio txtFechaNacimiento
End Sub

Private Sub txtFechaNacimiento_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsParaFecha(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
End Sub


Private Sub txtPrimerNombre_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, txtPrimerNombre
RaiseEvent SePresionoTeclaEspecial(KeyCode)
End Sub


Private Sub txtPrimerNombre_LostFocus()
    
    If txtPrimerNombre.Text <> "NN" Then
        txtPrimerNombre.Text = mo_Teclado.CapitalizarNombres(txtPrimerNombre.Text)
    End If
    mo_Formulario.MarcarComoVacio txtPrimerNombre
    
End Sub

Private Sub txtPrimerNombre_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsValidoParaNombre(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
End Sub


Private Sub txtApellidoMaterno_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, txtApellidoMaterno
RaiseEvent SePresionoTeclaEspecial(KeyCode)
End Sub


Private Sub txtApellidoMaterno_LostFocus()
    If txtApellidoMaterno.Text <> "NN" Then
        txtApellidoMaterno.Text = mo_Teclado.CapitalizarNombres(txtApellidoMaterno.Text)
    End If
   mo_Formulario.MarcarComoVacio txtApellidoMaterno
End Sub

Private Sub txtApellidoMaterno_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsValidoParaNombre(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
End Sub


Private Sub txtApellidoPaterno_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, txtApellidoPaterno
RaiseEvent SePresionoTeclaEspecial(KeyCode)
End Sub


Private Sub txtApellidoPaterno_LostFocus()
    If txtApellidoPaterno.Text <> "NN" Then
        txtApellidoPaterno.Text = mo_Teclado.CapitalizarNombres(txtApellidoPaterno.Text)
    End If
    mo_Formulario.MarcarComoVacio txtApellidoPaterno
End Sub

Private Sub txtApellidoPaterno_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsValidoParaNombre(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
End Sub

Public Sub ConfigurarComboBoxes()
Dim sMensaje As String
        
        'CARGA COMBO BOXES DE PACIENTE
        mo_CmbIdTipoSexo.BoundColumn = "IdtipoSexo"
        mo_CmbIdTipoSexo.ListField = "DescripcionLarga"
        Set mo_CmbIdTipoSexo.RowSource = mo_AdminServiciosComunes.TiposSexoSeleccionarTodos()
        sMensaje = sMensaje + mo_AdminServiciosComunes.MensajeError
End Sub

Public Function ValidarDatosObligatorios() As String
Dim sMensajeLocal As String

    '---------------------------------------------------------------------------------
  '           VALIDA DATOS DE PACIENTES
    '---------------------------------------------------------------------------------
   
    If txtApellidoPaterno.Text = "" Then
        sMensajeLocal = sMensajeLocal + "Ingrese el apellido paterno" + Chr(13)
    End If
    If txtApellidoMaterno.Text = "" Then
        sMensajeLocal = sMensajeLocal + "Ingrese el apellido materno" + Chr(13)
    End If
    If txtPrimerNombre.Text = "" Then
        sMensajeLocal = sMensajeLocal + "Ingrese el primer nombre" + Chr(13)
    End If
    If Val(mo_CmbIdTipoSexo.BoundText) = 0 Then
       sMensajeLocal = sMensajeLocal + "Ingrese el sexo" + Chr(13)
    End If
    If txtFechaNacimiento.Text = "__/__/____" Then
       sMensajeLocal = sMensajeLocal + "Ingrese la Fecha de Nacimiento" + Chr(13)
    End If
   
   ValidarDatosObligatorios = sMensajeLocal

End Function
Public Function ValidarReglas(oDOPaciente As doPaciente) As Boolean
Dim rspacientes As ADODB.Recordset

    ValidarReglas = False
    
    'Si el paciente aun no existe (IdPaciente = 0) se verifica que no haya duplicados
    If oDOPaciente.idPaciente = 0 Then
        Set rspacientes = mo_AdminAdmision.PacientesObtenerConElMismoNombre(oDOPaciente)
        If Not (rspacientes.EOF And rspacientes.BOF) Then
            rspacientes.MoveFirst
            MsgBox "Existe un paciente con el mismo nombre: " + rspacientes!ApellidoPaterno + " " + rspacientes!ApellidoMaterno + " " + rspacientes!PrimerNombre, vbExclamation, "Datos de paciente"
            rspacientes.Close
            Exit Function
        End If
        rspacientes.Close
         
         Set rspacientes = mo_AdminAdmision.PacientesObtenerConElAutogenerado(oDOPaciente)
         If Not (rspacientes.EOF And rspacientes.BOF) Then
             rspacientes.MoveFirst
             If MsgBox("Existe un paciente con el mismo n�mero autogenerado: " + rspacientes!ApellidoPaterno + " " + rspacientes!ApellidoMaterno + " " + rspacientes!PrimerNombre + Chr(13) + "Desea continuar?", vbQuestion + vbYesNo, "Datos de paciente") = vbNo Then
                 rspacientes.Close
                 Exit Function
             End If
         End If
         rspacientes.Close
        
    End If
    
   If txtFechaNacimiento.Text <> SIGHComun.FECHA_VACIA_DMY Then
        If CDate(txtFechaNacimiento.Text) > Date Then
            MsgBox "La fecha de nacimiento no puede ser mayor que la fecha de creaci�n de la historia", vbExclamation, "Registro de pacientes"
            Exit Function
        End If
    End If
   
   ValidarReglas = True

End Function

Public Function CargarDatosAlObjetoDatos(oDOPaciente As doPaciente)
    
    '---------------------------------------------------------------------------------
    '           CARGA DATOS DEL PACIENTE
    '---------------------------------------------------------------------------------
   With oDOPaciente
        .idPaciente = Me.idPaciente
        .ApellidoPaterno = txtApellidoPaterno.Text
        .ApellidoMaterno = txtApellidoMaterno.Text
        .PrimerNombre = txtPrimerNombre.Text
        .SegundoNombre = txtSegundoNombre.Text
        .FechaNacimiento = IIf(txtFechaNacimiento.Text = SIGHComun.FECHA_VACIA_DMY, 0, txtFechaNacimiento.Text)
        .NroHistoriaClinica = Me.NroHistoriaClinica
        .IdTipoSexo = Val(mo_CmbIdTipoSexo.BoundText)
        .IdTipoNumeracion = mo_cmbIdTipoGenHistoriaClinica.BoundText
        .Autogenerado = mo_AdminAdmision.PacienteCrearNroAutogenerado(oDOPaciente)
         Autogenerado = .Autogenerado
   End With
   Set CargarDatosAlObjetoDatos = oDOPaciente
End Function

Public Sub CargarDatosDePacienteALosControles()
Dim oPacientes  As New doPaciente
'MsgBox "pacientes"
        
        'CARGAR DATOS DEL PACIENTE
        Set oPacientes = mo_AdminAdmision.PacientesSeleccionarPorId(ml_idPaciente)
        If mo_AdminAdmision.MensajeError <> "" Then
             MsgBox "No se pudo obtener los datos" + Chr(13) + mo_AdminAdmision.MensajeError, vbCritical, "Datos de paciente"
             mb_ExistenDatos = False
             Exit Sub
        End If
        If Not oPacientes Is Nothing Then
           With oPacientes
                Me.idPaciente = .idPaciente
                Autogenerado = .Autogenerado
                txtApellidoPaterno.Text = Trim(.ApellidoPaterno)
                txtApellidoMaterno.Text = Trim(.ApellidoMaterno)
                txtPrimerNombre.Text = Trim(.PrimerNombre)
                txtSegundoNombre.Text = Trim(.SegundoNombre)
                txtFechaNacimiento.Text = IIf(.FechaNacimiento = 0, SIGHComun.FECHA_VACIA_DMY, .FechaNacimiento)
                RaiseEvent SeModificoFechaNacimiento(txtFechaNacimiento.Text)
                mo_CmbIdTipoSexo.BoundText = .IdTipoSexo
                RaiseEvent SeModificoSexo(.IdTipoSexo)
                
                mo_cmbIdTipoGenHistoriaClinica.BoundText = .IdTipoNumeracion
                cmbIdTipoGenHistoriaClinica.Tag = .IdTipoNumeracion         'lo guarda para luego comparar
                txtIdNroHistoria.Text = .NroHistoriaClinica          'esto tiene que ir luego del tipo de generacion, por que sino se borra con el change del combo box
                txtIdNroHistoria.Tag = .NroHistoriaClinica
                txtEdad.Text = Trim(Str(EdadActual(.FechaNacimiento, ml_FechaRegistro)))
                mb_ExistenDatos = True
            End With
            
        Else
            mb_ExistenDatos = False
            Exit Sub
        End If

End Sub

Public Sub ConfigurarValoresPorDefecto()
   
    mo_Formulario.HabilitarDeshabilitar cmbIdTipoGenHistoriaClinica, False
    mo_Formulario.HabilitarDeshabilitar txtIdNroHistoria, False
    mo_Formulario.HabilitarDeshabilitar txtEdad, False
    
    mo_cmbIdTipoGenHistoriaClinica.BoundColumn = "IdTipoNumeracion"
    mo_cmbIdTipoGenHistoriaClinica.ListField = "DescripcionLarga"
    'Set mo_cmbIdTipoGenHistoriaClinica.RowSource = mo_AdminArchivoClinico.TiposGeneracionHistoriasSeleccionarDeConsultaExterna()
    Set mo_cmbIdTipoGenHistoriaClinica.RowSource = mo_AdminArchivoClinico.TiposGeneracionHistoriasSeleccionarTodos
    mo_cmbIdTipoGenHistoriaClinica.BoundText = sghHistoriaTemporalCOnsultaExterna

    mo_CmbIdTipoSexo.BoundColumn = "IdtipoSexo"
    mo_CmbIdTipoSexo.ListField = "DescripcionLarga"
    Set mo_CmbIdTipoSexo.RowSource = mo_AdminServiciosComunes.TiposSexoSeleccionarTodos()
    
End Sub

Public Sub LimpiarDatosDePaciente()
           
           'LIMPIAR DATOS DEL PACIENTE
           idPaciente = 0
           Autogenerado = 0
           txtApellidoPaterno.Text = ""
           txtApellidoMaterno.Text = ""
           txtPrimerNombre.Text = ""
           txtSegundoNombre.Text = ""
           txtFechaNacimiento.Text = SIGHComun.FECHA_VACIA_DMY
           txtEdad.Text = ""
           mo_cmbIdTipoGenHistoriaClinica.BoundText = sghHistoriaTemporalCOnsultaExterna
           txtIdNroHistoria.Text = ""

End Sub

Public Sub DeshabilitarFrames(lbDesHabilita As Boolean)
    
    fraDatosPaciente.Enabled = Not lbDesHabilita

End Sub

Public Sub SetFocusOnApellidoPaterno()
         txtApellidoPaterno.SetFocus
End Sub


Public Function Inicializar()
    Set mo_cmbIdTipoGenHistoriaClinica.MiComboBox = cmbIdTipoGenHistoriaClinica
    Set mo_CmbIdTipoSexo.MiComboBox = cmbIdTipoSexo
    ConfigurarValoresPorDefecto
'    mo_Formulario.HabilitarDeshabilitar Me.NroHistoriaClinica, False
'    mo_Formulario.HabilitarDeshabilitar cmbIdTipoGenHistoriaClinica, False
End Function


Private Sub txtSegundoNombre_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, txtSegundoNombre
   RaiseEvent SePresionoTeclaEspecial(KeyCode)
End Sub

Private Sub txtSegundoNombre_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsValidoParaNombre(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
End Sub

Private Sub txtSegundoNombre_LostFocus()
    txtSegundoNombre.Text = mo_Teclado.CapitalizarNombres(txtSegundoNombre.Text)
    cmdBuscaPacientePorApellidos_Click
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    AdministrarKeyPreview KeyCode
End Sub
Sub AdministrarKeyPreview(KeyCode As Integer)
    
    Select Case KeyCode
    Case vbKeyEscape
    Case vbKeyF2
    Case vbKeyF3
     Case vbKeyF4
     Case vbKeyF5
     Case vbKeyF6
     Case vbKeyF7
     Case vbKeyF8
    Case vbKey9
    End Select
       
End Sub


Public Sub CargaAlgunosDatosDesdeBoleta(lcRazonSocial As String)
    Dim lnLen As Integer
    Dim lnPos As Integer
    lnPos = 1
    txtApellidoPaterno.Text = ""
    txtApellidoMaterno.Text = ""
    txtPrimerNombre.Text = ""
    txtSegundoNombre.Text = ""
    If lcRazonSocial <> "" Then
        For lnLen = 1 To Len(lcRazonSocial)
            If Mid(lcRazonSocial, lnLen, 1) = " " Then
               lnPos = lnPos + 1
            Else
               Select Case lnPos
               Case 1
                    txtApellidoPaterno.Text = txtApellidoPaterno.Text & Mid(lcRazonSocial, lnLen, 1)
               Case 2
                    txtApellidoMaterno.Text = txtApellidoMaterno.Text & Mid(lcRazonSocial, lnLen, 1)
               Case 3
                    txtPrimerNombre.Text = txtPrimerNombre.Text & Mid(lcRazonSocial, lnLen, 1)
               End Select
            End If
        Next
        txtApellidoPaterno.Text = Left(txtApellidoPaterno.Text, 20)
        txtApellidoMaterno.Text = Left(txtApellidoMaterno.Text, 20)
        txtPrimerNombre.Text = Left(txtPrimerNombre.Text, 20)
        txtApellidoPaterno.Text = mo_Teclado.CapitalizarNombres(txtApellidoPaterno.Text)
        txtApellidoMaterno.Text = mo_Teclado.CapitalizarNombres(txtApellidoMaterno.Text)
        txtPrimerNombre.Text = mo_Teclado.CapitalizarNombres(txtPrimerNombre.Text)
    End If
    mo_cmbIdTipoGenHistoriaClinica.BoundText = sghHistoriaTemporalCOnsultaExterna
End Sub



