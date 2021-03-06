VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{5A9433E9-DD7B-4529-91B6-A5E8CA054615}#2.0#0"; "IGULTR~1.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form ImagTomog 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   8790
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13755
   Icon            =   "ImagTomog.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8790
   ScaleWidth      =   13755
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      Height          =   840
      Left            =   75
      TabIndex        =   30
      Top             =   7860
      Width           =   13665
      Begin VB.CommandButton btnImprimir 
         Caption         =   "Imprime"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Left            =   120
         Picture         =   "ImagTomog.frx":0CCA
         Style           =   1  'Graphical
         TabIndex        =   48
         Top             =   150
         Visible         =   0   'False
         Width           =   1365
      End
      Begin VB.CommandButton btnAceptar 
         Caption         =   "Aceptar (F2)"
         DisabledPicture =   "ImagTomog.frx":11A3
         DownPicture     =   "ImagTomog.frx":1603
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   5430
         Picture         =   "ImagTomog.frx":1A78
         Style           =   1  'Graphical
         TabIndex        =   32
         Top             =   150
         Width           =   1365
      End
      Begin VB.CommandButton btnCancelar 
         Caption         =   "Cancelar (ESC)"
         DisabledPicture =   "ImagTomog.frx":1EED
         DownPicture     =   "ImagTomog.frx":23B1
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   6840
         Picture         =   "ImagTomog.frx":289D
         Style           =   1  'Graphical
         TabIndex        =   31
         Top             =   135
         Width           =   1365
      End
   End
   Begin VB.Frame fraDatosAtencion 
      Caption         =   "Datos de Cabecera"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4215
      Left            =   75
      TabIndex        =   5
      Top             =   15
      Width           =   13665
      Begin VB.TextBox txtNcita 
         Alignment       =   1  'Right Justify
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
         Left            =   5250
         MaxLength       =   30
         TabIndex        =   52
         Top             =   945
         Width           =   1380
      End
      Begin VB.CommandButton cmbBuscaReceta 
         Height          =   330
         Left            =   2775
         Picture         =   "ImagTomog.frx":2D89
         Style           =   1  'Graphical
         TabIndex        =   50
         Top             =   945
         Width           =   300
      End
      Begin VB.CommandButton cmdBuscaCuentaPorApellidos 
         Height          =   330
         Left            =   2775
         Picture         =   "ImagTomog.frx":3313
         Style           =   1  'Graphical
         TabIndex        =   49
         Top             =   585
         Width           =   300
      End
      Begin UltraGrid.SSUltraGrid grdConsumoPaciente 
         Height          =   2625
         Left            =   8595
         TabIndex        =   44
         Top             =   1050
         Visible         =   0   'False
         Width           =   540
         _ExtentX        =   953
         _ExtentY        =   4630
         _Version        =   131072
         GridFlags       =   17040384
         LayoutFlags     =   71303188
         BorderStyle     =   5
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   "ImagTomog.frx":389D
         Caption         =   "Ex?menes hist?ricos del Paciente (Consulta Externa, Hospitalizaci?n, Emergencia)"
      End
      Begin VB.CheckBox chkMuestraHistorico 
         Alignment       =   1  'Right Justify
         Caption         =   "Muestra Hist?rico de ex?menes"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   10680
         TabIndex        =   45
         Top             =   165
         Width           =   2895
      End
      Begin VB.TextBox txtNreceta 
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
         Left            =   1500
         MaxLength       =   30
         TabIndex        =   42
         Top             =   945
         Width           =   1245
      End
      Begin VB.TextBox txtDerechoI 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   7470
         MaxLength       =   30
         TabIndex        =   40
         Top             =   2010
         Width           =   825
      End
      Begin VB.CheckBox chkPlanNoCubre 
         Alignment       =   1  'Right Justify
         Caption         =   "IAFA NO cubre"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   6750
         TabIndex        =   38
         Top             =   600
         Width           =   1545
      End
      Begin VB.TextBox txtPorcIR 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   12735
         MaxLength       =   30
         TabIndex        =   36
         Top             =   3765
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.CheckBox chkIonico 
         Alignment       =   1  'Right Justify
         Caption         =   "Es Ionico"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   9150
         TabIndex        =   35
         Top             =   3795
         Value           =   1  'Checked
         Visible         =   0   'False
         Width           =   1035
      End
      Begin VB.CheckBox chkContraste 
         Alignment       =   1  'Right Justify
         Caption         =   "Con Contraste"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   4245
         TabIndex        =   34
         Top             =   1650
         Value           =   1  'Checked
         Visible         =   0   'False
         Width           =   1515
      End
      Begin VB.TextBox txtEstado 
         Enabled         =   0   'False
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
         Left            =   2910
         MaxLength       =   30
         TabIndex        =   17
         Top             =   210
         Width           =   645
      End
      Begin VB.TextBox txtNmovimiento 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1500
         MaxLength       =   30
         TabIndex        =   16
         Top             =   210
         Width           =   735
      End
      Begin VB.TextBox txtPlan 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   9195
         TabIndex        =   14
         Top             =   510
         Width           =   4365
      End
      Begin VB.TextBox txtNroOrden 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   6900
         MaxLength       =   30
         TabIndex        =   13
         Top             =   1650
         Visible         =   0   'False
         Width           =   1395
      End
      Begin VB.TextBox txtDatosDeCuenta 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   3120
         TabIndex        =   12
         Top             =   600
         Width           =   3495
      End
      Begin VB.TextBox txtNcuenta 
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
         Left            =   1500
         MaxLength       =   30
         TabIndex        =   0
         Top             =   600
         Width           =   1245
      End
      Begin VB.TextBox txtProcedencia 
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
         Left            =   1500
         TabIndex        =   11
         Top             =   1290
         Width           =   4365
      End
      Begin VB.TextBox txtNserie 
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
         Left            =   1500
         MaxLength       =   4
         TabIndex        =   9
         Top             =   1650
         Width           =   585
      End
      Begin VB.TextBox txtNboleta 
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
         Left            =   2070
         MaxLength       =   30
         TabIndex        =   10
         Top             =   1650
         Width           =   1125
      End
      Begin VB.ComboBox cmbResponsable 
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
         Left            =   1500
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   2040
         Width           =   4290
      End
      Begin VB.TextBox txtDx 
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
         Left            =   1500
         MaxLength       =   30
         TabIndex        =   2
         ToolTipText     =   "Ingrese el Dx (4 d?gitos)"
         Top             =   2400
         Width           =   1065
      End
      Begin VB.TextBox txtNombreDx 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2970
         TabIndex        =   8
         Top             =   2400
         Width           =   3765
      End
      Begin VB.CommandButton cmdBuscaDx 
         Caption         =   "..."
         Height          =   315
         Left            =   2610
         TabIndex        =   7
         Top             =   2400
         Width           =   315
      End
      Begin VB.TextBox txtResultadoFinal 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1140
         Left            =   1500
         MaxLength       =   3000
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   3
         Top             =   2760
         Width           =   7095
      End
      Begin VB.CheckBox chkDxDefinitivo 
         Alignment       =   1  'Right Justify
         Caption         =   "Dx Definitivo"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   6810
         TabIndex        =   6
         Top             =   2370
         Value           =   1  'Checked
         Visible         =   0   'False
         Width           =   1485
      End
      Begin SIGHImagen.UcPacienteDatos UcPacienteDatos1 
         Height          =   2700
         Left            =   9195
         TabIndex        =   15
         Top             =   960
         Width           =   4425
         _ExtentX        =   7805
         _ExtentY        =   4763
      End
      Begin MSMask.MaskEdBox txtFregistro 
         Height          =   315
         Left            =   4080
         TabIndex        =   18
         Top             =   210
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSDataListLib.DataCombo cmbFormaPago 
         Height          =   330
         Left            =   5910
         TabIndex        =   39
         Top             =   1290
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   582
         _Version        =   393216
         MatchEntry      =   -1  'True
         Style           =   2
         Text            =   "DataCombo1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSMask.MaskEdBox txtFrealizaCpt 
         Height          =   315
         Left            =   6540
         TabIndex        =   46
         Top             =   210
         Width           =   1770
         _ExtentX        =   3122
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   16
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##/##/#### ##:##"
         PromptChar      =   "_"
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "N? Cita"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   4635
         TabIndex        =   53
         Top             =   1020
         Width           =   600
      End
      Begin VB.Label lblOrdenPago 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "N? Orden de Pago"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   240
         Left            =   7380
         TabIndex        =   51
         Top             =   3915
         Width           =   1200
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "F.Realiza CPT"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   5490
         TabIndex        =   47
         Top             =   270
         Width           =   1065
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "N? Receta"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   210
         TabIndex        =   43
         Top             =   990
         Width           =   870
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Derecho Informe"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   5010
         TabIndex        =   41
         Top             =   2040
         Width           =   2430
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "% para Informe Radiol?gico"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   10335
         TabIndex        =   37
         Top             =   3795
         Visible         =   0   'False
         Width           =   2400
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Estado"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   2370
         TabIndex        =   29
         Top             =   240
         Width           =   555
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "N? Movimiento"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   210
         TabIndex        =   28
         Top             =   255
         Width           =   1245
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "F.Reg"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   3630
         TabIndex        =   27
         Top             =   240
         Width           =   465
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Plan"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   8805
         TabIndex        =   26
         Top             =   540
         Width           =   330
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "N? Orden"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   6030
         TabIndex        =   25
         Top             =   1710
         Visible         =   0   'False
         Width           =   885
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "N? Boleta"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   210
         TabIndex        =   24
         Top             =   1700
         Width           =   780
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "N? Cuenta"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   210
         TabIndex        =   23
         Top             =   645
         Width           =   855
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Procedencia"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   210
         TabIndex        =   22
         Top             =   1320
         Width           =   990
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
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
         Height          =   210
         Left            =   210
         TabIndex        =   21
         Top             =   2080
         Width           =   1005
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         Caption         =   "Resultado Final"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   210
         TabIndex        =   20
         Top             =   2850
         Width           =   1200
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Diagn?stico"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   210
         TabIndex        =   19
         Top             =   2460
         Width           =   930
      End
   End
   Begin SIGHImagen.ucInsumoYcpt ucProductos 
      Height          =   3570
      Left            =   75
      TabIndex        =   4
      Top             =   4290
      Width           =   13665
      _ExtentX        =   24104
      _ExtentY        =   6297
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      Caption         =   "N? Boleta"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   8250
      TabIndex        =   33
      Top             =   1830
      Width           =   780
   End
End
Attribute VB_Name = "ImagTomog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'------------------------------------------------------------------------------------
'        Organizaci?n: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Mantenimiento de Tomograf?a
'        Programado por: Barrantes D
'        Fecha: Julio 2009
'
'------------------------------------------------------------------------------------
Option Explicit
Dim mo_ReporteUtil As New ReporteUtil
Dim ml_idMovimiento As Long
Dim mi_Opcion As SIGHEntidades.sghOpciones
Dim ms_MensajeError As String
Dim ml_idUsuario As Long
Dim mb_ExistenDatos As Boolean
Dim mo_ReglasImagenes As New SIGHNegocios.ReglasImagenes
Dim mo_AdminAdmision As New SIGHNegocios.ReglasAdmision
Dim mo_ReglasFacturacion As New SIGHNegocios.ReglasFacturacion
Dim mo_ReglasFarmacia As New SIGHNegocios.ReglasFarmacia
Dim mo_AdminCaja As New SIGHNegocios.ReglasCaja
Dim mo_reglasComunes As New SIGHNegocios.ReglasComunes
Dim mo_ReglasSeguridad As New SIGHNegocios.ReglasDeSeguridad
Dim mo_AdminArchivoClinico As New SIGHNegocios.ReglasArchivoClinico
Dim mo_ReglasSISgalenhos As New SIGHSis.ReglasSISgalenhos
Dim wxParametro302 As String, lnIdTipoServicio As Long
Dim mo_Apariencia As New SIGHEntidades.GridInfragistic
Dim mo_cmbIdEstado As New SIGHEntidades.ListaDespleglable
Dim mo_cmbResponsable As New SIGHEntidades.ListaDespleglable
Dim lcBuscaParametro As New SIGHDatos.Parametros
Dim mo_Teclado As New SIGHEntidades.Teclado
Dim mo_Formulario As New SIGHEntidades.Formulario
Dim lbPrimeraVez As Boolean
Dim ml_IdTipoFinanciamiento As Long
Dim ml_IdPaciente As Long
Dim ml_IdComprobantePago As Long
Dim ml_IdFuenteFinanciamiento  As Long
Dim ml_IdServicioPaciente As Long
Dim ml_IdDiagnostico As Long
Dim oDOPaciente As New doPaciente
Dim oDoImagMovimiento As New DoImagMovimiento
Dim oDoImagMovimientoImagenes As New DoImagMovimientoImagenes
Dim oDoFactOrdenServ As New DoFactOrdenServ
Dim rsProductosCPT As Recordset
Dim rsProductos As Recordset
Dim oRsFormaPago As New Recordset
Dim ml_IdFuenteFinanciamientoDespacho As Long
Const ml_PuntoCarga As Long = 22      'Tomografia
Const lcConstanteMovimientoSalida As String = "S"
Dim ml_IdTipoVentaSeleccionada As Long
Dim mo_lcNombrePc As String
Dim mo_lnIdTablaLISTBARITEMS   As Long
Dim ml_MontoDerechoInforme As Double, ml_idDerechoInforme As Long
Dim lnIdReceta As Long
Dim lnUltimaBusqueda As sghUltimaBusqueda
Dim lnIdPacienteHistorico As Long
Dim ml_SeEligioGridBoleta As Boolean
Dim lbPideInsumo  As Boolean
Dim wxParametro509 As String
Dim lnEpsPorcentaje As Double
Dim lcMedicoDNI As String, lcCama As String, lcMedico As String, lnMedicoId As Long
Dim lbCuentaDeEmergenciaCerrada As Boolean
Dim wxParametro578 As String

Property Let SeEligioGridBoleta(lValue As Boolean)
    ml_SeEligioGridBoleta = lValue
End Property
Property Get SeEligioGridBoleta() As Boolean
    SeEligioGridBoleta = ml_SeEligioGridBoleta
End Property

Property Let lcNombrePc(lValue As String)
   mo_lcNombrePc = lValue
End Property
Property Let lnIdTablaLISTBARITEMS(lValue As Long)
   mo_lnIdTablaLISTBARITEMS = lValue
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

Property Let idMovimiento(lValue As Long)
    ml_idMovimiento = lValue
End Property

Property Get idMovimiento() As Long
    idMovimiento = ml_idMovimiento
End Property


Private Sub btnAceptar_Click()
   If btnAceptar.Enabled = False Then
      Exit Sub
   End If
   mo_reglasComunes.DevuelveCamaYdniMedico lcMedico, lcMedicoDNI, lcCama, 0, lnMedicoId, ml_IdPaciente
   Select Case mi_Opcion
   Case sghAgregar
       If ValidarDatosObligatorios() Then
           CargaDatosAlObjetosDeDatos
           If ValidarReglas() Then
               If AgregarDatos() Then
                    lblOrdenPago.Caption = mo_ReglasImagenes.DevuelveNombreArchivoImagenes(oDoImagMovimiento.idMovimiento, _
                       mo_ReglasImagenes.DevuelveIdPacienteParaLeerImagenes(oDOPaciente, oDoImagMovimientoImagenes), _
                       rsProductosCPT) & lblOrdenPago.Caption
                    Me.txtNmovimiento = oDoImagMovimiento.idMovimiento
                    MsgBox "Se agreg? correctamente el Movimiento N? " & oDoImagMovimiento.idMovimiento & Chr(13) & _
                    lblOrdenPago.Caption, vbInformation, Me.Caption
                    LimpiarFormulario
                Else
                    MsgBox "No se pudo agregar los datos" & Chr(13) & ms_MensajeError, vbExclamation, Me.Caption
               End If
           End If
       End If
   Case sghModificar
       If ValidarDatosObligatorios() Then
            CargaDatosAlObjetosDeDatos
           If ValidarReglas() Then
               If ModificarDatos() Then
                    lblOrdenPago.Caption = mo_ReglasImagenes.DevuelveNombreArchivoImagenes(oDoImagMovimiento.idMovimiento, _
                       mo_ReglasImagenes.DevuelveIdPacienteParaLeerImagenes(oDOPaciente, oDoImagMovimientoImagenes), _
                       rsProductosCPT) & lblOrdenPago.Caption
                    MsgBox "Se Modific? correctamente el Movimiento N? " & oDoImagMovimiento.idMovimiento & Chr(13) & _
                    lblOrdenPago.Caption, vbInformation, Me.Caption
                    Me.Visible = False
                    LimpiarVariablesDeMemoria
                Else
                    MsgBox "No se pudo modificar los datos" & Chr(13) & ms_MensajeError, vbExclamation, Me.Caption
               End If
           End If
       End If
   Case sghEliminar
            If MsgBox("?Realmente desea Anular?", vbQuestion + vbYesNo, Me.Caption) = vbNo Then
                 Exit Sub
            End If
           If ValidarReglas() Then
                CargaDatosAlObjetosDeDatos
               If EliminarDatos() Then
                    MsgBox "Los datos se Anularon correctamente", vbInformation, Me.Caption
                    Me.Visible = False
                    LimpiarVariablesDeMemoria
                Else
                    MsgBox "No se pudo anular los datos" & Chr(13) & ms_MensajeError, vbExclamation, Me.Caption
               End If
           End If
   End Select
        
End Sub

Sub LimpiarFormulario()
    lbCuentaDeEmergenciaCerrada = False
    lnEpsPorcentaje = 0
    lblOrdenPago.Caption = ""
    txtNcuenta.Text = ""
    txtDatosDeCuenta.Text = ""
    txtProcedencia.Text = ""
    txtNserie.Text = ""
    txtNboleta.Text = ""
    txtNroOrden.Text = ""
   ' txtNcuenta.SetFocus
    Me.ucProductos.LimpiarGrilla
    
    UcPacienteDatos1.LimpiarDatosDePaciente
    ml_IdPaciente = 0
    ml_IdTipoFinanciamiento = 0
    ml_IdFuenteFinanciamiento = 0
    ml_IdServicioPaciente = 0
    ml_IdComprobantePago = 0
    chkDxDefinitivo.Visible = False
    chkDxDefinitivo.Value = 1
    ml_IdDiagnostico = 0
    txtDx.Text = ""
    txtNombreDx.Text = ""
    txtResultadoFinal.Text = ""
    chkIonico.Value = 1
    chkContraste.Value = 1
    Me.chkPlanNoCubre.Value = 0
    cmbFormaPago.Text = ""
    txtDerechoI.Text = "30.00"
    ml_MontoDerechoInforme = 0
    txtNreceta.Text = ""
    txtFrealizaCpt.Text = lcBuscaParametro.RetornaFechaHoraServidorSQL
    mo_Formulario.HabilitarDeshabilitar txtDx, True
    ucProductos.PermiteAgregarItems = True
    If mi_Opcion = sghAgregar Then
        On Error Resume Next
        If lnUltimaBusqueda = sghEnBoleta Then
           Me.txtNserie.SetFocus
        Else
           Me.txtNcuenta.SetFocus
        End If
    End If
    lnIdPacienteHistorico = 0: Set grdConsumoPaciente.DataSource = Nothing
    Me.chkMuestraHistorico.Value = 0: chkMuestraHistorico_Click
    
End Sub

Function ValidarDatosObligatorios() As Boolean
    On Error GoTo ErrVald
    Dim lnTabError As Integer
    Dim lbNuevo As Boolean
    ValidarDatosObligatorios = False
    If txtNcuenta.Text = "" And txtNboleta.Text = "" Then
       Exit Function
    End If
    ms_MensajeError = ""
    UcPacienteDatos1.CargarDatosAlObjetoDatos oDOPaciente
    If txtDatosDeCuenta.Text = "" Then
       If oDOPaciente.ApellidoPaterno = "" Then
           ms_MensajeError = ms_MensajeError & "Tiene que registrar el Apellido Paterno" & Chr(13)
           lnTabError = 1
       End If
       If oDOPaciente.ApellidoMaterno = "" Then
           ms_MensajeError = ms_MensajeError & "Tiene que registrar el Apellido Materno" & Chr(13)
           lnTabError = 1
       End If
       If oDOPaciente.PrimerNombre = "" Then
           ms_MensajeError = ms_MensajeError & "Tiene que registrar el Primer Nombre" & Chr(13)
           lnTabError = 1
       End If
    Else
    End If
    If cmbResponsable.Text = "" Then
       ms_MensajeError = ms_MensajeError & "Tiene que elegir el Responsable " & Chr(13)
       lnTabError = 2
    End If
    Select Case mi_Opcion
    Case sghAgregar, sghModificar
        'Cpt
        lbNuevo = True
        Set rsProductosCPT = Me.ucProductos.FacturacionProductos
        If Not (rsProductosCPT.EOF And rsProductosCPT.BOF) Then
            rsProductosCPT.MoveFirst
            txtNroOrden.Text = rsProductosCPT.Fields!IdOrden
            Do While Not rsProductosCPT.EOF
                If rsProductosCPT!idProducto = 0 Then
                   rsProductosCPT.Delete
                   rsProductosCPT.Update
                Else
                   If rsProductosCPT!Cantidad <= 0 Then
                      ms_MensajeError = ms_MensajeError & "El producto CPT: " & rsProductosCPT!codigo & " " & Trim(rsProductosCPT!nombreProducto) & "   Tiene problemas con la Cantidad" & Chr(13)
                   End If
                   If rsProductosCPT!PrecioUnitario <= 0 And rsProductosCPT!SeUsaSinPrecio = False Then
                      If Val(Me.txtNboleta.Text) = 0 Then  'debb-05/04/2011
                         ms_MensajeError = ms_MensajeError & "El producto CPT: " & rsProductosCPT!codigo & " " & Trim(rsProductosCPT!nombreProducto) & "   Tiene problemas con el Precio" & Chr(13)
                      End If
                   End If
                   If rsProductosCPT!Cantidad < rsProductosCPT!cantidadFallada Then
                      ms_MensajeError = ms_MensajeError & "El producto CPT: " & rsProductosCPT!codigo & " " & Trim(rsProductosCPT!nombreProducto) & "   la CANTIDAD FALLADA debe ser menor a la CANTIDAD" & Chr(13)
                   End If
                   If txtNboleta.Text = "" Then
                      'chequeo solo para pacientes con  Nro Cuenta
                      rsProductosCPT.Fields!totalPorPagar = Round(rsProductosCPT!Cantidad * rsProductosCPT!PrecioUnitario, 2)
                   End If
                End If
                On Error Resume Next
                If ml_idDerechoInforme = rsProductosCPT.Fields!idProducto Then
                   lbNuevo = False
                End If
                rsProductosCPT.MoveNext
            Loop
            If ml_idDerechoInforme > 0 And lbNuevo = True Then
                With rsProductosCPT
                    .AddNew
                    .Fields!IdFacturacionProducto = 0
                    .Fields!idProducto = ml_idDerechoInforme
                    .Fields!codigo = ""
                    .Fields!nombreProducto = ""
                    .Fields!Cantidad = 1
                    .Fields!PrecioUnitario = ml_MontoDerechoInforme
                    .Fields!totalPorPagar = ml_MontoDerechoInforme
                    .Fields!IdTipoFinanciamiento = ml_IdTipoFinanciamiento
                    .Fields!IdPuntoCarga = ml_PuntoCarga
                    .Fields!EstadoLocal = "A"   'Agregar
                    .Fields!FechaAutorizaPendiente = 0
                    .Fields!IdUsuarioAutorizaPendiente = 0
                    Select Case ml_IdTipoFinanciamiento
                    Case 2, 3, 4
                        .Fields!idEstadoFacturacion = 4
                        .Fields!FechaAutorizaSeguro = Now
                        .Fields!IdUsuarioAutorizaSeguro = ml_idUsuario
                    Case Else
                        .Fields!idEstadoFacturacion = 1
                        .Fields!FechaAutorizaSeguro = 0
                        .Fields!IdUsuarioAutorizaSeguro = 0
                    End Select
                    .Fields!idFuenteFinanciamiento = 1
                    .Fields!IdServicioInternamiento = 0
                    .Fields!IdUsuarioAuditoria = ml_idUsuario
                    .Fields!idComprobantePago = 0
                    .Fields!IdComprobantePagoDevolucion = 0
                    .Fields!IdOrden = 0
                    .Update
                End With
            End If
        End If
        If Me.ucProductos.DevuelveTotalPagar <= 0 Then
           If txtNboleta.Text = "" Then
             'chequeo solo para pacientes con  Nro Cuenta
             ' ms_MensajeError = ms_MensajeError & "El Importe Total es 0.....verifique" & Chr(13)
           End If
        End If
        'Insumos
        Set rsProductos = Me.ucProductos.FacturacionInsumos
        If Not (rsProductos.EOF And rsProductos.BOF) Then
            rsProductos.MoveFirst
            Do While Not rsProductos.EOF
                If rsProductos!idProducto = 0 Or rsProductos!idProductoCpt = 0 Then
                   rsProductos.Delete
                   rsProductos.Update
                Else
                   If rsProductos!Cantidad <= 0 Then
                      ms_MensajeError = ms_MensajeError & "El INSUMO: " & rsProductos!codigo & " " & Trim(rsProductos!nombreProducto) & "   Tiene problemas con la Cantidad" & Chr(13)
                   End If
                   If rsProductos!PrecioUnitario <= 0 Then
                         ms_MensajeError = ms_MensajeError & "El INSUMO: " & rsProductos!codigo & " " & Trim(rsProductos!nombreProducto) & "   Tiene problemas con el Precio" & Chr(13)
                   End If
                   If rsProductos!Cantidad < rsProductos!cantidadFallada Then
                      ms_MensajeError = ms_MensajeError & "El INSUMO: " & rsProductos!codigo & " " & Trim(rsProductos!nombreProducto) & "   la CANTIDAD FALLADA debe ser menor a la CANTIDAD" & Chr(13)
                   End If
                   rsProductosCPT.MoveFirst
                   rsProductosCPT.Find "idProducto=" & rsProductos!idProductoCpt
                   If rsProductosCPT.EOF Then
                      ms_MensajeError = ms_MensajeError & "El INSUMO: " & rsProductos!codigo & " " & Trim(rsProductos!nombreProducto) & "   no tiene C?digo CPT" & Chr(13)
                   End If
                End If
                rsProductos.MoveNext
            Loop
        End If
        If rsProductos.RecordCount = 0 And lbPideInsumo = True Then
           ms_MensajeError = ms_MensajeError & "Tiene que registrar INSUMOS.....verifique" & Chr(13)
        End If
    End Select
    If ms_MensajeError = "" Then
       ValidarDatosObligatorios = True
    Else
       MsgBox ms_MensajeError, vbInformation, Me.Caption
       Select Case lnTabError
       Case 1
           UcPacienteDatos1.SetFocusOnApellidoPaterno
       Case 2
           cmbResponsable.SetFocus
       End Select
    End If
ErrVald:
End Function

Sub CargaDatosAlObjetosDeDatos()
    Select Case mi_Opcion
    Case sghAgregar
        With oDoImagMovimiento
            .fecha = lcBuscaParametro.RetornaFechaHoraServidorSQL
            .IdImagEstado = sghEstadoTabla.sghRegistrado    'Registrado
            .IdPuntoCarga = ml_PuntoCarga
            .IdTipoConcepto = sghTipoConceptoImagen.sghImgTCsalida  'Salidas con Orden de Pago
            .idUsuario = ml_idUsuario
            .IdUsuarioAuditoria = ml_idUsuario
            .MovTipo = lcConstanteMovimientoSalida
        End With
        With oDoImagMovimientoImagenes
            '.CorrelativoAnual
            .idComprobantePago = ml_IdComprobantePago
            .idCuentaAtencion = Val(txtNcuenta.Text)
            .IdOrden = Val(txtNroOrden.Text)
            .IdPersonaTomaImagen = Val(mo_cmbResponsable.BoundText)
            .IdUsuarioAuditoria = ml_idUsuario
            '.PersonaRecoge
            .PorcInformeRadiolog = CCur(txtPorcIR.Text)
            .ResultadoFinal = txtResultadoFinal.Text
            '.zonaRayosX
            .EsContraste = IIf(chkContraste.Value = 1, 1, 2)
            If chkContraste.Value = 1 Then
              .EsContrasteIonico = IIf(chkIonico.Value = 1, 1, 2)
            End If
            .idDiagnostico = ml_IdDiagnostico
            If ml_IdDiagnostico > 0 Then
               .EsDiagnosticoDefinitivo = IIf(chkDxDefinitivo.Value = 1, sghTipoDx.sghTipoDxDefinitivo, sghTipoDx.sghTipoDxPresuntivo)    '1-definitivo, 2-presuntivo
            Else
               .EsDiagnosticoDefinitivo = sghTipoDx.sghTipoDxNINGUNO
            End If
            .Paciente = Trim(oDOPaciente.ApellidoPaterno) & " " & Trim(oDOPaciente.ApellidoMaterno) & " " & oDOPaciente.PrimerNombre
            .idTipoSexo = oDOPaciente.idTipoSexo
            .FechaNacimiento = oDOPaciente.FechaNacimiento
        End With
        With oDOPaciente  'ya lo cargo en Validacion de Datos
            .IdUsuarioAuditoria = ml_idUsuario
        End With
        With oDoFactOrdenServ
            .FechaCreacion = oDoImagMovimiento.fecha
            .FechaDespacho = oDoImagMovimiento.fecha
            .idCuentaAtencion = Val(txtNcuenta.Text)
            .idEstadoFacturacion = sghEstadoFacturacion.sghAtendido    '1=Registrado, 11=despachado
            .idFuenteFinanciamiento = ml_IdFuenteFinanciamiento
            .idPaciente = ml_IdPaciente
            .IdPuntoCarga = ml_PuntoCarga
            .idServicioPaciente = ml_IdServicioPaciente
            .IdTipoFinanciamiento = ml_IdTipoFinanciamiento
            .idUsuario = ml_idUsuario
            .IdUsuarioAuditoria = ml_idUsuario
            .IdUsuarioDespacho = ml_idUsuario
            .FechaHoraRealizaCpt = txtFrealizaCpt.Text
        End With
    Case sghModificar
        With oDoImagMovimiento
            .IdUsuarioAuditoria = ml_idUsuario
        End With
        With oDoImagMovimientoImagenes
            '.CorrelativoAnual
            '.IdComprobantePago = ml_IdComprobantePago
            '.idCuentaAtencion = Val(txtNcuenta.Text)
            '.IdOrden = Val(txtNroOrden.Text)
            .IdPersonaTomaImagen = Val(mo_cmbResponsable.BoundText)
            .IdUsuarioAuditoria = ml_idUsuario
            '.PersonaRecoge
            .PorcInformeRadiolog = CCur(txtPorcIR.Text)
            .ResultadoFinal = txtResultadoFinal.Text
            .EsContraste = IIf(chkContraste.Value = 1, 1, 2)
            If chkContraste.Value = 1 Then
               .EsContrasteIonico = IIf(chkIonico.Value = 1, 1, 2)
            Else
               .EsContrasteIonico = 0
            End If
            .idDiagnostico = ml_IdDiagnostico
            If ml_IdDiagnostico > 0 Then
               .EsDiagnosticoDefinitivo = IIf(chkDxDefinitivo.Value = 1, sghTipoDx.sghTipoDxDefinitivo, sghTipoDx.sghTipoDxPresuntivo)    '1-definitivo, 2-presuntivo
            Else
               .EsDiagnosticoDefinitivo = sghTipoDx.sghTipoDxNINGUNO
            End If
            .Paciente = Trim(oDOPaciente.ApellidoPaterno) & " " & Trim(oDOPaciente.ApellidoMaterno) & " " & oDOPaciente.PrimerNombre
            .idTipoSexo = oDOPaciente.idTipoSexo
            .FechaNacimiento = oDOPaciente.FechaNacimiento
        End With
        With oDOPaciente  'ya lo cargo en Validacion de Datos
            .IdUsuarioAuditoria = ml_idUsuario
        End With
        With oDoFactOrdenServ
            .FechaHoraRealizaCpt = txtFrealizaCpt.Text
        End With
    Case sghEliminar
        With oDoImagMovimiento
            .IdUsuarioAuditoria = ml_idUsuario
        End With
        With oDoImagMovimientoImagenes
            .IdUsuarioAuditoria = ml_idUsuario
        End With
        With oDOPaciente  'ya lo cargo en Validacion de Datos
            .IdUsuarioAuditoria = ml_idUsuario
        End With
    End Select
End Sub

Function ValidarReglas() As Boolean
   ValidarReglas = False
    

    
   ValidarReglas = True
End Function
Function AgregarDatos() As Boolean
    
    AgregarDatos = mo_ReglasImagenes.ImagMovimientoImagenesAgregar(oDoImagMovimiento, oDoImagMovimientoImagenes, _
                   oDOPaciente, oDoFactOrdenServ, rsProductos, Val(txtNcuenta.Text), rsProductosCPT, _
                   mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, lnIdReceta, txtNserie.Text & " " & txtNboleta.Text, lcMedicoDNI, lcMedico, lcCama, _
                   txtNombreDx.Text, txtDx.Text, "", lnEpsPorcentaje, Val(txtNcita.Text))
    If mo_ReglasImagenes.IdOrdenPago > 0 Then
       lblOrdenPago.Caption = "N? Orden de Pago: " & mo_ReglasImagenes.IdOrdenPago
    End If
                   
    ms_MensajeError = mo_ReglasImagenes.MensajeError
    ml_idMovimiento = oDoImagMovimiento.idMovimiento
    If oDoImagMovimientoImagenes.idCuentaAtencion > 0 Then
       mo_ReglasFacturacion.FacturacionCuentasAtencionPtosActualizar oDoImagMovimientoImagenes.idCuentaAtencion, False, 0
       mo_ReglasSISgalenhos.FuaActualizaDespachosEnServicios oDoImagMovimientoImagenes.idCuentaAtencion, wxParametro302, lnIdTipoServicio, ml_IdFuenteFinanciamiento
    End If
End Function

Function ModificarDatos() As Boolean
    ModificarDatos = mo_ReglasImagenes.ImagMovimientoImagenesModificar(oDoImagMovimiento, oDoImagMovimientoImagenes, _
                     oDOPaciente, oDoFactOrdenServ, rsProductos, Val(txtNcuenta.Text), rsProductosCPT, _
                     mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, lnIdReceta, txtNserie.Text & " " & txtNboleta.Text, lcMedicoDNI, lcMedico, lcCama, _
                   txtNombreDx.Text, txtDx.Text, "", lnEpsPorcentaje, Val(lblOrdenPago.Tag))
    If mo_ReglasImagenes.IdOrdenPago > 0 Then
       lblOrdenPago.Caption = "N? Orden de Pago: " & mo_ReglasImagenes.IdOrdenPago
    Else
       lblOrdenPago.Caption = ""
    End If
    ms_MensajeError = mo_ReglasImagenes.MensajeError
    If oDoImagMovimientoImagenes.idCuentaAtencion > 0 Then
       mo_ReglasFacturacion.FacturacionCuentasAtencionPtosActualizar oDoImagMovimientoImagenes.idCuentaAtencion, False, 0
       mo_ReglasSISgalenhos.FuaActualizaDespachosEnServicios oDoImagMovimientoImagenes.idCuentaAtencion, wxParametro302, lnIdTipoServicio, ml_IdFuenteFinanciamiento
    End If
End Function

Function EliminarDatos() As Boolean
    Set rsProductosCPT = Me.ucProductos.FacturacionProductos
    EliminarDatos = mo_ReglasImagenes.ImagMovimientoImagenesAnular(oDoImagMovimiento, oDoImagMovimientoImagenes, oDOPaciente, _
                    oDoFactOrdenServ, rsProductos, Val(txtNcuenta.Text), rsProductosCPT, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, _
                    lnIdReceta, txtNserie.Text & " " & txtNboleta.Text, lcMedicoDNI, lcMedico, lcCama, _
                   txtNombreDx.Text, txtDx.Text, "", Val(lblOrdenPago.Tag))
    ms_MensajeError = mo_ReglasImagenes.MensajeError
    If oDoImagMovimientoImagenes.idCuentaAtencion > 0 Then
       mo_ReglasFacturacion.FacturacionCuentasAtencionPtosActualizar oDoImagMovimientoImagenes.idCuentaAtencion, False, 0
       mo_ReglasSISgalenhos.FuaActualizaDespachosEnServicios oDoImagMovimientoImagenes.idCuentaAtencion, wxParametro302, lnIdTipoServicio, ml_IdFuenteFinanciamiento
    End If
End Function





Private Sub btnCancelar_Click()
    Me.Visible = False
    LimpiarVariablesDeMemoria
End Sub




Private Sub chkContraste_Click()
    If chkContraste.Value = 1 Then
       chkIonico.Visible = True
    Else
       chkIonico.Visible = False
    End If
End Sub

Private Sub chkMuestraHistorico_Click()
    If chkMuestraHistorico.Value = 1 Then
       grdConsumoPaciente.Visible = True
       If lnIdPacienteHistorico > 0 Then
          If mi_Opcion = sghAgregar Then
             Set grdConsumoPaciente.DataSource = mo_ReglasImagenes.CptHistoricosPorPaciente(lnIdPacienteHistorico, ml_PuntoCarga, 0)
          Else
             Set grdConsumoPaciente.DataSource = mo_ReglasImagenes.CptHistoricosPorPaciente(lnIdPacienteHistorico, ml_PuntoCarga, ml_idMovimiento)
          End If
          'grdConsumoPaciente.Top = Me.UcPacienteDatos1.Top
          grdConsumoPaciente.Left = Me.UcPacienteDatos1.Left
          grdConsumoPaciente.Width = Me.UcPacienteDatos1.Width
          grdConsumoPaciente.Caption = "Hist?ricos de ex?menes: " & Me.UcPacienteDatos1.DevuelvePaciente
       Else
          Set grdConsumoPaciente.DataSource = Nothing
          grdConsumoPaciente.Caption = ""
       End If
    Else
       grdConsumoPaciente.Visible = False
    End If

End Sub

Private Sub chkPlanNoCubre_Click()
    If chkPlanNoCubre.Value = 1 Then
       ml_IdTipoFinanciamiento = 1
       cmbFormaPago.BoundText = ml_IdTipoFinanciamiento
       ucProductos.IdTipoFinanciamiento = ml_IdTipoFinanciamiento
       ucProductos.LimpiarGrilla
    Else
       txtNcuenta_LostFocus
       ucProductos.LimpiarGrilla
    End If
End Sub

Private Sub cmbResponsable_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, cmbResponsable
    AdministrarKeyPreview KeyCode
End Sub

Private Sub cmdBuscaDx_Click()
    Dim oBusqueda As New SIGHNegocios.BuscaDiagnosticos
    Dim oDODiagnostico As DODiagnostico
    oBusqueda.SoloMuestraDxGalenHos = True
    oBusqueda.MostrarFormulario
    If oBusqueda.BotonPresionado = sghAceptar Then
        Set oDODiagnostico = mo_reglasComunes.DiagnosticosSeleccionarPorId(oBusqueda.IdRegistroSeleccionado)
        If Not oDODiagnostico Is Nothing Then
            ml_IdDiagnostico = oDODiagnostico.idDiagnostico
            txtDx.Text = oDODiagnostico.CodigoCIE2004
            txtNombreDx.Text = oDODiagnostico.descripcion
            chkDxDefinitivo.Visible = True
        End If
    End If
    Set oBusqueda = Nothing
    Set oDODiagnostico = Nothing

End Sub

Private Sub Form_Activate()
    If ml_SeEligioGridBoleta = True And mi_Opcion = sghAgregar And ml_idMovimiento > 0 Then
       ml_SeEligioGridBoleta = False
       On Error Resume Next
       cmbResponsable.SetFocus
    End If
End Sub
Sub CargaBoletaAutomaticamente()
    If ml_SeEligioGridBoleta = True And ml_idMovimiento > 0 Then
        Dim oRsTmp1 As New Recordset
        Set oRsTmp1 = mo_AdminCaja.CajaComprobantesSeleccionarPorId(ml_idMovimiento)
        If oRsTmp1.RecordCount > 0 Then
            Me.txtNserie = oRsTmp1.Fields!NroSerie
            Me.txtNboleta = oRsTmp1.Fields!NroDocumento
            txtNboleta_LostFocus
        End If
        Set oRsTmp1 = Nothing
    End If
End Sub

Private Sub Form_Initialize()
    Set mo_cmbResponsable.MiComboBox = cmbResponsable
End Sub

Private Sub Form_Load()
    wxParametro578 = lcBuscaParametro.SeleccionaFilaParametro(578)
    lblOrdenPago.Caption = ""
    txtFregistro.Text = lcBuscaParametro.RetornaFechaServidorSQL
    txtEstado.Text = "Registrado"
    txtFrealizaCpt.Text = lcBuscaParametro.RetornaFechaHoraServidorSQL
    lbPideInsumo = IIf(lcBuscaParametro.SeleccionaFilaParametro(508) = "S", True, False)
    
    CargaDataCombos
    
    Me.ucProductos.HabilitaIngresoDePrecio = False
    Me.ucProductos.PermiteVerColumnaCantidadFallada = True
    Me.ucProductos.idUsuario = ml_idUsuario
    Me.ucProductos.Inicializar
    Me.ucProductos.IdTipoFinanciamiento = ml_IdTipoFinanciamiento
    Me.ucProductos.TipoProducto = sghServicio
    Me.ucProductos.IdPuntoCarga = ml_PuntoCarga

    Select Case mi_Opcion
    Case sghAgregar
        Me.Caption = "Agregar Tomograf?a"
    Case sghModificar
        Me.Caption = "Modificar Tomograf?a"
    Case sghConsultar
        Me.Caption = "Consultar Tomograf?a"
        btnImprimir.Visible = True
        fraDatosAtencion.Enabled = False
    Case sghEliminar
        Me.Caption = "Eliminar Tomograf?a"
    End Select
    
    CargarDatosAlFormulario
End Sub
Sub CargaMontoDerechoInforme(lnIdTipoFinanciamiento As Long)
    Dim oRsTmp As New Recordset
    Dim lcSql As String
    Set oRsTmp = mo_ReglasFacturacion.CatalogoServiciosHospSeleccionarXidProductoIdTipoFinanciamiento(ml_idDerechoInforme, lnIdTipoFinanciamiento)
    ml_MontoDerechoInforme = 0
    If oRsTmp.RecordCount > 0 Then
       ml_MontoDerechoInforme = oRsTmp.Fields!PrecioUnitario
    End If
    oRsTmp.Close
    Set oRsTmp = Nothing
End Sub
Sub CargarDatosAlFormulario()
 mo_Formulario.HabilitarDeshabilitar Me.txtNmovimiento, False
 mo_Formulario.HabilitarDeshabilitar Me.txtFregistro, False
 mo_Formulario.HabilitarDeshabilitar Me.txtEstado, False
 mo_Formulario.HabilitarDeshabilitar Me.txtDatosDeCuenta, False
 mo_Formulario.HabilitarDeshabilitar Me.txtPlan, False
 mo_Formulario.HabilitarDeshabilitar Me.txtNroOrden, False
 mo_Formulario.HabilitarDeshabilitar Me.txtProcedencia, False
 mo_Formulario.HabilitarDeshabilitar Me.txtNombreDx, False
 mo_Formulario.HabilitarDeshabilitar Me.txtPorcIR, False
 mo_Formulario.HabilitarDeshabilitar Me.cmbFormaPago, False
 mo_Formulario.HabilitarDeshabilitar Me.txtDerechoI, False
 wxParametro302 = lcBuscaParametro.SeleccionaFilaParametro(302)
 wxParametro509 = lcBuscaParametro.SeleccionaFilaParametro(509)
 Me.UcPacienteDatos1.Inicializar

 Select Case mi_Opcion
     Case sghAgregar
        Me.ucProductos.IdOrden = 0
        Me.ucProductos.CargaProductosPorIdOrden
        CargaBoletaAutomaticamente
     Case sghModificar
        CargarDatosALosControles
     Case sghConsultar
        CargarDatosALosControles
     Case sghEliminar
        CargarDatosALosControles
 End Select
End Sub

Sub CargarDatosALosControles()
        lnMedicoId = 0
        mo_Formulario.HabilitarDeshabilitar Me.txtNcuenta, False
        mo_Formulario.HabilitarDeshabilitar Me.txtNserie, False
        mo_Formulario.HabilitarDeshabilitar Me.txtNboleta, False
        Me.chkPlanNoCubre.Visible = False: txtDatosDeCuenta.Width = txtDatosDeCuenta.Width + chkPlanNoCubre.Width + 210
        cmdBuscaCuentaPorApellidos.Enabled = False
        
        'Carga datos de la orden
        Dim oRsTmp As New Recordset
        Dim lbSigue As Boolean, lbSeguirConCuentaCerrada As Boolean
        Dim oConexion As New Connection
        Dim oFactOrdenServicio As New FactOrdenServicio
        oConexion.CommandTimeout = 300
        oConexion.CursorLocation = adUseClient
        oConexion.Open SIGHEntidades.CadenaConexion
        Set oRsTmp = mo_ReglasImagenes.ImagMovimientoImagenesSeleccionarPorIdMovimiento(ml_idMovimiento)
        If oRsTmp.RecordCount > 0 Then
            With oDoImagMovimiento
                .idMovimiento = ml_idMovimiento
                .fecha = oRsTmp.Fields!fecha
                .IdImagEstado = oRsTmp.Fields!IdImagEstado
                .IdPuntoCarga = oRsTmp.Fields!IdPuntoCarga
                .IdTipoConcepto = oRsTmp.Fields!IdTipoConcepto
                .MovTipo = oRsTmp.Fields!MovTipo
                .idUsuario = oRsTmp.Fields!idUsuario
            End With
            With oDoImagMovimientoImagenes
                .idMovimiento = ml_idMovimiento
                .CorrelativoAnual = IIf(IsNull(oRsTmp.Fields!CorrelativoAnual), 0, oRsTmp.Fields!CorrelativoAnual)
                .idComprobantePago = IIf(IsNull(oRsTmp.Fields!idComprobantePago), 0, oRsTmp.Fields!idComprobantePago)
                .idCuentaAtencion = IIf(IsNull(oRsTmp.Fields!idCuentaAtencion), 0, oRsTmp.Fields!idCuentaAtencion)
                .IdOrden = oRsTmp.Fields!IdOrden
                .IdPersonaTomaImagen = IIf(IsNull(oRsTmp.Fields!IdPersonaTomaImagen), 0, oRsTmp.Fields!IdPersonaTomaImagen)
                .idPersonaRecoge = IIf(IsNull(oRsTmp.Fields!idPersonaRecoge), 0, oRsTmp.Fields!idPersonaRecoge)
                .PorcInformeRadiolog = IIf(IsNull(oRsTmp.Fields!PorcInformeRadiolog), 0, oRsTmp.Fields!PorcInformeRadiolog)
                .ResultadoFinal = IIf(IsNull(oRsTmp.Fields!ResultadoFinal), "", oRsTmp.Fields!ResultadoFinal)
                .ZonaRayosX = IIf(IsNull(oRsTmp.Fields!ZonaRayosX), "", oRsTmp.Fields!ZonaRayosX)
                .EsContraste = IIf(IsNull(oRsTmp!EsContraste), 0, oRsTmp!EsContraste)
                .EsContrasteIonico = IIf(IsNull(oRsTmp!EsContrasteIonico), 0, oRsTmp!EsContrasteIonico)
                .idDiagnostico = IIf(IsNull(oRsTmp!idDiagnostico), 0, oRsTmp!idDiagnostico)
                .EsDiagnosticoDefinitivo = IIf(IsNull(oRsTmp!EsDiagnosticoDefinitivo), 0, oRsTmp!EsDiagnosticoDefinitivo)
            End With
            '
            oDoFactOrdenServ.IdOrden = oDoImagMovimientoImagenes.IdOrden
            Set oFactOrdenServicio.Conexion = oConexion
            If oFactOrdenServicio.SeleccionarPorId(oDoFactOrdenServ) Then
               Me.txtFrealizaCpt.Text = Format(oDoFactOrdenServ.FechaHoraRealizaCpt, SIGHEntidades.DevuelveFechaSoloFormato_DMY_HM)
            End If
            '
            txtFregistro.Text = Format(oDoImagMovimiento.fecha, SIGHEntidades.DevuelveFechaSoloFormato_DMY)
            txtEstado.Text = oRsTmp.Fields!destado
            txtNmovimiento.Text = ml_idMovimiento
            txtNcuenta.Text = oDoImagMovimientoImagenes.idCuentaAtencion
            txtNroOrden.Text = oDoImagMovimientoImagenes.IdOrden
            txtResultadoFinal.Text = oDoImagMovimientoImagenes.ResultadoFinal
            chkContraste.Value = IIf(oDoImagMovimientoImagenes.EsContraste = 1, 1, 0)
            If oDoImagMovimientoImagenes.EsContraste = 1 Then
               chkIonico.Value = IIf(oDoImagMovimientoImagenes.EsContrasteIonico = 1, 1, 0)
            End If
            txtPorcIR.Text = oDoImagMovimientoImagenes.PorcInformeRadiolog
            'Dx
            Dim mo_Diagnostico As New DODiagnostico
            ml_IdDiagnostico = oDoImagMovimientoImagenes.idDiagnostico
            If ml_IdDiagnostico > 0 Then
                Set mo_Diagnostico = mo_reglasComunes.DiagnosticosSeleccionarPorId(ml_IdDiagnostico)
                txtDx.Text = mo_Diagnostico.CodigoCIE2004
                txtNombreDx.Text = mo_Diagnostico.descripcion
                chkDxDefinitivo.Visible = True
                chkDxDefinitivo.Value = IIf(oDoImagMovimientoImagenes.EsDiagnosticoDefinitivo = 1, 1, 0)
                mo_Formulario.HabilitarDeshabilitar txtDx, False
            End If
            '
            mo_cmbResponsable.BoundText = oDoImagMovimientoImagenes.IdPersonaTomaImagen
            ml_IdServicioPaciente = IIf(IsNull(oRsTmp.Fields!idServicioPaciente), 0, oRsTmp.Fields!idServicioPaciente)
            ml_IdPaciente = IIf(IsNull(oRsTmp.Fields!idPaciente), 0, oRsTmp.Fields!idPaciente)
            ml_IdFuenteFinanciamiento = oRsTmp.Fields!idFuenteFinanciamiento
            ml_IdTipoFinanciamiento = oRsTmp.Fields!IdTipoFinanciamiento
            ml_IdFuenteFinanciamientoDespacho = oRsTmp.Fields!idFuenteFinanciamiento
            lnIdPacienteHistorico = ml_IdPaciente
            '
            UcPacienteDatos1.FechaRegistro = CDate(txtFregistro.Text)
            If ml_IdPaciente = 0 Then
                If Not IsNull(oRsTmp.Fields!FechaNacimiento) Then
                   UcPacienteDatos1.FechaNacimiento = oRsTmp.Fields!FechaNacimiento
                End If
                If Not IsNull(oRsTmp.Fields!idTipoSexo) Then
                   UcPacienteDatos1.idTipoSexo = oRsTmp.Fields!idTipoSexo
                End If
                UcPacienteDatos1.CargaAlgunosDatosDesdeBoleta oRsTmp.Fields!Paciente
            Else
                UcPacienteDatos1.idPaciente = ml_IdPaciente
                UcPacienteDatos1.CargarDatosDePacienteALosControles
            End If
            '
            If mo_Teclado.TextoEsSoloNumeros(txtNcuenta.Text) Then
                Set oRsTmp = mo_ReglasFarmacia.AtencionesSelecionarPorCuenta(txtNcuenta.Text, oConexion)
                lbSigue = True
                If oRsTmp.RecordCount > 0 Then
                   lnMedicoId = IIf(IsNull(oRsTmp!idMedicoIngreso), 0, oRsTmp!idMedicoIngreso)
                   If oRsTmp.Fields!idEstado <> 1 Then
                      If mi_Opcion <> sghConsultar Then
                         '
                         lbSeguirConCuentaCerrada = True
                         If mi_Opcion = sghModificar And oRsTmp!idTipoServicio = sghTipoServicio.sghEmergenciaConsultorios Then
                           If mo_reglasComunes.HospitalizadoConCtaEmergNOabierta(ml_IdPaciente, _
                              Format(oRsTmp!FechaEgreso & " " & oRsTmp!horaEgreso, SIGHEntidades.DevuelveFechaSoloFormato_DMY_HM), _
                              oRsTmp!idDestinoAtencion) = True Then
                              lbSeguirConCuentaCerrada = False
                              ucProductos.habilitar False
                              cmbResponsable.Enabled = False
                              UcPacienteDatos1.habilitar False
                              MsgBox "Ese estado de Cuenta no se encuentra ABIERTA" & Chr(13) & _
                                     "    solo podr? registrar RESULTADO FINAL    ", vbInformation, Me.Caption
                           End If
                         End If
                         '
                         If lbSeguirConCuentaCerrada = True Then
                            MsgBox "Ese estado de Cuenta no se encuentra ABIERTA", vbInformation, Me.Caption
                            If mi_Opcion = sghModificar Or mi_Opcion = sghEliminar Then
                               btnAceptar.Enabled = False
                            Else
                               lbSigue = False
                            End If
                         End If
                      End If
                   End If
                   If lbSigue Then
                         lnEpsPorcentaje = mo_ReporteUtil.DevuelveEpsPorcentaje(oRsTmp!EpsPorcentaje)
                         mo_Formulario.HabilitarAlerta txtPlan, IIf(lnEpsPorcentaje > 0, True, False)
                         If lnEpsPorcentaje > 0 Then
                            Dim lcBoletaEPS As String
                            lblOrdenPago.Tag = mo_ReglasFacturacion.DevuelveOrdenPago(oRsTmp!idAtencion, sghPtoCargaCaja, oDoImagMovimiento.fecha, oConexion, lcBoletaEPS)
                            lblOrdenPago.Caption = "N? Orden de Pago: " & lblOrdenPago.Tag
                            If lcBoletaEPS <> "" Then
                                lblOrdenPago.Caption = lcBoletaEPS
                                MsgBox "El SEGURO tiene EPS, No podr? MODIFICAR/ELIMINAR porque ya pag? en CAJA" & Chr(13) & _
                                       "Tendr?a que ANULAR (o NOTA DE CREDITO) la BOLETA para usar MODIFICAR/ELIMINAR", vbInformation, ""
                                Me.btnAceptar.Enabled = False
                            End If
                            
                         End If
                         lnIdTipoServicio = oRsTmp.Fields!idTipoServicio
                         txtDatosDeCuenta.Text = "F.Ing: " & oRsTmp.Fields!fechaingreso & " - " & IIf(oRsTmp.Fields!idTipoServicio = 1, "Consultorios Externos", IIf(oRsTmp.Fields!idTipoServicio = 3, "Hospitalizaci?n", "Emergencia")) & " - (Est: " & Trim(oRsTmp.Fields!estadoCta) & ")"
                         txtPlan.Text = "IAFA Act.: " & oRsTmp.Fields!dFuenteFinanciamiento & mo_ReporteUtil.DevuelveEPScubre(lnEpsPorcentaje)
                         'debb-14/04/2011
                         If mi_Opcion = sghModificar And oRsTmp.Fields!idFuenteFinanciamiento <> ml_IdFuenteFinanciamientoDespacho Then
                            MsgBox "No se podr? modificar datos, porque el despacho tubo otro PRODUCTO/PLAN" & Chr(13) & "hubo RECALCULO", vbInformation, Me.Caption
                            btnAceptar.Enabled = False
                         End If
                   End If
               End If
               '
               Set oRsTmp = mo_reglasComunes.RecetaCabeceraFiltraXcuentaYDocumentodespacho(txtNmovimiento.Text, Val(txtNcuenta.Text))
               lnIdReceta = 0
               ucProductos.PermiteAgregarItems = True
               If oRsTmp.RecordCount > 0 Then
                  lnIdReceta = oRsTmp.Fields!IdReceta
                  ucProductos.PermiteAgregarItems = False
               End If
               '
               txtProcedencia.Text = mo_ReglasFacturacion.BuscaServicioActualDelPaciente(ml_IdServicioPaciente)
               UcPacienteDatos1.DeshabilitarFrames True
               '
               chkMuestraHistorico.Value = 1
               chkMuestraHistorico_Click
            Else
                Dim oDOCajaComprobantesPago As New DOCajaComprobantesPago
                Set oDOCajaComprobantesPago = mo_AdminCaja.ComprobantePagoSeleccionarPorId(oRsTmp.Fields!idComprobantePago, oConexion)
                txtNserie.Text = oDOCajaComprobantesPago.NroSerie
                txtNboleta.Text = oDOCajaComprobantesPago.NroDocumento
                ucProductos.PermiteAgregarItems = False
                UcPacienteDatos1.DeshabilitarFrames False
                If ml_IdServicioPaciente > 0 Then
                   'Paciente contado, con cuenta (CE), pago en CAJA
                   ml_IdServicioPaciente = mo_ReglasFarmacia.DevuelveServicioDondeSeEncuentraElPacienteSegunFechaHora(oDOCajaComprobantesPago.idCuentaAtencion, CDate(txtFregistro.Text), lcBuscaParametro.RetornaHoraServidorSQL)
                   txtProcedencia.Text = mo_ReglasFacturacion.BuscaServicioActualDelPaciente(ml_IdServicioPaciente)
                   UcPacienteDatos1.DeshabilitarFrames True
                End If
            End If
            If oDoImagMovimiento.IdImagEstado = 0 Or mi_Opcion = sghConsultar Then
               btnAceptar.Enabled = False
            End If
            cmbFormaPago.BoundText = ml_IdTipoFinanciamiento
            mb_ExistenDatos = True
         Else
            mb_ExistenDatos = False
            Exit Sub
        End If
        
        
        
        'Cargar datos de los servicios
        Me.ucProductos.LimpiarGrilla
        Me.ucProductos.idMovimiento = ml_idMovimiento
        Me.ucProductos.IdTipoFinanciamiento = ml_IdTipoFinanciamiento
        Me.ucProductos.CargaProductosPorIdMovimiento
        Me.ucProductos.CargaObservacionesDeReceta lnIdReceta, oConexion
        ml_MontoDerechoInforme = Me.ucProductos.MontoDerechoInforme
        txtDerechoI.Text = ml_MontoDerechoInforme
        
        UcPacienteDatos1.CargarDatosAlObjetoDatos oDOPaciente
        lblOrdenPago.Caption = mo_ReglasImagenes.DevuelveNombreArchivoImagenes(oDoImagMovimiento.idMovimiento, _
           mo_ReglasImagenes.DevuelveIdPacienteParaLeerImagenes(oDOPaciente, oDoImagMovimientoImagenes), _
           Me.ucProductos.FacturacionProductos) & _
           lblOrdenPago.Caption
        
        mo_ReglasImagenes.YaExisteUnResultado btnAceptar, oDoImagMovimiento.idMovimiento
        
        Select Case mi_Opcion
        Case sghModificar
        Case sghEliminar
        Case sghConsultar
        End Select
        oConexion.Close
        Set oFactOrdenServicio = Nothing
        Set oConexion = Nothing
End Sub


Private Sub cmdBuscaCuentaPorApellidos_Click()
    Dim oBusqueda As New SIGHNegocios.BuscaPacientes
    Dim oDOPaciente As New doPaciente
    Dim oConexion As New Connection
    oConexion.Open SIGHEntidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    oBusqueda.TipoFiltro = sghFiltrarTodos
    oBusqueda.MostrarFormulario
    If oBusqueda.BotonPresionado = sghAceptar Then
        Set oDOPaciente = mo_AdminAdmision.PacientesSeleccionarPorId(oBusqueda.IdRegistroSeleccionado, oConexion)
        If Not oDOPaciente Is Nothing Then
            ml_IdPaciente = oDOPaciente.idPaciente
            Dim oRsTmp As New Recordset
            Set oRsTmp = mo_ReglasFarmacia.FacturacionCuentasAtencionSeleccionarPorIdPaciente(ml_IdPaciente, oConexion, True)
            If oRsTmp.RecordCount > 0 Then
               txtNcuenta.Text = oRsTmp.Fields!idCuentaAtencion
            End If
            oRsTmp.Close
            Set oRsTmp = Nothing
            txtNcuenta_LostFocus
        End If
    End If
    oConexion.Close
    Set oConexion = Nothing
End Sub




Private Sub Form_Unload(Cancel As Integer)
    LimpiarVariablesDeMemoria
End Sub



Private Sub grdConsumoPaciente_InitializeLayout(ByVal Context As UltraGrid.Constants_Context, ByVal Layout As UltraGrid.SSLayout)
     Layout.Override.HeaderClickAction = ssHeaderClickActionSortMulti
     grdConsumoPaciente.Top = fraDatosAtencion.Top + 1000
     grdConsumoPaciente.Bands(0).Columns("Fecha").Width = 800
     grdConsumoPaciente.Bands(0).Columns("idMovimiento").Width = 700
     grdConsumoPaciente.Bands(0).Columns("Codigo").Width = 500
     grdConsumoPaciente.Bands(0).Columns("Nombre").Width = 2500
     grdConsumoPaciente.Bands(0).Columns("Cantidad").Width = 300

End Sub

Private Sub txtFrealizaCpt_LostFocus()
If Not IsDate(txtFrealizaCpt.Text) Then
        MsgBox "La fecha ingresada no es v?lida", vbInformation, ""
        txtFrealizaCpt.Text = SIGHEntidades.FECHA_VACIA_DMY_HM
        Exit Sub
    End If
End Sub

Private Sub txtNboleta_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtNboleta
End Sub

Private Sub txtNboleta_KeyPress(KeyAscii As Integer)
       If Not mo_Teclado.CodigoAsciiEsNumero(KeyAscii) Then
           KeyAscii = 0
       End If
End Sub

Private Sub txtNboleta_LostFocus()
    If Trim(txtNserie.Text) <> "" And Trim(txtNboleta.Text) <> "" Then
        lnMedicoId = 0
        lnEpsPorcentaje = 0
        lnUltimaBusqueda = sghEnBoleta
        Dim rsBuscaBoleta As New Recordset
        Dim rsBuscaBoletaEnImagenes As New Recordset
        Dim oConexion As New Connection
        oConexion.CommandTimeout = 300
        oConexion.Open SIGHEntidades.CadenaConexion
        oConexion.CursorLocation = adUseClient
        Set rsBuscaBoleta = mo_AdminCaja.CajaComprobantePagoServiciosPorNroSerieNroDocumentoConexion(txtNserie.Text, Trim(txtNboleta.Text), oConexion)
        If rsBuscaBoleta.RecordCount > 0 Then
            '
            lnIdPacienteHistorico = 0
            If rsBuscaBoleta.Fields!idPaciente > 0 Then
               lnIdPacienteHistorico = rsBuscaBoleta.Fields!idPaciente
               chkMuestraHistorico_Click
            End If
            '
            If rsBuscaBoleta.Fields!idEstadoComprobante <> sghEstadosComprobante.sighEstadosComprobantePagado Then
                MsgBox "Esa Boleta est? ANULADA", vbInformation, Me.Caption
                txtNboleta.Text = ""
                txtNserie.Text = ""
                ml_IdComprobantePago = 0
            Else
                Set rsBuscaBoletaEnImagenes = mo_ReglasImagenes.ImagMovimientoImagenesSeleccionarPorIdComprobantePago(rsBuscaBoleta.Fields!idComprobantePago)
                If rsBuscaBoletaEnImagenes.RecordCount > 0 Then
                    MsgBox "Esa Boleta ya fu? DESPACHADA con N? Movimiento: " & Chr(13) & rsBuscaBoletaEnImagenes.Fields!idMovimiento & "      y fecha: " & rsBuscaBoletaEnImagenes.Fields!fecha, vbInformation, Me.Caption
                    txtNboleta.Text = ""
                    txtNserie.Text = ""
                    ml_IdComprobantePago = 0
                Else
                    Set rsBuscaBoletaEnImagenes = mo_AdminCaja.FactOrdenServicioSeleccionarPuntoCargaPorIdOrden(rsBuscaBoleta.Fields!IdOrden)
                    If rsBuscaBoletaEnImagenes.RecordCount > 0 Then
                        ml_IdTipoFinanciamiento = rsBuscaBoletaEnImagenes.Fields!IdTipoFinanciamiento     'Contado
                        ml_IdFuenteFinanciamiento = rsBuscaBoletaEnImagenes.Fields!idFuenteFinanciamiento 'contado
                    End If
                    ml_IdComprobantePago = rsBuscaBoleta.Fields!idComprobantePago
                    txtNroOrden.Text = rsBuscaBoleta.Fields!IdOrden
                    If rsBuscaBoleta.Fields!idPaciente > 0 And rsBuscaBoleta.Fields!idCuentaAtencion > 0 Then
                       'Paciente contado, con cuenta (CE), pago en CAJA
                       ml_IdServicioPaciente = mo_ReglasFarmacia.DevuelveServicioDondeSeEncuentraElPacienteSegunFechaHora(rsBuscaBoleta.Fields!idCuentaAtencion, CDate(txtFregistro.Text), lcBuscaParametro.RetornaHoraServidorSQL)
                       txtProcedencia.Text = mo_ReglasFacturacion.BuscaServicioActualDelPaciente(ml_IdServicioPaciente)
                       UcPacienteDatos1.FechaRegistro = CDate(txtFregistro.Text)
                       UcPacienteDatos1.idPaciente = rsBuscaBoleta.Fields!idPaciente
                       UcPacienteDatos1.CargarDatosDePacienteALosControles
                       UcPacienteDatos1.DeshabilitarFrames True
                    ElseIf rsBuscaBoleta.Fields!idPaciente > 0 Then
                       'Paciente con Nro Historia
                       UcPacienteDatos1.FechaRegistro = CDate(txtFregistro.Text)
                       UcPacienteDatos1.idPaciente = rsBuscaBoleta.Fields!idPaciente
                       UcPacienteDatos1.CargarDatosDePacienteALosControles
                       UcPacienteDatos1.DeshabilitarFrames True
                    Else
                       'Paciente contado, EXTERNO
                       UcPacienteDatos1.CargaAlgunosDatosDesdeBoleta (rsBuscaBoleta.Fields!razonSocial)
                       UcPacienteDatos1.DeshabilitarFrames False
                       UcPacienteDatos1.FechaRegistro = CDate(txtFregistro.Text)
                    End If
                    ucProductos.NoPermiteCargarCantidadFallada = True
                    ucProductos.IdTipoFinanciamiento = ml_IdTipoFinanciamiento
                    ucProductos.PermiteAgregarItems = False
                    ucProductos.LimpiarGrilla
                    ucProductos.CargarItemsALaGrillaCPT rsBuscaBoleta, True
                    ml_MontoDerechoInforme = ucProductos.MontoDerechoInforme
                    txtDerechoI.Text = ml_MontoDerechoInforme
                    txtNcuenta.Text = ""
                    txtDatosDeCuenta.Text = ""
                    txtPlan.Text = ""
                    On Error Resume Next
                    cmbResponsable.SetFocus
                End If
            End If
        End If
        Set rsBuscaBoleta = Nothing
        Set rsBuscaBoletaEnImagenes = Nothing
        oConexion.Close
        Set oConexion = Nothing
    End If
End Sub

Private Sub txtNcuenta_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtNcuenta
End Sub


Private Sub txtNcuenta_LostFocus()
   If Val(txtNcuenta.Text) = 0 And txtNcuenta.Locked = False Then
      txtNserie.SetFocus
      Exit Sub
   End If
   If mo_Teclado.TextoEsSoloNumeros(txtNcuenta.Text) And txtNcuenta.Locked = False Then
       lnMedicoId = 0
       lnUltimaBusqueda = sghEnNroCuenta
       Dim oRsTmp As New Recordset
       Dim lbSigue As Boolean
       Dim oConexion As New Connection
       oConexion.Open SIGHEntidades.CadenaConexion
       oConexion.CursorLocation = adUseClient
       Set oRsTmp = mo_ReglasFarmacia.AtencionesSelecionarPorCuenta(txtNcuenta.Text, oConexion)
       lbSigue = True
       If oRsTmp.RecordCount > 0 Then
          lnMedicoId = IIf(IsNull(oRsTmp!idMedicoIngreso), 0, oRsTmp!idMedicoIngreso)
          If oRsTmp.Fields!idEstado <> 1 Then
             If mi_Opcion <> sghConsultar And lbCuentaDeEmergenciaCerrada = False Then
                MsgBox "Ese estado de Cuenta no se encuentra ABIERTA", vbInformation, Me.Caption
                If mi_Opcion = sghModificar Or mi_Opcion = sghEliminar Then
                   btnAceptar.Enabled = False
                Else
                   lbSigue = False
                End If
             End If
          End If
          If mi_Opcion = sghAgregar And _
             mo_AdminAdmision.AtencionesDatosAdicionalesSItieneCodigoPrestacionSIS(Val(txtNcuenta.Text), wxParametro302, _
                                                                          oRsTmp.Fields!idFuenteFinanciamiento) = False Then
                                                                       
             lbSigue = False
          End If
          '
          If lbSigue = True And oRsTmp!esPacienteExterno <> True And wxParametro509 = "S" And mi_Opcion = sghAgregar Then
             If Val(txtNreceta.Text) = 0 Then
                MsgBox "No puede usar N? CUENTA, tiene que generar RECETA", vbInformation, Me.Caption
                lbSigue = False
             End If
          End If
          '
          
          If mi_Opcion = sghAgregar And oRsTmp.Fields!idTipoServicio = sghTipoServicio.sghConsultaExterna _
                                                                                And oRsTmp.Fields!IdFormaPago = 1 Then
                MsgBox "Es un Paciente PAGANTE y viene por CONSULTORIO EXTERNO" & Chr(13) & _
                        "Debe pagar antes en CAJA", vbInformation, "Im?genes"
                lbSigue = False
          End If
          If mi_Opcion = sghAgregar And _
                                    mo_AdminAdmision.LaFechaDespachoEsMenorAfechaCita(CDate(Format(oRsTmp!fechaingreso, _
                                    SIGHEntidades.DevuelveFechaSoloFormato_DMY) & " " & oRsTmp!horaIngreso)) = True Then
             lbSigue = False
          End If
          
          If lbSigue Then
                lnEpsPorcentaje = mo_ReporteUtil.DevuelveEpsPorcentaje(oRsTmp!EpsPorcentaje)
                mo_Formulario.HabilitarAlerta txtPlan, IIf(lnEpsPorcentaje > 0, True, False)
                lnIdTipoServicio = oRsTmp.Fields!idTipoServicio
                txtDatosDeCuenta.Text = "F.Ing: " & oRsTmp.Fields!fechaingreso & " - " & _
                            IIf(oRsTmp!esPacienteExterno = True, "Externo", _
                            IIf(oRsTmp.Fields!idTipoServicio = 1, "Consultorios Externos", _
                            IIf(oRsTmp.Fields!idTipoServicio = 3, "Hospitalizaci?n", "Emergencia"))) & _
                            " - (Est: " & Trim(oRsTmp.Fields!estadoCta) & ")"
                txtPlan.Text = "IAFA Act.: " & oRsTmp.Fields!dFuenteFinanciamiento & mo_ReporteUtil.DevuelveEPScubre(lnEpsPorcentaje)
                ml_IdPaciente = oRsTmp.Fields!idPaciente
                ml_IdFuenteFinanciamiento = oRsTmp.Fields!idFuenteFinanciamiento
                ml_IdTipoFinanciamiento = oRsTmp.Fields!IdFormaPago
                UcPacienteDatos1.idPaciente = ml_IdPaciente
                UcPacienteDatos1.FechaRegistro = CDate(txtFregistro.Text)
                UcPacienteDatos1.CargarDatosDePacienteALosControles
                UcPacienteDatos1.DeshabilitarFrames True
                ucProductos.IdTipoFinanciamiento = oRsTmp.Fields!IdFormaPago
                ucProductos.PermiteAgregarItems = True
                '
                ml_IdServicioPaciente = mo_ReglasFarmacia.DevuelveServicioDondeSeEncuentraElPacienteSegunFechaHora(Val(txtNcuenta.Text), CDate(txtFregistro.Text), lcBuscaParametro.RetornaHoraServidorSQL)
                txtProcedencia.Text = mo_ReglasFacturacion.BuscaServicioActualDelPaciente(ml_IdServicioPaciente)
                '
                txtNserie.Text = ""
                txtNboleta.Text = ""
                ml_IdComprobantePago = 0
                txtNroOrden.Text = ""
                If mi_Opcion <> sghAgregar And ml_IdFuenteFinanciamiento <> ml_IdFuenteFinanciamientoDespacho And ml_IdFuenteFinanciamientoDespacho > 0 Then
                   txtPlan.Text = "Plan Desp: " & Trim(mo_ReglasFacturacion.FuentesFinanciamientoDevuelveNombrePlan(ml_IdFuenteFinanciamientoDespacho)) & " - " & txtPlan.Text
                End If
                cmbFormaPago.BoundText = ml_IdTipoFinanciamiento
                If mi_Opcion = sghAgregar Then
                   CargaMontoDerechoInforme ml_IdTipoFinanciamiento
                   ucProductos.MontoDerechoInforme = ml_MontoDerechoInforme
                   Me.txtDerechoI.Text = ml_MontoDerechoInforme
                End If
                '
                lnIdPacienteHistorico = ml_IdPaciente
                chkMuestraHistorico.Value = 1
                chkMuestraHistorico_Click
                '
                mo_Formulario.HabilitarDeshabilitar txtDx, True
                Set oRsTmp = mo_AdminAdmision.AtencionesDiagnosticosSeleccionarTodosPorIdAtencion(oRsTmp.Fields!idAtencion)
                If oRsTmp.RecordCount > 0 Then
                   txtDx.Text = oRsTmp.Fields!CodigoCIE2004
                   txtNombreDx.Text = oRsTmp.Fields!descripcion
                   Me.chkDxDefinitivo.Value = 0
                   mo_Formulario.HabilitarDeshabilitar txtDx, False
                   ml_IdDiagnostico = oRsTmp!idDiagnostico
                ElseIf wxParametro578 = "S" Then
                     MsgBox "No se puede despachar sino tiene Dx", vbInformation, ""
                     txtDatosDeCuenta.Text = ""
                     txtPlan.Text = ""
                     txtProcedencia.Text = ""
                End If
          Else
                txtNreceta.Text = ""
          
          End If
       End If
       oRsTmp.Close
       Set oRsTmp = Nothing
       oConexion.Close
       Set oConexion = Nothing
   End If
End Sub



Private Sub txtNreceta_KeyDown(KeyCode As Integer, Shift As Integer)
       mo_Teclado.RealizarNavegacion KeyCode, txtNreceta
       AdministrarKeyPreview KeyCode
End Sub

Private Sub txtNserie_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtNserie
End Sub

Private Sub txtNserie_KeyPress(KeyAscii As Integer)
'       If Not mo_Teclado.CodigoAsciiEsNumero(KeyAscii) Then
'           KeyAscii = 0
'       End If
End Sub


Sub CargaDataCombos()
    '
    mo_cmbResponsable.BoundColumn = "idEmpleado"
    mo_cmbResponsable.ListField = "ApNom"
    Set mo_cmbResponsable.RowSource = mo_ReglasFarmacia.EmpleadosDeImagen("dbo.EmpleadosCargos.idCargo =" & mo_ReglasFarmacia.EmpleadosDevuelveIdCargoSegunPuntoCarga(ml_PuntoCarga))
    'debb-09/08/2016
    If mo_reglasComunes.NOpuedeModificarResponsable(mi_Opcion, ml_idUsuario, mo_cmbResponsable.RowSource) Then
       If mi_Opcion = sghAgregar Then
          mo_cmbResponsable.BoundText = Trim(Str(ml_idUsuario))
       End If
       mo_Formulario.HabilitarDeshabilitar Me.cmbResponsable, False
    End If
    '
    Set oRsFormaPago = mo_reglasComunes.TiposFinanciamientoSegunFiltro("esFarmacia=1")
    Set cmbFormaPago.RowSource = oRsFormaPago
    cmbFormaPago.ListField = "Descripcion"
    cmbFormaPago.BoundColumn = "idTipoFinanciamiento"
    '
    txtPorcIR.Text = lcBuscaParametro.SeleccionaFilaParametro(230)
    ml_idDerechoInforme = lcBuscaParametro.SeleccionaFilaParametro(250)
End Sub


Private Sub txtDx_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtDx
    AdministrarKeyPreview KeyCode
End Sub


Private Sub txtDx_LostFocus()
        Dim oDODiagnostico As DODiagnostico
        Set oDODiagnostico = mo_reglasComunes.DiagnosticosSeleccionarPorCodigoCIE2004(txtDx.Text, True)
        If Not oDODiagnostico Is Nothing Then
            ml_IdDiagnostico = oDODiagnostico.idDiagnostico
            txtNombreDx.Text = oDODiagnostico.descripcion
            chkDxDefinitivo.Visible = True
        Else
            ml_IdDiagnostico = 0
            txtNombreDx.Text = ""
            chkDxDefinitivo.Visible = False
        End If

End Sub




Private Sub txtResultadoFinal_KeyDown(KeyCode As Integer, Shift As Integer)
      ' mo_Teclado.RealizarNavegacion KeyCode, txtResultadoFinal
       AdministrarKeyPreview KeyCode
End Sub

Private Sub txtResultadoFinal_LostFocus()
    If Val(txtNreceta.Text) = 0 Then
       ucProductos.TabEnDescripcion
    End If
End Sub



Private Sub UcPacienteDatos1_SePresionoTeclaEspecial(KeyCode As Integer)
    If KeyCode = vbKeyF2 Then
       AdministrarKeyPreview KeyCode
    End If
End Sub

Private Sub ucProductos_SePresionoTeclaEspecial(KeyCode As Integer)
     If KeyCode = vbKeyF2 Then
        AdministrarKeyPreview KeyCode
     End If

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    AdministrarKeyPreview KeyCode
End Sub

Sub AdministrarKeyPreview(KeyCode As Integer)
   Select Case KeyCode
        Case vbKeyF6
        Case vbKeyEscape
'           btnCancelar_Click
        Case vbKeyF2
           btnAceptar_Click
       End Select
End Sub


Sub LimpiarVariablesDeMemoria()
    On Error Resume Next
    Set mo_ReglasImagenes = Nothing
    Set mo_AdminAdmision = Nothing
    Set mo_ReglasFacturacion = Nothing
    Set mo_ReglasFarmacia = Nothing
    Set mo_AdminCaja = Nothing
    Set mo_reglasComunes = Nothing
    Set mo_ReglasSeguridad = Nothing
    Set mo_AdminArchivoClinico = Nothing
    Set mo_Apariencia = Nothing
    Set mo_cmbIdEstado = Nothing
    Set mo_cmbResponsable = Nothing
    Set lcBuscaParametro = Nothing
    Set mo_Teclado = Nothing
    Set mo_Formulario = Nothing
    Set oDOPaciente = Nothing
    Set oDoImagMovimiento = Nothing
    Set oDoImagMovimientoImagenes = Nothing
    Set oDoFactOrdenServ = Nothing
    Set oRsFormaPago = Nothing
End Sub




Private Sub txtNreceta_LostFocus()
    If Val(txtNreceta.Text) > 0 Then
       Dim lcSql As String
       Dim oRsTmp1 As New Recordset, lnRecetaProcesada As Long, lnCuenta As Long
       lnRecetaProcesada = Val(txtNreceta.Text)
       '
       ucProductos.LimpiarGrilla
       Set oRsTmp1 = mo_reglasComunes.RecetasConCabeceraYdetalleSoloCpt(lnRecetaProcesada, sghRecetaEstados.sighRecetaRegistrada)
       If oRsTmp1.RecordCount > 0 Then
            If oRsTmp1.Fields!idEstado <> sghRecetaEstados.sighRecetaRegistrada Then
                mo_reglasComunes.RecetaChequeaEstadoActual oRsTmp1.Fields!idCuentaAtencion, _
                                                           oRsTmp1.Fields!idEstado, _
                                                           0, oRsTmp1.Fields!DocumentoDespacho
                txtNreceta.Text = ""
            Else
                If oRsTmp1.Fields!IdPuntoCarga <> sghPtoCargaTomografia Then
                     MsgBox "Esa receta no es de TOMOGRAFIA", vbInformation, "Im?genes"
                     txtNreceta.Text = ""
                Else
                     lbCuentaDeEmergenciaCerrada = mo_reglasComunes.CuentaDeEmergenciaCerrada(oRsTmp1!idCuentaAtencion, sghPtoCargaTomografia)
                     txtNcuenta.Text = oRsTmp1.Fields!idCuentaAtencion
                     txtNcuenta_LostFocus
                     ucProductos.PermiteAgregarItems = False
                     ucProductos.CargaProductosPorIdReceta oRsTmp1
                     lnIdReceta = lnRecetaProcesada
                     On Error Resume Next
                     Me.cmbResponsable.SetFocus
                End If
            End If
       Else
            MsgBox "Ese N? Receta NO EXISTE", vbInformation, "Caja"
            txtNreceta.Text = ""
       End If
       oRsTmp1.Close
       Set oRsTmp1 = Nothing
    End If
End Sub

Private Sub cmbBuscaReceta_Click()
    Dim oBusqueda As New SIGHNegocios.clBuscaReceta
    oBusqueda.IdPuntoCarga = sghPtoCargaTomografia
    oBusqueda.MostrarFormulario
    If oBusqueda.BotonPresionado = sghAceptar Then
       txtNreceta.Text = oBusqueda.IdRecetaSeleccionada
       txtNreceta_LostFocus
    End If
    Set oBusqueda = Nothing
End Sub


Private Sub btnImprimir_Click()
    Dim oRep As New RayosX
    oRep.ImpresionDelResultado ml_idMovimiento, Me.cmbResponsable.Text, Me.txtFrealizaCpt.Text, 0
    Set oRep = Nothing
End Sub
Private Sub txtNcita_KeyDown(KeyCode As Integer, Shift As Integer)
       mo_Teclado.RealizarNavegacion KeyCode, txtNcita
       AdministrarKeyPreview KeyCode
End Sub

Private Sub txtNcita_LostFocus()
     If Val(txtNcita.Text) > 0 Then
        Dim oConexion As New Connection
        Dim oSiCitas As New SiCitas
        Dim DoSiCitas As New DoSiCitas
        Dim oDOCajaComprobantesPago As New DOCajaComprobantesPago
        oConexion.CursorLocation = adUseClient
        oConexion.CommandTimeout = 900
        oConexion.Open SIGHEntidades.CadenaConexion
        Set oSiCitas.Conexion = oConexion
        DoSiCitas.IdUsuarioAuditoria = SIGHEntidades.Usuario
        DoSiCitas.idCitaSI = Val(txtNcita.Text)
        Me.ucProductos.LimpiarGrilla
        If oSiCitas.SeleccionarPorId(DoSiCitas) = True Then
           If DoSiCitas.IdPuntoCarga <> ml_PuntoCarga Then
                MsgBox "La Cita existe pero NO pertenece al PUNTO DE CARGA", vbInformation, ""
           ElseIf DoSiCitas.idMovimiento > 0 Then
                MsgBox "La CITA ya tiene MOVIMIENTO N? " & DoSiCitas.idMovimiento, vbInformation, ""
           Else
                If DoSiCitas.IdReceta > 0 Then
                   txtNreceta.Text = DoSiCitas.IdReceta
                   txtNreceta_LostFocus
                ElseIf DoSiCitas.idCuentaAtencion > 0 Then
                   txtNcuenta.Text = DoSiCitas.idCuentaAtencion
                   txtNcuenta_LostFocus
                   Me.ucProductos.CargaProductosPorIdCita Val(txtNcita.Text)
                ElseIf DoSiCitas.idComprobantePago > 0 Then
                   Set oDOCajaComprobantesPago = mo_AdminCaja.ComprobantePagoSeleccionarPorId(DoSiCitas.idComprobantePago, oConexion)
                   txtNserie.Text = oDOCajaComprobantesPago.NroSerie
                   txtNboleta.Text = oDOCajaComprobantesPago.NroDocumento
                   txtNboleta_LostFocus
                   UcPacienteDatos1.FechaRegistro = CDate(txtFregistro.Text)
                   If DoSiCitas.idPaciente = 0 Then
                          If DoSiCitas.FechaNacimiento <> 0 Then
                             UcPacienteDatos1.FechaNacimiento = DoSiCitas.FechaNacimiento
                          End If
                          If DoSiCitas.idTipoSexo > 0 Then
                             UcPacienteDatos1.idTipoSexo = DoSiCitas.idTipoSexo
                          End If
                          UcPacienteDatos1.CargaAlgunosDatosDesdeBoleta DoSiCitas.Paciente
                          chkMuestraHistorico.Value = 0
                          lnIdPacienteHistorico = 0
                   Else
                          UcPacienteDatos1.idPaciente = DoSiCitas.idPaciente
                          UcPacienteDatos1.CargarDatosDePacienteALosControles
                          chkMuestraHistorico.Value = 1
                   End If
                   chkMuestraHistorico_Click
                End If
           End If
        End If
        oConexion.Close
        Set oConexion = Nothing
        Set oSiCitas = Nothing
        Set DoSiCitas = Nothing
        Set oDOCajaComprobantesPago = Nothing
     End If
End Sub

