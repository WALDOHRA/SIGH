VERSION 5.00
Object = "{5A9433E9-DD7B-4529-91B6-A5E8CA054615}#2.0#0"; "IGULTR~1.OCX"
Begin VB.Form PaquetesBuscar 
   Caption         =   "Busqueda de Paquetes"
   ClientHeight    =   8355
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10785
   Icon            =   "PaquetesBuscar.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8355
   ScaleWidth      =   10785
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7095
      Left            =   30
      TabIndex        =   3
      Top             =   30
      Width           =   10665
      Begin VB.TextBox txtConsideraciones 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   90
         MultiLine       =   -1  'True
         TabIndex        =   6
         Text            =   "PaquetesBuscar.frx":000C
         Top             =   240
         Width           =   10395
      End
      Begin UltraGrid.SSUltraGrid grdPreVentaCab 
         Height          =   2625
         Left            =   90
         TabIndex        =   4
         Top             =   690
         Width           =   10410
         _ExtentX        =   18362
         _ExtentY        =   4630
         _Version        =   131072
         GridFlags       =   17040384
         LayoutFlags     =   67108884
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Lista de Paquetes"
      End
      Begin UltraGrid.SSUltraGrid grdPreVentaDet 
         Height          =   3105
         Left            =   90
         TabIndex        =   5
         Top             =   3420
         Width           =   10470
         _ExtentX        =   18468
         _ExtentY        =   5477
         _Version        =   131072
         GridFlags       =   17040384
         LayoutFlags     =   67108884
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Detalle"
      End
      Begin VB.Label lblTotal 
         Alignment       =   1  'Right Justify
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   8160
         TabIndex        =   7
         Top             =   6630
         Width           =   2295
      End
   End
   Begin VB.Frame Frame3 
      Height          =   1110
      Left            =   30
      TabIndex        =   0
      Top             =   7200
      Width           =   10665
      Begin VB.CommandButton btnCancelar 
         Cancel          =   -1  'True
         Caption         =   "Cancelar (ESC)"
         DisabledPicture =   "PaquetesBuscar.frx":0051
         DownPicture     =   "PaquetesBuscar.frx":0515
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
         Left            =   5430
         Picture         =   "PaquetesBuscar.frx":0A01
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   225
         Width           =   1365
      End
      Begin VB.CommandButton btnAceptar 
         Caption         =   "Aceptar (F2)"
         DisabledPicture =   "PaquetesBuscar.frx":0EED
         DownPicture     =   "PaquetesBuscar.frx":134D
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
         Left            =   3900
         Picture         =   "PaquetesBuscar.frx":17C2
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   225
         Width           =   1365
      End
   End
End
Attribute VB_Name = "PaquetesBuscar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim mo_Formulario As New sighcomun.Formulario
Dim mo_Teclado As New sighcomun.Teclado
Dim mo_Apariencia As New sighcomun.GridInfragistic
'
Dim oRsPreVentaCab As New Recordset
Dim oRsPreVentaDet As New Recordset
'
Dim lcSql As String
Dim mi_BotonPresionado As sghBotonDetallePresionado
Dim mi_idFactPaquete As Long
Dim mi_DebeConsiderarPaquete As sghTipoPaquetes

Property Let DebeConsiderarPaquete(iValue As sghTipoPaquetes)
  mi_DebeConsiderarPaquete = iValue
End Property
Property Let BotonPresionado(iValue As sghBotonDetallePresionado)
  mi_BotonPresionado = iValue
End Property

Property Get BotonPresionado() As sghBotonDetallePresionado
  BotonPresionado = mi_BotonPresionado
End Property

Property Let idFactPaquete(iValue As Long)
  mi_idFactPaquete = iValue
End Property

Property Get idFactPaquete() As Long
  idFactPaquete = mi_idFactPaquete
End Property

Private Sub btnAceptar_Click()
    mi_BotonPresionado = sghAceptar
    Me.Visible = False
End Sub

Private Sub btnCancelar_Click()
        mi_BotonPresionado = sghCancelar
        Me.Visible = False
End Sub

Private Sub Form_Load()
    CargaPreventas
    mo_Apariencia.ConfigurarFilasBiColores Me.grdPreVentaCab, sighcomun.GrillaConFilasBicolor
    mo_Apariencia.ConfigurarFilasBiColores Me.grdPreVentaDet, sighcomun.GrillaConFilasBicolor
End Sub


Private Sub grdPreVentaCab_AfterRowActivate()
Dim rsRecordset As ADODB.Recordset
    Set rsRecordset = grdPreVentaCab.DataSource
    On Error Resume Next
    mi_idFactPaquete = rsRecordset("IdFactPaquete")
End Sub

Private Sub grdPreVentaCab_Click()
        On Error GoTo errDet
        Dim lnTotal As Double
        lcSql = "SELECT      dbo.Especialidades.Nombre AS Especialidad, dbo.FactCatalogoServicios.Codigo, " & _
                "      dbo.FactCatalogoServicios.Nombre AS Procedimiento, dbo.FacturacionCatalogoPaquetes.Cantidad, dbo.FacturacionCatalogoPaquetes.Precio," & _
                "      dbo.FacturacionCatalogoPaquetes.Importe" & _
                " FROM         dbo.FactCatalogoServicios RIGHT OUTER JOIN" & _
                "      dbo.FacturacionCatalogoPaquetes ON dbo.FactCatalogoServicios.IdProducto = dbo.FacturacionCatalogoPaquetes.idProducto LEFT OUTER JOIN" & _
                "      dbo.Especialidades ON dbo.FacturacionCatalogoPaquetes.idEspecialidadServicio = dbo.Especialidades.IdEspecialidad" & _
                " WHERE dbo.FacturacionCatalogoPaquetes.idFactPaquete=" & oRsPreVentaCab.Fields!idFactPaquete
        oRsPreVentaDet.Open lcSql, sighcomun.CadenaConexionShape, adOpenKeyset, adLockOptimistic
        lnTotal = 0
        If oRsPreVentaDet.RecordCount > 0 Then
           oRsPreVentaDet.MoveFirst
           Do While Not oRsPreVentaDet.EOF
              lnTotal = lnTotal + oRsPreVentaDet.Fields!Importe
              oRsPreVentaDet.MoveNext
           Loop
           oRsPreVentaDet.MoveFirst
        End If
        Me.lblTotal.Caption = Format(lnTotal, "####,###,##0.00")
        Set Me.grdPreVentaDet.DataSource = oRsPreVentaDet
        Exit Sub
errDet:
        If Err.Number = 3705 Then
           oRsPreVentaDet.Close
           Resume
        End If
End Sub



Private Sub grdPreVentaCab_InitializeLayout(ByVal Context As UltraGrid.Constants_Context, ByVal Layout As UltraGrid.SSLayout)
    grdPreVentaCab.Bands(0).Columns("idFactPaquete").Hidden = True
    grdPreVentaCab.Bands(0).Columns("Codigo").Activation = ssActivationActivateNoEdit
    grdPreVentaCab.Bands(0).Columns("Codigo").Width = 800
    grdPreVentaCab.Bands(0).Columns("Descripcion").Width = 9000
    grdPreVentaCab.Bands(0).Columns("Descripcion").Activation = ssActivationActivateNoEdit
End Sub

Private Sub grdPreVentaCab_KeyPress(KeyAscii As UltraGrid.SSReturnShort)
    If KeyAscii = 13 Then
       mi_idFactPaquete = 0
       On Error GoTo ErrCab
       mi_idFactPaquete = oRsPreVentaCab.Fields!idFactPaquete
       btnAceptar_Click
    End If
ErrCab:
End Sub

Private Sub grdPreVentaDet_InitializeLayout(ByVal Context As UltraGrid.Constants_Context, ByVal Layout As UltraGrid.SSLayout)
    grdPreVentaDet.Bands(0).Columns("Especialidad").Width = 2000
    grdPreVentaDet.Bands(0).Columns("Especialidad").Activation = ssActivationActivateNoEdit
    grdPreVentaDet.Bands(0).Columns("Codigo").Width = 800
    grdPreVentaDet.Bands(0).Columns("Codigo").Activation = ssActivationActivateNoEdit
    grdPreVentaDet.Bands(0).Columns("Procedimiento").Width = 4500
    grdPreVentaDet.Bands(0).Columns("Procedimiento").Activation = ssActivationActivateNoEdit
    grdPreVentaDet.Bands(0).Columns("Cantidad").Width = 800
    grdPreVentaDet.Bands(0).Columns("Cantidad").Format = "###0"
    grdPreVentaDet.Bands(0).Columns("Cantidad").Activation = ssActivationActivateNoEdit
    grdPreVentaDet.Bands(0).Columns("Precio").Width = 600
    grdPreVentaDet.Bands(0).Columns("Precio").Format = "#0.00"
    grdPreVentaDet.Bands(0).Columns("Precio").Activation = ssActivationActivateNoEdit
    grdPreVentaDet.Bands(0).Columns("Importe").Width = 1200
    grdPreVentaDet.Bands(0).Columns("Importe").Activation = ssActivationActivateNoEdit
End Sub


Sub CargaPreventas()
    lcSql = "SELECT     idFactPaquete, Codigo, Descripcion" & _
            " From dbo.FactCatalogoPaquete"
    Select Case mi_DebeConsiderarPaquete
    Case sghTipoPaqueteSoloFarmacia
         lcSql = lcSql & " Where TipoPaquete=" & sghTipoPaqueteSoloFarmacia
    Case sghTipoPaqueteSoloServicio
         lcSql = lcSql & " Where TipoPaquete=" & sghTipoPaqueteSoloServicio
    End Select
    lcSql = lcSql & " ORDER BY Descripcion"
    oRsPreVentaCab.Open lcSql, sighcomun.CadenaConexion, adOpenKeyset, adLockOptimistic
    Set Me.grdPreVentaCab.DataSource = oRsPreVentaCab
    grdPreVentaCab_Click
End Sub




