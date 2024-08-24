VERSION 5.00
Begin VB.Form frmMenuPrincipal 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   Caption         =   "Menú Principal"
   ClientHeight    =   9015
   ClientLeft      =   1890
   ClientTop       =   1530
   ClientWidth     =   11910
   DrawMode        =   1  'Blackness
   LinkTopic       =   "Form1"
   Palette         =   "frmMenuPrincipal.frx":0000
   Picture         =   "frmMenuPrincipal.frx":066A
   ScaleHeight     =   9015
   ScaleWidth      =   11910
   StartUpPosition =   2  'CenterScreen
   WhatsThisHelp   =   -1  'True
   Begin VB.Menu mnAlumnos 
      Caption         =   "Alumnos"
      Begin VB.Menu mnAlumnos2 
         Caption         =   "Alumnos"
      End
      Begin VB.Menu mnMensajeriaElectrónica 
         Caption         =   "Mensajeria Electrónica"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
   End
   Begin VB.Menu mnPersonal 
      Caption         =   "Personal"
      Begin VB.Menu mnPersonal2 
         Caption         =   "Personal"
      End
      Begin VB.Menu mnMovimientos 
         Caption         =   "Movimientos"
         Visible         =   0   'False
      End
   End
   Begin VB.Menu mnCarreras 
      Caption         =   "Carreras"
      Begin VB.Menu mnPlanes 
         Caption         =   "Planes"
      End
   End
   Begin VB.Menu mnHorarios 
      Caption         =   "Horarios"
      Begin VB.Menu mnCursadas 
         Caption         =   "Cursadas"
      End
      Begin VB.Menu mnMesas 
         Caption         =   "Mesas"
         Begin VB.Menu mnArmado 
            Caption         =   "Armado"
         End
         Begin VB.Menu mnInformacion 
            Caption         =   "Información"
         End
      End
      Begin VB.Menu mnAsistencia 
         Caption         =   "Asistencia"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
      Begin VB.Menu mnMatriculacion 
         Caption         =   "Matriculación"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
   End
   Begin VB.Menu mnExamenes 
      Caption         =   "Exámenes"
      Begin VB.Menu mnCursadas2 
         Caption         =   "Cursadas"
      End
      Begin VB.Menu mnFinales 
         Caption         =   "Situación Académica"
      End
      Begin VB.Menu mnActas 
         Caption         =   "Actas"
         Begin VB.Menu mnImpresion 
            Caption         =   "Impresión"
         End
         Begin VB.Menu mnIngreso 
            Caption         =   "Ingreso"
         End
      End
      Begin VB.Menu mnEquivalencias 
         Caption         =   "Equivalencias"
      End
   End
   Begin VB.Menu mnCertificados 
      Caption         =   "Certificados"
      Begin VB.Menu mnAnalitico 
         Caption         =   "Certificados Varios"
      End
      Begin VB.Menu mnTitulo 
         Caption         =   "Tìtulo"
      End
      Begin VB.Menu mnAlumRegular 
         Caption         =   "Alumno Regular"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
   End
   Begin VB.Menu mnCooperadora 
      Caption         =   "Cooperadora"
      Enabled         =   0   'False
      Visible         =   0   'False
      Begin VB.Menu mnPeriodosyRecibos 
         Caption         =   "Periodos y Recibos"
      End
      Begin VB.Menu mnIngresoDeCobranza 
         Caption         =   "Ingreso de Cobranza"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnPlan 
         Caption         =   "Plan"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnPagos 
         Caption         =   "Pagos"
         Enabled         =   0   'False
      End
   End
   Begin VB.Menu mnConfiguracion 
      Caption         =   "Configuración"
      Begin VB.Menu mnParametros 
         Caption         =   "Parámetros"
      End
      Begin VB.Menu mnUsuarios 
         Caption         =   "Usuarios"
      End
      Begin VB.Menu mnInformesYControles 
         Caption         =   "Informes y Procesos"
      End
      Begin VB.Menu mnNotificaciones 
         Caption         =   "Notificaciones"
      End
      Begin VB.Menu mnVersion 
         Caption         =   "Versión"
      End
   End
   Begin VB.Menu mnOtros 
      Caption         =   "Sistema Alumnos"
      Begin VB.Menu mnConexionAlumnos 
         Caption         =   "Conexión Alumnos"
      End
      Begin VB.Menu mnSistemaDeTitulosFederales 
         Caption         =   "Sistema de Títulos Federales"
         Enabled         =   0   'False
         Visible         =   0   'False
         Begin VB.Menu mnEnviarInformacionDeTitulos 
            Caption         =   "Enviar Inoformación de Titulos"
         End
      End
   End
   Begin VB.Menu mnSalir 
      Caption         =   "Salir"
   End
End
Attribute VB_Name = "frmMenuPrincipal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Conexion As New Connection

Private Sub Command1_Click()

    
End Sub

Private Sub Form_Load()
    Conexion.ConnectionString = ("DSN=Instituto")
    'para usarlo como demo limito la cargha de alumnos a 20
    Dim rs As ADODB.Recordset
    ' Crear y abrir un Recordset
    
    'codigo para habilitar el exe como demo
    'Conexion.Open  ' abre
    'Set rs = Conexion.Execute("SELECT count(permiso) as Total from alumnos")
    'If rs!Total >= 10 Then
    '   MsgBox ("Este sistema es un Demo"): rs.Close: Set rs = Nothing: Conexion.Close: End
    'End If
    'rs.Close
    'Set rs = Nothing
    'Conexion.Close
    
    'Consulto parametro para saber si usa certificación de servicio o no
    On Error GoTo Gestionaerror
        Conexion.Open
        Set rs = Conexion.Execute("SELECT UsarCertificacionServicio from Parametros")
        'MsgBox (rs!UsarCertificacionServicio)
        If rs!UsarCertificacionServicio = True Then
           mnMovimientos.Visible = True
        Else
            mnMovimientos.Visible = False
        End If
        rs.Close
        Set rs = Nothing
        Conexion.Close
Gestionaerror:

   

End Sub

Private Sub mnAlumnos2_Click()
    frmAlumnos.Show 1
End Sub

Private Sub mnAnalitico_Click()
    If frmIdentificacion.Permisos!ImprimirAnalitico = False Then MsgBox ("Usted no tiene permiso para imprimir Analíticos"): Exit Sub
    frmAnalitico.Show 1
End Sub

Private Sub mnArmado_Click()
    frmMesasArmado.Show 1
End Sub
Private Sub mnConexionAlumnos_Click()
    frmConexionAlumnos.SalirPorCompleto = "No"
    frmConexionAlumnos.Show 1
End Sub

Private Sub mnCursadas_Click()
    frmCursadas.Show 1
End Sub

Private Sub mnCursadas2_Click()
    frmParciales.Show 1
End Sub

Private Sub mnEnviarInformacionDeTitulos_Click()
    frmEnviarInformacionDeTitulos.Show 1
End Sub

Private Sub mnEquivalencias_Click()
    frmEquivalencias.Show 1
End Sub

Private Sub mnFinales_Click()
    frmFinales.Show 1
End Sub

Private Sub mnImpresion_Click()
    frmActasImpresion.Show 1
End Sub

Private Sub mnInformacion_Click()
    frmInformacionMesas.Show 1
End Sub

Private Sub mnInformesyReportes_Click()
    frmInformesyReportes.Show 1
End Sub

Private Sub mnInformesYControles_Click()
    frmInformesyReportes.Show 1
End Sub

Private Sub mnIngreso_Click()
    frmActasIngreso.Show 1
End Sub

Private Sub mnIngresoDeCobranza_Click()
    frmIngresoDeCobranza.Show 1
End Sub

Private Sub mnMensajeriaElectrónica_Click()
    frmMensajeriaElectronica.Show 1
End Sub

Private Sub mnMovimientos_Click()
    frmMovimientos.Show 1
End Sub

Private Sub mnNotificaciones_Click()
    frmNotificaciones.Show 1
End Sub

Private Sub mnPagos_Click()
   frmCooperadoraPagos.Show 1
End Sub

Private Sub mnParametros_Click()
    If frmIdentificacion.Permisos!IngresarParametros = True Then
        frmParametros.Show 1
    Else
        MsgBox ("Usted no tiene permiso para ingresar a este formulario")
    End If
End Sub

Private Sub mnPeriodosyRecibos_Click()
    Respuesta = InputBox("Contraseña")
    If Respuesta <> "asociacion" Then Exit Sub
    frmCooperadora.Show 1
End Sub

Private Sub mnPersonal2_Click()
   frmPersonal.Show 1
End Sub

Private Sub mnPlan_Click()
    If frmIdentificacion.Permisos!IngresarCooperadoraPlan = False Then MsgBox ("Usted no tiene permiso para ingresar al Plan de Cooperadora"): Exit Sub
    frmCooperadoraPlan.Show 1
End Sub

Private Sub mnPlanes_Click()
    frmPlanes.Show 1
End Sub

Private Sub mnSalir_Click()
    Unload Me
End Sub

Private Sub mnTitulo_Click()
    If frmIdentificacion.Permisos!ImprimirTitulo = False Then MsgBox ("Usted no tiene permiso para imprimir Títulos"): Exit Sub
    frmTitulo.Show 1
End Sub

Private Sub mnUsuarios_Click()
    If frmIdentificacion.Permisos!IngresarUsuarios = True Then
        frmUsuarios.Show 1
    Else
        MsgBox ("Usted no tiene permiso para ingresar a este formulario")
    End If
End Sub

Private Sub mnVersion_Click()
    frmVersion.Show 1
End Sub
