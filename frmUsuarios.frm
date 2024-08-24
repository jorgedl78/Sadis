VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "msadodc.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Begin VB.Form frmUsuarios 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Usuarios"
   ClientHeight    =   5700
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7680
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   5700
   ScaleWidth      =   7680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame frComandos 
      Height          =   4575
      Left            =   6000
      TabIndex        =   24
      Top             =   1080
      Width           =   1575
      Begin VB.CommandButton cmdModificar 
         Caption         =   "Modificar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   26
         Top             =   960
         Width           =   1335
      End
      Begin VB.CommandButton cmdSalir 
         Height          =   705
         Left            =   480
         Picture         =   "frmUsuarios.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   25
         ToolTipText     =   "Salir"
         Top             =   3720
         Width           =   720
      End
   End
   Begin VB.Frame frPermisos 
      Caption         =   "Permisos"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4575
      Left            =   0
      TabIndex        =   1
      Top             =   1080
      Width           =   5895
      Begin VB.CheckBox chkCerrarEquivalencias 
         Caption         =   "Cerrar Equivalencias"
         DataField       =   "CerrarEquivalencias"
         DataSource      =   "adoPermisos"
         Height          =   255
         Left            =   3240
         TabIndex        =   30
         Top             =   3600
         Width           =   2415
      End
      Begin VB.CheckBox chkIngresarParametros 
         Caption         =   "Ingresar Parámetros"
         DataField       =   "IngresarParametros"
         DataSource      =   "adoPermisos"
         Height          =   255
         Left            =   3240
         TabIndex        =   29
         Top             =   3240
         Width           =   2415
      End
      Begin VB.CheckBox chkIngresarUsuarios 
         Caption         =   "Ingresar Usuarios"
         DataField       =   "IngresarUsuarios"
         DataSource      =   "adoPermisos"
         Height          =   255
         Left            =   3240
         TabIndex        =   28
         Top             =   2880
         Width           =   2415
      End
      Begin MSAdodcLib.Adodc adoPermisos 
         Height          =   330
         Left            =   3240
         Top             =   3960
         Visible         =   0   'False
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   582
         ConnectMode     =   0
         CursorLocation  =   3
         IsolationLevel  =   -1
         ConnectionTimeout=   15
         CommandTimeout  =   30
         CursorType      =   3
         LockType        =   3
         CommandType     =   1
         CursorOptions   =   0
         CacheSize       =   50
         MaxRecords      =   0
         BOFAction       =   0
         EOFAction       =   0
         ConnectStringType=   3
         Appearance      =   1
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Orientation     =   0
         Enabled         =   -1
         Connect         =   "DSN=Instituto"
         OLEDBString     =   ""
         OLEDBFile       =   ""
         DataSourceName  =   "Instituto"
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   "SELECT * FROM Permisos WHERE Usuario = 0"
         Caption         =   "Permisos"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _Version        =   393216
      End
      Begin VB.CheckBox chkIngresarCooperadoraPlan 
         Caption         =   "Ingresar Cooperadora Plan"
         DataField       =   "IngresarCooperadoraPlan"
         DataSource      =   "adoPermisos"
         Height          =   255
         Left            =   3240
         TabIndex        =   23
         Top             =   2160
         Width           =   2415
      End
      Begin VB.CheckBox chkAgregarPagosCooperadora 
         Caption         =   "Agregar Pagos Cooperadora"
         DataField       =   "AgregarPagosCooperadora"
         DataSource      =   "adoPermisos"
         Height          =   255
         Left            =   3240
         TabIndex        =   22
         Top             =   2520
         Width           =   2415
      End
      Begin VB.CheckBox chkModificarMatriculados 
         Caption         =   "Modificar Matriculados"
         DataField       =   "ModificarMatriculados"
         DataSource      =   "adoPermisos"
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   3240
         Width           =   2175
      End
      Begin VB.CheckBox chkModificarParciales 
         Caption         =   "Modificar Parciales"
         DataField       =   "ModificarParciales"
         DataSource      =   "adoPermisos"
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   3600
         Width           =   1815
      End
      Begin VB.CheckBox chkIngresarAsistencia 
         Caption         =   "Ingresar Asistencia"
         DataField       =   "IngresarAsistencia"
         DataSource      =   "adoPermisos"
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   3960
         Width           =   1815
      End
      Begin VB.CheckBox chkModificarFinales 
         Caption         =   "Modificar Finales"
         DataField       =   "ModificarFinales"
         DataSource      =   "adoPermisos"
         Height          =   255
         Left            =   3240
         TabIndex        =   18
         Top             =   360
         Width           =   1815
      End
      Begin VB.CheckBox chkImprimirActas 
         Caption         =   "Imprimir Actas"
         DataField       =   "ImprimirActas"
         DataSource      =   "adoPermisos"
         Height          =   255
         Left            =   3240
         TabIndex        =   17
         Top             =   720
         Width           =   1815
      End
      Begin VB.CheckBox chkIngresarActas 
         Caption         =   "Ingresar Actas"
         DataField       =   "IngresarActas"
         DataSource      =   "adoPermisos"
         Height          =   255
         Left            =   3240
         TabIndex        =   16
         Top             =   1080
         Width           =   1815
      End
      Begin VB.CheckBox chkImprimirAnalitico 
         Caption         =   "Imprimir Analitico"
         DataField       =   "ImprimirAnalitico"
         DataSource      =   "adoPermisos"
         Height          =   255
         Left            =   3240
         TabIndex        =   15
         Top             =   1440
         Width           =   1815
      End
      Begin VB.CheckBox chkImprimirTitulo 
         Caption         =   "Imprimir Titulo"
         DataField       =   "ImprimirTitulo"
         DataSource      =   "adoPermisos"
         Height          =   255
         Left            =   3240
         TabIndex        =   14
         Top             =   1800
         Width           =   1815
      End
      Begin VB.CheckBox chkModificarDatosAcademicos 
         Caption         =   "Modificar Datos Académicos"
         DataField       =   "ModificarDatosAcademicos"
         DataSource      =   "adoPermisos"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   720
         Width           =   2415
      End
      Begin VB.CheckBox chkModificarPlanes 
         Caption         =   "Modificar Planes"
         DataField       =   "ModificarPlanes"
         DataSource      =   "adoPermisos"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   1080
         Width           =   1815
      End
      Begin VB.CheckBox chkModificarCorrelativas 
         Caption         =   "Modificar Correlativas"
         DataField       =   "ModificarCorrelativas"
         DataSource      =   "adoPermisos"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   1440
         Width           =   1815
      End
      Begin VB.CheckBox chkModificarDivisiones 
         Caption         =   "Modificar Divisiones"
         DataField       =   "ModificarDivisiones"
         DataSource      =   "adoPermisos"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   1800
         Width           =   1815
      End
      Begin VB.CheckBox chkModificarEncuentros 
         Caption         =   "Modificar Encuentros"
         DataField       =   "ModificarEncuentros"
         DataSource      =   "adoPermisos"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   2160
         Width           =   1815
      End
      Begin VB.CheckBox chkModificarMesas 
         Caption         =   "Modificar Mesas"
         DataField       =   "ModificarMesas"
         DataSource      =   "adoPermisos"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   2520
         Width           =   1815
      End
      Begin VB.CheckBox chkModificarSuplentes 
         Caption         =   "Modificar Suplentes"
         DataField       =   "ModificarSuplentes"
         DataSource      =   "adoPermisos"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   2880
         Width           =   1815
      End
      Begin VB.CheckBox chkModificarAlumnos 
         Caption         =   "Modificar Alumnos"
         DataField       =   "ModificarAlumnos"
         DataSource      =   "adoPermisos"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   360
         Width           =   1815
      End
   End
   Begin VB.Frame frUsuarios 
      Height          =   1095
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7695
      Begin VB.TextBox txtIdentificacion 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   27
         Top             =   480
         Visible         =   0   'False
         Width           =   3375
      End
      Begin VB.CommandButton cmdQuitar 
         Caption         =   "Quitar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   5880
         TabIndex        =   5
         Top             =   360
         Width           =   1335
      End
      Begin VB.CommandButton cmdNuevo 
         Caption         =   "Nuevo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3960
         TabIndex        =   4
         Top             =   360
         Width           =   1335
      End
      Begin MSDataListLib.DataCombo dtcUsuarios 
         Bindings        =   "frmUsuarios.frx":0442
         Height          =   360
         Left            =   120
         TabIndex        =   3
         Top             =   480
         Width           =   3375
         _ExtentX        =   5953
         _ExtentY        =   741
         _Version        =   393216
         Style           =   2
         ListField       =   "Identificacion"
         BoundColumn     =   "Usuario"
         Text            =   "DataCombo1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSAdodcLib.Adodc adoUsuarios 
         Height          =   330
         Left            =   360
         Top             =   600
         Visible         =   0   'False
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   582
         ConnectMode     =   0
         CursorLocation  =   3
         IsolationLevel  =   -1
         ConnectionTimeout=   15
         CommandTimeout  =   30
         CursorType      =   3
         LockType        =   3
         CommandType     =   1
         CursorOptions   =   0
         CacheSize       =   50
         MaxRecords      =   0
         BOFAction       =   0
         EOFAction       =   0
         ConnectStringType=   3
         Appearance      =   1
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Orientation     =   0
         Enabled         =   -1
         Connect         =   "DSN=Instituto"
         OLEDBString     =   ""
         OLEDBFile       =   ""
         DataSourceName  =   "Instituto"
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   "SELECT * FROM Usuarios WHERE Eliminado = 0 ORDER BY Identificacion"
         Caption         =   "Usuarios"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _Version        =   393216
      End
      Begin VB.Label lblUsuarios 
         Caption         =   "Usuarios:"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   735
      End
   End
End
Attribute VB_Name = "frmUsuarios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Conexion As New Connection
Dim Auxiliar As New Recordset

Private Sub cmdModificar_Click()
    If cmdModificar.Caption = "Modificar" Then
        frUsuarios.Enabled = False
        cmdSalir.Enabled = False
        frPermisos.Enabled = True
        cmdModificar.Caption = "Guardar"
    Else
        adoPermisos.Recordset.Update
        adoPermisos.Refresh
        dtcUsuarios_Change
        frPermisos.Enabled = False
        cmdModificar.Caption = "Modificar"
        cmdSalir.Enabled = True
        frUsuarios.Enabled = True
    End If
End Sub

Private Sub cmdNuevo_Click()
    If cmdNuevo.Caption = "Nuevo" Then
        frComandos.Enabled = False
        cmdQuitar.Enabled = False
        txtIdentificacion.Visible = True
        cmdNuevo.Caption = "Guardar"
        lblUsuarios = "Ingrese el nuevo usuario"
        txtIdentificacion = ""
        txtIdentificacion.SetFocus
    Else
        If txtIdentificacion = "" Then MsgBox ("La identificación no puede estar en blanco"): txtIdentificacion.SetFocus: Exit Sub
        Conexion.Open
        Set Auxiliar = Conexion.Execute("SELECT MAX(Usuario) + 1 AS NuevoUsuario FROM Usuarios")
        NuevoUsuario = Auxiliar!NuevoUsuario
        Conexion.Execute ("INSERT INTO Usuarios ( Usuario,Identificacion, Contraseña,Modo ) Values (" & NuevoUsuario & ",'" & txtIdentificacion & "','u',1)")
        Conexion.Execute ("INSERT INTO Permisos ( Usuario ) Values (" & NuevoUsuario & ")")
        Conexion.Close
        adoUsuarios.Refresh
        txtIdentificacion.Visible = False
        dtcUsuarios.BoundText = NuevoUsuario
        frComandos.Enabled = True
        cmdQuitar.Enabled = True
        cmdNuevo.Caption = "Nuevo"
        lblUsuarios = "Usuarios"
        MsgBox ("La contraseña del nuevo usuario es: 1")
        frmIdentificacion.adoUsuarios.Refresh
        frmIdentificacion.dtcUsuarios.Refresh
    End If
End Sub

Private Sub cmdQuitar_Click()
    Respuesta = MsgBox("Está seguro de quitar al usuario: " & dtcUsuarios, vbYesNo, "Quitar Usuario")
    If Respuesta = vbYes Then
        Conexion.Open
        Conexion.Execute ("UPDATE Usuarios SET Eliminado = True WHERE Usuario=" & dtcUsuarios.BoundText)
        Conexion.Close
        adoUsuarios.Refresh
        dtcUsuarios.BoundText = adoUsuarios.Recordset!Usuario
    End If
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub dtcUsuarios_Change()
    adoPermisos.RecordSource = "SELECT * FROM Permisos WHERE Usuario =" & dtcUsuarios.BoundText
    adoPermisos.Refresh
End Sub

Private Sub Form_Load()
    Conexion.ConnectionString = ("DSN=Instituto")
    dtcUsuarios.BoundText = adoUsuarios.Recordset!Usuario
End Sub
