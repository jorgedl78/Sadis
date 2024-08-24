VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Begin VB.Form frmAnalitico 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Certificado Analítico"
   ClientHeight    =   7770
   ClientLeft      =   60
   ClientTop       =   300
   ClientWidth     =   11265
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   7770
   ScaleWidth      =   11265
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdAsistenciaExamanes 
      Caption         =   "Asistencia a Exámenes"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   9480
      Picture         =   "frmAnalitico.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   38
      Top             =   4320
      Width           =   1455
   End
   Begin VB.CommandButton cmdTituloEnTramite 
      Caption         =   "Tìtulo en Trámite"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   7920
      Picture         =   "frmAnalitico.frx":066A
      Style           =   1  'Graphical
      TabIndex        =   37
      Top             =   4320
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      Height          =   5175
      Left            =   0
      TabIndex        =   21
      Top             =   2400
      Width           =   11175
      Begin VB.CommandButton cmdPorcentajeDeMaterias 
         Caption         =   "Porcentaje de Materias"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   1680
         Picture         =   "frmAnalitico.frx":0CD4
         Style           =   1  'Graphical
         TabIndex        =   36
         Top             =   1920
         Width           =   1455
      End
      Begin VB.CommandButton cmdAlumnoRegular 
         Caption         =   "Alumno Regular"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   120
         Picture         =   "frmAnalitico.frx":133E
         Style           =   1  'Graphical
         TabIndex        =   33
         Top             =   1920
         Width           =   1455
      End
      Begin VB.ComboBox cbCurso 
         Height          =   315
         Left            =   2040
         Style           =   2  'Dropdown List
         TabIndex        =   31
         Top             =   600
         Width           =   615
      End
      Begin Crystal.CrystalReport CrystalReport3 
         Left            =   10200
         Top             =   4440
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   262150
         ReportFileName  =   "analicom.rpt"
         WindowTitle     =   "Situación Académica"
         WindowControlBox=   -1  'True
         WindowMaxButton =   -1  'True
         WindowMinButton =   -1  'True
      End
      Begin VB.CommandButton cmdCompleto 
         Caption         =   "Situación Académica"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   6360
         Picture         =   "frmAnalitico.frx":19A8
         Style           =   1  'Graphical
         TabIndex        =   30
         Top             =   1920
         Width           =   1455
      End
      Begin Crystal.CrystalReport CrystalReport2 
         Left            =   10440
         Top             =   4080
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   262150
         ReportFileName  =   "analidet.rpt"
         WindowTitle     =   "Analítico Detallado"
         WindowControlBox=   -1  'True
         WindowMaxButton =   -1  'True
         WindowMinButton =   -1  'True
      End
      Begin VB.CommandButton cmdDetallado 
         Caption         =   "Detallado"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   4800
         Picture         =   "frmAnalitico.frx":2012
         Style           =   1  'Graphical
         TabIndex        =   29
         Top             =   1920
         Width           =   1455
      End
      Begin VB.CommandButton cmdSalir 
         Height          =   840
         Left            =   8640
         Picture         =   "frmAnalitico.frx":267C
         Style           =   1  'Graphical
         TabIndex        =   28
         ToolTipText     =   "Salir"
         Top             =   4080
         Width           =   1080
      End
      Begin VB.CommandButton cmdImprimir 
         Caption         =   "Analítico"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   3240
         Picture         =   "frmAnalitico.frx":2ABE
         Style           =   1  'Graphical
         TabIndex        =   27
         Top             =   1920
         Width           =   1455
      End
      Begin VB.TextBox Text1 
         DataField       =   "Numero"
         DataSource      =   "adoMeses"
         Height          =   285
         Left            =   2280
         TabIndex        =   23
         Text            =   "Text1"
         Top             =   120
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.TextBox txtObservaciones 
         Height          =   1095
         Left            =   120
         MaxLength       =   255
         MultiLine       =   -1  'True
         TabIndex        =   22
         Top             =   3960
         Width           =   7695
      End
      Begin MSAdodcLib.Adodc adoMeses 
         Height          =   330
         Left            =   480
         Top             =   0
         Visible         =   0   'False
         Width           =   1815
         _ExtentX        =   3201
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
         RecordSource    =   "SELECT * FROM Meses ORDER BY Numero"
         Caption         =   "Meses"
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
      Begin Crystal.CrystalReport CrystalReport1 
         Left            =   9960
         Top             =   4080
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   262150
         ReportFileName  =   "analitic.rpt"
         WindowLeft      =   50
         WindowWidth     =   600
         WindowHeight    =   400
         WindowTitle     =   "Titulo"
         WindowControlBox=   -1  'True
         WindowMaxButton =   -1  'True
         WindowMinButton =   -1  'True
      End
      Begin MSComCtl2.DTPicker dtFecha 
         Height          =   375
         Left            =   120
         TabIndex        =   24
         Top             =   600
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         Format          =   122093569
         CurrentDate     =   37672
      End
      Begin MSAdodcLib.Adodc adoTitulo 
         Height          =   330
         Left            =   2640
         Top             =   0
         Visible         =   0   'False
         Width           =   1695
         _ExtentX        =   2990
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
         RecordSource    =   "SELECT * FROM Titulo"
         Caption         =   "Titulo"
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
      Begin MSDataListLib.DataCombo dtcTitulos 
         Bindings        =   "frmAnalitico.frx":3128
         Height          =   315
         Left            =   120
         TabIndex        =   34
         Top             =   1320
         Width           =   10815
         _ExtentX        =   19076
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "Titulo"
         BoundColumn     =   ""
         Text            =   ""
      End
      Begin MSAdodcLib.Adodc adoTitulos 
         Height          =   330
         Left            =   2280
         Top             =   960
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
         RecordSource    =   "SELECT * FROM TitulosPosibles WHERE TitulosPosibles.Carrera=0"
         Caption         =   "Titulos"
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
      Begin MSComCtl2.DTPicker dtpFechaExamen 
         Height          =   375
         Left            =   9480
         TabIndex        =   39
         Top             =   3000
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         _Version        =   393216
         Format          =   122093569
         CurrentDate     =   37672
      End
      Begin VB.Label Label2 
         Caption         =   "Elija el título correspondiente:"
         Height          =   255
         Left            =   120
         TabIndex        =   35
         Top             =   1080
         Width           =   2295
      End
      Begin VB.Label lblCurso 
         Caption         =   "Hasta año:"
         Height          =   255
         Left            =   2040
         TabIndex        =   32
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label3 
         Caption         =   "Fecha de Impresión:"
         Height          =   255
         Left            =   120
         TabIndex        =   26
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label Label6 
         Caption         =   "Obsevaciones:"
         Height          =   255
         Left            =   120
         TabIndex        =   25
         Top             =   3600
         Width           =   1335
      End
   End
   Begin VB.Frame frAlumnos 
      Caption         =   "Alumno"
      Height          =   2415
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11175
      Begin VB.TextBox txtPermiso 
         Height          =   315
         Left            =   120
         MaxLength       =   5
         TabIndex        =   1
         Top             =   480
         Width           =   615
      End
      Begin VB.CommandButton cmdBuscarAlumno 
         Caption         =   "Buscar..."
         Height          =   315
         Left            =   960
         TabIndex        =   11
         Top             =   480
         Width           =   855
      End
      Begin VB.TextBox txtAlumnoNombre 
         DataField       =   "Nombre"
         DataSource      =   "adoAlumnos"
         Enabled         =   0   'False
         Height          =   315
         Left            =   2040
         TabIndex        =   9
         Top             =   480
         Width           =   5415
      End
      Begin VB.TextBox txtAlumnoDocumento 
         DataField       =   "Documento"
         DataSource      =   "adoAlumnos"
         Enabled         =   0   'False
         Height          =   315
         Left            =   8280
         TabIndex        =   8
         Top             =   480
         Width           =   1335
      End
      Begin VB.TextBox txtAlumnoIngreso 
         DataField       =   "Ingreso"
         DataSource      =   "adoCarreras"
         Enabled         =   0   'False
         Height          =   315
         Left            =   120
         TabIndex        =   7
         Top             =   1920
         Width           =   1095
      End
      Begin VB.TextBox txtAlumnoCondicion 
         DataField       =   "Condicion"
         DataSource      =   "adoCarreras"
         Enabled         =   0   'False
         Height          =   315
         Left            =   1440
         TabIndex        =   6
         Top             =   1920
         Width           =   1935
      End
      Begin VB.TextBox txtAlumnoLibro 
         DataField       =   "Libro"
         DataSource      =   "adoCarreras"
         Enabled         =   0   'False
         Height          =   315
         Left            =   5160
         TabIndex        =   5
         Top             =   1920
         Width           =   615
      End
      Begin VB.TextBox txtAlumnoFecha 
         DataField       =   "Fecha"
         DataSource      =   "adoCarreras"
         Enabled         =   0   'False
         Height          =   315
         Left            =   3480
         TabIndex        =   4
         Top             =   1920
         Width           =   1095
      End
      Begin VB.TextBox txtAlumnoFolio 
         DataField       =   "Folio"
         DataSource      =   "adoCarreras"
         Enabled         =   0   'False
         Height          =   315
         Left            =   6240
         TabIndex        =   3
         Top             =   1920
         Width           =   615
      End
      Begin VB.TextBox txtAlumnoTipo 
         DataField       =   "Tipo"
         DataSource      =   "adoAlumnos"
         Enabled         =   0   'False
         Height          =   315
         Left            =   7800
         TabIndex        =   2
         Top             =   480
         Width           =   495
      End
      Begin MSDataListLib.DataCombo dtcCarreras 
         Bindings        =   "frmAnalitico.frx":3141
         Height          =   315
         Left            =   120
         TabIndex        =   10
         Top             =   1320
         Width           =   10815
         _ExtentX        =   19076
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "Nombre"
         BoundColumn     =   "Carrera"
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSAdodcLib.Adodc adoAlumnos 
         Height          =   330
         Left            =   120
         Top             =   960
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
         RecordSource    =   "SELECT Permiso,Nombre,Tipo,Documento FROM Alumnos WHERE Permiso = 0"
         Caption         =   "Alumnos"
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
      Begin MSAdodcLib.Adodc adoCarreras 
         Height          =   330
         Left            =   2160
         Top             =   960
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
         RecordSource    =   $"frmAnalitico.frx":315B
         Caption         =   "Carreras"
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
      Begin VB.Label Label1 
         Caption         =   "Nº Permiso"
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   240
         Width           =   855
      End
      Begin VB.Label lblFechaAlumno 
         Caption         =   "Fecha:"
         Height          =   255
         Left            =   3480
         TabIndex        =   19
         Top             =   1680
         Width           =   615
      End
      Begin VB.Label lblCondicionAlumno 
         Caption         =   "Condición:"
         Height          =   255
         Left            =   1440
         TabIndex        =   18
         Top             =   1680
         Width           =   855
      End
      Begin VB.Label lblIngresoAlumno 
         Caption         =   "Ingresó:"
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   1680
         Width           =   615
      End
      Begin VB.Label Label5 
         Caption         =   "Carrera"
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   960
         Width           =   855
      End
      Begin VB.Label lblFolioAlumno 
         Caption         =   "Folio:"
         Height          =   255
         Left            =   6240
         TabIndex        =   15
         Top             =   1680
         Width           =   495
      End
      Begin VB.Label lblLibroAlumno 
         Caption         =   "Libro:"
         Height          =   255
         Left            =   5160
         TabIndex        =   14
         Top             =   1680
         Width           =   495
      End
      Begin VB.Label Label8 
         Caption         =   "Documento:"
         Height          =   255
         Left            =   7800
         TabIndex        =   13
         Top             =   240
         Width           =   975
      End
      Begin VB.Label lblAlumnoNombre 
         Caption         =   "Nombre"
         Height          =   255
         Left            =   2040
         TabIndex        =   12
         Top             =   240
         Width           =   615
      End
   End
End
Attribute VB_Name = "frmAnalitico"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Conexion As New Connection
Dim Resultado As New Recordset
Dim CantidadPlan As New Recordset
Dim CantidadAprobadas As New Recordset
Dim TotalPlan As Integer
Dim TotalAprobadas As Integer
Dim porcentaje As Long

Private Sub cbCurso_Change()
    MsgBox ("curso change")
End Sub

Private Sub cbCurso_Click()
    CalcularPorcentaje
End Sub

Private Sub cmdAlumnoRegular_Click()
    If txtPermiso = "" Then
        MsgBox ("Debe especificar un número de permiso"): txtPermiso.SetFocus: Exit Sub
    End If
    
    
    adoMeses.Recordset.MoveFirst
    adoMeses.Recordset.Find ("Numero=" & dtFecha.Month)

    cn.Open
    Set rs = cn.Execute("SELECT NombreInstitucion, Localidad, Provincia, PathLogoReporte, Domicilio, Telefono from Parametros")
    Set CertificadoAlumnoRegular.DataSource = rs
    With CertificadoAlumnoRegular.Sections("Sección4")
        .Controls("lblInstitucion").Caption = rs!NombreInstitucion
        .Controls("lblDomicilio_Localidad").Caption = rs!Domicilio & " - " & rs!Localidad
        .Controls("lblTelefonos").Caption = rs!Telefono
        On Error Resume Next
        Set .Controls("imgLogo").Picture = LoadPicture(rs!PathLogoReporte)
        Set .Controls("imgSello").Picture = LoadPicture("sello.jpg")
        
        .Controls("lblTextoCompleto").Caption = "Se deja constancia de que, a la fecha, " & txtAlumnoNombre & ", DNI: " & txtAlumnoDocumento & " es alumno/a regular " & rs!NombreInstitucion & " DISTRITO " & rs!Localidad & ", y cursa la carrera " & dtcCarreras.Text & Chr(13) & Chr(10) & Chr(13) & Chr(10) & "A pedido del interesado/a y para ser presentado ante las autoridades que correspondan, se extiende la presente en la ciudad de " & rs!Localidad & " Prov. de " & rs!Provincia & " a los " & dtFecha.Day & " dias del mes de " & adoMeses.Recordset!Nombre & " del año " & dtFecha.Year & ".-"
        
        '.Controls("lblTextoCompleto").Caption = "LA DIRECCIÓN DEL NIVEL SUPERIOR del " & rs!NombreINstitucion & " DISTRITO " & rs!Localidad & " CERTIFICA que " & txtAlumnoNombre & ", DNI: " & txtAlumnoDocumento & " es alumno regular de la carrera " & dtcCarreras.Text & Chr(13) & Chr(10) & Chr(13) & Chr(10) & "A pedido del interesado y para ser presentado ante las autoridades que correspondan, se extiende la presente en " & rs!Localidad & " Prov. de " & rs!Provincia & " a los " & dtFecha.Day & " dias del mes de " & adoMeses.Recordset!Nombre & " del año " & dtFecha.Year & ".-"
    End With
    
    
    CertificadoAlumnoRegular.WindowState = 2
    CertificadoAlumnoRegular.Show 1
    cn.Close
    Me.Refresh
    
End Sub

Private Sub cmdAsistenciaExamanes_Click()
    If txtPermiso = "" Then
        MsgBox ("Debe especificar un número de permiso"): txtPermiso.SetFocus: Exit Sub
    End If
    
    
    adoMeses.Recordset.MoveFirst
    adoMeses.Recordset.Find ("Numero=" & dtFecha.Month)

    cn.Open
    Set rs = cn.Execute("SELECT NombreInstitucion, Localidad, Provincia, PathLogoReporte, Domicilio, Telefono from Parametros")
    Set rptConstanciaAsistenciaExamenes.DataSource = rs
    With rptConstanciaAsistenciaExamenes.Sections("Sección4")
        .Controls("lblInstitucion").Caption = rs!NombreInstitucion
        .Controls("lblDomicilio_Localidad").Caption = rs!Domicilio & " - " & rs!Localidad
        .Controls("lblTelefonos").Caption = rs!Telefono
        On Error Resume Next
        Set .Controls("imgLogo").Picture = LoadPicture(rs!PathLogoReporte)
        Set .Controls("imgSello").Picture = LoadPicture("sello.jpg")
        
        .Controls("lblTextoCompleto").Caption = "Se deja constancia de que, a la fecha, " & txtAlumnoNombre & ", DNI: " & txtAlumnoDocumento & " es alumno/a del " & rs!NombreInstitucion & " DISTRITO " & rs!Localidad & ", en la carrera " & dtcCarreras.Text & " y ha rendido examen parcial/final el día " & dtpFechaExamen & "." & Chr(13) & Chr(10) & Chr(13) & Chr(10) & "A pedido del interesado/a y para ser presentado ante las autoridades que correspondan, se extiende la presente en la ciudad de " & rs!Localidad & " Prov. de " & rs!Provincia & " a los " & dtFecha.Day & " dias del mes de " & adoMeses.Recordset!Nombre & " del año " & dtFecha.Year & ".-"
        
        '.Controls("lblTextoCompleto").Caption = "LA DIRECCIÓN DEL NIVEL SUPERIOR del " & rs!NombreINstitucion & " DISTRITO " & rs!Localidad & " CERTIFICA que " & txtAlumnoNombre & ", DNI: " & txtAlumnoDocumento & " es alumno regular de la carrera " & dtcCarreras.Text & Chr(13) & Chr(10) & Chr(13) & Chr(10) & "A pedido del interesado y para ser presentado ante las autoridades que correspondan, se extiende la presente en " & rs!Localidad & " Prov. de " & rs!Provincia & " a los " & dtFecha.Day & " dias del mes de " & adoMeses.Recordset!Nombre & " del año " & dtFecha.Year & ".-"
    End With
    
    
    rptConstanciaAsistenciaExamenes.WindowState = 2
    rptConstanciaAsistenciaExamenes.Show 1
    cn.Close
    Me.Refresh
End Sub

Private Sub cmdBuscarAlumno_Click()
    Respuesta = InputBox("Ingrese Nº de Documento", "Buscar Alumno")
    If Respuesta = "" Then Exit Sub
    Conexion.Open
    Set Resultado = Conexion.Execute("SELECT Permiso FROM Alumnos WHERE Documento = " & Respuesta & " AND Eliminado = False")
    If Resultado.EOF = False Then
        txtPermiso = Resultado!Permiso: txtPermiso.SetFocus
    Else
        MsgBox ("El documento no existe"): txtPermiso = ""
    End If
    Conexion.Close
End Sub

Private Sub cmdCompleto_Click()
    If txtPermiso = "" Then
        MsgBox ("Debe especificar un número de permiso"): txtPermiso.SetFocus: Exit Sub
    End If
    If txtPermiso = "" Then MsgBox ("Debe definir el Nº de permiso"): txtPermiso.SetFocus: Exit Sub
    adoMeses.Recordset.MoveFirst
    adoMeses.Recordset.Find ("Numero=" & dtFecha.Month)
    Conexion.Open
    Conexion.Execute ("DELETE * from Titulo")
    Conexion.Execute ("INSERT INTO Titulo ( Permiso, Nombre, Tipo, Documento, CodigoMateria, Materia, CodigoCarrera, Fecha, Nota, LibroFinal, FolioFinal, Equivalencia, Curso, Carrera, Medida, Detalle, Caracteristica, FechaDe, Resolucion, AnoCursada, Libro, Folio, Lugar, Nacio, idEstablecimiento, Libre, AprobadaEn) " _
    & "SELECT Alumnos.Permiso, Alumnos.Nombre, Alumnos.Tipo, Alumnos.Documento, Materias.Codigo AS CodigoMateria, Materias.Nombre AS Materia, Carreras.Codigo AS CodigoCarrera, Finales.Fecha, Finales.Nota, Finales.Libro as LibroFinal, Finales.Folio as FolioFinal, Finales.Equivalencia, Materias.CursoAnalitico, Carreras.Nombre AS Carrera, Modalidad.Medida, Materias.Detalle, [Caracteristica de carrera].Caracteristica, [Caracteristica de carrera].FechaDe, Carreras.Resolucion, Finales.Ano, CarrerasHechas.Libro, CarrerasHechas.Folio, Alumnos.Lugar, Alumnos.FechaNacimiento, Finales.Establecimiento, Finales.Libre, Instituciones.Institucion " _
    & "FROM ((((((Alumnos INNER JOIN Finales ON Alumnos.[Permiso] = Finales.[Alumno]) INNER JOIN Materias ON Finales.[Materia] = Materias.[Codigo]) INNER JOIN Carreras ON Materias.[Carrera] = Carreras.[Codigo]) INNER JOIN Modalidad ON Carreras.[Modalidad] = Modalidad.[Codigo]) INNER JOIN [Caracteristica de carrera] ON Carreras.[Caracteristica] = [Caracteristica de carrera].[Codigo]) INNER JOIN CarrerasHechas ON (Carreras.Codigo = CarrerasHechas.Carrera) AND (Alumnos.Permiso = CarrerasHechas.Permiso)) INNER JOIN Instituciones ON Finales.Establecimiento = Instituciones.Codigo " _
    & "Where [Alumnos].[Permiso] = " & txtPermiso & " And [Carreras].[Codigo] = " & dtcCarreras.BoundText & " AND Materias.Detalle <> 5 AND Finales.Habilitada = 1 AND ((Finales.Cursada = 1 AND Materias.Detalle BETWEEN 1 AND 3) OR Materias.Detalle=4) and Materias.Eliminada = 0 ORDER BY [Materias].[OrdenPlan], [Materias].[CursoAnalitico], [Materias].[Codigo]")
    Set Resultado = Conexion.Execute("SELECT NombreInstitucion, Localidad, Provincia FROM Parametros")
    Localidad = Resultado!Localidad
    Provincia = Resultado!Provincia
    PieDePagina = "En fe de lo cual se expide el presente certificado, en la ciudad de " & Localidad & " (" & Provincia & ") a los " & dtFecha.Day & " dias del mes de " & adoMeses.Recordset!Nombre & " del año " & dtFecha.Year & ".-"
    Conexion.Execute ("UPDATE Titulo SET PieDePagina = '" & PieDePagina & "'")
    Conexion.Execute ("UPDATE Titulo SET Establecimiento = '" & Resultado!NombreInstitucion & "'")
    Set Resultado = Conexion.Execute("SELECT * FROM Titulo")
    If Resultado.EOF = True Then MsgBox ("El alumno no rindió ninguna materia para esta carrera"): Conexion.Close: Exit Sub
    CarrerayTitulo = "Conste que el alumno " & Resultado!Nombre & ", " & Resultado!Tipo & " Nº " & Resultado!documento & "  ha aprobado las " & Resultado!Caracteristica & " que, con las respectivas calificaciones, que abajo se registran, correspondientes"
    If Resultado!CodigoCarrera = 51 Then
        'TituloDeBase = InputBox("Ingrese el título de base")
        CarrerayTitulo = CarrerayTitulo & " al POST TÍTULO FORMACIÓN DOCENTE, con especialización en EGB 3 y Polimodal.-" ' & dtcTitulos.Text & " en conjunción con su título de base: " & TituloDeBase & " que lo habilita, de acuerdo al Nomenclador vigente, en cada nivel mencionado."
    Else
        CarrerayTitulo = CarrerayTitulo & " a la carrera " & Resultado!Carrera & ".-" ', haciéndose acreedor/a al título de " & dtcTitulos.Text
    End If
    'CarrerayTitulo = CarrerayTitulo & "-Resolución " & Resultado!Resolucion
    Conexion.Execute ("UPDATE Titulo SET CarrerayTitulo = '" & CarrerayTitulo & "'")
    Conexion.Execute ("UPDATE Titulo SET Equivalencia = '*' WHERE Equivalencia = '-1'")
    Conexion.Execute ("UPDATE Titulo SET Equivalencia = Null WHERE Equivalencia = '0'")
    Conexion.Execute ("UPDATE Titulo SET AprobadaEn='Este establecimiento' WHERE idEstablecimiento = 0")
    Conexion.Execute ("UPDATE Titulo SET Condicion = 'Regular'")
    Conexion.Execute ("UPDATE Titulo SET Condicion = 'Equivalencia' WHERE idEstablecimiento>0")
    Conexion.Execute ("UPDATE Titulo SET Condicion = 'Libre' WHERE Libre=1")
    
    adoTitulo.RecordSource = "SELECT * FROM Titulo WHERE Detalle = 1 OR DEtalle = 2"
    adoTitulo.Refresh
    While adoTitulo.Recordset.EOF = False
    Decimo = Int(adoTitulo.Recordset!Nota)
    centesimo = Right(Format(adoTitulo.Recordset!Nota, "0.00"), 3)
    If centesimo = ".00" Then centesimo = "" 'si es entero no se muestra decimales
    If Decimo = 0 Then EnLetras = "CERO" & centesimo
    If Decimo = 1 Then EnLetras = "UNO" & centesimo
    If Decimo = 2 Then EnLetras = "DOS" & centesimo
    If Decimo = 3 Then EnLetras = "TRES" & centesimo
    If Decimo = 4 Then EnLetras = "CUATRO" & centesimo
    If Decimo = 5 Then EnLetras = "CINCO" & centesimo
    If Decimo = 6 Then EnLetras = "SEIS" & centesimo
    If Decimo = 7 Then EnLetras = "SIETE" & centesimo
    If Decimo = 8 Then EnLetras = "OCHO" & centesimo
    If Decimo = 9 Then EnLetras = "NUEVE" & centesimo
    If Decimo = 10 Then EnLetras = "DIEZ" & centesimo
    'If adoTitulo.Recordset!Nota > 0 Then
        Conexion.Execute ("UPDATE Titulo SET EnLetras = '" & EnLetras & "' WHERE CodigoMateria = " & adoTitulo.Recordset!CodigoMateria)
    'End If
    'If adoTitulo.Recordset!Nota = 0 Then
    '    Conexion.Execute ("UPDATE Titulo SET EnLetras = 'Regularizada' WHERE CodigoMateria = " & adoTitulo.Recordset!CodigoMateria)
    'End If
    adoTitulo.Recordset.MoveNext
    Wend
    Set Resultado = Conexion.Execute("SELECT Avg(Nota) AS Promedio FROM Titulo WHERE (Detalle = 1 OR Detalle = 2) AND Nota <>0")
    Conexion.Execute ("UPDATE Titulo SET Promedio = " & Replace(Resultado!Promedio, ",", "."))
    Decimo = Int(Format(Resultado!Promedio, "0.00"))
    centesimo = Right(Format(Resultado!Promedio, "0.00"), 3)
    If Decimo = 1 Then PromedioLetras = "UNO" & centesimo
    If Decimo = 2 Then PromedioLetras = "DOS" & centesimo
    If Decimo = 3 Then PromedioLetras = "TRES" & centesimo
    If Decimo = 4 Then PromedioLetras = "CUATRO" & centesimo
    If Decimo = 5 Then PromedioLetras = "CINCO" & centesimo
    If Decimo = 6 Then PromedioLetras = "SEIS" & centesimo
    If Decimo = 7 Then PromedioLetras = "SIETE" & centesimo
    If Decimo = 8 Then PromedioLetras = "OCHO" & centesimo
    If Decimo = 9 Then PromedioLetras = "NUEVE" & centesimo
    If Decimo = 10 Then PromedioLetras = "DIEZ" & centesimo
    Conexion.Execute ("UPDATE Titulo SET Fecha = Null, Nota = Null, EnLetras = Null WHERE (Detalle = 3 or Detalle = 4)")
    Conexion.Execute ("UPDATE Titulo SET EnLetras = 'Regularizada' WHERE (Detalle = 1 OR Detalle = 2) AND Nota =0")
    Conexion.Execute ("UPDATE Titulo SET PromedioLetras = '" & PromedioLetras & "'")
    
    If txtObservaciones <> "" Then Conexion.Execute ("UPDATE Titulo SET Observaciones = 'Observaciones: " & txtObservaciones & "'")
    Conexion.Execute ("INSERT INTO Entradas ( Permiso, Fecha, Hora, Profesor ) VALUES (" & txtPermiso & ",#" & Format(Date, "mm/dd/yyyy") & "#,# " & Format(Time, "hh:mm") & " #," & frmIdentificacion.dtcUsuarios.BoundText & ")")
    Conexion.Close
    CrystalReport3.PrintReport

   
End Sub

Private Sub cmdDetallado_Click()
    If txtPermiso = "" Then
        MsgBox ("Debe especificar un número de permiso"): txtPermiso.SetFocus: Exit Sub
    End If
    If txtPermiso = "" Then MsgBox ("Debe definir el Nº de permiso"): txtPermiso.SetFocus: Exit Sub
    adoMeses.Recordset.MoveFirst
    adoMeses.Recordset.Find ("Numero=" & dtFecha.Month)
    Conexion.Open
    Conexion.Execute ("DELETE * from Titulo")
    Conexion.Execute ("INSERT INTO Titulo ( Permiso, Nombre, Tipo, Documento, CodigoMateria, Materia, CodigoCarrera, Fecha, Nota, Equivalencia, Curso, Carrera, Medida, Detalle, Caracteristica, FechaDe, Resolucion, AnoCursada) SELECT [Alumnos].[Permiso], [Alumnos].[Nombre], [Alumnos].[Tipo], [Alumnos].[Documento],[Materias].[Codigo] AS  CodigoMateria,[Materias].[Nombre] AS Materia, Carreras.Codigo AS CodigoCarrera, [Finales].[Fecha], [Finales].[Nota],[Finales].[Equivalencia], [Materias].[CursoAnalitico],[Carreras].[Nombre] AS Carrera, [Modalidad].[Medida], [Materias].[Detalle], [Caracteristica de carrera].[Caracteristica],[Caracteristica de carrera].[FechaDe], [Carreras].[Resolucion], [Finales].[Ano] FROM ((((Alumnos INNER JOIN Finales ON [Alumnos].[Permiso]=[Finales].[Alumno]) INNER JOIN Materias ON [Finales].[Materia]=[Materias].[Codigo])INNER JOIN Carreras ON [Materias].[Carrera]=[Carreras].[Codigo])INNER JOIN Modalidad ON [Carreras].[Modalidad]=[Modalidad].[Codigo])" _
    & " INNER JOIN [Caracteristica de carrera] ON [Carreras].[Caracteristica]=[Caracteristica de carrera].[Codigo] Where [Alumnos].[Permiso] = " & txtPermiso & " And [Carreras].[Codigo] = " & dtcCarreras.BoundText & " AND Materias.Detalle <> 5 AND Finales.Habilitada = 1 AND ((Finales.Cursada = 1 AND Materias.Detalle BETWEEN 1 AND 3) OR Materias.Detalle=4) and Materias.Eliminada=0 ORDER BY [Materias].[OrdenPlan], [Materias].[CursoAnalitico], [Materias].[Codigo]")
    Set Resultado = Conexion.Execute("SELECT NombreInstitucion, Localidad, Provincia FROM Parametros")
    Localidad = Resultado!Localidad
    Provincia = Resultado!Provincia
    PieDePagina = "En fe de lo cual se expide el presente certificado, en la ciudad de " & Localidad & " (" & Provincia & ") a los " & dtFecha.Day & " dias del mes de " & adoMeses.Recordset!Nombre & " del año " & dtFecha.Year & ".-"
    Conexion.Execute ("UPDATE Titulo SET PieDePagina = '" & PieDePagina & "'")
    Conexion.Execute ("UPDATE Titulo SET Establecimiento = '" & Resultado!NombreInstitucion & "'")
    Set Resultado = Conexion.Execute("SELECT * FROM Titulo")
    If Resultado.EOF = True Then MsgBox ("El alumno no rindió ninguna materia para esta carrera"): Conexion.Close: Exit Sub
    CarrerayTitulo = "Conste que el alumno " & Resultado!Nombre & ", " & Resultado!Tipo & " Nº " & Resultado!documento & "  ha aprobado las " & Resultado!Caracteristica & " que, con las respectivas calificaciones, que abajo se registran, correspondientes"
    If Resultado!CodigoCarrera = 51 Then
        'TituloDeBase = InputBox("Ingrese el título de base")
        CarrerayTitulo = CarrerayTitulo & " al POST TÍTULO FORMACIÓN DOCENTE, con especialización en EGB 3 y Polimodal.-" ' & dtcTitulos.Text & " en conjunción con su título de base: " & TituloDeBase & " que lo habilita, de acuerdo al Nomenclador vigente, en cada nivel mencionado."
    Else
        CarrerayTitulo = CarrerayTitulo & " a la carrera " & Resultado!Carrera & ".-" ', haciéndose acreedor/a al título de " & dtcTitulos.Text
    End If
    'CarrerayTitulo = CarrerayTitulo & "-Resolución " & Resultado!Resolucion
    Conexion.Execute ("UPDATE Titulo SET CarrerayTitulo = '" & CarrerayTitulo & "'")
    Conexion.Execute ("UPDATE Titulo SET Equivalencia = '*' WHERE Equivalencia = '-1'")
    Conexion.Execute ("UPDATE Titulo SET Equivalencia = Null WHERE Equivalencia = '0'")
    adoTitulo.RecordSource = "SELECT * FROM Titulo WHERE Detalle = 1 OR DEtalle = 2"
    adoTitulo.Refresh
    While adoTitulo.Recordset.EOF = False
    Decimo = Int(adoTitulo.Recordset!Nota)
    centesimo = Right(Format(adoTitulo.Recordset!Nota, "0.00"), 3)
    If centesimo = ".00" Then centesimo = "" 'si es entero no se muestra decimales
    If Decimo = 0 Then EnLetras = "CERO" & centesimo
    If Decimo = 1 Then EnLetras = "UNO" & centesimo
    If Decimo = 2 Then EnLetras = "DOS" & centesimo
    If Decimo = 3 Then EnLetras = "TRES" & centesimo
    If Decimo = 4 Then EnLetras = "CUATRO" & centesimo
    If Decimo = 5 Then EnLetras = "CINCO" & centesimo
    If Decimo = 6 Then EnLetras = "SEIS" & centesimo
    If Decimo = 7 Then EnLetras = "SIETE" & centesimo
    If Decimo = 8 Then EnLetras = "OCHO" & centesimo
    If Decimo = 9 Then EnLetras = "NUEVE" & centesimo
    If Decimo = 10 Then EnLetras = "DIEZ" & centesimo
    'If adoTitulo.Recordset!Nota > 0 Then
        Conexion.Execute ("UPDATE Titulo SET EnLetras = '" & EnLetras & "' WHERE CodigoMateria = " & adoTitulo.Recordset!CodigoMateria)
    'End If
    'If adoTitulo.Recordset!Nota = 0 Then
    '    Conexion.Execute ("UPDATE Titulo SET EnLetras = 'Regularizada' WHERE CodigoMateria = " & adoTitulo.Recordset!CodigoMateria)
    'End If
    adoTitulo.Recordset.MoveNext
    Wend
    Set Resultado = Conexion.Execute("SELECT Avg(Nota) AS Promedio FROM Titulo WHERE (Detalle = 1 OR Detalle = 2) AND Nota <>0")
    
    Conexion.Execute ("UPDATE Titulo SET Promedio = " & Replace(Resultado!Promedio, ",", "."))
    Decimo = Int(Format(Resultado!Promedio, "0.00"))
    centesimo = Right(Format(Resultado!Promedio, "0.00"), 3)
    If Decimo = 1 Then PromedioLetras = "UNO" & centesimo
    If Decimo = 2 Then PromedioLetras = "DOS" & centesimo
    If Decimo = 3 Then PromedioLetras = "TRES" & centesimo
    If Decimo = 4 Then PromedioLetras = "CUATRO" & centesimo
    If Decimo = 5 Then PromedioLetras = "CINCO" & centesimo
    If Decimo = 6 Then PromedioLetras = "SEIS" & centesimo
    If Decimo = 7 Then PromedioLetras = "SIETE" & centesimo
    If Decimo = 8 Then PromedioLetras = "OCHO" & centesimo
    If Decimo = 9 Then PromedioLetras = "NUEVE" & centesimo
    If Decimo = 10 Then PromedioLetras = "DIEZ" & centesimo
    Conexion.Execute ("UPDATE Titulo SET Fecha = Null, Nota = Null, EnLetras = Null WHERE (Detalle = 3 or Detalle = 4)")
    Conexion.Execute ("UPDATE Titulo SET EnLetras = 'Regularizada' WHERE (Detalle = 1 OR Detalle = 2) AND Nota =0")
    Conexion.Execute ("UPDATE Titulo SET PromedioLetras = '" & PromedioLetras & "'")
    
    If txtObservaciones <> "" Then Conexion.Execute ("UPDATE Titulo SET Observaciones = 'Observaciones: " & txtObservaciones & "'")
    Conexion.Execute ("INSERT INTO Entradas ( Permiso, Fecha, Hora, Profesor ) VALUES (" & txtPermiso & ",#" & Format(Date, "mm/dd/yyyy") & "#,# " & Format(Time, "hh:mm") & " #," & frmIdentificacion.dtcUsuarios.BoundText & ")")
    Conexion.Close
    CrystalReport2.PrintReport

End Sub

Private Sub cmdImprimir_Click()
    If txtPermiso = "" Then
        MsgBox ("Debe especificar un número de permiso"): txtPermiso.SetFocus: Exit Sub
    End If
    If txtPermiso = "" Then MsgBox ("Debe definir el Nº de permiso"): txtPermiso.SetFocus: Exit Sub
    adoMeses.Recordset.MoveFirst
    adoMeses.Recordset.Find ("Numero=" & dtFecha.Month)
    Conexion.Open
    Conexion.Execute ("DELETE * from Titulo")
    Conexion.Execute ("INSERT INTO Titulo ( Permiso, Nombre, Tipo, Documento, CodigoMateria, Materia, CodigoCarrera, Fecha, Nota, Equivalencia, Curso, Carrera, Medida, Detalle, Caracteristica, FechaDe, Resolucion)SELECT [Alumnos].[Permiso], [Alumnos].[Nombre], [Alumnos].[Tipo], [Alumnos].[Documento],[Materias].[Codigo] AS  CodigoMateria,[Materias].[Nombre] AS Materia, Carreras.Codigo AS CodigoCarrera, [Finales].[Fecha], [Finales].[Nota],[Finales].[Equivalencia], [Materias].[CursoAnalitico],[Carreras].[Nombre] AS Carrera, [Modalidad].[Medida], [Materias].[Detalle], [Caracteristica de carrera].[Caracteristica],[Caracteristica de carrera].[FechaDe], [Carreras].[Resolucion] FROM ((((Alumnos INNER JOIN Finales ON [Alumnos].[Permiso]=[Finales].[Alumno]) INNER JOIN Materias ON [Finales].[Materia]=[Materias].[Codigo])INNER JOIN Carreras ON [Materias].[Carrera]=[Carreras].[Codigo])INNER JOIN Modalidad ON [Carreras].[Modalidad]=[Modalidad].[Codigo])" _
    & " INNER JOIN [Caracteristica de carrera] ON [Carreras].[Caracteristica]=[Caracteristica de carrera].[Codigo] Where [Alumnos].[Permiso] = " & txtPermiso & " And [Carreras].[Codigo] = " & dtcCarreras.BoundText & " AND Materias.Detalle <> 5  AND Materias.Curso <=" & cbCurso.Text & " AND Finales.Habilitada = 1 AND ((Finales.Aprobada = 1 AND Materias.Detalle BETWEEN 1 AND 3) OR Materias.Detalle=4) AND Materias.Eliminada=0 ORDER BY [Materias].[OrdenPlan], [Materias].[CursoAnalitico], [Materias].[Codigo]")
    Set Resultado = Conexion.Execute("SELECT NombreInstitucion, Localidad, Provincia FROM Parametros")
    Localidad = Resultado!Localidad
    Provincia = Resultado!Provincia
    PieDePagina = "En fe de lo cual se expide el presente certificado, en la ciudad de " & Localidad & " (" & Provincia & ") a los " & dtFecha.Day & " dias del mes de " & adoMeses.Recordset!Nombre & " del año " & dtFecha.Year & ".-"
    Conexion.Execute ("UPDATE Titulo SET PieDePagina = '" & PieDePagina & "'")
    Conexion.Execute ("UPDATE Titulo SET Establecimiento = '" & Resultado!NombreInstitucion & "'")
    Set Resultado = Conexion.Execute("SELECT * FROM Titulo")
    If Resultado.EOF = True Then MsgBox ("El alumno no rindió ninguna materia para esta carrera"): Conexion.Close: Exit Sub
    CarrerayTitulo = "Conste que el alumno " & Resultado!Nombre & ", " & Resultado!Tipo & " Nº " & Resultado!documento & "  ha aprobado las " & Resultado!Caracteristica & " que, con las respectivas calificaciones, que abajo se registran, correspondientes"
    If Resultado!CodigoCarrera = 51 Then
        'TituloDeBase = InputBox("Ingrese el título de base")
        CarrerayTitulo = CarrerayTitulo & " al POST TÍTULO FORMACIÓN DOCENTE, con especialización en EGB 3 y Polimodal.-" ' & dtcTitulos.Text & " en conjunción con su título de base: " & TituloDeBase & " que lo habilita, de acuerdo al Nomenclador vigente, en cada nivel mencionado."
    Else
        CarrerayTitulo = CarrerayTitulo & " a la carrera " & Resultado!Carrera & ".-" ', haciéndose acreedor/a al título de " & dtcTitulos.Text
    End If
    'CarrerayTitulo = CarrerayTitulo & "-Resolución " & Resultado!Resolucion
    Conexion.Execute ("UPDATE Titulo SET CarrerayTitulo = '" & CarrerayTitulo & "'")
    Conexion.Execute ("UPDATE Titulo SET Equivalencia = '*' WHERE Equivalencia = '-1'")
    Conexion.Execute ("UPDATE Titulo SET Equivalencia = Null WHERE Equivalencia = '0'")
    adoTitulo.RecordSource = "SELECT * FROM Titulo WHERE Detalle = 1 OR DEtalle = 2"
    adoTitulo.Refresh
    While adoTitulo.Recordset.EOF = False
    Decimo = Int(adoTitulo.Recordset!Nota)
    centesimo = Right(Format(adoTitulo.Recordset!Nota, "0.00"), 3)
    If centesimo = ".00" Then centesimo = "" 'si es entero no se muestra decimales
    If Decimo = 1 Then EnLetras = "UNO" & centesimo
    If Decimo = 2 Then EnLetras = "DOS" & centesimo
    If Decimo = 3 Then EnLetras = "TRES" & centesimo
    If Decimo = 4 Then EnLetras = "CUATRO" & centesimo
    If Decimo = 5 Then EnLetras = "CINCO" & centesimo
    If Decimo = 6 Then EnLetras = "SEIS" & centesimo
    If Decimo = 7 Then EnLetras = "SIETE" & centesimo
    If Decimo = 8 Then EnLetras = "OCHO" & centesimo
    If Decimo = 9 Then EnLetras = "NUEVE" & centesimo
    If Decimo = 10 Then EnLetras = "DIEZ" & centesimo
    If adoTitulo.Recordset!Nota > 0 Then
        Conexion.Execute ("UPDATE Titulo SET EnLetras = '" & EnLetras & "' WHERE CodigoMateria = " & adoTitulo.Recordset!CodigoMateria)
    End If
    adoTitulo.Recordset.MoveNext
    Wend
    Set Resultado = Conexion.Execute("SELECT Avg(Nota) AS Promedio FROM Titulo WHERE Detalle = 1 OR Detalle = 2")
    PromedioCalculado = Format(Resultado!Promedio, "#.#0")
    Conexion.Execute ("UPDATE Titulo SET Promedio = " & Replace(PromedioCalculado, ",", "."))
    Decimo = Int(Format(PromedioCalculado, "0.00"))
    centesimo = Right(Format(PromedioCalculado, "0.00"), 3)
    If Decimo = 1 Then PromedioLetras = "UNO" & centesimo
    If Decimo = 2 Then PromedioLetras = "DOS" & centesimo
    If Decimo = 3 Then PromedioLetras = "TRES" & centesimo
    If Decimo = 4 Then PromedioLetras = "CUATRO" & centesimo
    If Decimo = 5 Then PromedioLetras = "CINCO" & centesimo
    If Decimo = 6 Then PromedioLetras = "SEIS" & centesimo
    If Decimo = 7 Then PromedioLetras = "SIETE" & centesimo
    If Decimo = 8 Then PromedioLetras = "OCHO" & centesimo
    If Decimo = 9 Then PromedioLetras = "NUEVE" & centesimo
    If Decimo = 10 Then PromedioLetras = "DIEZ" & centesimo
    Conexion.Execute ("UPDATE Titulo SET Fecha = Null, Nota = Null, EnLetras = Null WHERE (Detalle = 3 or Detalle = 4)")
    Conexion.Execute ("UPDATE Titulo SET PromedioLetras = '" & PromedioLetras & "'")
    If txtObservaciones <> "" Then Conexion.Execute ("UPDATE Titulo SET Observaciones = 'Observaciones: " & txtObservaciones & "'")
    Conexion.Execute ("INSERT INTO Entradas ( Permiso, Fecha, Hora, Profesor ) VALUES (" & txtPermiso & ",#" & Format(Date, "mm/dd/yyyy") & "#,# " & Format(Time, "hh:mm") & " #," & frmIdentificacion.dtcUsuarios.BoundText & ")")
    Conexion.Close
    CrystalReport1.PrintReport
End Sub

Private Sub cmdPorcentajeDeMaterias_Click()
    If txtPermiso = "" Then
        MsgBox ("Debe especificar un número de permiso"): txtPermiso.SetFocus: Exit Sub
    End If
    
    
    adoMeses.Recordset.MoveFirst
    adoMeses.Recordset.Find ("Numero=" & dtFecha.Month)

    cn.Open
    Set rs = cn.Execute("SELECT NombreInstitucion, Localidad, Provincia, PathLogoReporte, Domicilio, Telefono from Parametros")
    Set Porcentaje_De_Materias.DataSource = rs
    With Porcentaje_De_Materias.Sections("Sección4")
        .Controls("lblInstitucion").Caption = rs!NombreInstitucion
        .Controls("lblDomicilio_Localidad").Caption = rs!Domicilio & " - " & rs!Localidad
        .Controls("lblTelefonos").Caption = rs!Telefono
        On Error Resume Next
        Set .Controls("imgLogo").Picture = LoadPicture(rs!PathLogoReporte)
        Set .Controls("imgSello").Picture = LoadPicture("sello.jpg")
        .Controls("lblTextoCompleto").Caption = "Se deja constancia de que, a la fecha, " & txtAlumnoNombre & ", DNI: " & txtAlumnoDocumento & " es alumno/a regular " & rs!NombreInstitucion & " DISTRITO " & rs!Localidad & ", de la carrera " & dtcCarreras.Text & " y ha rendido y aprobado el siguiente nùmero de asignaturas:"
        .Controls("lblTotalDelPlan").Caption = TotalPlan
        .Controls("lblRendidas").Caption = TotalAprobadas
        .Controls("lblAdeudadas").Caption = TotalPlan - TotalAprobadas
        .Controls("lblPlanDeEstudios").Caption = "PLAN DE ESTUDIOS: " & cbCurso.Text & " AÑOS"
        .Controls("lblPorcentaje").Caption = "PORCENTAJE DE MATERIAS APROBADAS: " & porcentaje & "%"
        .Controls("lblTextoPie").Caption = "A pedido del interesado/a y para ser presentado ante quien correspondan, se extiende la presente en la ciudad de " & rs!Localidad & " Prov. de " & rs!Provincia & " a los " & dtFecha.Day & " dias del mes de " & adoMeses.Recordset!Nombre & " del año " & dtFecha.Year & ".-"
        '"A pedido del interesado/a y para ser presentado ante las autoridades que correspondan, se extiende la presente en la ciudad de " & rs!Localidad & " Prov. de " & rs!Provincia & " a los " & dtFecha.Day & " dias del mes de " & adoMeses.Recordset!Nombre & " del año " & dtFecha.Year & ".-"
        '.Controls("lblTextoCompleto").Caption = "LA DIRECCIÓN DEL NIVEL SUPERIOR del " & rs!NombreINstitucion & " DISTRITO " & rs!Localidad & " CERTIFICA que " & txtAlumnoNombre & ", DNI: " & txtAlumnoDocumento & " tiene aprobado a la fecha un " & Format(porcentaje, "0") & " % del total de las asignaturas para obtener el título de " & dtcTitulos.Text & Chr(13) & Chr(10) & Chr(13) & Chr(10) & "A pedido del interesado y para ser presentado ante las autoridades que correspondan, se extiende la presente en " & rs!Localidad & " Prov. de " & rs!Provincia & " a los " & dtFecha.Day & " dias del mes de " & adoMeses.Recordset!Nombre & " del año " & dtFecha.Year & ".-"
    End With
    
    
    Porcentaje_De_Materias.WindowState = 2
    Porcentaje_De_Materias.Show 1
    cn.Close
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Command1_Click()

End Sub

Private Sub cmdTituloEnTramite_Click()
    If txtPermiso = "" Then
        MsgBox ("Debe especificar un número de permiso"): txtPermiso.SetFocus: Exit Sub
    End If
    
    
    adoMeses.Recordset.MoveFirst
    adoMeses.Recordset.Find ("Numero=" & dtFecha.Month)

    cn.Open
    
    'calculo el promedio
    Set rs = cn.Execute("SELECT avg(Finales.Nota) as Promedio FROM Finales INNER JOIN Materias ON Finales.Materia = Materias.Codigo WHERE (((Finales.Alumno)=" & txtPermiso.Text & ") AND ((Materias.Carrera)=" & dtcCarreras.BoundText & ") AND ((Finales.Aprobada)=True) AND ((Materias.Detalle) Between 1 And 2) AND ((Finales.Habilitada)=True))")

    'Conexion.Execute ("UPDATE Titulo SET Promedio = " & Replace(Resultado!promedio, ",", "."))
    Promedio = Format(rs!Promedio, "0.00")
    Decimo = Int(Format(rs!Promedio, "0.00"))
    centesimo = Right(Format(rs!Promedio, "0.00"), 3)
    If Decimo = 1 Then PromedioLetras = "UNO" & centesimo
    If Decimo = 2 Then PromedioLetras = "DOS" & centesimo
    If Decimo = 3 Then PromedioLetras = "TRES" & centesimo
    If Decimo = 4 Then PromedioLetras = "CUATRO" & centesimo
    If Decimo = 5 Then PromedioLetras = "CINCO" & centesimo
    If Decimo = 6 Then PromedioLetras = "SEIS" & centesimo
    If Decimo = 7 Then PromedioLetras = "SIETE" & centesimo
    If Decimo = 8 Then PromedioLetras = "OCHO" & centesimo
    If Decimo = 9 Then PromedioLetras = "NUEVE" & centesimo
    If Decimo = 10 Then PromedioLetras = "DIEZ" & centesimo

    
    Set rs = cn.Execute("SELECT NombreInstitucion, Localidad, Provincia, PathLogoReporte, Domicilio, Telefono from Parametros")
    Set rptCertificadoTituloEnTramite.DataSource = rs
    With rptCertificadoTituloEnTramite.Sections("Sección4")
        .Controls("lblInstitucion").Caption = rs!NombreInstitucion
        .Controls("lblDomicilio_Localidad").Caption = rs!Domicilio & " - " & rs!Localidad
        .Controls("lblTelefonos").Caption = rs!Telefono
        On Error Resume Next
        Set .Controls("imgLogo").Picture = LoadPicture(rs!PathLogoReporte)
        Set .Controls("imgSello").Picture = LoadPicture("sello.jpg")
        .Controls("lblTextoCompleto").Caption = "El " & rs!NombreInstitucion & " certifica que el alumno/a " & txtAlumnoNombre & ", DNI: " & txtAlumnoDocumento & " ha acreditado todas las asignaturas correspondientes al Plan de Estudios vigente de la carrera " & dtcCarreras.Text & " con un promedio general de " & Promedio & " (" & PromedioLetras & "), estando el respectivo tìtulo en trámite " & Chr(13) & Chr(10) & Chr(13) & Chr(10) & "A pedido del interesado/a y para ser presentado ante las autoridades que correspondan, se extiende la presente en la ciudad de " & rs!Localidad & " Prov. de " & rs!Provincia & " a los " & dtFecha.Day & " dias del mes de " & adoMeses.Recordset!Nombre & " del año " & dtFecha.Year & ".-"
        
        '.Controls("lblTextoCompleto").Caption = "LA DIRECCIÓN DEL NIVEL SUPERIOR del " & rs!NombreINstitucion & " DISTRITO " & rs!Localidad & " CERTIFICA que " & txtAlumnoNombre & ", DNI: " & txtAlumnoDocumento & " es alumno regular de la carrera " & dtcCarreras.Text & Chr(13) & Chr(10) & Chr(13) & Chr(10) & "A pedido del interesado y para ser presentado ante las autoridades que correspondan, se extiende la presente en " & rs!Localidad & " Prov. de " & rs!Provincia & " a los " & dtFecha.Day & " dias del mes de " & adoMeses.Recordset!Nombre & " del año " & dtFecha.Year & ".-"
    End With
    
    
    rptCertificadoTituloEnTramite.WindowState = 2
    rptCertificadoTituloEnTramite.Show 1
    cn.Close
End Sub

Private Sub dtcCarreras_Change()
    If dtcCarreras.Text <> "" Then
        adoCarreras.Recordset.MoveFirst
        adoCarreras.Recordset.Find ("Carrera=" & dtcCarreras.BoundText)
        cbCurso.Clear
        For i = 0 To adoCarreras.Recordset!Años - 1
            cbCurso.List(i) = i + 1
        Next i
        cbCurso.Text = cbCurso.List(cbCurso.ListCount - 1)
        CalcularPorcentaje
    End If
    If dtcCarreras.Text <> "" Then
        adoCarreras.Recordset.MoveFirst
        adoCarreras.Recordset.Find ("Carrera=" & dtcCarreras.BoundText)
        adoTitulos.RecordSource = "SELECT * From TitulosPosibles Where TitulosPosibles.Carrera = " & dtcCarreras.BoundText & " ORDER BY Titulo"
        adoTitulos.Refresh
        If adoTitulos.Recordset.RecordCount = 0 Then
            MsgBox ("El plan seleccionado no contiene su descripción de título"): txtPermiso.SetFocus: Exit Sub
        End If
        dtcTitulos.Text = adoTitulos.Recordset!Titulo
    End If


End Sub



Private Sub Form_Load()
    Conexion.ConnectionString = ("DSN=Instituto")
    dtFecha.Value = Date
    dtpFechaExamen.Value = Date
       
    Conexion.Open
    On Error GoTo hErr
    Conexion.Execute ("ALTER TABLE  Titulo ADD Libro int, Folio int, Lugar char(20), Nacio char(20), AprobadaEn  char(70), Condicion char(20), idEstablecimiento int, Libre Logical, AnoCursada int")
    Conexion.Close

hErr:
   'MsgBox Err.Number & " " & Err.Description
   Conexion.Close
   Exit Sub
End Sub

Private Sub txtPermiso_Click()
    dtcCarreras.Text = ""
    txtPermiso = ""
End Sub

Private Sub txtPermiso_KeyPress(KeyAscii As Integer)
    If InStr(1, "0123456789" & Chr(13) & Chr(8), Chr(KeyAscii)) = 0 Then
        KeyAscii = 0
    End If
    If txtPermiso = "" Then Exit Sub
    If KeyAscii = 13 Then dtcCarreras.Text = "": BuscarDatos
End Sub

Private Function BuscarDatos()
    adoAlumnos.RecordSource = "SELECT Permiso,Nombre,Tipo,Documento FROM Alumnos WHERE Eliminado = 0 AND Permiso = " & txtPermiso
    adoAlumnos.Refresh
    If adoAlumnos.Recordset.RecordCount = 1 Then
        adoCarreras.RecordSource = "SELECT CarrerasHechas.Permiso, CarrerasHechas.Carrera, CarrerasHechas.Ingreso, Condicion.Condicion, CarrerasHechas.Fecha, CarrerasHechas.Libro, CarrerasHechas.Folio, Carreras.Nombre, Carreras.Años FROM (CarrerasHechas INNER JOIN Carreras ON CarrerasHechas.Carrera = Carreras.Codigo) INNER JOIN Condicion ON CarrerasHechas.Condición = Condicion.Codigo WHERE CarrerasHechas.Permiso=" & txtPermiso
        adoCarreras.Refresh
        If adoCarreras.Recordset.RecordCount <= 0 Then
           MsgBox ("El alumno no tiene ningún plan de estudio asociado"): Exit Function
        End If
        dtcCarreras.BoundText = adoCarreras.Recordset!Carrera
    Else
        MsgBox ("El alumno no existe")
        txtPermiso = ""
    End If
End Function

Private Function CalcularPorcentaje()
        If Conexion.State = 0 Then Conexion.Open
        Set CantidadPlan = Conexion.Execute("SELECT Count(Materias.Nombre) AS CantidadPlan From Materias WHERE Materias.Carrera=" & dtcCarreras.BoundText & " AND (Materias.Detalle=1 OR Materias.Detalle=2) and Materias.Curso<=" & cbCurso.Text & " AND Eliminada=0")
        TotalPlan = CantidadPlan!CantidadPlan
        Set CantidadAprobadas = Conexion.Execute("SELECT Count(Materias.Detalle) As CantidadAprobadas FROM Finales INNER JOIN Materias ON Finales.Materia = Materias.Codigo WHERE (((Materias.Detalle)=1 Or (Materias.Detalle)=2) AND ((Finales.Alumno)=" & adoAlumnos.Recordset!Permiso & ") AND ((Materias.Carrera)=" & dtcCarreras.BoundText & ") AND ((Finales.Aprobada)=1) and Materias.Eliminada=0 and Materias.Curso<=" & cbCurso.Text & ")")
        TotalAprobadas = CantidadAprobadas!CantidadAprobadas
        Conexion.Close
        porcentaje = (TotalAprobadas * 100) / TotalPlan
        txtObservaciones = "Tiene aprobadas el " & Format(porcentaje, "0") & " % de las asignaturas hasta el " & cbCurso.Text & " º año del plan de estudios.-"
End Function
