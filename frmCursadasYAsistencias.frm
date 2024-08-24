VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmCursadasyAsistencias 
   ClientHeight    =   8880
   ClientLeft      =   210
   ClientTop       =   525
   ClientWidth     =   11880
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MousePointer    =   1  'Arrow
   Moveable        =   0   'False
   ScaleHeight     =   8880
   ScaleWidth      =   11880
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame frInscripcionFinales 
      BackColor       =   &H00C0FFFF&
      Height          =   7815
      Left            =   0
      TabIndex        =   0
      Top             =   1200
      Width           =   11895
      Begin MSDataGridLib.DataGrid dtgMaterias 
         Bindings        =   "frmCursadasYAsistencias.frx":0000
         Height          =   4815
         Left            =   120
         TabIndex        =   12
         Top             =   2640
         Width           =   11655
         _ExtentX        =   20558
         _ExtentY        =   8493
         _Version        =   393216
         AllowUpdate     =   0   'False
         AllowArrows     =   -1  'True
         HeadLines       =   1
         RowHeight       =   15
         RowDividerStyle =   0
         FormatLocked    =   -1  'True
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   10
         BeginProperty Column00 
            DataField       =   "Curso"
            Caption         =   "Curso"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   11274
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   "Nombre"
            Caption         =   "Nombre"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   11274
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column02 
            DataField       =   "Division"
            Caption         =   "Div."
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   11274
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column03 
            DataField       =   "Parcial1"
            Caption         =   "1º Parcial"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "#0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   11274
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column04 
            DataField       =   "Recuperatorio1"
            Caption         =   "1º Recup."
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "#0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   11274
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column05 
            DataField       =   "Parcial2"
            Caption         =   "2º Parcial"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "#0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   11274
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column06 
            DataField       =   "Recuperatorio2"
            Caption         =   "2º Recup."
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "#0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   11274
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column07 
            DataField       =   "AsistenciaPorcentaje"
            Caption         =   "% Asist."
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "% ##"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   11274
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column08 
            DataField       =   "Asistencia"
            Caption         =   "Aprob."
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   5
               Format          =   ""
               HaveTrueFalseNull=   1
               TrueValue       =   "Si"
               FalseValue      =   ""
               NullValue       =   ""
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   11274
               SubFormatType   =   7
            EndProperty
         EndProperty
         BeginProperty Column09 
            DataField       =   "AsistenciaHasta"
            Caption         =   "Hasta el"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            MarqueeStyle    =   4
            RecordSelectors =   0   'False
            BeginProperty Column00 
               ColumnWidth     =   645.165
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   4844.977
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   390.047
            EndProperty
            BeginProperty Column03 
               Alignment       =   1
               ColumnWidth     =   840.189
            EndProperty
            BeginProperty Column04 
               Alignment       =   1
               ColumnWidth     =   794.835
            EndProperty
            BeginProperty Column05 
               Alignment       =   1
               ColumnWidth     =   780.095
            EndProperty
            BeginProperty Column06 
               Alignment       =   1
               ColumnWidth     =   824.882
            EndProperty
            BeginProperty Column07 
               Alignment       =   1
               ColumnWidth     =   645.165
            EndProperty
            BeginProperty Column08 
               Alignment       =   2
               ColumnWidth     =   524.976
            EndProperty
            BeginProperty Column09 
               ColumnWidth     =   1124.787
            EndProperty
         EndProperty
      End
      Begin VB.CommandButton cmdSalir 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Salir"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   960
         Left            =   10320
         MouseIcon       =   "frmCursadasYAsistencias.frx":001A
         MousePointer    =   99  'Custom
         Picture         =   "frmCursadasYAsistencias.frx":0324
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Salir"
         Top             =   240
         Width           =   1200
      End
      Begin MSAdodcLib.Adodc adoMaterias 
         Height          =   375
         Left            =   960
         Top             =   3480
         Visible         =   0   'False
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   661
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
         RecordSource    =   $"frmCursadasYAsistencias.frx":0766
         Caption         =   "Materias"
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
         Left            =   240
         Top             =   480
         Visible         =   0   'False
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   582
         ConnectMode     =   1
         CursorLocation  =   3
         IsolationLevel  =   -1
         ConnectionTimeout=   15
         CommandTimeout  =   30
         CursorType      =   2
         LockType        =   1
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
         RecordSource    =   $"frmCursadasYAsistencias.frx":08D6
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
      Begin MSDataListLib.DataCombo dtcCarrera 
         Bindings        =   "frmCursadasYAsistencias.frx":0A54
         Height          =   360
         Left            =   120
         TabIndex        =   6
         Top             =   1440
         Width           =   11535
         _ExtentX        =   20346
         _ExtentY        =   741
         _Version        =   393216
         Enabled         =   0   'False
         Style           =   2
         ListField       =   "Nombre"
         BoundColumn     =   "Carrera"
         Text            =   ""
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
      Begin VB.Label lblAño 
         BackStyle       =   0  'Transparent
         Caption         =   "año"
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
         Left            =   4920
         TabIndex        =   10
         Top             =   2250
         Width           =   975
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Carrera:"
         DataField       =   "Nombre"
         DataSource      =   "adoAlumnos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   8
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Lista de materias que cursa en el año: "
         DataField       =   "Nombre"
         DataSource      =   "adoAlumnos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   7
         Top             =   2280
         Width           =   4695
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Cursadas y Asistencias"
         DataField       =   "Nombre"
         DataSource      =   "adoAlumnos"
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
         Left            =   4440
         TabIndex        =   5
         Top             =   360
         Width           =   3255
      End
   End
   Begin MSComCtl2.MonthView Calendario 
      Height          =   2370
      Left            =   7440
      TabIndex        =   9
      Top             =   1440
      Visible         =   0   'False
      Width           =   2595
      _ExtentX        =   4577
      _ExtentY        =   4180
      _Version        =   393216
      ForeColor       =   -2147483630
      BackColor       =   12632256
      Appearance      =   1
      StartOfWeek     =   72155138
      CurrentDate     =   37600
   End
   Begin VB.Frame frDatosAlumno 
      BackColor       =   &H0080FFFF&
      Height          =   1215
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   11895
      Begin VB.Label lblNombre 
         BackStyle       =   0  'Transparent
         Caption         =   "Nombre del Alumno"
         DataField       =   "Nombre"
         DataSource      =   "adoAlumnos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   360
         Width           =   7575
      End
      Begin VB.Label lblDocumento 
         BackStyle       =   0  'Transparent
         Caption         =   "Documento"
         DataField       =   "Documento"
         DataSource      =   "adoAlumnos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   8880
         TabIndex        =   3
         Top             =   360
         Width           =   1695
      End
      Begin VB.Label lblTipo 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo:"
         DataField       =   "Tipo"
         DataSource      =   "adoAlumnos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   7920
         TabIndex        =   2
         Top             =   360
         Width           =   855
      End
   End
End
Attribute VB_Name = "frmCursadasyAsistencias"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Conexion As New Connection

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Public Sub dtcCarrera_Change()
    adoMaterias.RecordSource = "SELECT Materias.Curso, Materias.Nombre, Finales.Division, Finales.Parcial1, Finales.Recuperatorio1, Finales.Parcial2, Finales.Recuperatorio2, Finales.AsistenciaPorcentaje, Finales.Asistencia, Finales.AsistenciaHasta FROM Finales INNER JOIN Materias ON Finales.Materia = Materias.Codigo Where (((Finales.Alumno) = " & frmConexionAlumnos.adoAlumnos.Recordset!Permiso & ") And ((Finales.Ano) = " & frmConexionAlumnos.adoParametros.Recordset!AñoMatriculacion & ")) AND Materias.Carrera = " & dtcCarrera.BoundText & " AND PerdioCursada = False ORDER BY Finales.Materia"
    adoMaterias.Refresh
End Sub

Private Sub Form_Activate()
    dtgMaterias.SetFocus
End Sub

Private Sub Form_Load()
    Conexion.ConnectionString = ("DSN=Instituto")
    With frmConexionAlumnos
    lblNombre = .lblNombre
    lblTipo = .lblTipo
    lblDocumento = .lblDocumento
    lblAño = .adoParametros.Recordset!AñoMatriculacion
    adoCarreras.RecordSource = "SELECT CarrerasHechas.Permiso, CarrerasHechas.Carrera, CarrerasHechas.Ingreso, Condicion.Condicion, CarrerasHechas.Fecha, CarrerasHechas.Libro, CarrerasHechas.Folio, Carreras.Nombre FROM (CarrerasHechas INNER JOIN Carreras ON CarrerasHechas.Carrera = Carreras.Codigo) INNER JOIN Condicion ON CarrerasHechas.Condición = Condicion.Codigo WHERE CarrerasHechas.Permiso=" & .adoAlumnos.Recordset!Permiso & " AND ((CarrerasHechas.Condición)=1 Or (CarrerasHechas.Condición)=4 Or (CarrerasHechas.Condición)=6)"
    adoCarreras.Refresh
    dtcCarrera.BoundText = adoCarreras.Recordset!Carrera
    If adoCarreras.Recordset.RecordCount > 1 Then dtcCarrera.Enabled = True
    End With
End Sub

