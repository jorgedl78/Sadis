VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmSituacionAcademica 
   Caption         =   "Situación Académica"
   ClientHeight    =   9105
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11895
   LinkTopic       =   "Form1"
   ScaleHeight     =   9105
   ScaleWidth      =   11895
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame frInscripcionFinales 
      BackColor       =   &H00C0FFFF&
      Height          =   7815
      Left            =   0
      TabIndex        =   4
      Top             =   1200
      Width           =   11895
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
         Left            =   9960
         MouseIcon       =   "frmSituacionAcademica.frx":0000
         MousePointer    =   99  'Custom
         Picture         =   "frmSituacionAcademica.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Salir"
         Top             =   1200
         Width           =   1080
      End
      Begin VB.TextBox Text1 
         DataField       =   "Mesa"
         DataSource      =   "adoInscripcion"
         Height          =   285
         Left            =   6360
         TabIndex        =   5
         Text            =   "Text1"
         Top             =   1320
         Visible         =   0   'False
         Width           =   255
      End
      Begin MSAdodcLib.Adodc adoFinales 
         Height          =   330
         Left            =   3840
         Top             =   1320
         Visible         =   0   'False
         Width           =   2295
         _ExtentX        =   4048
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
         RecordSource    =   $"frmSituacionAcademica.frx":074C
         Caption         =   "adoFinales"
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
         Top             =   120
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
         RecordSource    =   $"frmSituacionAcademica.frx":0892
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
      Begin MSDataGridLib.DataGrid dtgSituacionAcademica 
         Bindings        =   "frmSituacionAcademica.frx":0A10
         Height          =   5295
         Left            =   120
         TabIndex        =   6
         Top             =   2280
         Width           =   11655
         _ExtentX        =   20558
         _ExtentY        =   9340
         _Version        =   393216
         AllowUpdate     =   0   'False
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
         ColumnCount     =   7
         BeginProperty Column00 
            DataField       =   "Curso"
            Caption         =   "Curso"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   "0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   "Nombre"
            Caption         =   "Asignatura"
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
         BeginProperty Column02 
            DataField       =   "Cursada"
            Caption         =   "Cursada"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   5
               Format          =   ""
               HaveTrueFalseNull=   1
               TrueValue       =   "Si"
               FalseValue      =   ""
               NullValue       =   ""
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   7
            EndProperty
         EndProperty
         BeginProperty Column03 
            DataField       =   "AsistenciaPorcentaje"
            Caption         =   "%"
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
         BeginProperty Column04 
            DataField       =   "Asistencia"
            Caption         =   "Aprob. Asist."
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   5
               Format          =   ""
               HaveTrueFalseNull=   1
               TrueValue       =   "Si"
               FalseValue      =   ""
               NullValue       =   ""
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   7
            EndProperty
         EndProperty
         BeginProperty Column05 
            DataField       =   "Fecha"
            Caption         =   "Fecha Final"
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
         BeginProperty Column06 
            DataField       =   "Nota"
            Caption         =   "Nota Final"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "#0.00"
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
               ColumnWidth     =   585.071
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   6765.166
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   750.047
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   540.284
            EndProperty
            BeginProperty Column04 
               ColumnWidth     =   989.858
            EndProperty
            BeginProperty Column05 
               Alignment       =   1
               ColumnWidth     =   945.071
            EndProperty
            BeginProperty Column06 
               ColumnWidth     =   854.929
            EndProperty
         EndProperty
      End
      Begin MSDataListLib.DataCombo dtcCarrera 
         Bindings        =   "frmSituacionAcademica.frx":0A29
         Height          =   360
         Left            =   120
         TabIndex        =   7
         Top             =   840
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
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Inscripción a Exámenes Finales"
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
         Left            =   3120
         TabIndex        =   10
         Top             =   120
         Width           =   4695
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Exámenes finales aprobados"
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
         TabIndex        =   9
         Top             =   1800
         Width           =   6375
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
         Top             =   480
         Width           =   1095
      End
   End
   Begin VB.Frame frDatosAlumno 
      BackColor       =   &H0080FFFF&
      Height          =   1215
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11895
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
         TabIndex        =   3
         Top             =   360
         Width           =   855
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
         TabIndex        =   2
         Top             =   360
         Width           =   1695
      End
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
         TabIndex        =   1
         Top             =   360
         Width           =   7575
      End
   End
End
Attribute VB_Name = "frmSituacionAcademica"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Conexion As New Connection
Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub dtcCarrera_Change()
    Me.MousePointer = 11
    adoFinales.RecordSource = "SELECT Materias.Curso, Materias.Nombre, Finales.Cursada, Finales.AsistenciaPorcentaje, Finales.Asistencia, Finales.Fecha, Finales.Nota FROM Finales INNER JOIN Materias ON Finales.Materia = Materias.Codigo" _
    & " WHERE (((Finales.Cursada)=True) AND ((Materias.Detalle)<>4) AND ((Finales.Alumno)=" & frmConexionAlumnos.adoAlumnos.Recordset!Permiso & ") AND ((Materias.Carrera)=" & dtcCarrera.BoundText & ")) ORDER BY Materias.Curso"
    adoFinales.Refresh
    Me.MousePointer = 0
End Sub

Private Sub Form_Load()
    Conexion.ConnectionString = ("DSN=Instituto")
    With frmConexionAlumnos
    lblNombre = .lblNombre
    lblTipo = .lblTipo
    lblDocumento = .lblDocumento
    adoCarreras.RecordSource = "SELECT CarrerasHechas.Permiso, CarrerasHechas.Carrera, CarrerasHechas.Ingreso, Condicion.Condicion, CarrerasHechas.Fecha, CarrerasHechas.Libro, CarrerasHechas.Folio, Carreras.Nombre FROM (CarrerasHechas INNER JOIN Carreras ON CarrerasHechas.Carrera = Carreras.Codigo) INNER JOIN Condicion ON CarrerasHechas.Condición = Condicion.Codigo WHERE CarrerasHechas.Permiso=" & .adoAlumnos.Recordset!Permiso & " AND ((CarrerasHechas.Condición)=1 Or (CarrerasHechas.Condición)=4 Or (CarrerasHechas.Condición)=6)"
    adoCarreras.Refresh
    dtcCarrera.BoundText = adoCarreras.Recordset!Carrera
    If adoCarreras.Recordset.RecordCount > 1 Then dtcCarrera.Enabled = True
    End With
End Sub

