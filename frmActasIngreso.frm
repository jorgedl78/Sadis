VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmActasIngreso 
   Caption         =   "Ingreso de Actas Volantes de Exámenes Finales"
   ClientHeight    =   8595
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   13905
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   8595
   ScaleWidth      =   13905
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtLibro 
      Alignment       =   2  'Center
      Enabled         =   0   'False
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
      Left            =   960
      MaxLength       =   4
      TabIndex        =   33
      Top             =   3720
      Width           =   855
   End
   Begin VB.Frame frNotas 
      Caption         =   "Notas"
      Height          =   6255
      Left            =   120
      TabIndex        =   2
      Top             =   2280
      Width           =   13695
      Begin VB.CommandButton cmdModificarFolio 
         Caption         =   "Modificar Libro y Folio"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3720
         TabIndex        =   36
         Top             =   1440
         Visible         =   0   'False
         Width           =   2895
      End
      Begin MSAdodcLib.Adodc adoInscriptos 
         Height          =   330
         Left            =   7440
         Top             =   4080
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
         RecordSource    =   $"frmActasIngreso.frx":0000
         Caption         =   "Inscriptos"
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
      Begin VB.TextBox txtFolio 
         Alignment       =   2  'Center
         Enabled         =   0   'False
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
         Left            =   2640
         MaxLength       =   4
         TabIndex        =   26
         Top             =   1440
         Width           =   855
      End
      Begin VB.CommandButton cmdAceptar 
         Enabled         =   0   'False
         Height          =   495
         Left            =   4080
         Picture         =   "frmActasIngreso.frx":00FB
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   3240
         Width           =   735
      End
      Begin VB.CheckBox chkAusente 
         Caption         =   "Ausente"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   2280
         TabIndex        =   16
         Top             =   3420
         Width           =   1335
      End
      Begin VB.TextBox txtNota 
         Alignment       =   2  'Center
         Enabled         =   0   'False
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
         Left            =   1200
         MaxLength       =   4
         TabIndex        =   15
         Top             =   3360
         Width           =   855
      End
      Begin VB.CommandButton cmdIngresar 
         Caption         =   "Ingresar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2040
         TabIndex        =   12
         Top             =   360
         Width           =   2415
      End
      Begin MSDataGridLib.DataGrid dtgResultados 
         Bindings        =   "frmActasIngreso.frx":053D
         Height          =   5895
         Left            =   6720
         TabIndex        =   4
         Top             =   240
         Width           =   6735
         _ExtentX        =   11880
         _ExtentY        =   10398
         _Version        =   393216
         AllowUpdate     =   0   'False
         Enabled         =   0   'False
         ColumnHeaders   =   -1  'True
         HeadLines       =   1
         RowHeight       =   15
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
         Caption         =   "Resultados"
         ColumnCount     =   5
         BeginProperty Column00 
            DataField       =   "Permiso"
            Caption         =   "Permiso"
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
            DataField       =   "Documento"
            Caption         =   "Documento"
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
         BeginProperty Column03 
            DataField       =   "Nota"
            Caption         =   "Nota"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   11274
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column04 
            DataField       =   "Ausente"
            Caption         =   "Ausente"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   5
               Format          =   ""
               HaveTrueFalseNull=   1
               TrueValue       =   "Ausente"
               FalseValue      =   ""
               NullValue       =   ""
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   11274
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            MarqueeStyle    =   4
            SizeMode        =   1
            AllowRowSizing  =   0   'False
            AllowSizing     =   0   'False
            BeginProperty Column00 
               ColumnAllowSizing=   -1  'True
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   1080
            EndProperty
            BeginProperty Column02 
            EndProperty
            BeginProperty Column03 
               Alignment       =   1
            EndProperty
            BeginProperty Column04 
            EndProperty
         EndProperty
      End
      Begin VB.Line Line3 
         X1              =   0
         X2              =   6720
         Y1              =   2040
         Y2              =   2040
      End
      Begin VB.Line Line4 
         X1              =   0
         X2              =   6720
         Y1              =   1080
         Y2              =   1080
      End
      Begin VB.Label Label3 
         Caption         =   "Libro:"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   35
         Top             =   1440
         Width           =   615
      End
      Begin VB.Label lblFolio 
         Caption         =   "Folio:"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2040
         TabIndex        =   27
         Top             =   1440
         Width           =   615
      End
      Begin VB.Line Line2 
         BorderWidth     =   3
         X1              =   2040
         X2              =   2760
         Y1              =   5565
         Y2              =   5565
      End
      Begin VB.Label lblTotal 
         Alignment       =   1  'Right Justify
         Caption         =   "0"
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
         Left            =   2160
         TabIndex        =   25
         Top             =   5640
         Width           =   495
      End
      Begin VB.Label lblAusentes 
         Alignment       =   1  'Right Justify
         Caption         =   "0"
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
         Left            =   2160
         TabIndex        =   24
         Top             =   5160
         Width           =   495
      End
      Begin VB.Label lblDesaprobados 
         Alignment       =   1  'Right Justify
         Caption         =   "0"
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
         Left            =   2160
         TabIndex        =   23
         Top             =   4800
         Width           =   495
      End
      Begin VB.Label lblAprobados 
         Alignment       =   1  'Right Justify
         Caption         =   "0"
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
         Left            =   2160
         TabIndex        =   22
         Top             =   4440
         Width           =   495
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         Caption         =   "Total:"
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
         Left            =   480
         TabIndex        =   21
         Top             =   5640
         Width           =   1575
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         Caption         =   "Ausentes:"
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
         Left            =   480
         TabIndex        =   20
         Top             =   5160
         Width           =   1575
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Caption         =   "Desaprobados:"
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
         Left            =   240
         TabIndex        =   19
         Top             =   4800
         Width           =   1815
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "Aprobados:"
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
         Left            =   480
         TabIndex        =   18
         Top             =   4440
         Width           =   1575
      End
      Begin VB.Label lblNota 
         Caption         =   "Nota:"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   600
         TabIndex        =   14
         Top             =   3360
         Width           =   615
      End
      Begin VB.Label lblAlumno 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   600
         TabIndex        =   13
         Top             =   2280
         Width           =   4455
      End
      Begin VB.Line Line1 
         X1              =   0
         X2              =   6720
         Y1              =   4200
         Y2              =   4200
      End
   End
   Begin VB.Frame frActa 
      Caption         =   "Acta"
      Height          =   975
      Left            =   120
      TabIndex        =   1
      Top             =   1320
      Width           =   13695
      Begin MSAdodcLib.Adodc adoActa 
         Height          =   330
         Left            =   4200
         Top             =   360
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
         RecordSource    =   "SELECT * FROM Actas WHERE Mesa=0"
         Caption         =   "Acta"
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
      Begin VB.ComboBox cbNumeroActa 
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   29
         Top             =   480
         Width           =   735
      End
      Begin VB.TextBox txtAñoCursada 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1680
         TabIndex        =   28
         Top             =   480
         Width           =   615
      End
      Begin VB.CommandButton cmdSalir 
         Height          =   675
         Left            =   12720
         Picture         =   "frmActasIngreso.frx":0559
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Salir"
         Top             =   240
         Width           =   675
      End
      Begin VB.Label Label11 
         Caption         =   "Acta Nº:"
         Height          =   255
         Left            =   120
         TabIndex        =   31
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label9 
         Caption         =   "Año Cursada:"
         Height          =   255
         Left            =   1560
         TabIndex        =   30
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Frame frMesa 
      Caption         =   "Mesa"
      Height          =   1335
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   13695
      Begin VB.TextBox Text1 
         DataField       =   "NotaAprobacionFinal"
         DataSource      =   "adoParametros"
         Height          =   285
         Left            =   6240
         TabIndex        =   32
         Text            =   "Text1"
         Top             =   600
         Visible         =   0   'False
         Width           =   150
      End
      Begin MSAdodcLib.Adodc adoParametros 
         Height          =   330
         Left            =   3960
         Top             =   600
         Visible         =   0   'False
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   582
         ConnectMode     =   0
         CursorLocation  =   3
         IsolationLevel  =   -1
         ConnectionTimeout=   15
         CommandTimeout  =   30
         CursorType      =   3
         LockType        =   3
         CommandType     =   2
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
         RecordSource    =   "Parametros"
         Caption         =   "Parametros"
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
      Begin MSAdodcLib.Adodc adoMesa 
         Height          =   330
         Left            =   3840
         Top             =   120
         Visible         =   0   'False
         Width           =   1935
         _ExtentX        =   3413
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
         RecordSource    =   "SELECT FROM Mesa"
         Caption         =   "Mesa"
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
      Begin VB.TextBox txtNumeroDeMesa 
         Height          =   315
         Left            =   840
         TabIndex        =   5
         Top             =   240
         Width           =   735
      End
      Begin VB.Label lblAño 
         Caption         =   "Año:"
         Height          =   375
         Left            =   12600
         TabIndex        =   11
         Top             =   360
         Width           =   735
      End
      Begin VB.Label lblTurno 
         Caption         =   "Turno:"
         Height          =   375
         Left            =   10680
         TabIndex        =   10
         Top             =   360
         Width           =   735
      End
      Begin VB.Label lblMateria 
         Caption         =   "Materia:"
         Height          =   375
         Left            =   2160
         TabIndex        =   9
         Top             =   840
         Width           =   10935
      End
      Begin VB.Label lblDivision 
         Caption         =   "División:"
         Height          =   375
         Left            =   120
         TabIndex        =   8
         Top             =   840
         Width           =   855
      End
      Begin VB.Label lblCarrera 
         Caption         =   "Carrera:"
         Height          =   375
         Left            =   2160
         TabIndex        =   7
         Top             =   300
         Width           =   8415
      End
      Begin VB.Label Label1 
         Caption         =   "Nº Mesa:"
         Height          =   375
         Left            =   120
         TabIndex        =   6
         Top             =   300
         Width           =   735
      End
   End
   Begin VB.Label Label2 
      Caption         =   "Folio:"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2880
      TabIndex        =   34
      Top             =   2760
      Width           =   615
   End
End
Attribute VB_Name = "frmActasIngreso"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Conexion As New Connection
Dim Auxiliar As Recordset
Dim Aprobados As Integer
Dim Desaprobados As Integer
Dim Ausentes As Integer
Dim EstadoNota As String
Dim EstadoAusente As String
Private Sub cbNumeroActa_Click()
    adoActa.RecordSource = "SELECT * From Actas WHERE Actas.Mesa=" & adoMesa.Recordset!Numero & " AND Actas.Acta= " & cbNumeroActa.Text
    adoActa.Refresh
    txtAñoCursada = adoActa.Recordset!Ano
    adoInscriptos.RecordSource = "SELECT Alumnos.Permiso, Alumnos.Tipo & ' ' & Alumnos.Documento AS Documento, Alumnos.Nombre, Inscripciones.Nota, Inscripciones.Ausente FROM Alumnos INNER JOIN Inscripciones ON Alumnos.Permiso = Inscripciones.Alumno WHERE Inscripciones.Mesa=" & adoMesa.Recordset!Numero & " AND Inscripciones.Acta=" & cbNumeroActa.Text & " ORDER BY Alumnos.Nombre"
    adoInscriptos.Refresh
    Aprobados = 0
    Desaprobados = 0
    Ausentes = 0
    If adoActa.Recordset!Ingresada = True Then
        If IsNull(adoActa.Recordset!Libro) Then
            txtLibro = 0
        Else
            txtLibro = adoActa.Recordset!Libro
        End If
        txtFolio = adoActa.Recordset!Folio
        Aprobados = adoActa.Recordset!Aprobados
        Desaprobados = adoActa.Recordset!Desaprobados
        Ausentes = adoActa.Recordset!Ausentes
        cmdIngresar.Caption = "Modificar"
        dtgResultados.Enabled = True
        cmdModificarFolio.Visible = True
    Else
        cmdIngresar.Caption = "Ingresar"
        dtgResultados.Enabled = False
        cmdModificarFolio.Visible = False
    End If
    ConteoDeTotales
    cmdIngresar.SetFocus
End Sub

Private Sub chkAusente_KeyPress(KeyAscii As Integer)
    If keyascci = 13 Then cmdAceptar_Click
End Sub

Private Sub cmdAceptar_Click()
    If Val(txtNota) > 10 Or Val(txtNota) < 0 Then
        Respuesta = MsgBox("Nota Incorrecta", 0, "Error")
        txtNota = ""
        txtNota.SetFocus
        Exit Sub
    Else
        Conexion.Open
        If Val(txtNota) >= adoParametros.Recordset!NotaAprobacionFinal Then
            'aprobo el final(se actualiza Aprobada=true, nota, fecha, folio, etc)
            chkAusente.Value = 0
            Aprobados = Aprobados + 1
            If cmdIngresar.Caption = "Ingresar" Then
                Set Auxiliar = Conexion.Execute("UPDATE Finales SET Finales.Nota = " & Replace(Val(txtNota), ",", ".") & ", Finales.Fecha = '" & DateValue(adoMesa.Recordset!Fecha) & "', Finales.Libro = " & txtLibro & ", Finales.Folio = " & txtFolio & " , Finales.Mesa = " & txtNumeroDeMesa & ", Finales.Acta = " & cbNumeroActa.Text & ", Finales.Aprobada = True, Finales.Habilitada=True, Finales.CantidadMesas = [CantidadMesas]+1, Finales.PerdioTurno=False WHERE (((Finales.Alumno)=" & adoInscriptos.Recordset!Permiso & ") AND ((Finales.Materia)=" & adoMesa.Recordset!CodigoMateria & ") AND ((Finales.Cursada)=True))")
            Else
                Set Auxiliar = Conexion.Execute("UPDATE Finales SET Finales.Nota = " & Replace(Val(txtNota), ",", ".") & ", Finales.Fecha = '" & DateValue(adoMesa.Recordset!Fecha) & "', Finales.Libro = " & txtLibro & ", Finales.Folio = " & txtFolio & " , Finales.Mesa = " & txtNumeroDeMesa & ", Finales.Acta = " & cbNumeroActa.Text & ", Finales.Aprobada = True, Finales.Habilitada=True WHERE Finales.Alumno=" & adoInscriptos.Recordset!Permiso & " AND Finales.Materia=" & adoMesa.Recordset!CodigoMateria & " AND Finales.Cursada=True")
            End If
        Else
            If chkAusente.Value = 1 Then
                Ausentes = Ausentes + 1
                'perder turno siguiente
                If cmdIngresar.Caption = "Modificar" Then
                    Set Auxiliar = Conexion.Execute("UPDATE Finales SET Finales.Nota = 0, Finales.Fecha = Null, Finales.Folio = 0 , Finales.Libro=0,  Finales.Mesa = 0, Finales.Acta = 0, Finales.Aprobada = False WHERE (((Finales.Alumno)=" & adoInscriptos.Recordset!Permiso & ") AND ((Finales.Materia)=" & adoMesa.Recordset!CodigoMateria & ") AND ((Finales.Cursada)=True))")
                Else
                    Set Auxiliar = Conexion.Execute("UPDATE Finales SET Finales.CantidadMesas = [CantidadMesas]+1 WHERE (((Finales.Alumno)=" & adoInscriptos.Recordset!Permiso & ") AND ((Finales.Materia)=" & adoMesa.Recordset!CodigoMateria & ") AND ((Finales.Cursada)=True))")
                End If
            Else
                If cmdIngresar.Caption = "Ingresar" Then
                    Set Auxiliar = Conexion.Execute("UPDATE Finales SET Finales.Nota = 0, Finales.Fecha = Null, Finales.Folio = 0, Finales.Libro=0, Finales.Mesa = 0, Finales.Acta = 0, Finales.Aprobada = False, Finales.CantidadMesas = [CantidadMesas]+1, Finales.PerdioTurno=False WHERE (((Finales.Alumno)=" & adoInscriptos.Recordset!Permiso & ") AND ((Finales.Materia)=" & adoMesa.Recordset!CodigoMateria & ") AND ((Finales.Cursada)=True))")
                Else
                    Set Auxiliar = Conexion.Execute("UPDATE Finales SET Finales.Nota = 0, Finales.Fecha = Null, Finales.Folio = 0, Finales.Libro=0, Finales.Mesa = 0, Finales.Acta = 0, Finales.Aprobada = False, Finales.PerdioTurno = False WHERE (((Finales.Alumno)=" & adoInscriptos.Recordset!Permiso & ") AND ((Finales.Materia)=" & adoMesa.Recordset!CodigoMateria & ") AND ((Finales.Cursada)=True))")
                End If
                Desaprobados = Desaprobados + 1
            End If
        End If
        Set Auxiliar = Conexion.Execute("UPDATE Inscripciones SET Inscripciones.Nota = " & Replace(Val(txtNota), ",", ".") & " , Inscripciones.Ausente = " & chkAusente.Value & " WHERE Inscripciones.Mesa=" & adoMesa.Recordset!Numero & " AND Inscripciones.Acta=" & cbNumeroActa.Text & " AND Inscripciones.Alumno=" & adoInscriptos.Recordset!Permiso)
        Conexion.Close
        Registro = adoInscriptos.Recordset.Bookmark
        adoInscriptos.Refresh
        If cmdIngresar.Caption = "Ingresar" Then
            adoInscriptos.Recordset.Move Registro
            'adoInscriptos.Recordset.MoveNext
            ConteoDeTotales
            If adoInscriptos.Recordset.EOF = True Then 'ya terminó el acta
                GrabaActa
                frMesa.Enabled = True
                frActa.Enabled = True
                cmdIngresar.Caption = "Modificar"
                cmdIngresar.Enabled = True
                lblNota.Enabled = False
                txtNota.Enabled = False
                chkAusente.Enabled = False
                cmdAceptar.Enabled = False
            Else
                PedirNota
            End If
        Else
            cmdIngresar.Caption = ""
            Aprobados = 0
            Desaprobados = 0
            Ausentes = 0
            adoInscriptos.Recordset.MoveFirst
            For i = 1 To adoInscriptos.Recordset.RecordCount
                If adoInscriptos.Recordset!Nota >= adoParametros.Recordset!NotaAprobacionFinal Then
                    Aprobados = Aprobados + 1
                ElseIf adoInscriptos.Recordset!Ausente = True Then Ausentes = Ausentes + 1
                Else
                    Desaprobados = Desaprobados + 1
                End If
                adoInscriptos.Recordset.MoveNext
            Next i
            ConteoDeTotales
            GrabaActa
            cmdIngresar.Caption = "Modificar"
            adoInscriptos.Recordset.MoveFirst
            adoInscriptos.Recordset.Move Registro - 1
            cmdIngresar.Enabled = True
            lblNota.Enabled = False
            txtNota.Enabled = False
            chkAusente.Enabled = False
            cmdAceptar.Enabled = False
            frMesa.Enabled = True
            frActa.Enabled = True
        End If
    End If
End Sub

Private Sub cmdIngresar_Click()
    If frmIdentificacion.Permisos!IngresarActas = False Then MsgBox ("Usted no tiene permiso para ingresar o modificar Actas"): Exit Sub
    cmdIngresar.Enabled = False
    frMesa.Enabled = False
    frActa.Enabled = False
    cmdIngresar.Enabled = False
    If cmdIngresar.Caption = "Ingresar" Then
        txtLibro.Enabled = True
        txtFolio.Enabled = True
        txtLibro.SetFocus
    Else
        If adoInscriptos.Recordset!Nota >= adoParametros.Recordset!NotaAprobacionFinal Then
            EstadoNota = "Aprobado"
        Else
            EstadoNota = "Desaprobado"
        End If
        If adoInscriptos.Recordset!Ausente = True Then
            EstadoAusente = "Si"
        Else
            EstadoAusente = "No"
        End If
        PedirNota
        lblAlumno.Enabled = True
        lblNota.Enabled = True
        txtNota.Enabled = True
        chkAusente.Enabled = True
        cmdAceptar.Enabled = True
        txtNota.SetFocus
    End If
End Sub

Private Sub cmdModificarFolio_Click()
    frmModificarFolioActa.lblMesa = txtNumeroDeMesa
    frmModificarFolioActa.txtLibroActual = txtLibro
    frmModificarFolioActa.txtFolioActual = txtFolio
    frmModificarFolioActa.Show 1
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub dtgResultados_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    If cmdIngresar.Caption = "Modificar" Then PedirNota
End Sub

Private Sub Form_Activate()
    Conexion.ConnectionString = ("DSN=Instituto")
    txtNumeroDeMesa.SetFocus
End Sub

Private Sub txtFolio_GotFocus()
    txtFolio = ""
End Sub

Private Sub txtFolio_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Val(txtFolio) > 0 Then
            lblFolio.Enabled = False
            txtFolio.Enabled = False
            txtLibro.Enabled = False
            lblNota.Enabled = True
            txtNota.Enabled = True
            chkAusente.Enabled = True
            cmdAceptar.Enabled = True
            PedirNota
            Exit Sub
        Else
            MsgBox ("El Nro. de Folio no corresponde")
            txtFolio.SetFocus
        End If
    End If
    If (KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 8 Then
   
    Else
        KeyAscii = 0
    End If
End Sub

Private Sub txtLibro_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtFolio.SetFocus
    End If
End Sub

Private Sub txtNota_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cmdAceptar_Click
    End If
    If (KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 8 Or KeyAscii = 46 Then
    Else
        KeyAscii = 0
    End If
End Sub

Private Sub txtNumeroDeMesa_Click()
    txtNumeroDeMesa = ""
End Sub

Private Sub txtNumeroDeMesa_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        adoMesa.RecordSource = "SELECT Mesas.Numero, Mesas.Materia AS CodigoMateria, Mesas.Division, Carreras.Abreviatura AS Carrera, Materias.Abreviatura AS Materia, Meses.Nombre, Mesas.Ano, Mesas.Fecha, Mesas.Actas, Mesas.Impresas FROM ((Mesas INNER JOIN Materias ON Mesas.Materia = Materias.Codigo) INNER JOIN Carreras ON Materias.Carrera = Carreras.Codigo) INNER JOIN Meses ON Mesas.Turno = Meses.Numero WHERE Mesas.Numero=" & txtNumeroDeMesa
        adoMesa.Refresh
        If adoMesa.Recordset.RecordCount = 1 Then
            cbNumeroActa.Clear
            lblCarrera = adoMesa.Recordset!Carrera
            lblMateria = adoMesa.Recordset!Materia
            lblDivision = "División: " & adoMesa.Recordset!Division
            lblTurno = adoMesa.Recordset!Nombre
            lblAño = adoMesa.Recordset!Ano
            If adoMesa.Recordset!Impresas = False Then
                MsgBox ("Las actas no fueron impresas")
                txtNumeroDeMesa.Text = ""
                txtNumeroDeMesa.SetFocus
                Exit Sub
            End If
            For i = 0 To adoMesa.Recordset!Actas - 1
                cbNumeroActa.List(i) = i + 1
            Next i
            cbNumeroActa.Text = cbNumeroActa.List(0)
        Else
            MsgBox ("No existe ninguna mesa con ese número")
            txtNumeroDeMesa.Text = ""
            txtNumeroDeMesa.SetFocus
        End If
    End If
    If (KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 8 Then
    Else
        KeyAscii = 0
    End If
End Sub
Private Function PedirNota()
    lblAlumno = adoInscriptos.Recordset!Nombre
    txtNota = ""
    'txtNota = Format(adoInscriptos.Recordset!Nota, "0.00")
    If adoInscriptos.Recordset!Ausente = True Then
        chkAusente.Value = 1
    Else
        chkAusente.Value = 0
    End If
End Function

Private Function ConteoDeTotales()
    lblAprobados = Aprobados
    lblDesaprobados = Desaprobados
    lblAusentes = Ausentes
    lblTotal = adoInscriptos.Recordset.RecordCount
End Function

Private Function GrabaActa()
    adoActa.Recordset!Aprobados = Aprobados
    adoActa.Recordset!Desaprobados = Desaprobados
    adoActa.Recordset!Ausentes = Ausentes
    adoActa.Recordset!Libro = txtLibro
    adoActa.Recordset!Folio = txtFolio
    adoActa.Recordset!Ingresada = True
    adoActa.Recordset.Update
End Function

