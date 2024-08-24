VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmPlanes 
   BackColor       =   &H8000000B&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Planes de Estudio"
   ClientHeight    =   10620
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   14910
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   10620
   ScaleWidth      =   14910
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtAbrebiatura1Carrera 
      Enabled         =   0   'False
      Height          =   315
      Left            =   120
      MaxLength       =   40
      TabIndex        =   13
      Top             =   1800
      Width           =   4095
   End
   Begin VB.Frame frMaterias 
      Caption         =   "Materias"
      Enabled         =   0   'False
      Height          =   7815
      Left            =   0
      TabIndex        =   1
      Top             =   2760
      Width           =   14775
      Begin VB.TextBox txtCursoAnalitico 
         Height          =   315
         Left            =   8040
         TabIndex        =   65
         Top             =   1440
         Width           =   735
      End
      Begin VB.CommandButton cmdCompuestaPor 
         Caption         =   "Compuesta Por"
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
         Height          =   375
         Left            =   6360
         TabIndex        =   58
         Top             =   120
         Width           =   1455
      End
      Begin VB.CommandButton cmdAgregarCorrelativa 
         Height          =   615
         Left            =   9840
         Picture         =   "frmPlanes.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   57
         ToolTipText     =   "Agregar"
         Top             =   6840
         Width           =   615
      End
      Begin VB.CommandButton cmdEliminarCorrelativa 
         Enabled         =   0   'False
         Height          =   615
         Left            =   10920
         Picture         =   "frmPlanes.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   56
         ToolTipText     =   "Eliminar"
         Top             =   6840
         Width           =   615
      End
      Begin MSAdodcLib.Adodc adoCorrelativas 
         Height          =   330
         Left            =   9840
         Top             =   7080
         Visible         =   0   'False
         Width           =   2775
         _ExtentX        =   4895
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
         RecordSource    =   $"frmPlanes.frx":0884
         Caption         =   "Correlativas"
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
      Begin MSAdodcLib.Adodc adoModalidadMaterias 
         Height          =   330
         Left            =   240
         Top             =   2880
         Visible         =   0   'False
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   582
         ConnectMode     =   0
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
         RecordSource    =   "SELECT * FROM Modalidad ORDER BY Codigo"
         Caption         =   "ModalidadMateria"
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
      Begin MSDataListLib.DataCombo dtcModalidadMateria 
         Bindings        =   "frmPlanes.frx":0954
         Height          =   315
         Left            =   9360
         TabIndex        =   54
         Top             =   2760
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         Style           =   2
         ListField       =   "Modalidad"
         BoundColumn     =   "Codigo"
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo dtcDetalleMateria 
         Bindings        =   "frmPlanes.frx":0977
         Height          =   315
         Left            =   9120
         TabIndex        =   53
         Top             =   1440
         Width           =   4095
         _ExtentX        =   7223
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         Style           =   2
         ListField       =   "Detalle"
         BoundColumn     =   "Codigo"
         Text            =   ""
      End
      Begin MSAdodcLib.Adodc adoDetalles 
         Height          =   330
         Left            =   240
         Top             =   3240
         Visible         =   0   'False
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   582
         ConnectMode     =   0
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
         RecordSource    =   "SELECT * FROM Detalles"
         Caption         =   "Detalles"
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
      Begin MSAdodcLib.Adodc adoMaterias 
         Height          =   330
         Left            =   240
         Top             =   2520
         Visible         =   0   'False
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   582
         ConnectMode     =   0
         CursorLocation  =   3
         IsolationLevel  =   -1
         ConnectionTimeout=   15
         CommandTimeout  =   30
         CursorType      =   2
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
         RecordSource    =   "SELECT * FROM Materias WHERE Codigo = 0"
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
      Begin VB.TextBox txtCursoMateria 
         Enabled         =   0   'False
         Height          =   315
         Left            =   120
         TabIndex        =   52
         Top             =   480
         Width           =   375
      End
      Begin MSDataGridLib.DataGrid dtgCorrelativas 
         Bindings        =   "frmPlanes.frx":0991
         Height          =   2535
         Left            =   8040
         TabIndex        =   45
         Top             =   3840
         Width           =   5415
         _ExtentX        =   9551
         _ExtentY        =   4471
         _Version        =   393216
         AllowUpdate     =   0   'False
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
         ColumnCount     =   4
         BeginProperty Column00 
            DataField       =   "Correlativa"
            Caption         =   "Correlativa"
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
         BeginProperty Column01 
            DataField       =   "Nombre"
            Caption         =   "Nombre"
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
            DataField       =   "Curso"
            Caption         =   "Curso"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column03 
            DataField       =   "PorFinal"
            Caption         =   "PorFinal"
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
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            MarqueeStyle    =   4
            BeginProperty Column00 
               ColumnWidth     =   915,024
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   2534,74
            EndProperty
            BeginProperty Column02 
               Alignment       =   1
               ColumnWidth     =   615,118
            EndProperty
            BeginProperty Column03 
               Alignment       =   2
               ColumnWidth     =   705,26
            EndProperty
         EndProperty
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H8000000C&
         BorderStyle     =   0  'None
         Height          =   7245
         Left            =   13680
         TabIndex        =   38
         Top             =   480
         Width           =   975
         Begin VB.CommandButton cmdBorrarMateria 
            Enabled         =   0   'False
            Height          =   615
            Left            =   240
            Picture         =   "frmPlanes.frx":09AF
            Style           =   1  'Graphical
            TabIndex        =   44
            ToolTipText     =   "Borrar"
            Top             =   3480
            Width           =   615
         End
         Begin VB.CommandButton cmdCancelarMateria 
            Enabled         =   0   'False
            Height          =   615
            Left            =   240
            Picture         =   "frmPlanes.frx":0DF1
            Style           =   1  'Graphical
            TabIndex        =   43
            ToolTipText     =   "Cancelar"
            Top             =   4080
            Width           =   615
         End
         Begin VB.CommandButton cmdModificarMateria 
            Height          =   615
            Left            =   240
            MaskColor       =   &H00808080&
            Picture         =   "frmPlanes.frx":1233
            Style           =   1  'Graphical
            TabIndex        =   42
            ToolTipText     =   "Modificar"
            Top             =   2280
            UseMaskColor    =   -1  'True
            Width           =   615
         End
         Begin VB.CommandButton cmdGuardarMateria 
            Enabled         =   0   'False
            Height          =   615
            Left            =   240
            Picture         =   "frmPlanes.frx":1675
            Style           =   1  'Graphical
            TabIndex        =   41
            ToolTipText     =   "Guardar"
            Top             =   2880
            Width           =   615
         End
         Begin VB.CommandButton cmdAgregarMateria 
            Height          =   615
            Left            =   240
            Picture         =   "frmPlanes.frx":1AB7
            Style           =   1  'Graphical
            TabIndex        =   40
            ToolTipText     =   "Nuevo"
            Top             =   1680
            Width           =   615
         End
         Begin VB.CommandButton cmdSalir 
            Height          =   615
            Left            =   240
            Picture         =   "frmPlanes.frx":1EF9
            Style           =   1  'Graphical
            TabIndex        =   39
            ToolTipText     =   "Salir"
            Top             =   4680
            Width           =   615
         End
      End
      Begin MSDataGridLib.DataGrid dtgMaterias 
         Bindings        =   "frmPlanes.frx":233B
         Height          =   6615
         Left            =   120
         TabIndex        =   31
         Top             =   1080
         Width           =   7725
         _ExtentX        =   13626
         _ExtentY        =   11668
         _Version        =   393216
         AllowUpdate     =   0   'False
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
         ColumnCount     =   3
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
            DataField       =   "Codigo"
            Caption         =   "Codigo"
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
            DataField       =   "Abreviatura"
            Caption         =   "Abreviatura"
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
         SplitCount      =   1
         BeginProperty Split0 
            MarqueeStyle    =   4
            BeginProperty Column00 
               ColumnWidth     =   510,236
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   824,882
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   5699,906
            EndProperty
         EndProperty
      End
      Begin VB.TextBox txtAbreviatura2Materia 
         Enabled         =   0   'False
         Height          =   315
         Left            =   8040
         MaxLength       =   6
         TabIndex        =   28
         Top             =   2760
         Width           =   1095
      End
      Begin VB.TextBox txtAbreviatura1Materia 
         Enabled         =   0   'False
         Height          =   315
         Left            =   8040
         MaxLength       =   30
         TabIndex        =   27
         Top             =   2160
         Width           =   5175
      End
      Begin VB.TextBox txtNombreMateria 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1560
         MaxLength       =   123
         TabIndex        =   26
         Top             =   480
         Width           =   11895
      End
      Begin VB.TextBox txtCodigoDesbloqueadoMateria 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1155
         MaxLength       =   2
         TabIndex        =   25
         Top             =   480
         Width           =   270
      End
      Begin VB.TextBox txtCodigoBloqueadoMateria 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   315
         Left            =   600
         MaxLength       =   5
         TabIndex        =   24
         Top             =   480
         Width           =   550
      End
      Begin VB.Label Label20 
         Caption         =   "Curso Analítco:"
         Height          =   255
         Left            =   8040
         TabIndex        =   64
         Top             =   1200
         Width           =   1335
      End
      Begin VB.Label Label18 
         Caption         =   "Materias:"
         Height          =   255
         Left            =   120
         TabIndex        =   47
         Top             =   840
         Width           =   735
      End
      Begin VB.Label Label12 
         Caption         =   "Correlativas:"
         Height          =   255
         Left            =   8160
         TabIndex        =   46
         Top             =   3600
         Width           =   975
      End
      Begin VB.Label Label11 
         Caption         =   "Modalidad:"
         Height          =   255
         Left            =   9360
         TabIndex        =   30
         Top             =   2520
         Width           =   975
      End
      Begin VB.Label Label10 
         Caption         =   "Detalle:"
         Height          =   255
         Left            =   9360
         TabIndex        =   29
         Top             =   1200
         Width           =   615
      End
      Begin VB.Label Label17 
         Caption         =   "Curso:"
         Height          =   255
         Left            =   120
         TabIndex        =   23
         Top             =   240
         Width           =   495
      End
      Begin VB.Label Label16 
         Caption         =   "Codigo:"
         Height          =   255
         Left            =   840
         TabIndex        =   22
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label15 
         Caption         =   "Nombre:"
         Height          =   255
         Left            =   1560
         TabIndex        =   21
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label14 
         Caption         =   "Nombre Usual:"
         Height          =   255
         Left            =   8040
         TabIndex        =   20
         Top             =   1920
         Width           =   1215
      End
      Begin VB.Label Label13 
         Caption         =   "Abrebiatura:"
         Height          =   255
         Left            =   8040
         TabIndex        =   19
         Top             =   2520
         Width           =   975
      End
   End
   Begin VB.Frame frCarreras 
      Caption         =   "Carreras"
      Height          =   2775
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   14775
      Begin VB.CommandButton cmdOrdenPlan 
         Caption         =   "Orden"
         Height          =   495
         Left            =   11040
         TabIndex        =   63
         Top             =   840
         Width           =   615
      End
      Begin VB.CommandButton cmdAgregarTitulo 
         Caption         =   "Agregar Título"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   9840
         TabIndex        =   60
         Top             =   1440
         Width           =   1575
      End
      Begin VB.CommandButton cmdQuitarTitulo 
         Caption         =   "Quitar Título"
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
         Height          =   300
         Left            =   11520
         TabIndex        =   59
         Top             =   1440
         Width           =   1575
      End
      Begin MSDataListLib.DataCombo dtcModalidadCarrera 
         Bindings        =   "frmPlanes.frx":2355
         Height          =   315
         Left            =   3480
         TabIndex        =   55
         Top             =   1080
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         Style           =   2
         ListField       =   "Modalidad"
         BoundColumn     =   "Codigo"
         Text            =   "DataCombo1"
      End
      Begin MSAdodcLib.Adodc Auxiliar 
         Height          =   330
         Left            =   9240
         Top             =   120
         Visible         =   0   'False
         Width           =   2055
         _ExtentX        =   3625
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
         RecordSource    =   "SELECT * FROM Dias"
         Caption         =   "Auxiliar"
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
      Begin MSAdodcLib.Adodc adoTitulos 
         Height          =   330
         Left            =   11280
         Top             =   120
         Visible         =   0   'False
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   582
         ConnectMode     =   0
         CursorLocation  =   3
         IsolationLevel  =   -1
         ConnectionTimeout=   15
         CommandTimeout  =   30
         CursorType      =   2
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
      Begin VB.TextBox txtNombreCarrera 
         Height          =   315
         Left            =   120
         MaxLength       =   120
         TabIndex        =   51
         Top             =   480
         Visible         =   0   'False
         Width           =   10935
      End
      Begin VB.Frame frMostrar 
         Caption         =   "Mostrar"
         Height          =   855
         Left            =   11880
         TabIndex        =   48
         Top             =   360
         Width           =   1215
         Begin VB.OptionButton optMostar 
            Caption         =   "Todas"
            Height          =   195
            Index           =   1
            Left            =   120
            TabIndex        =   50
            Top             =   480
            Width           =   975
         End
         Begin VB.OptionButton optMostar 
            Caption         =   "Vigentes"
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   49
            Top             =   240
            Value           =   -1  'True
            Width           =   975
         End
      End
      Begin VB.Frame frBotonesCarreras 
         BackColor       =   &H8000000C&
         BorderStyle     =   0  'None
         Height          =   2445
         Left            =   13320
         TabIndex        =   32
         Top             =   195
         Width           =   1335
         Begin VB.CommandButton cmdAgregarCarrera 
            Height          =   615
            Left            =   120
            Picture         =   "frmPlanes.frx":2378
            Style           =   1  'Graphical
            TabIndex        =   37
            ToolTipText     =   "Nuevo"
            Top             =   0
            Width           =   615
         End
         Begin VB.CommandButton cmdGuardarCarrera 
            Enabled         =   0   'False
            Height          =   615
            Left            =   120
            Picture         =   "frmPlanes.frx":27BA
            Style           =   1  'Graphical
            TabIndex        =   36
            ToolTipText     =   "Guardar"
            Top             =   1200
            Width           =   615
         End
         Begin VB.CommandButton cmdModificarCarrera 
            Enabled         =   0   'False
            Height          =   615
            Left            =   120
            MaskColor       =   &H00808080&
            Picture         =   "frmPlanes.frx":2BFC
            Style           =   1  'Graphical
            TabIndex        =   35
            ToolTipText     =   "Modificar"
            Top             =   600
            UseMaskColor    =   -1  'True
            Width           =   615
         End
         Begin VB.CommandButton cmdCancelarCarrera 
            Enabled         =   0   'False
            Height          =   615
            Left            =   720
            Picture         =   "frmPlanes.frx":303E
            Style           =   1  'Graphical
            TabIndex        =   34
            ToolTipText     =   "Cancelar"
            Top             =   1800
            Width           =   615
         End
         Begin VB.CommandButton cmdBorrarCarrera 
            Enabled         =   0   'False
            Height          =   615
            Left            =   120
            Picture         =   "frmPlanes.frx":3480
            Style           =   1  'Graphical
            TabIndex        =   33
            ToolTipText     =   "Borrar"
            Top             =   1800
            Width           =   615
         End
      End
      Begin VB.TextBox txtCodigoCarrera 
         Enabled         =   0   'False
         Height          =   315
         Left            =   120
         TabIndex        =   18
         Top             =   1080
         Width           =   615
      End
      Begin MSDataListLib.DataList dtlTitulos 
         Bindings        =   "frmPlanes.frx":38C2
         Height          =   795
         Left            =   5400
         TabIndex        =   16
         Top             =   1800
         Width           =   7815
         _ExtentX        =   13785
         _ExtentY        =   1402
         _Version        =   393216
         IntegralHeight  =   0   'False
         ListField       =   "Titulo"
         BoundColumn     =   "Carrera"
      End
      Begin VB.CheckBox chkVigente 
         Enabled         =   0   'False
         Height          =   315
         Left            =   9360
         TabIndex        =   15
         Top             =   1080
         Width           =   255
      End
      Begin VB.TextBox txtAbrebiatura2Carrera 
         Enabled         =   0   'False
         Height          =   315
         Left            =   4320
         MaxLength       =   6
         TabIndex        =   14
         Top             =   1800
         Width           =   975
      End
      Begin VB.TextBox txtAños 
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   11274
            SubFormatType   =   0
         EndProperty
         Enabled         =   0   'False
         Height          =   315
         Left            =   2760
         TabIndex        =   12
         Top             =   1080
         Width           =   495
      End
      Begin VB.TextBox txtResolucion 
         Enabled         =   0   'False
         Height          =   315
         Left            =   840
         MaxLength       =   80
         TabIndex        =   3
         Top             =   1080
         Width           =   1815
      End
      Begin MSDataListLib.DataCombo dtcCarreras 
         Bindings        =   "frmPlanes.frx":38DB
         Height          =   315
         Left            =   120
         TabIndex        =   2
         Top             =   480
         Width           =   11655
         _ExtentX        =   20558
         _ExtentY        =   556
         _Version        =   393216
         MatchEntry      =   -1  'True
         Style           =   2
         ListField       =   "Nombre"
         BoundColumn     =   "Codigo"
         Text            =   ""
      End
      Begin MSAdodcLib.Adodc adoCarreras 
         Height          =   330
         Left            =   4680
         Top             =   120
         Visible         =   0   'False
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   582
         ConnectMode     =   0
         CursorLocation  =   3
         IsolationLevel  =   -1
         ConnectionTimeout=   15
         CommandTimeout  =   30
         CursorType      =   2
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
         RecordSource    =   "SELECT* FROM Carreras WHERE Eliminada = 0 AND Vigente = 1 ORDER BY Nombre"
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
      Begin MSAdodcLib.Adodc adoModalidadCarreras 
         Height          =   330
         Left            =   1800
         Top             =   120
         Visible         =   0   'False
         Width           =   2775
         _ExtentX        =   4895
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
         RecordSource    =   "SELECT * FROM Modalidad ORDER BY Codigo"
         Caption         =   "ModalidadCarreras"
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
      Begin MSAdodcLib.Adodc adoCaracteristica 
         Height          =   330
         Left            =   6840
         Top             =   120
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
         RecordSource    =   "SELECT * FROM [Caracteristica de carrera]"
         Caption         =   "Caracteristica"
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
      Begin MSDataListLib.DataCombo dtcCaracteristica 
         Bindings        =   "frmPlanes.frx":38F5
         Height          =   315
         Left            =   5640
         TabIndex        =   62
         Top             =   1080
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         Style           =   2
         ListField       =   "Caracteristica"
         BoundColumn     =   "Codigo"
         Text            =   "DataCombo1"
      End
      Begin VB.Label Label19 
         Caption         =   "Característica de las materias:"
         Height          =   255
         Left            =   5640
         TabIndex        =   61
         Top             =   840
         Width           =   2295
      End
      Begin VB.Label Label9 
         Caption         =   "Diferentes Posibilidades de Títulos:"
         Height          =   255
         Left            =   5400
         TabIndex        =   17
         Top             =   1560
         Width           =   2535
      End
      Begin VB.Label Label8 
         Caption         =   "Vigente:"
         Height          =   255
         Left            =   9240
         TabIndex        =   11
         Top             =   840
         Width           =   615
      End
      Begin VB.Label Label7 
         Caption         =   "Código:"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   840
         Width           =   735
      End
      Begin VB.Label Label6 
         Caption         =   "Resolución:"
         Height          =   255
         Left            =   840
         TabIndex        =   9
         Top             =   840
         Width           =   855
      End
      Begin VB.Label Label5 
         Caption         =   "Años:"
         Height          =   255
         Left            =   2760
         TabIndex        =   8
         Top             =   840
         Width           =   495
      End
      Begin VB.Label Label4 
         Caption         =   "Nombre Usual:"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   1560
         Width           =   1215
      End
      Begin VB.Label Label3 
         Caption         =   "Abrebiatura:"
         Height          =   255
         Left            =   4320
         TabIndex        =   6
         Top             =   1560
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "Modalidad:"
         Height          =   255
         Left            =   3600
         TabIndex        =   5
         Top             =   840
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Carrera:"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   735
      End
   End
End
Attribute VB_Name = "frmPlanes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim EstadoCarreras As String
Dim EstadoMaterias As String
Dim i As Integer
Dim MateriaActual As String
Dim Conexion As New Connection

Private Sub cmdAgregarCarrera_Click()
    EstadoCarreras = "Nuevo"
    LimpiarCarreras
    HabilitarCarreras
    frMostrar.Enabled = False
    Auxiliar.RecordSource = "SELECT IIf(IsNull(MAX(Codigo)), 0, Max(Codigo)) AS Ultimo FROM Carreras WHERE Eliminada =0"
    Auxiliar.Refresh
    txtCodigoCarrera = Auxiliar.Recordset!Ultimo + 1
    cmdAgregarCarrera.Enabled = False
    cmdModificarCarrera.Enabled = False
    cmdGuardarCarrera.Enabled = True
    cmdBorrarCarrera.Enabled = False
    cmdCancelarCarrera.Enabled = True
    txtNombreCarrera.SetFocus
End Sub

Private Sub cmdAgregarCorrelativa_Click()
    frmCorrelativas.Show 1
End Sub

Private Sub cmdAgregarMateria_Click()
    EstadoMaterias = "Nuevo"
    frCarreras.Enabled = False
    cmdAgregarMateria.Enabled = False
    cmdModificarMateria.Enabled = False
    cmdGuardarMateria.Enabled = True
    cmdBorrarMateria.Enabled = False
    cmdCancelarMateria.Enabled = True
    cmdSalir.Enabled = False
    LimpiarMaterias
    For i = 0 To adoCarreras.Recordset!Años - 1
        frmAgregarMateria.cbAño.List(i) = i + 1
    Next i
    HabilitarMaterias
    frmAgregarMateria.Show 1
End Sub

Private Sub cmdAgregarTitulo_Click()
    PorDefecto = dtlTitulos.Text
    Respuesta = InputBox("Ingrese el titulo correspondiente a la carrera", "Agregar Título", PorDefecto)
    If Respuesta <> "" Then
        adoTitulos.Recordset.AddNew
        adoTitulos.Recordset!Carrera = dtcCarreras.BoundText
        adoTitulos.Recordset!Titulo = Respuesta
        adoTitulos.Recordset.Update
        adoTitulos.Recordset.Requery
        cmdQuitarTitulo.Enabled = True
    End If
End Sub

Private Sub cmdBorrarCarrera_Click()
    Respuesta = MsgBox("¿Está seguro de eliminar la carrera?" & Chr(13) & adoCarreras.Recordset!Abreviatura, vbYesNo, "Eliminar Carrera")
    If Respuesta = vbYes Then
        adoCarreras.Recordset!Eliminada = 1
        adoCarreras.Recordset.Update
        adoCarreras.Recordset.Requery
        LimpiarCarreras
    End If
End Sub

Private Sub cmdBorrarMateria_Click()
    If adoMaterias.Recordset.BOF = False Then
        Respuesta = MsgBox("¿Está seguro de eliminar la Materia?" & Chr(13) & adoMaterias.Recordset!Abreviatura, vbYesNo, "Eliminar Materia")
        If Respuesta = vbYes Then
            adoMaterias.Recordset!Eliminada = 1
            adoMaterias.Recordset.Requery
            MostrarMaterias
        End If
    End If
End Sub

Private Sub cmdCancelarCarrera_Click()
    cmdAgregarCarrera.Enabled = True
    cmdModificarCarrera.Enabled = True
    cmdGuardarCarrera.Enabled = False
    cmdBorrarCarrera.Enabled = True
    cmdCancelarCarrera.Enabled = False
    frMostrar.Enabled = True
    frMaterias.Enabled = True
    DeshabilitarCarreras
    dtcCarreras.Text = adoCarreras.Recordset!Nombre
    EstadoCarreras = ""
End Sub

Public Sub cmdCancelarMateria_Click()
    frCarreras.Enabled = True
    cmdAgregarMateria.Enabled = True
    cmdModificarMateria.Enabled = True
    cmdGuardarMateria.Enabled = False
    cmdBorrarMateria.Enabled = True
    cmdCancelarMateria.Enabled = False
    cmdSalir.Enabled = True
    DeshabilitarMaterias
    MostrarMaterias
End Sub

Private Sub cmdCompuestaPor_Click()
    With frmMateriasPromedio
        .lblNombreMateria = adoMaterias.Recordset!Codigo & " - " & txtNombreMateria
        .adoMateriasPromedio.RecordSource = "SELECT MateriasPromedio.Componente, Materias.Curso, Materias.Abreviatura, Detalles.Detalle FROM Detalles INNER JOIN (MateriasPromedio INNER JOIN Materias ON MateriasPromedio.Componente = Materias.Codigo) ON Detalles.Codigo = Materias.Detalle Where (((MateriasPromedio.Principal) = " & adoMaterias.Recordset!Codigo & "))ORDER BY Materias.Curso, MateriasPromedio.Componente"
        .adoMateriasPromedio.Refresh
        If .adoMateriasPromedio.Recordset.RecordCount > 0 Then .cmdEliminarComponente.Enabled = True
        .Show 1
    End With
End Sub

Private Sub cmdEliminarCorrelativa_Click()
    Respuesta = MsgBox("¿Está seguro de eliminar la correlativa?" & Chr(13) & adoCorrelativas.Recordset!Nombre, vbYesNo, "Eliminar Correlativa")
    If Respuesta = vbYes Then
        Conexion.Open
        
        'para SQL Server
        'Conexion.Execute ("DELETE Correlativas WHERE Correlativas.Principal = " & adoMaterias.Recordset!Codigo & " AND Correlativas.Correlativa = " & adoCorrelativas.Recordset!Correlativa)
        
        'para Acces
        Conexion.Execute ("DELETE * FROM Correlativas WHERE Correlativas.Principal = " & adoMaterias.Recordset!Codigo & " AND Correlativas.Correlativa = " & adoCorrelativas.Recordset!Correlativa)
        
        Conexion.Close
        adoCorrelativas.Refresh
        MostrarMaterias
    End If
End Sub

Private Sub cmdGuardarCarrera_Click()
    DeshabilitarCarreras
    PasarCarreras
    cmdAgregarCarrera.Enabled = True
    cmdModificarCarrera.Enabled = True
    cmdGuardarCarrera.Enabled = False
    cmdBorrarCarrera.Enabled = True
    cmdCancelarCarrera.Enabled = False
End Sub

Private Sub cmdGuardarMateria_Click()
    frCarreras.Enabled = True
    cmdAgregarMateria.Enabled = True
    cmdModificarMateria.Enabled = True
    cmdGuardarMateria.Enabled = False
    cmdBorrarMateria.Enabled = True
    cmdCancelarMateria.Enabled = False
    cmdSalir.Enabled = True
    PasarMaterias
    DeshabilitarMaterias
    MostrarMaterias
End Sub

Private Sub cmdModificarCarrera_Click()
    frMostrar.Enabled = False
    frMaterias.Enabled = False
    HabilitarCarreras
    cmdAgregarCarrera.Enabled = False
    cmdModificarCarrera.Enabled = False
    cmdGuardarCarrera.Enabled = True
    cmdBorrarCarrera.Enabled = False
    cmdCancelarCarrera.Enabled = True
End Sub


Private Sub cmdModificarMateria_Click()
    frCarreras.Enabled = False
    cmdAgregarMateria.Enabled = False
    cmdModificarMateria.Enabled = False
    cmdGuardarMateria.Enabled = True
    cmdBorrarMateria.Enabled = False
    cmdCancelarMateria.Enabled = True
    cmdSalir.Enabled = False
    HabilitarMaterias
    txtNombreMateria.SetFocus
End Sub

Private Sub cmdOrdenPlan_Click()
    frmOrdenPlan.Show 1
End Sub

Private Sub cmdQuitarTitulo_Click()
    If dtlTitulos.Text = "" Then MsgBox ("Debe elegir un titulo"): Exit Sub
    Respuesta = MsgBox("¿Está seguro de quitar el Título " & dtlTitulos.Text & "?", vbYesNo, "Quitar Título")
    If Respuesta = vbYes Then
        Conexion.Open
        
        'para SQL Server
        'Conexion.Execute ("DELETE  TitulosPosibles WHERE Carrera= " & dtcCarreras.BoundText & " AND Titulo = '" & dtlTitulos.Text & "'")
        
        'para Acces
        Conexion.Execute ("DELETE  * FROM TitulosPosibles WHERE Carrera= " & dtcCarreras.BoundText & " AND Titulo = '" & dtlTitulos.Text & "'")
        
        Conexion.Close
        adoTitulos.Refresh
        If adoTitulos.Recordset.RecordCount < 1 Then cmdQuitarTitulo.Enabled = False
    End If
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub dtcCarreras_Change()
    If dtcCarreras.BoundText <> "" Then
        adoCarreras.Recordset.MoveFirst
        adoCarreras.Recordset.Find ("Codigo=" & dtcCarreras.BoundText)
        MostrarCarreras
        adoTitulos.RecordSource = "SELECT * From TitulosPosibles Where TitulosPosibles.Carrera = " & dtcCarreras.BoundText & " ORDER BY Titulo"
        adoTitulos.Refresh
        If frmIdentificacion.Permisos!ModificarPlanes = True Then
            cmdAgregarTitulo.Enabled = True
        Else
            cmdAgregarTitulo.Enabled = False
        End If
        If adoTitulos.Recordset.RecordCount > 0 And frmIdentificacion.Permisos!ModificarPlanes = True Then
            cmdQuitarTitulo.Enabled = True
        Else
            cmdQuitarTitulo.Enabled = False
        End If
    End If
    If dtcCarreras.Text <> "" Then
        If frmIdentificacion.Permisos!ModificarPlanes = True Then
            cmdAgregarCarrera.Enabled = True
            cmdBorrarCarrera.Enabled = True
            cmdModificarCarrera.Enabled = True
        Else
            cmdAgregarCarrera.Enabled = False
            cmdBorrarCarrera.Enabled = False
            cmdModificarCarrera.Enabled = False
        End If
        adoMaterias.RecordSource = "SELECT * FROM Materias WHERE Carrera = " & dtcCarreras.BoundText & " AND Eliminada = 0 ORDER BY OrdenPlan,Codigo"
        adoMaterias.Refresh
        MostrarMaterias
        frMaterias.Enabled = True
    Else
        cmdBorrarCarrera.Enabled = False
        cmdModificarCarrera.Enabled = False
        adoMaterias.RecordSource = "SELECT * FROM Materias WHERE Carrera =0"
        adoMaterias.Refresh
        frMaterias.Enabled = False
    End If
End Sub

Private Sub dtgMaterias_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    If EstadoMaterias <> "Nuevo" Then MostrarMaterias
End Sub

Private Sub Form_Load()

    Conexion.ConnectionString = ("DSN=Instituto")
On Error GoTo hErr
   Conexion.Open
   Conexion.Execute ("ALTER TABLE  Materias ADD OrdenPlan int")
On Error GoTo hErr2
   Conexion.Execute ("ALTER TABLE  Materias ADD CursoAnalitico int")
On Error GoTo hErr3
   Conexion.Execute ("Update Materias set OrdenPlan=0 where OrdenPlan is Null")
On Error GoTo hErr4
   Conexion.Execute ("Update Materias set CursoAnalitico=Curso where CursoAnalitico is Null")
   Conexion.Close
   
    If adoCarreras.Recordset.RecordCount > 0 Then
        dtcCarreras.BoundText = adoCarreras.Recordset!Codigo
    End If
   
   Exit Sub
hErr:
   'MsgBox Err.Number & " " & Err.Description
 '  Conexion.Close
   Resume Next
hErr2:
   Resume Next
hErr3:
   Resume Next
hErr4:
   Resume Next
   
End Sub

Private Sub optMostar_Click(Index As Integer)
    If Index = 0 Then
        adoCarreras.RecordSource = "SELECT * FROM Carreras WHERE Eliminada = 0 AND Vigente = 1 ORDER BY Nombre"
    Else
        adoCarreras.RecordSource = "SELECT * FROM Carreras WHERE Eliminada = 0 ORDER BY Nombre"
    End If
    adoCarreras.Refresh
    dtcCarreras.Text = ""
    LimpiarCarreras
End Sub

Private Function MostrarCarreras()
    txtNombreCarrera = adoCarreras.Recordset!Nombre
    txtCodigoCarrera = adoCarreras.Recordset!Codigo
    If adoCarreras.Recordset!Resolucion <> vacio Then
        txtResolucion = adoCarreras.Recordset!Resolucion
    Else
        txtResolucion = ""
    End If
    txtAños = adoCarreras.Recordset!Años
    dtcModalidadCarrera.BoundText = adoCarreras.Recordset!Modalidad
    If adoCarreras.Recordset!Abreviatura <> "" Then
        txtAbrebiatura1Carrera = adoCarreras.Recordset!Abreviatura
    Else
        txtAbrebiatura1Carrera = ""
    End If
    If adoCarreras.Recordset!AbreviaturaPof <> "" Then
        txtAbrebiatura2Carrera = adoCarreras.Recordset!AbreviaturaPof
    Else
        txtAbrebiatura2Carrera = ""
    End If
    If adoCarreras.Recordset!Vigente = "Verdadero" Then
        chkVigente = 1
    Else
        chkVigente = 0
    End If
    dtcCaracteristica.BoundText = adoCarreras.Recordset!Caracteristica
End Function

Private Function LimpiarCarreras()
    dtcCarreras.Text = ""
    txtNombreCarrera = ""
    txtCodigoCarrera = ""
    txtResolucion = ""
    txtAños = ""
    txtAbrebiatura1Carrera = ""
    txtAbrebiatura2Carrera = ""
    chkVigente = 1
    adoTitulos.RecordSource = "SELECT * From TitulosPosibles Where TitulosPosibles.Carrera = 0"
    adoTitulos.Refresh
End Function

Private Function HabilitarCarreras()
    txtResolucion.Enabled = True
    txtAños.Enabled = True
    dtcModalidadCarrera.Enabled = True
    txtAbrebiatura1Carrera.Enabled = True
    txtAbrebiatura2Carrera.Enabled = True
    chkVigente.Enabled = True
    txtNombreCarrera.Visible = True
    dtcCaracteristica.Enabled = True
End Function

Private Function DeshabilitarCarreras()
    txtNombreCarrera.Visible = False
    txtResolucion.Enabled = False
    txtAños.Enabled = False
    dtcModalidadCarrera.Enabled = False
    txtAbrebiatura1Carrera.Enabled = False
    txtAbrebiatura2Carrera.Enabled = False
    chkVigente.Enabled = False
    dtcCaracteristica.Enabled = False
End Function

Private Function PasarCarreras()
    If EstadoCarreras = "Nuevo" Then
        adoCarreras.Recordset.AddNew
        EstadoCarreras = ""
    End If
    adoCarreras.Recordset!Nombre = txtNombreCarrera
    adoCarreras.Recordset!Codigo = txtCodigoCarrera
    adoCarreras.Recordset!Resolucion = txtResolucion
    adoCarreras.Recordset!Años = Val(txtAños)
    adoCarreras.Recordset!Modalidad = dtcModalidadCarrera.BoundText
    adoCarreras.Recordset!Abreviatura = txtAbrebiatura1Carrera
    adoCarreras.Recordset!AbreviaturaPof = txtAbrebiatura2Carrera
    adoCarreras.Recordset!Vigente = chkVigente
    adoCarreras.Recordset!Caracteristica = dtcCaracteristica.BoundText
    adoCarreras.Recordset.Update
    dtcCarreras.Text = adoCarreras.Recordset!Nombre
    frMostrar.Enabled = True
    frMaterias.Enabled = True
End Function
Public Function MostrarMaterias()
    If adoMaterias.Recordset.BOF = False Then
        txtCursoMateria = adoMaterias.Recordset!Curso
        txtCodigoBloqueadoMateria = Mid(adoMaterias.Recordset!Codigo, 1, Len(adoMaterias.Recordset!Codigo) - 2)
        txtCodigoDesbloqueadoMateria = Mid(adoMaterias.Recordset!Codigo, Len(adoMaterias.Recordset!Codigo) - 1, 2)
        txtCursoAnalitico = adoMaterias.Recordset!CursoAnalitico
        txtNombreMateria = adoMaterias.Recordset!Nombre
        txtAbreviatura1Materia = adoMaterias.Recordset!Abreviatura
        If adoMaterias.Recordset!AbreviaturaPof <> "" Then
            txtAbreviatura2Materia = adoMaterias.Recordset!AbreviaturaPof
        Else
            txtAbreviatura2Materia = ""
        End If
        dtcModalidadMateria.BoundText = adoMaterias.Recordset!Modalidad
        dtcDetalleMateria.BoundText = adoMaterias.Recordset!Detalle
        If frmIdentificacion.Permisos!ModificarPlanes = True Then
            cmdAgregarMateria.Enabled = True
            cmdModificarMateria.Enabled = True
            cmdBorrarMateria.Enabled = True
        Else
            cmdAgregarMateria.Enabled = False
            cmdModificarMateria.Enabled = False
            cmdBorrarMateria.Enabled = False
        End If
        adoCorrelativas.RecordSource = "SELECT Correlativas.Correlativa, Materias.Nombre, Materias.Curso, Correlativas.PorFinal FROM Correlativas INNER JOIN Materias ON Correlativas.Correlativa = Materias.Codigo WHERE Correlativas.Principal = " & adoMaterias.Recordset!Codigo & " ORDER BY Correlativa"
        adoCorrelativas.Refresh
        If frmIdentificacion.Permisos!ModificarPlanes = True Then
            cmdAgregarCorrelativa.Enabled = True
            If adoCorrelativas.Recordset.RecordCount = 0 Then
                cmdEliminarCorrelativa.Enabled = False
            Else
                cmdEliminarCorrelativa.Enabled = True
            End If
        Else
            cmdAgregarCorrelativa.Enabled = False
            cmdEliminarCorrelativa.Enabled = False
        End If
        If dtcDetalleMateria.BoundText = 2 Then
        
            cmdCompuestaPor.Enabled = True
        Else
            cmdCompuestaPor.Enabled = False
        End If
    Else
        cmdAgregarCorrelativa.Enabled = False
        cmdEliminarCorrelativa.Enabled = False
        cmdBorrarMateria.Enabled = False
        LimpiarMaterias
    End If
End Function
Private Function HabilitarMaterias()
    'txtCodigoDesbloqueadoMateria.Enabled = True
    txtNombreMateria.Enabled = True
    txtAbreviatura1Materia.Enabled = True
    txtAbreviatura2Materia.Enabled = True
    txtCursoAnalitico.Enabled = True
    dtcModalidadMateria.Enabled = True
    dtcDetalleMateria.Enabled = True
    frCarreras.Enabled = False
    dtgMaterias.Enabled = False
    dtgCorrelativas.Enabled = False
    cmdAgregarCorrelativa.Enabled = False
    cmdEliminarCorrelativa.Enabled = False
End Function

Public Function DeshabilitarMaterias()
    txtCodigoDesbloqueadoMateria.Enabled = False
    txtNombreMateria.Enabled = False
    txtAbreviatura1Materia.Enabled = False
    txtAbreviatura2Materia.Enabled = False
    txtCursoAnalitico.Enabled = False
    dtcModalidadMateria.Enabled = False
    dtcDetalleMateria.Enabled = False
    frCarreras.Enabled = True
    dtgMaterias.Enabled = True
    dtgCorrelativas.Enabled = True
    cmdAgregarCorrelativa.Enabled = True
    cmdEliminarCorrelativa.Enabled = True
End Function
Private Function PasarMaterias()
    If EstadoMaterias = "Nuevo" Then
        adoMaterias.Recordset.AddNew
        EstadoMaterias = ""
        adoMaterias.Recordset!Codigo = txtCodigoBloqueadoMateria & txtCodigoDesbloqueadoMateria
        adoMaterias.Recordset!Carrera = adoCarreras.Recordset!Codigo
    End If
    MateriaActual = adoMaterias.Recordset!Codigo
    adoMaterias.Recordset!Curso = txtCursoMateria
    adoMaterias.Recordset!Nombre = txtNombreMateria
    adoMaterias.Recordset!Abreviatura = txtAbreviatura1Materia
    adoMaterias.Recordset!AbreviaturaPof = txtAbreviatura2Materia
    adoMaterias.Recordset!CursoAnalitico = txtCursoAnalitico
    adoMaterias.Recordset!Modalidad = dtcModalidadMateria.BoundText
    adoMaterias.Recordset!Detalle = dtcDetalleMateria.BoundText
    adoMaterias.Recordset.Update
    adoMaterias.RecordSource = "SELECT * FROM Materias WHERE Carrera = " & dtcCarreras.BoundText & " AND Eliminada = 0 ORDER BY Codigo"
    adoMaterias.Recordset.Requery
    adoMaterias.Recordset.Find ("Codigo = " & MateriaActual)
    cmdBorrarMateria.Enabled = False
End Function
Private Function LimpiarMaterias()
    txtCursoMateria = ""
    txtCodigoBloqueadoMateria = ""
    txtCodigoDesbloqueadoMateria = ""
    txtNombreMateria = ""
    txtAbreviatura1Materia = ""
    txtAbreviatura2Materia = ""
    txtCursoAnalitico = ""
    dtcModalidadMateria.BoundText = 1
    dtcDetalleMateria.BoundText = 1
End Function
Private Sub txtAbrebiatura1Carrera_KeyPress(KeyAscii As Integer)
    If KeyAscii >= 97 And KeyAscii <= 122 Then KeyAscii = KeyAscii - 32
End Sub

Private Sub txtAbrebiatura2Carrera_KeyPress(KeyAscii As Integer)
    If KeyAscii >= 97 And KeyAscii <= 122 Then KeyAscii = KeyAscii - 32
End Sub

Private Sub txtAbreviatura1Materia_KeyPress(KeyAscii As Integer)
    If KeyAscii >= 97 And KeyAscii <= 122 Then KeyAscii = KeyAscii - 32
End Sub

Private Sub txtAbreviatura2Materia_KeyPress(KeyAscii As Integer)
    If KeyAscii >= 97 And KeyAscii <= 122 Then KeyAscii = KeyAscii - 32
End Sub

Private Sub txtNombreCarrera_KeyPress(KeyAscii As Integer)
    'If KeyAscii >= 97 And KeyAscii <= 122 Then KeyAscii = KeyAscii - 32
End Sub

Private Sub txtNombreMateria_KeyPress(KeyAscii As Integer)
    'If KeyAscii >= 97 And KeyAscii <= 122 Then KeyAscii = KeyAscii - 32
End Sub
