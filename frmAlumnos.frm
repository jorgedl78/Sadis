VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmAlumnos 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Alumnos"
   ClientHeight    =   10620
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   14910
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   10620
   ScaleWidth      =   14910
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   WhatsThisHelp   =   -1  'True
   Begin VB.Frame frDocumentacion 
      Caption         =   "Documentación"
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
      Height          =   1695
      Left            =   0
      TabIndex        =   77
      Top             =   2880
      Width           =   2055
      Begin VB.CheckBox chkBloquearAutogestion 
         Caption         =   "Bloquear Autogestión"
         Height          =   495
         Left            =   240
         TabIndex        =   81
         Top             =   1080
         Width           =   1575
      End
      Begin VB.CheckBox chkTitulo 
         Caption         =   "Titulo Secundario"
         Height          =   255
         Left            =   240
         TabIndex        =   80
         Top             =   720
         Width           =   1575
      End
      Begin VB.CheckBox chkDocumento 
         Caption         =   "Documento"
         Height          =   255
         Left            =   240
         TabIndex        =   79
         Top             =   480
         Width           =   1455
      End
      Begin VB.CheckBox chkFotos 
         Caption         =   "Fotos"
         Height          =   255
         Left            =   240
         TabIndex        =   78
         Top             =   240
         Width           =   1335
      End
   End
   Begin MSComCtl2.MonthView clFechaCondicion 
      Height          =   2370
      Left            =   6360
      TabIndex        =   76
      Top             =   3000
      Visible         =   0   'False
      Width           =   2595
      _ExtentX        =   4577
      _ExtentY        =   4180
      _Version        =   393216
      ForeColor       =   -2147483630
      BackColor       =   -2147483633
      Appearance      =   1
      StartOfWeek     =   104792065
      CurrentDate     =   37725
   End
   Begin MSComCtl2.MonthView clFechaNacimiento 
      Height          =   2370
      Left            =   5160
      TabIndex        =   75
      Top             =   960
      Visible         =   0   'False
      Width           =   2595
      _ExtentX        =   4577
      _ExtentY        =   4180
      _Version        =   393216
      ForeColor       =   -2147483630
      BackColor       =   -2147483633
      Appearance      =   1
      StartOfWeek     =   128974849
      CurrentDate     =   37725
   End
   Begin VB.Frame frComandos 
      Height          =   4575
      Left            =   13920
      TabIndex        =   3
      Top             =   0
      Width           =   855
      Begin VB.CommandButton cmdCancelar 
         Enabled         =   0   'False
         Height          =   600
         Left            =   120
         Picture         =   "frmAlumnos.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Cancelar"
         Top             =   2760
         Width           =   600
      End
      Begin VB.CommandButton cmdGuardar 
         Enabled         =   0   'False
         Height          =   600
         Left            =   120
         Picture         =   "frmAlumnos.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Guardar"
         Top             =   1920
         Width           =   600
      End
      Begin VB.CommandButton cmdModificar 
         Enabled         =   0   'False
         Height          =   600
         Left            =   120
         Picture         =   "frmAlumnos.frx":0884
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Modificar"
         Top             =   1080
         Width           =   600
      End
      Begin VB.CommandButton cmdAgregar 
         Height          =   600
         Left            =   120
         Picture         =   "frmAlumnos.frx":0CC6
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Agregar"
         Top             =   240
         Width           =   600
      End
      Begin VB.CommandButton cmdSalir 
         Height          =   600
         Left            =   120
         Picture         =   "frmAlumnos.frx":1108
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Salir"
         Top             =   3720
         Width           =   600
      End
   End
   Begin VB.Frame frOpcionesDeBusqueda 
      Caption         =   "Opciones de Búsqueda"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5895
      Left            =   0
      TabIndex        =   2
      Tag             =   "Imprimir Inscripción"
      Top             =   4680
      Width           =   14775
      Begin MSAdodcLib.Adodc adoAlumnos 
         Height          =   330
         Left            =   2160
         Top             =   2040
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
         RecordSource    =   "SELECT * FROM Alumnos WHERE Eliminado = 0 ORDER BY Nombre"
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
      Begin MSDataGridLib.DataGrid dtgAlumnos 
         Bindings        =   "frmAlumnos.frx":154A
         Height          =   4455
         Left            =   120
         TabIndex        =   54
         Top             =   1320
         Width           =   14535
         _ExtentX        =   25638
         _ExtentY        =   7858
         _Version        =   393216
         AllowUpdate     =   0   'False
         HeadLines       =   1
         RowHeight       =   15
         RowDividerStyle =   1
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
            DataField       =   "Nombre"
            Caption         =   "Apellido y Nombre"
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
            DataField       =   "Tipo"
            Caption         =   "Tipo"
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
         BeginProperty Column04 
            DataField       =   "Domicilio"
            Caption         =   "Domicilio Origen"
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
         BeginProperty Column05 
            DataField       =   "Localidad"
            Caption         =   "Localidad"
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
         BeginProperty Column06 
            DataField       =   "DomicilioEnJunin"
            Caption         =   "Domicilio Local"
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
            Locked          =   -1  'True
            BeginProperty Column00 
               Locked          =   -1  'True
               ColumnWidth     =   734,74
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   4470,236
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   540,284
            EndProperty
            BeginProperty Column03 
            EndProperty
            BeginProperty Column04 
               ColumnWidth     =   1890,142
            EndProperty
            BeginProperty Column05 
               ColumnWidth     =   2489,953
            EndProperty
            BeginProperty Column06 
               ColumnWidth     =   2190,047
            EndProperty
         EndProperty
      End
      Begin VB.Frame frBuscarPor 
         Caption         =   "BuscarPor:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   120
         TabIndex        =   42
         Top             =   240
         Width           =   11655
         Begin MSDataListLib.DataCombo dtcCarrerasVigentes 
            Bindings        =   "frmAlumnos.frx":1563
            Height          =   315
            Left            =   6000
            TabIndex        =   52
            Top             =   480
            Width           =   3735
            _ExtentX        =   6588
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            Style           =   2
            ListField       =   "Abreviatura"
            BoundColumn     =   "Codigo"
            Text            =   "Carrera"
         End
         Begin VB.TextBox txtBuscarPorDomicilio 
            Enabled         =   0   'False
            Height          =   285
            Left            =   4200
            MaxLength       =   20
            TabIndex        =   51
            Top             =   480
            Width           =   1575
         End
         Begin VB.TextBox txtBuscarPorDocumento 
            Enabled         =   0   'False
            Height          =   285
            Left            =   3000
            MaxLength       =   8
            TabIndex        =   50
            Top             =   480
            Width           =   975
         End
         Begin VB.TextBox txtBuscarPorPermiso 
            Enabled         =   0   'False
            Height          =   285
            Left            =   2040
            MaxLength       =   6
            TabIndex        =   49
            Top             =   480
            Width           =   735
         End
         Begin VB.TextBox txtBuscarPorNombre 
            Enabled         =   0   'False
            Height          =   285
            Left            =   240
            MaxLength       =   20
            TabIndex        =   48
            Top             =   480
            Width           =   1575
         End
         Begin VB.OptionButton optBuscarPor 
            Caption         =   "Domicilio"
            Height          =   255
            Index           =   3
            Left            =   4200
            TabIndex        =   47
            Top             =   240
            Width           =   1095
         End
         Begin VB.OptionButton optBuscarPor 
            Caption         =   "Carreras Vigentes                                            y       Situación"
            Height          =   255
            Index           =   4
            Left            =   6240
            TabIndex        =   46
            Top             =   240
            Width           =   5175
         End
         Begin VB.OptionButton optBuscarPor 
            Caption         =   "Nº Doc."
            Height          =   255
            Index           =   2
            Left            =   3000
            TabIndex        =   45
            Top             =   240
            Width           =   1455
         End
         Begin VB.OptionButton optBuscarPor 
            Caption         =   "Permiso"
            Height          =   255
            Index           =   1
            Left            =   1920
            TabIndex        =   44
            Top             =   240
            Width           =   975
         End
         Begin VB.OptionButton optBuscarPor 
            Caption         =   "Apellido y Nombre"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   43
            Top             =   240
            Width           =   1695
         End
         Begin MSDataListLib.DataCombo dtcSituación 
            Bindings        =   "frmAlumnos.frx":1585
            Height          =   315
            Left            =   9840
            TabIndex        =   53
            Top             =   480
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            Style           =   2
            ListField       =   "Condicion"
            BoundColumn     =   "Codigo"
            Text            =   "Sitiación"
         End
         Begin MSAdodcLib.Adodc adoCarrerasVigentes 
            Height          =   330
            Left            =   6600
            Top             =   720
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
            RecordSource    =   $"frmAlumnos.frx":15A9
            Caption         =   "CarrerasVigentes"
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
         Begin MSAdodcLib.Adodc adoBuscarPorSituacion 
            Height          =   330
            Left            =   9360
            Top             =   720
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
            RecordSource    =   "SELECT * FROM Condicion ORDER BY Codigo"
            Caption         =   "BuscarPorSituacion"
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
      End
   End
   Begin VB.Frame frDatosAcademicos 
      Caption         =   "Datos Académicos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   2160
      TabIndex        =   1
      Top             =   2880
      Width           =   11655
      Begin VB.CommandButton cmdAgregarCarrera 
         Height          =   600
         Left            =   10080
         Picture         =   "frmAlumnos.frx":1658
         Style           =   1  'Graphical
         TabIndex        =   72
         ToolTipText     =   "Agregar"
         Top             =   240
         Width           =   600
      End
      Begin VB.CommandButton cmdModificarCarrera 
         Enabled         =   0   'False
         Height          =   600
         Left            =   10920
         Picture         =   "frmAlumnos.frx":1A9A
         Style           =   1  'Graphical
         TabIndex        =   71
         ToolTipText     =   "Modificar"
         Top             =   240
         Width           =   600
      End
      Begin VB.CommandButton cmdGuardarCarrera 
         Enabled         =   0   'False
         Height          =   600
         Left            =   10080
         Picture         =   "frmAlumnos.frx":1EDC
         Style           =   1  'Graphical
         TabIndex        =   70
         ToolTipText     =   "Guardar"
         Top             =   960
         Width           =   600
      End
      Begin VB.CommandButton cmdCancelarCarrera 
         Enabled         =   0   'False
         Height          =   600
         Left            =   10920
         Picture         =   "frmAlumnos.frx":231E
         Style           =   1  'Graphical
         TabIndex        =   69
         ToolTipText     =   "Cancelar"
         Top             =   960
         Width           =   600
      End
      Begin VB.TextBox txtFolio 
         Enabled         =   0   'False
         Height          =   285
         Left            =   7320
         MaxLength       =   4
         TabIndex        =   67
         Top             =   1200
         Width           =   615
      End
      Begin VB.TextBox txtLibro 
         Enabled         =   0   'False
         Height          =   285
         Left            =   6480
         MaxLength       =   4
         TabIndex        =   65
         Top             =   1200
         Width           =   495
      End
      Begin VB.TextBox txtFechaCondicion 
         Enabled         =   0   'False
         Height          =   285
         Left            =   5280
         MaxLength       =   10
         TabIndex        =   63
         Top             =   1200
         Width           =   975
      End
      Begin VB.TextBox txtDivision 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1440
         MaxLength       =   4
         TabIndex        =   60
         Top             =   1200
         Width           =   495
      End
      Begin VB.TextBox txtAñoIngreso 
         Enabled         =   0   'False
         Height          =   285
         Left            =   240
         MaxLength       =   4
         TabIndex        =   58
         Top             =   1200
         Width           =   735
      End
      Begin MSDataListLib.DataCombo dtcCarrerasQueCursa 
         Bindings        =   "frmAlumnos.frx":2760
         Height          =   315
         Left            =   120
         TabIndex        =   56
         Top             =   480
         Width           =   9735
         _ExtentX        =   17171
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "Abreviatura"
         BoundColumn     =   "Codigo"
         Text            =   "Carrera"
      End
      Begin MSAdodcLib.Adodc adoCarrerasQueCursa 
         Height          =   330
         Left            =   1680
         Top             =   240
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
         RecordSource    =   $"frmAlumnos.frx":2782
         Caption         =   "CarrerasQueCursa"
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
      Begin MSDataListLib.DataCombo dtcCondicion 
         Bindings        =   "frmAlumnos.frx":28BA
         Height          =   315
         Left            =   2280
         TabIndex        =   61
         Top             =   1200
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         Style           =   2
         ListField       =   "Condicion"
         BoundColumn     =   "Codigo"
         Text            =   "Condicion"
      End
      Begin VB.Label Label24 
         AutoSize        =   -1  'True
         Caption         =   "Folio:"
         Height          =   195
         Left            =   7320
         TabIndex        =   68
         Top             =   960
         Width           =   375
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         Caption         =   "Libro:"
         Height          =   195
         Left            =   6480
         TabIndex        =   66
         Top             =   960
         Width           =   390
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         Caption         =   "Fecha"
         Height          =   195
         Left            =   5280
         TabIndex        =   64
         Top             =   960
         Width           =   450
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         Caption         =   "Condición:"
         Height          =   195
         Left            =   2280
         TabIndex        =   62
         Top             =   960
         Width           =   750
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         Caption         =   "División:"
         Height          =   195
         Left            =   1440
         TabIndex        =   59
         Top             =   960
         Width           =   600
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         Caption         =   "Año Ingreso:"
         Height          =   195
         Left            =   240
         TabIndex        =   57
         Top             =   960
         Width           =   900
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         Caption         =   "Carreras que cursa:"
         Height          =   195
         Left            =   120
         TabIndex        =   55
         Top             =   240
         Width           =   1380
      End
   End
   Begin VB.Frame frDatosPersonales 
      Caption         =   "Datos Personales"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2775
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   13815
      Begin VB.CommandButton cmdNotificacionPersonal 
         Caption         =   "Notificación Personal"
         Height          =   960
         Left            =   11520
         Picture         =   "frmAlumnos.frx":28DE
         Style           =   1  'Graphical
         TabIndex        =   83
         ToolTipText     =   "Modificar"
         Top             =   1440
         Width           =   1680
      End
      Begin VB.CommandButton cmdImprimir 
         Caption         =   "Imprimir Inscripción"
         Height          =   735
         Left            =   11520
         Picture         =   "frmAlumnos.frx":2D20
         Style           =   1  'Graphical
         TabIndex        =   82
         ToolTipText     =   "Imprimir Acta"
         Top             =   480
         Width           =   1755
      End
      Begin VB.TextBox txtNombre 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1080
         MaxLength       =   50
         TabIndex        =   74
         Top             =   480
         Width           =   4575
      End
      Begin VB.TextBox txtFechaNacimiento 
         Enabled         =   0   'False
         Height          =   285
         Left            =   5640
         TabIndex        =   73
         Top             =   1680
         Width           =   1095
      End
      Begin VB.CommandButton cmdRestaurarPSW 
         Caption         =   "Restaurar PSW"
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
         Left            =   7440
         TabIndex        =   41
         Tag             =   "Restaurar Contraseña"
         Top             =   2160
         Width           =   1695
      End
      Begin VB.TextBox txtComentario 
         Enabled         =   0   'False
         Height          =   2085
         Left            =   9360
         MaxLength       =   100
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   39
         Top             =   480
         Width           =   1455
      End
      Begin VB.TextBox txtRegistro 
         Enabled         =   0   'False
         Height          =   285
         Left            =   6000
         MaxLength       =   12
         TabIndex        =   37
         Top             =   2280
         Width           =   1215
      End
      Begin VB.TextBox txtEstablecimiento 
         Enabled         =   0   'False
         Height          =   285
         Left            =   3120
         MaxLength       =   30
         TabIndex        =   35
         Top             =   2280
         Width           =   2775
      End
      Begin VB.TextBox txtTitulo 
         Enabled         =   0   'False
         Height          =   285
         Left            =   120
         MaxLength       =   30
         TabIndex        =   33
         Top             =   2280
         Width           =   2895
      End
      Begin VB.ComboBox cbSexo 
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "frmAlumnos.frx":338A
         Left            =   7800
         List            =   "frmAlumnos.frx":3394
         Style           =   2  'Dropdown List
         TabIndex        =   31
         Top             =   480
         Width           =   735
      End
      Begin VB.TextBox txtLugar 
         Enabled         =   0   'False
         Height          =   285
         Left            =   7080
         MaxLength       =   20
         TabIndex        =   30
         Top             =   1680
         Width           =   2055
      End
      Begin VB.TextBox txtCorreo 
         Enabled         =   0   'False
         Height          =   285
         Left            =   2400
         MaxLength       =   50
         TabIndex        =   27
         Top             =   1680
         Width           =   3135
      End
      Begin VB.TextBox txtTelefono 
         Enabled         =   0   'False
         Height          =   285
         Left            =   120
         MaxLength       =   20
         TabIndex        =   25
         Top             =   1680
         Width           =   2055
      End
      Begin VB.TextBox txtDomicilioEnJunin 
         Enabled         =   0   'False
         Height          =   285
         Left            =   6360
         MaxLength       =   30
         TabIndex        =   23
         Top             =   1080
         Width           =   2895
      End
      Begin VB.TextBox txtPostal 
         Enabled         =   0   'False
         Height          =   285
         Left            =   5160
         MaxLength       =   10
         TabIndex        =   21
         Top             =   1080
         Width           =   1095
      End
      Begin VB.TextBox txtLocalidad 
         Enabled         =   0   'False
         Height          =   285
         Left            =   3120
         MaxLength       =   20
         TabIndex        =   19
         Top             =   1080
         Width           =   1935
      End
      Begin VB.TextBox txtDomicilio 
         Enabled         =   0   'False
         Height          =   285
         Left            =   120
         MaxLength       =   30
         TabIndex        =   17
         Top             =   1080
         Width           =   2895
      End
      Begin VB.ComboBox cbTipo 
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "frmAlumnos.frx":339E
         Left            =   5880
         List            =   "frmAlumnos.frx":33AB
         Style           =   2  'Dropdown List
         TabIndex        =   15
         Top             =   480
         Width           =   735
      End
      Begin VB.TextBox txtDocumento 
         Enabled         =   0   'False
         Height          =   285
         Left            =   6600
         MaxLength       =   8
         TabIndex        =   13
         Top             =   480
         Width           =   855
      End
      Begin VB.TextBox txtPermiso 
         Enabled         =   0   'False
         Height          =   285
         Left            =   120
         TabIndex        =   10
         Top             =   480
         Width           =   855
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "Comentario:"
         Height          =   195
         Left            =   9720
         TabIndex        =   40
         Top             =   240
         Width           =   840
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "Nº Registro:"
         Height          =   195
         Left            =   6000
         TabIndex        =   38
         Top             =   2040
         Width           =   855
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "Establecimiento que lo otorgó:"
         Height          =   195
         Left            =   3120
         TabIndex        =   36
         Top             =   2040
         Width           =   2130
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "Título Secundario o Polimodal:"
         Height          =   195
         Left            =   120
         TabIndex        =   34
         Top             =   2040
         Width           =   2175
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "Sexo:"
         Height          =   195
         Left            =   7800
         TabIndex        =   32
         Top             =   240
         Width           =   405
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Lugar de Nacimiento:"
         Height          =   195
         Left            =   7080
         TabIndex        =   29
         Top             =   1440
         Width           =   1515
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Nacimiento:"
         Height          =   195
         Left            =   5640
         TabIndex        =   28
         Top             =   1440
         Width           =   1335
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Correo Electrónico:"
         Height          =   195
         Left            =   2400
         TabIndex        =   26
         Top             =   1440
         Width           =   1350
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Teléfonos:"
         Height          =   195
         Left            =   120
         TabIndex        =   24
         Top             =   1440
         Width           =   750
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Domicilio Local:"
         Height          =   195
         Left            =   6360
         TabIndex        =   22
         Top             =   840
         Width           =   1110
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Cod. Postal:"
         Height          =   195
         Left            =   5160
         TabIndex        =   20
         Top             =   840
         Width           =   855
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Localidad"
         Height          =   195
         Left            =   3120
         TabIndex        =   18
         Top             =   840
         Width           =   690
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Domicilio ciudad de origen"
         Height          =   195
         Left            =   120
         TabIndex        =   16
         Top             =   840
         Width           =   1860
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Tipo:"
         Height          =   195
         Left            =   5880
         TabIndex        =   14
         Top             =   240
         Width           =   360
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Nº Documento"
         Height          =   195
         Left            =   6480
         TabIndex        =   12
         Top             =   240
         Width           =   1050
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Apellido y Nombres:"
         Height          =   195
         Left            =   1080
         TabIndex        =   11
         Top             =   240
         Width           =   1395
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Permiso:"
         Height          =   195
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   600
      End
   End
End
Attribute VB_Name = "frmAlumnos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Conexion As New Connection
Dim Resultado As New Recordset
Dim Estado As String
Dim EstadoCarreras As String

Private Sub cbSexo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtDomicilio.SetFocus
End Sub

Private Sub cbTipo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtDocumento.SetFocus
End Sub

Private Sub clFechaCondicion_DateClick(ByVal DateClicked As Date)
    frDatosAcademicos.Enabled = True
    txtFechaCondicion = clFechaCondicion.Value
    clFechaCondicion.Visible = False
End Sub

Private Sub clFechaNacimiento_DateClick(ByVal DateClicked As Date)
    txtFechaNacimiento = clFechaNacimiento.Value
    frDatosPersonales.Enabled = True
    frComandos.Enabled = True
    frDatosAcademicos.Enabled = True
    dtgAlumnos.Enabled = True
    clFechaNacimiento.Visible = False
    txtLugar.SetFocus
End Sub

Private Sub clFechaNacimiento_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtFechaNacimiento = clFechaNacimiento.Value
        frDatosPersonales.Enabled = True
        frComandos.Enabled = True
        frDatosAcademicos.Enabled = True
        dtgAlumnos.Enabled = True
        clFechaNacimiento.Visible = False
        txtLugar.SetFocus
    End If
End Sub

Private Sub cmdAgregar_Click()
    cmdAgregar.Enabled = False
    cmdModificar.Enabled = False
    cmdGuardar.Enabled = True
    cmdCancelar.Enabled = True
    cmdSalir.Enabled = False
    cmdImprimir.Enabled = False
    cmdRestaurarPSW.Enabled = False
    frDatosAcademicos.Enabled = False
    frOpcionesDeBusqueda.Enabled = False
    LimpiarDatos
    HabilitarDatos
    Estado = "Agregando"
    txtNombre.SetFocus
End Sub

Private Sub cmdAgregarCarrera_Click()
    EstadoCarreras = "Agregando"
    cmdAgregarCarrera.Enabled = False
    cmdModificarCarrera.Enabled = False
    cmdCancelarCarrera.Enabled = True
    cmdGuardarCarrera.Enabled = True
    frDatosPersonales.Enabled = False
    frComandos.Enabled = False
    frOpcionesDeBusqueda.Enabled = False
    adoCarrerasQueCursa.RecordSource = "SELECT Codigo, Abreviatura FROM Carreras Where Vigente = 1 ORDER BY Abreviatura"
    adoCarrerasQueCursa.Refresh
    dtcCarrerasQueCursa.BoundText = ""
    'dtcCarrerasQueCursa.BoundText = adoCarrerasQueCursa.Recordset!Codigo
    txtAñoIngreso.Enabled = True
    txtDivision.Enabled = True
    dtcCondicion.Enabled = True
    txtFechaCondicion.Enabled = True
    txtLibro.Enabled = True
    txtFolio.Enabled = True
    txtAñoIngreso = Format(Date, "yyyy")
    txtDivision = "1"
    txtFechaCondicion = Date
    txtLibro = ""
    txtFolio = ""
    dtcCondicion.BoundText = 1 'cursando
End Sub

Private Sub cmdCancelar_Click()
    cmdAgregar.Enabled = True
    cmdModificar.Enabled = True
    cmdGuardar.Enabled = False
    cmdCancelar.Enabled = False
    cmdSalir.Enabled = True
    cmdImprimir.Enabled = True
    cmdRestaurarPSW.Enabled = True
    frDatosAcademicos.Enabled = True
    frOpcionesDeBusqueda.Enabled = True
    DeshabilitarDatos
    MostrarDatos
End Sub

Private Sub cmdCancelarCarrera_Click()
    cmdAgregarCarrera.Enabled = True
    cmdModificarCarrera.Enabled = True
    cmdCancelarCarrera.Enabled = False
    cmdGuardarCarrera.Enabled = False
    frDatosPersonales.Enabled = True
    frComandos.Enabled = True
    frOpcionesDeBusqueda.Enabled = True
    dtcCarrerasQueCursa.Enabled = True
    EstadoCarreras = ""
    txtAñoIngreso.Enabled = False
    txtDivision.Enabled = False
    dtcCondicion.Enabled = False
    txtFechaCondicion.Enabled = False
    txtLibro.Enabled = False
    txtFolio.Enabled = False
    MostrarDatos
End Sub

Private Sub cmdGuardar_Click()
    If Estado = "Agregando" Then
        Conexion.Open
        'busco si el alumno ya estaba cargado con otra carrera para no duplicar
        Set Resultado = Conexion.Execute("SELECT Permiso FROM Alumnos Where Documento = " & txtDocumento & " AND Eliminado = 0")
        If Resultado.EOF = False Then
            Respuesta = MsgBox("El alumno ya existe con el permiso Nº " & Resultado!Permiso, , "Imposible agregar"): cmdCancelar_Click: Conexion.Close: Exit Sub
        End If
        Set Resultado = Conexion.Execute("SELECT Permiso FROM Alumnos")
        If Resultado.EOF = True Then 'no hay alumnos cargados
            RegistroActual = 1
        Else
            Set Resultado = Conexion.Execute("SELECT (MAX(Permiso) + 1) AS NuevoCodigo FROM Alumnos WHERE Permiso < 30000")
            RegistroActual = Resultado!NuevoCodigo
        End If
        Conexion.Execute ("INSERT INTO Alumnos ( Permiso, Nombre, Tipo, Documento, Domicilio, Localidad, Postal, DomicilioEnJunin, Telefono, Correo, FechaNacimiento, Lugar, Sexo, Titulo, Establecimiento, Registro, Comentario, Fotos, FotocopiaDNI,TituloSecundario, BloquearAutogestion ) VALUES (" & RegistroActual & ",'" & txtNombre & "','" & cbTipo & "'," & txtDocumento & ",'" & txtDomicilio & "','" & txtLocalidad & "','" & txtPostal & "','" & txtDomicilioEnJunin & "','" & txtTelefono & "','" & txtCorreo & "','" & DateValue(txtFechaNacimiento) & "','" & txtLugar & "','" & cbSexo & "','" & txtTitulo & "','" & txtEstablecimiento & "','" & txtRegistro & "','" & txtComentario & "'," & chkFotos & "," & chkDocumento & "," & chkTitulo & "," & chkBloquearAutogestion & ")")
        Conexion.Close
    Else
        RegistroActual = adoAlumnos.Recordset!Permiso
        Conexion.Open
        Conexion.Execute ("UPDATE Alumnos SET Alumnos.Nombre = '" & txtNombre & "', Alumnos.Tipo = '" & cbTipo & "', Alumnos.Documento = " & Val(txtDocumento) & ", Alumnos.Domicilio = '" & txtDomicilio & "', Alumnos.Localidad = '" & txtLocalidad & "', Alumnos.Postal = '" & txtPostal & "', Alumnos.DomicilioEnJunin = '" & txtDomicilioEnJunin & "', Alumnos.Telefono = '" & txtTelefono & "', Alumnos.Correo = '" & txtCorreo & "', Alumnos.FechaNacimiento = '" & DateValue(txtFechaNacimiento) & "', Alumnos.Lugar = '" & txtLugar & "', Alumnos.Sexo = '" & cbSexo & "', Alumnos.Titulo = '" & txtTitulo & "', Alumnos.Establecimiento = '" & txtEstablecimiento & "', Alumnos.Registro = '" & txtRegistro & "', Alumnos.Comentario = '" & txtComentario & "', Fotos = " & chkFotos.Value & ", FotocopiaDNI = " & chkDocumento & ", TituloSecundario=" & chkTitulo & ", BloquearAutogestion=" & chkBloquearAutogestion & " WHERE Alumnos.Permiso=" & RegistroActual)
        Conexion.Close
    End If
    adoAlumnos.Refresh
    adoAlumnos.Recordset.Find ("Permiso=" & RegistroActual)
    MostrarDatos
    cmdAgregar.Enabled = True
    cmdModificar.Enabled = True
    cmdGuardar.Enabled = False
    cmdCancelar.Enabled = False
    cmdSalir.Enabled = True
    cmdImprimir.Enabled = True
    cmdRestaurarPSW.Enabled = True
    frDatosAcademicos.Enabled = True
    frOpcionesDeBusqueda.Enabled = True
    DeshabilitarDatos
End Sub

Private Sub cmdGuardarCarrera_Click()
    If dtcCarrerasQueCursa.BoundText = "" Then
        MsgBox ("Debe elejir un plan de Estudios"): Exit Sub
    End If
    
    If EstadoCarreras = "Agregando" Then
        Dim rs As ADODB.Recordset
        ' Crear y abrir un Recordset
        Conexion.Open  ' abre
        Set rs = Conexion.Execute("SELECT CarrerasHechas.Permiso, CarrerasHechas.Carrera FROM CarrerasHechas WHERE CarrerasHechas.Permiso=" & adoAlumnos.Recordset!Permiso & " AND CarrerasHechas.Carrera=" & dtcCarrerasQueCursa.BoundText)
        If rs.EOF = False Then
           MsgBox ("Este plan ya se encuentra dado de alta para el Alumno"): rs.Close: Set rs = Nothing: Conexion.Close: Exit Sub
        End If
        rs.Close
        Set rs = Nothing
        Conexion.Close
    End If
    
    If EstadoCarreras = "Agregando" Then
        Conexion.Open
        Conexion.Execute ("INSERT INTO CarrerasHechas ( Permiso, Carrera, Ingreso, Division, Condición, Fecha, Libro, Folio ) VALUES (" & adoAlumnos.Recordset!Permiso & "," & dtcCarrerasQueCursa.BoundText & "," & Val(txtAñoIngreso) & "," & Val(txtDivision) & "," & dtcCondicion.BoundText & ",'" & DateValue(txtFechaCondicion) & "'," & Val(txtLibro) & "," & Val(txtFolio) & ")")
        Conexion.Close
    Else
        Conexion.Open
        Conexion.Execute ("UPDATE CarrerasHechas SET CarrerasHechas.Ingreso = " & Val(txtAñoIngreso) & ", CarrerasHechas.Division = " & Val(txtDivision) & ", CarrerasHechas.Condición = " & dtcCondicion.BoundText & ", CarrerasHechas.Fecha = '" & DateValue(txtFechaCondicion) & "', CarrerasHechas.Libro = " & Val(txtLibro) & ", CarrerasHechas.Folio = " & Val(txtFolio) & " WHERE CarrerasHechas.Permiso=" & adoAlumnos.Recordset!Permiso & " AND CarrerasHechas.Carrera=" & dtcCarrerasQueCursa.BoundText)
        Conexion.Close
    End If
    cmdAgregarCarrera.Enabled = True
    cmdModificarCarrera.Enabled = True
    cmdCancelarCarrera.Enabled = False
    cmdGuardarCarrera.Enabled = False
    frDatosPersonales.Enabled = True
    frComandos.Enabled = True
    frOpcionesDeBusqueda.Enabled = True
    dtcCarrerasQueCursa.Enabled = True
    EstadoCarreras = ""
    txtAñoIngreso.Enabled = False
    txtDivision.Enabled = False
    dtcCondicion.Enabled = False
    txtFechaCondicion.Enabled = False
    txtLibro.Enabled = False
    txtFolio.Enabled = False
    MostrarDatos
End Sub

Private Sub cmdMasDatos_Click()

End Sub

Private Sub cmdImprimir_Click()
    ReporteFichaPermiso = adoAlumnos.Recordset!Permiso
    ReporteFichaCarrera = dtcCarrerasQueCursa.BoundText
    
    cn.Open
    rptFichaAlumno.WindowState = 2
    rptFichaAlumno.Show 1
    cn.Close
End Sub

Private Sub cmdImprimirFicha_Click()

    
End Sub

Private Sub cmdModificar_Click()
    cmdAgregar.Enabled = False
    cmdModificar.Enabled = False
    cmdGuardar.Enabled = True
    cmdCancelar.Enabled = True
    cmdSalir.Enabled = False
    cmdImprimir.Enabled = False
    cmdRestaurarPSW.Enabled = False
    frDatosAcademicos.Enabled = False
    frOpcionesDeBusqueda.Enabled = False
    HabilitarDatos
    Estado = "Modificando"
End Sub

Private Sub cmdModificarCarrera_Click()
    EstadoCarreras = "Modificando"
    cmdAgregarCarrera.Enabled = False
    cmdModificarCarrera.Enabled = False
    cmdCancelarCarrera.Enabled = True
    cmdGuardarCarrera.Enabled = True
    frDatosPersonales.Enabled = False
    frComandos.Enabled = False
    frOpcionesDeBusqueda.Enabled = False
    dtcCarrerasQueCursa.Enabled = False
    txtAñoIngreso.Enabled = True
    txtDivision.Enabled = True
    dtcCondicion.Enabled = True
    txtFechaCondicion.Enabled = True
    txtLibro.Enabled = True
    txtFolio.Enabled = True
End Sub

Private Sub cmdNotificacionPersonal_Click()
    frmAgregarNotificacion.lblTipoDeNotificacion = "Personal"
    frmAgregarNotificacion.lblPermiso = adoAlumnos.Recordset!Permiso
    frmAgregarNotificacion.Show 1
End Sub

Private Sub cmdRestaurarPSW_Click()
    Respuesta = MsgBox("¿Está seguro de reestablecer la contraseña?", vbYesNo + vbQuestion, "Atención!!")
    If Respuesta = vbYes Then
        Conexion.Open
        Conexion.Execute ("update Alumnos set Contrasena='IOEL', Modo=4 where Permiso=" & adoAlumnos.Recordset!Permiso)
        Conexion.Close
        MsgBox ("La nueva contraseña es: 1234")
    End If

End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Function MostrarDatos()
   If adoAlumnos.Recordset.BOF = False And adoAlumnos.Recordset.EOF = False Then
        With adoAlumnos.Recordset
        txtPermiso = !Permiso
        txtNombre = !Nombre
        cbTipo = !Tipo
        If !documento <> vacio Then
            txtDocumento = !documento
        Else
            txtDocumento = ""
        End If
        If !Domicilio <> vacio Then
            txtDomicilio = !Domicilio
        Else
            txtDomicilio = ""
        End If
        If !Localidad <> vacio Then
            txtLocalidad = !Localidad
        Else
            txtLocalidad = ""
        End If
        If !Postal <> vacio Then
            txtPostal = !Postal
        Else
            txtPostal = ""
        End If
        If !DomicilioEnJunin <> vacio Then
            txtDomicilioEnJunin = !DomicilioEnJunin
        Else
            txtDomicilioEnJunin = ""
        End If
        If !Telefono <> vacio Then
            txtTelefono = !Telefono
        Else
            txtTelefono = ""
        End If
        If !Correo <> vacio Then
            txtCorreo = !Correo
        Else
            txtCorreo = ""
        End If
        If !FechaNacimiento <> vacio Then
            txtFechaNacimiento = !FechaNacimiento
        Else
            txtFechaNacimiento = ""
        End If
        If !Lugar <> vacio Then
            txtLugar = !Lugar
        Else
            txtLugar = ""
        End If
        cbSexo = !Sexo
        If !Titulo <> vacio Then
            txtTitulo = !Titulo
        Else
            txtTitulo = ""
        End If
        If !Establecimiento <> vacio Then
            txtEstablecimiento = !Establecimiento
        Else
            txtEstablecimiento = ""
        End If
        If !Registro <> vacio Then
            txtRegistro = !Registro
        Else
            txtRegistro = ""
        End If
        If !Comentario <> vacio Then
            txtComentario = !Comentario
        Else
            txtComentario = ""
        End If
        If !Fotos = True Then
            chkFotos = 1
        Else
            chkFotos = 0
        End If
        If !FotocopiaDNI = True Then
            chkDocumento = 1
        Else
            chkDocumento = 0
        End If
        If !TituloSecundario = True Then
            chkTitulo = 1
        Else
            chkTitulo = 0
        End If
        If !BloquearAutogestion = True Then
            chkBloquearAutogestion = 1
        Else
            chkBloquearAutogestion = 0
        End If

        End With
        'habilito o deshabilito los botones para modificar o agregar datos del alumno segun el permiso del usuario
        If frmIdentificacion.Permisos!ModificarAlumnos = True Then
            cmdAgregar.Enabled = True
            cmdModificar.Enabled = True
        Else
            cmdAgregar.Enabled = False
            cmdModificar.Enabled = False
        End If
        adoCarrerasQueCursa.RecordSource = "SELECT Carreras.Codigo, Carreras.Abreviatura, CarrerasHechas.Ingreso, CarrerasHechas.Division, CarrerasHechas.Condición, CarrerasHechas.Fecha, CarrerasHechas.Libro, CarrerasHechas.Folio FROM CarrerasHechas INNER JOIN Carreras ON CarrerasHechas.Carrera = Carreras.Codigo Where (((CarrerasHechas.Permiso) = " & adoAlumnos.Recordset!Permiso & ")) ORDER BY Carreras.Abreviatura"
        adoCarrerasQueCursa.Refresh
        If adoCarrerasQueCursa.Recordset.RecordCount > 0 Then
            'tiene asignada alguna carrera
            dtcCarrerasQueCursa.BoundText = adoCarrerasQueCursa.Recordset!Codigo
            MostrarDatosAcademicos
        Else
            dtcCarrerasQueCursa.Text = ""
            cmdModificarCarrera.Enabled = False
        End If
    End If
End Function

Private Sub Command1_Click()

End Sub

Private Sub dtcCarrerasQueCursa_Change()
If EstadoCarreras <> "Agregando" And EstadoCarreras <> "Modificando" Then
    cmdAgregarCarrera.Enabled = True
    If adoCarrerasQueCursa.Recordset.EOF = False And adoCarrerasQueCursa.Recordset.BOF = False Then
        adoCarrerasQueCursa.Recordset.MoveFirst
        adoCarrerasQueCursa.Recordset.Find ("Codigo=" & dtcCarrerasQueCursa.BoundText)
        MostrarDatosAcademicos
        cmdModificarCarrera.Enabled = True
    Else
        txtAñoIngreso = ""
        txtDivision = ""
        dtcCondicion.Text = ""
        txtFechaCondicion = ""
        txtLibro = ""
        txtFolio = ""
        cmdModificarCarrera.Enabled = False
    End If
Else 'se le está asignando o modificando una carrera al alumno
    cmdGuardarCarrera.SetFocus
End If
'habilito o deshabilito los botones para modificar o agregar datos academicos segun el permiso del usuario
If frmIdentificacion.Permisos!ModificarDatosAcademicos = True Then
    cmdModificarCarrera.Enabled = True
    cmdAgregarCarrera.Enabled = True
Else
    cmdModificarCarrera.Enabled = False
    cmdAgregarCarrera.Enabled = False
End If
End Sub


Private Sub dtcCarrerasVigentes_Click(Area As Integer)
    If dtcCarrerasVigentes.BoundText <> "" And dtcSituación.BoundText <> "" Then
        adoAlumnos.RecordSource = "SELECT Alumnos.* FROM Alumnos INNER JOIN CarrerasHechas ON Alumnos.Permiso = CarrerasHechas.Permiso WHERE (((CarrerasHechas.Carrera)=" & dtcCarrerasVigentes.BoundText & ") AND ((CarrerasHechas.Condición)=" & dtcSituación.BoundText & ") AND ((Alumnos.Eliminado)=False)) ORDER BY Nombre"
        adoAlumnos.Refresh
        MostrarDatos
    End If
End Sub

Private Sub dtcSituación_Click(Area As Integer)
    If dtcCarrerasVigentes.BoundText <> "" And dtcSituación.BoundText <> "" Then
        adoAlumnos.RecordSource = "SELECT Alumnos.* FROM Alumnos INNER JOIN CarrerasHechas ON Alumnos.Permiso = CarrerasHechas.Permiso WHERE (((CarrerasHechas.Carrera)=" & dtcCarrerasVigentes.BoundText & ") AND ((CarrerasHechas.Condición)=" & dtcSituación.BoundText & ") AND ((Alumnos.Eliminado)=False)) ORDER BY Nombre"
        adoAlumnos.Refresh
        MostrarDatos
    End If
End Sub

Private Sub dtgAlumnos_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    MostrarDatos
End Sub

Private Sub Form_Activate()
    Conexion.ConnectionString = ("DSN=Instituto")
    If adoAlumnos.Recordset.EOF = False And adoCarrerasQueCursa.Recordset.EOF = False Then
       dtcCarrerasVigentes.BoundText = adoCarrerasVigentes.Recordset!Codigo
       dtcSituación.BoundText = adoBuscarPorSituacion.Recordset!Codigo
    End If
    optBuscarPor(0).Value = True
   
    
    MostrarDatos
End Sub

Private Sub optBuscarPor_Click(Index As Integer)
    txtBuscarPorDocumento.Enabled = False
    txtBuscarPorDomicilio.Enabled = False
    txtBuscarPorNombre.Enabled = False
    txtBuscarPorPermiso.Enabled = False
    dtcCarrerasVigentes.Enabled = False
    dtcSituación.Enabled = False
    txtBuscarPorDocumento = ""
    txtBuscarPorDomicilio = ""
    txtBuscarPorNombre = ""
    txtBuscarPorPermiso = ""
    If adoAlumnos.Recordset.EOF = False Then
        If Index = 0 Then
            txtBuscarPorNombre.Enabled = True
            adoAlumnos.RecordSource = "SELECT * FROM Alumnos WHERE Eliminado = 0 ORDER BY Nombre"
            adoAlumnos.Refresh
            txtBuscarPorNombre.SetFocus
        ElseIf Index = 1 Then
            txtBuscarPorPermiso.Enabled = True
            adoAlumnos.RecordSource = "SELECT * FROM Alumnos WHERE Eliminado = False ORDER BY Permiso"
            adoAlumnos.Refresh
            txtBuscarPorPermiso.SetFocus
        ElseIf Index = 2 Then
            txtBuscarPorDocumento.Enabled = True
            adoAlumnos.RecordSource = "SELECT * FROM Alumnos WHERE Eliminado = False ORDER BY Documento"
            adoAlumnos.Refresh
            txtBuscarPorDocumento.SetFocus
        ElseIf Index = 3 Then
            txtBuscarPorDomicilio.Enabled = True
            adoAlumnos.RecordSource = "SELECT * FROM Alumnos WHERE Eliminado = False ORDER BY Domicilio"
            adoAlumnos.Refresh
            txtBuscarPorDomicilio.SetFocus
        ElseIf Index = 4 Then
            dtcCarrerasVigentes.Enabled = True
            dtcSituación.Enabled = True
        End If
    End If
    If adoAlumnos.Recordset.RecordCount > 0 Then
        cmdModificar.Enabled = True
    Else
        cmdModificar.Enabled = False
    End If
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 Then KeyAscii = 0
End Sub

Private Sub txtAñoIngreso_KeyPress(KeyAscii As Integer)
    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 Then KeyAscii = 0
End Sub

Private Sub txtBuscarPorDocumento_Change()
    adoAlumnos.Recordset.MoveFirst
    adoAlumnos.Recordset.Find ("Documento >=" & Val(txtBuscarPorDocumento))
    MostrarDatos
End Sub

Private Sub txtBuscarPorDocumento_KeyPress(KeyAscii As Integer)
    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 Then KeyAscii = 0
End Sub

Private Sub txtBuscarPorDomicilio_Change()
    adoAlumnos.Recordset.MoveFirst
    adoAlumnos.Recordset.Find ("Domicilio >='" & txtBuscarPorDomicilio & "'")
    MostrarDatos
End Sub

Private Sub txtBuscarPorNombre_Change()
    adoAlumnos.Recordset.MoveFirst
    adoAlumnos.Recordset.Find ("Nombre >='" & txtBuscarPorNombre & "'")
    MostrarDatos
End Sub

Private Sub txtBuscarPorPermiso_Change()
    adoAlumnos.Recordset.MoveFirst
    adoAlumnos.Recordset.Find ("Permiso >=" & Val(txtBuscarPorPermiso))
    MostrarDatos
End Sub

Private Sub txtBuscarPorPermiso_KeyPress(KeyAscii As Integer)
    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 Then KeyAscii = 0
End Sub

Private Function MostrarDatosAcademicos()
    With adoCarrerasQueCursa.Recordset
    txtAñoIngreso = !Ingreso
    txtDivision = !division
    dtcCondicion.BoundText = !Condición
    If !Fecha <> Nulo Then
        txtFechaCondicion = !Fecha
    Else
        txtFechaCondicion = ""
    End If
    txtLibro = !Libro
    txtFolio = !Folio
    End With
End Function

Private Function HabilitarDatos()
    txtNombre.Enabled = True
    cbTipo.Enabled = True
    txtDocumento.Enabled = True
    txtDomicilio.Enabled = True
    txtLocalidad.Enabled = True
    txtPostal.Enabled = True
    txtDomicilioEnJunin.Enabled = True
    txtTelefono.Enabled = True
    txtCorreo.Enabled = True
    txtFechaNacimiento.Enabled = True
    txtLugar.Enabled = True
    cbSexo.Enabled = True
    txtTitulo.Enabled = True
    txtEstablecimiento.Enabled = True
    txtRegistro.Enabled = True
    txtComentario.Enabled = True
    frDocumentacion.Enabled = True
End Function

Private Function DeshabilitarDatos()
    txtNombre.Enabled = False
    cbTipo.Enabled = False
    txtDocumento.Enabled = False
    txtDomicilio.Enabled = False
    txtLocalidad.Enabled = False
    txtPostal.Enabled = False
    txtDomicilioEnJunin.Enabled = False
    txtTelefono.Enabled = False
    txtCorreo.Enabled = False
    txtFechaNacimiento.Enabled = False
    txtLugar.Enabled = False
    cbSexo.Enabled = False
    txtTitulo.Enabled = False
    txtEstablecimiento.Enabled = False
    txtRegistro.Enabled = False
    txtComentario.Enabled = False
    frDocumentacion.Enabled = False
End Function

Private Function LimpiarDatos()
    txtPermiso = ""
    txtNombre = ""
    txtDocumento = ""
    txtDomicilio = ""
    txtLocalidad = ""
    txtPostal = ""
    txtDomicilioEnJunin = ""
    txtTelefono = ""
    txtCorreo = ""
    txtFechaNacimiento = ""
    txtLugar = ""
    txtTitulo = ""
    txtEstablecimiento = ""
    txtRegistro = ""
    txtComentario = ""
    chkFotos = 0
    chkDocumento = 0
    chkTitulo = 0
End Function

Private Sub txtCorreo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtFechaNacimiento.SetFocus
End Sub

Private Sub txtDivision_KeyPress(KeyAscii As Integer)
    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 Then KeyAscii = 0
End Sub

Private Sub txtDocumento_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cbSexo.SetFocus
    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 Then KeyAscii = 0
End Sub

Private Sub txtDomicilio_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtLocalidad.SetFocus
End Sub

Private Sub txtDomicilioEnJunin_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtTelefono.SetFocus
End Sub

Private Sub txtEstablecimiento_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtRegistro.SetFocus
End Sub

Private Sub txtFechaCondicion_GotFocus()
    If txtFechaCondicion <> "" Then clFechaCondicion.Value = txtFechaCondicion
    clFechaCondicion.Visible = True
    frDatosAcademicos.Enabled = False
End Sub

Private Sub txtFechaNacimiento_GotFocus()
    frDatosPersonales.Enabled = False
    frComandos.Enabled = False
    frDatosAcademicos.Enabled = False
    dtgAlumnos.Enabled = False
    If txtFechaNacimiento <> "" Then clFechaNacimiento.Value = txtFechaNacimiento
    clFechaNacimiento.Visible = True
End Sub

Private Sub txtFechaNacimiento_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtLugar.SetFocus
End Sub

Private Sub txtFolio_KeyPress(KeyAscii As Integer)
    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 Then KeyAscii = 0
End Sub

Private Sub txtLibro_KeyPress(KeyAscii As Integer)
    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 Then KeyAscii = 0
End Sub

Private Sub txtLocalidad_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtPostal.SetFocus
End Sub

Private Sub txtLugar_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtTitulo.SetFocus
End Sub

Private Sub txtNombre_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cbTipo.SetFocus
End Sub

Private Sub txtPostal_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtDomicilioEnJunin.SetFocus
End Sub

Private Sub txtRegistro_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cmdGuardar.SetFocus
End Sub

Private Sub txtTelefono_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtCorreo.SetFocus
End Sub

Private Sub txtTitulo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtEstablecimiento.SetFocus
End Sub
