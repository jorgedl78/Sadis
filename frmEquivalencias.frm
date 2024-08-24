VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmEquivalencias 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Equivalencias"
   ClientHeight    =   8595
   ClientLeft      =   45
   ClientTop       =   240
   ClientWidth     =   11910
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8595
   ScaleWidth      =   11910
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame4 
      Height          =   2295
      Left            =   10920
      TabIndex        =   46
      Top             =   6120
      Width           =   855
      Begin VB.CommandButton cmdIngresarActa 
         Height          =   550
         Left            =   120
         Picture         =   "frmEquivalencias.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   48
         ToolTipText     =   "Ingresar Acta"
         Top             =   1320
         Width           =   550
      End
      Begin VB.CommandButton cmdImprimir 
         Enabled         =   0   'False
         Height          =   550
         Left            =   120
         Picture         =   "frmEquivalencias.frx":0152
         Style           =   1  'Graphical
         TabIndex        =   47
         ToolTipText     =   "Imprimir Acta"
         Top             =   480
         Width           =   550
      End
   End
   Begin VB.Frame Frame3 
      Height          =   3735
      Left            =   10920
      TabIndex        =   40
      Top             =   2280
      Width           =   855
      Begin VB.CommandButton cmdCancelar 
         Enabled         =   0   'False
         Height          =   600
         Left            =   120
         Picture         =   "frmEquivalencias.frx":07BC
         Style           =   1  'Graphical
         TabIndex        =   45
         ToolTipText     =   "Cancelar"
         Top             =   2280
         Width           =   600
      End
      Begin VB.CommandButton cmdGuardar 
         Enabled         =   0   'False
         Height          =   600
         Left            =   120
         Picture         =   "frmEquivalencias.frx":0BFE
         Style           =   1  'Graphical
         TabIndex        =   44
         ToolTipText     =   "Guardar"
         Top             =   1560
         Width           =   600
      End
      Begin VB.CommandButton cmdModificar 
         Enabled         =   0   'False
         Height          =   600
         Left            =   120
         Picture         =   "frmEquivalencias.frx":1040
         Style           =   1  'Graphical
         TabIndex        =   43
         ToolTipText     =   "Modificar"
         Top             =   840
         Width           =   600
      End
      Begin VB.CommandButton cmdAgregar 
         Enabled         =   0   'False
         Height          =   550
         Left            =   120
         Picture         =   "frmEquivalencias.frx":1482
         Style           =   1  'Graphical
         TabIndex        =   42
         ToolTipText     =   "Agregar"
         Top             =   240
         Width           =   550
      End
      Begin VB.CommandButton cmdEliminarEquivalencia 
         Enabled         =   0   'False
         Height          =   550
         Left            =   120
         Picture         =   "frmEquivalencias.frx":18C4
         Style           =   1  'Graphical
         TabIndex        =   41
         ToolTipText     =   "Borrar"
         Top             =   3000
         Width           =   550
      End
   End
   Begin VB.Frame frSolicitudes 
      Caption         =   "Solicitudes"
      Height          =   6135
      Left            =   120
      TabIndex        =   14
      Top             =   2280
      Width           =   10695
      Begin VB.Frame frResolucion 
         Caption         =   "Resolución"
         Height          =   2055
         Left            =   120
         TabIndex        =   23
         Top             =   3960
         Width           =   10455
         Begin VB.TextBox txtLibro 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            Height          =   285
            Left            =   600
            TabIndex        =   50
            Top             =   960
            Width           =   495
         End
         Begin MSDataListLib.DataCombo dtcProfesor 
            Bindings        =   "frmEquivalencias.frx":1D06
            Height          =   315
            Left            =   120
            TabIndex        =   39
            Top             =   1560
            Width           =   2055
            _ExtentX        =   3625
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            Style           =   2
            ListField       =   "Nombre"
            BoundColumn     =   "Codigo"
            Text            =   ""
         End
         Begin VB.TextBox txtObservacion 
            Enabled         =   0   'False
            Height          =   1575
            Left            =   7560
            TabIndex        =   37
            Top             =   360
            Width           =   2775
         End
         Begin VB.TextBox txtFundamentacion 
            Enabled         =   0   'False
            Height          =   1575
            Left            =   2280
            TabIndex        =   30
            Top             =   360
            Width           =   5055
         End
         Begin VB.TextBox txtFolio 
            Enabled         =   0   'False
            Height          =   285
            Left            =   1680
            TabIndex        =   28
            Top             =   960
            Width           =   495
         End
         Begin VB.CheckBox chkOtorgada 
            Alignment       =   1  'Right Justify
            Caption         =   "Otorgada"
            Enabled         =   0   'False
            Height          =   195
            Left            =   120
            TabIndex        =   25
            Top             =   240
            Width           =   975
         End
         Begin MSComCtl2.DTPicker dtpFechaResolucion 
            DragIcon        =   "frmEquivalencias.frx":1D20
            Height          =   285
            Left            =   720
            TabIndex        =   35
            Top             =   600
            Visible         =   0   'False
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   503
            _Version        =   393216
            Enabled         =   0   'False
            Format          =   127139841
            CurrentDate     =   37550
            MinDate         =   -36522
         End
         Begin MSAdodcLib.Adodc adoProfesor 
            Height          =   330
            Left            =   480
            Top             =   1320
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
            RecordSource    =   "SELECT * FROM Personal ORDER BY Nombre"
            Caption         =   "Profesor"
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
         Begin VB.Label Label4 
            Caption         =   "Libro:"
            Height          =   255
            Left            =   120
            TabIndex        =   51
            Top             =   960
            Width           =   375
         End
         Begin VB.Label Label20 
            Caption         =   "Profesor:"
            Height          =   255
            Left            =   120
            TabIndex        =   38
            Top             =   1320
            Width           =   735
         End
         Begin VB.Label Label12 
            Caption         =   "Observación:"
            Height          =   255
            Left            =   7800
            TabIndex        =   36
            Top             =   120
            Width           =   975
         End
         Begin VB.Label Label11 
            Caption         =   "Fundamentación:"
            Height          =   255
            Left            =   2280
            TabIndex        =   29
            Top             =   120
            Width           =   1455
         End
         Begin VB.Label Label10 
            Caption         =   "Folio:"
            Height          =   255
            Left            =   1200
            TabIndex        =   27
            Top             =   960
            Width           =   375
         End
         Begin VB.Label Label9 
            Caption         =   "Fecha:"
            Height          =   255
            Left            =   120
            TabIndex        =   26
            Top             =   600
            Width           =   615
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Iniciada"
         Height          =   855
         Left            =   120
         TabIndex        =   20
         Top             =   3000
         Width           =   4095
         Begin VB.TextBox txtUsuario 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   285
            Left            =   2280
            TabIndex        =   33
            Top             =   240
            Width           =   1695
         End
         Begin MSComCtl2.DTPicker dtpFechaInicio 
            DragIcon        =   "frmEquivalencias.frx":2162
            Height          =   285
            Left            =   480
            TabIndex        =   32
            Top             =   240
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   503
            _Version        =   393216
            Enabled         =   0   'False
            Format          =   91881473
            CurrentDate     =   37550
            MinDate         =   -36522
         End
         Begin VB.Label lblpor 
            Caption         =   "Por"
            Height          =   255
            Left            =   1920
            TabIndex        =   22
            Top             =   240
            Width           =   255
         End
         Begin VB.Label Label7 
            Caption         =   "el"
            Height          =   255
            Left            =   120
            TabIndex        =   21
            Top             =   240
            Width           =   255
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Acreditación"
         Height          =   855
         Left            =   4320
         TabIndex        =   16
         Top             =   3000
         Width           =   6255
         Begin VB.CommandButton cmdInstituciones 
            Caption         =   "..."
            Height          =   300
            Left            =   5640
            TabIndex        =   49
            Top             =   480
            Width           =   495
         End
         Begin MSAdodcLib.Adodc adoInstituciones 
            Height          =   330
            Left            =   0
            Top             =   720
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
            RecordSource    =   $"frmEquivalencias.frx":25A4
            Caption         =   "Instituciones"
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
         Begin MSDataListLib.DataCombo dtcInstituciones 
            Bindings        =   "frmEquivalencias.frx":25F9
            Height          =   315
            Left            =   120
            TabIndex        =   31
            Top             =   480
            Width           =   5415
            _ExtentX        =   9551
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            Style           =   2
            ListField       =   "Institucion"
            BoundColumn     =   "Codigo"
            Text            =   ""
         End
         Begin VB.TextBox txtNota 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            Height          =   285
            Left            =   3000
            TabIndex        =   19
            Top             =   120
            Width           =   615
         End
         Begin MSComCtl2.DTPicker dtpFechaAcreditacion 
            DragIcon        =   "frmEquivalencias.frx":2618
            Height          =   285
            Left            =   4680
            TabIndex        =   34
            Top             =   120
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   503
            _Version        =   393216
            Enabled         =   0   'False
            Format          =   92078081
            CurrentDate     =   37550
            MinDate         =   -36522
         End
         Begin VB.Label Label8 
            Caption         =   "Institución:"
            Height          =   255
            Left            =   120
            TabIndex        =   24
            Top             =   240
            Width           =   855
         End
         Begin VB.Label Label6 
            Caption         =   "Fecha:"
            Height          =   255
            Left            =   4080
            TabIndex        =   18
            Top             =   120
            Width           =   615
         End
         Begin VB.Label Label5 
            Caption         =   "Nota:"
            Height          =   255
            Left            =   2520
            TabIndex        =   17
            Top             =   120
            Width           =   495
         End
      End
      Begin MSAdodcLib.Adodc adoEquivalencias 
         Height          =   375
         Left            =   1200
         Top             =   1800
         Visible         =   0   'False
         Width           =   2415
         _ExtentX        =   4260
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
         RecordSource    =   $"frmEquivalencias.frx":2A5A
         Caption         =   "Equivalencias"
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
      Begin MSDataGridLib.DataGrid dtgEquivalencias 
         Bindings        =   "frmEquivalencias.frx":2EC2
         Height          =   2655
         Left            =   120
         TabIndex        =   15
         Top             =   240
         Width           =   10455
         _ExtentX        =   18441
         _ExtentY        =   4683
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
            DataField       =   "Asignatura"
            Caption         =   "Asignatura"
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
            DataField       =   "NombreInstitucion"
            Caption         =   "NombreInstitucion"
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
               ColumnWidth     =   780,095
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   2129,953
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   5204,977
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   1800
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame frMaterias 
      Enabled         =   0   'False
      Height          =   855
      Left            =   0
      TabIndex        =   10
      Top             =   1200
      Width           =   11895
      Begin MSDataListLib.DataCombo dtcMaterias 
         Bindings        =   "frmEquivalencias.frx":2EE1
         Height          =   315
         Left            =   120
         TabIndex        =   11
         Top             =   360
         Width           =   11535
         _ExtentX        =   20346
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "Nombre"
         BoundColumn     =   "Codigo"
         Text            =   ""
      End
      Begin MSAdodcLib.Adodc adoMaterias 
         Height          =   330
         Left            =   5160
         Top             =   600
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
         RecordSource    =   $"frmEquivalencias.frx":2EFB
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
      Begin VB.Label lblCodigoProfesor 
         Height          =   135
         Left            =   8760
         TabIndex        =   13
         Top             =   120
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label Label3 
         Caption         =   "Materias"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   120
         Width           =   615
      End
   End
   Begin VB.Frame frCarrera 
      Height          =   1095
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11895
      Begin VB.ComboBox cbCurso 
         Height          =   315
         Left            =   8280
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   480
         Width           =   735
      End
      Begin VB.TextBox txtAño 
         Height          =   285
         Left            =   9120
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   480
         Width           =   615
      End
      Begin VB.CommandButton cmdMostrar 
         Caption         =   "Mostrar"
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
         Left            =   10080
         TabIndex        =   2
         Top             =   360
         Width           =   1095
      End
      Begin VB.CommandButton cmdSalir 
         Height          =   550
         Left            =   11280
         Picture         =   "frmEquivalencias.frx":2FFC
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Salir"
         Top             =   360
         Width           =   550
      End
      Begin MSComCtl2.UpDown UpDown1 
         Height          =   375
         Left            =   9720
         TabIndex        =   4
         Top             =   420
         Width           =   240
         _ExtentX        =   450
         _ExtentY        =   661
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSAdodcLib.Adodc adoCarreras 
         Height          =   330
         Left            =   1920
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
         RecordSource    =   $"frmEquivalencias.frx":343E
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
      Begin MSDataListLib.DataCombo dtcCarreras 
         Bindings        =   "frmEquivalencias.frx":3542
         Height          =   315
         Left            =   120
         TabIndex        =   6
         Top             =   480
         Width           =   8055
         _ExtentX        =   14208
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "Nombre"
         BoundColumn     =   "Codigo"
         Text            =   ""
      End
      Begin VB.Label Label1 
         Caption         =   "Carrera:"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label2 
         Caption         =   "Año:"
         Height          =   255
         Left            =   9120
         TabIndex        =   8
         Top             =   240
         Width           =   615
      End
      Begin VB.Label lblCurso 
         Caption         =   "Curso:"
         Height          =   255
         Left            =   8280
         TabIndex        =   7
         Top             =   240
         Width           =   495
      End
   End
End
Attribute VB_Name = "frmEquivalencias"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Conexion As New Connection
Dim Resultado As New Recordset
Dim Auxiliar As New Recordset
Dim Otorgadas As New Recordset
Dim Profesor(10) As String
Dim CodigoProfesor(10) As Integer
Public FechaDeLaMesa As Date
Public ProfesorCodigo As Integer
Public NombreProfesor As String
Dim AlumnoActual As Double

Private Sub cbCurso_Click()
    NuevasMaterias
End Sub

Private Sub chkOtorgada_Click()
If chkOtorgada.Value = 1 Then
    dtpFechaResolucion.Visible = True
Else
    dtpFechaResolucion.Visible = False
End If
End Sub

Private Sub cmdAgregar_Click()
    If adoEquivalencias.Recordset.RecordCount >= 25 Then MsgBox ("Solo se pueden ingresar hasta 25 solicitudes"): Exit Sub
    Conexion.Open
    'verifico que el acta no haya sido impresa
    Set Resultado = Conexion.Execute("SELECT Impresa FROM EquivalenciasResumen WHERE Asignatura = " & dtcMaterias.BoundText & " AND AnoSolicitud = " & txtAño & "")
    If Resultado.EOF = False Then
        If Resultado!Impresa = True Then Respuesta = MsgBox("El acta de equivalencias para esta asignatura ya se imprimió", vbOKOnly, "Imposible eliminar equivalencias"): Conexion.Close: Exit Sub
    End If
    Conexion.Close
    frmEquivalenciaNueva.dtcProfesor.BoundText = dtcProfesor.BoundText
    frmEquivalenciaNueva.dtpFechaSolicitud = Date
    frmEquivalenciaNueva.Show 1
End Sub

Private Sub cmdCancelar_Click()
    cmdAgregar.Enabled = True
    cmdModificar.Enabled = True
    cmdGuardar.Enabled = False
    cmdCancelar.Enabled = False
    frCarrera.Enabled = True
    frMaterias.Enabled = True
    DeshabilitarControles
End Sub


Private Sub cmdEliminarEquivalencia_Click()
    Conexion.Open
    'verifico que el acta no haya sido impresa
    Set Resultado = Conexion.Execute("SELECT Impresa FROM EquivalenciasResumen WHERE Asignatura = " & dtcMaterias.BoundText & " AND AnoSolicitud = " & txtAño & "")
    If Resultado.EOF = False Then
        If Resultado!Impresa = True Then Respuesta = MsgBox("El acta de equivalencias para esta asignatura ya se imprimió", vbOKOnly, "Imposible eliminar equivalencias"): Conexion.Close: Exit Sub
    End If
    
    If frmIdentificacion.dtcUsuarios.BoundText <> adoEquivalencias.Recordset!Usuario And frmIdentificacion.Permisos!CerrarEquivalencias = False Then
        Respuesta = MsgBox("Solo tiene permiso el usuario que inicio el trámite", vbOKOnly, "Imposible borrar la equivalencia"): Conexion.Close: Exit Sub
    End If
        
    Respuesta = MsgBox("Esta seguro de eliminar la equivalencia", vbYesNo, "Eliminar Equivalencia")
    If Respuesta = vbYes Then
        Conexion.Execute ("DELETE * FROM Equivalencias Where MateriaAReconocer = " & dtcMaterias.BoundText & " AND AnoSolicitud = " & txtAño & " AND Alumno =" & adoEquivalencias.Recordset!Permiso)
    End If
    Conexion.Close
    adoEquivalencias.Refresh
End Sub

Private Sub cmdGuardar_Click()
    cmdAgregar.Enabled = True
    cmdModificar.Enabled = True
    cmdGuardar.Enabled = False
    cmdCancelar.Enabled = False
    frCarrera.Enabled = True
    frMaterias.Enabled = True
    DeshabilitarControles
    Conexion.Open
    Conexion.Execute ("UPDATE Equivalencias SET FechaSolicitud = '" & DateValue(dtpFechaInicio) & "', FechaAprobacion='" & DateValue(dtpFechaAcreditacion) & "', Nota = " & Replace(txtNota, ",", ".") & ", Otorgada = " & chkOtorgada.Value & ", FechaResolucion='" & DateValue(dtpFechaResolucion) & "', Institucion=" & dtcInstituciones.BoundText & ", Profesor=" & dtcProfesor.BoundText & ", Observacion='" & txtObservacion & "', Fundamentacion='" & txtFundamentacion & "' WHERE Alumno=" & adoEquivalencias.Recordset!Permiso & " AND MateriaAReconocer = " & dtcMaterias.BoundText & " AND AnoSolicitud = " & txtAño)
    Conexion.Close
    adoEquivalencias.Refresh
    adoEquivalencias.Recordset.Find ("Permiso=" & AlumnoActual)
    AlumnoActual = 0
    dtgEquivalencias.SetFocus
End Sub

Private Sub cmdImprimir_Click()
    Conexion.Open
    'verifico que el acta no haya sido impresa
    Set Resultado = Conexion.Execute("SELECT * FROM EquivalenciasResumen WHERE Asignatura = " & dtcMaterias.BoundText & " AND AnoSolicitud = " & txtAño & "")
    If Resultado.EOF = False Then Respuesta = MsgBox("El acta de Equivalencia ya fue impresa", vbOKOnly, "Imposible imprimir"): Conexion.Close: Exit Sub
    
    'levanto las equivalencias otorgadas
    Set Otorgadas = Conexion.Execute("SELECT Equivalencias.Alumno, Equivalencias.Otorgada, Alumnos.Nombre, ([Alumnos].[Tipo] & ' ' & [Alumnos].[Documento]) AS Documento, Equivalencias.Nota FROM Equivalencias INNER JOIN Alumnos ON Equivalencias.Alumno = Alumnos.Permiso WHERE (((Equivalencias.Otorgada)=True) AND ((Equivalencias.MateriaAReconocer)=" & dtcMaterias.BoundText & ") AND ((Equivalencias.AnoSolicitud)=" & txtAño & "))")
    If Otorgadas.EOF = True Then MsgBox ("Ningún alumno tiene la equivalencia otorgada"): Conexion.Close: Exit Sub
    
    Set Resultado = Conexion.Execute("SELECT NombreInstitucion FROM Parametros")
    NombreInstitucion = Resultado!NombreInstitucion
    
    
    frmEquivalenciaImprimirActa.Show 1
    
    
    'armo e imprimo las actas de equivalencia
    Otorgadas.MoveFirst
    OrdenActa = 1
    TotalActas = 1
    ArmarEncabezado = "Si"
    With frmImprimeActas
    .lblEstablecimiento = NombreInstitucion
    While Otorgadas.EOF = False
        ImprimioActa = "No"
        If ArmarEncabezado = "Si" Then ArmarEncabezado = "No": MsgBox ("Imprimir acta Nº " & TotalActas): ArmaEncabezado
        .lblOrden(OrdenActa).Visible = True
        .lblPermiso(OrdenActa).Visible = True
        .lblEscritoNota(OrdenActa).Visible = True
        .lnlEscritoLetras(OrdenActa).Visible = True
        .lblOralNota(OrdenActa).Visible = True
        .lnlOralLetras(OrdenActa).Visible = True
        .lblFinalNota(OrdenActa).Visible = True
        .lnlFinalLetras(OrdenActa).Visible = True
        .lblAlumno(OrdenActa).Visible = True
        .lblDocumento(OrdenActa).Visible = True
        .lblPermiso(OrdenActa) = Otorgadas!Alumno
        .lblAlumno(OrdenActa) = Otorgadas!Nombre
        .lblDocumento(OrdenActa) = Format(Otorgadas!documento, "##,###,###")
        .lblEscritoNota(OrdenActa) = Format(Otorgadas!Nota, "0.00")
        .lblFinalNota(OrdenActa) = Format(Otorgadas!Nota, "0.00")
        'Agrega nota en letras
        Decimo = Int(Otorgadas!Nota)
        centesimo = Right(Format(Otorgadas!Nota, "0.00"), 3)
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
        .lnlEscritoLetras(OrdenActa) = EnLetras
        .lnlFinalLetras(OrdenActa) = EnLetras
        OrdenActa = OrdenActa + 1
        Otorgadas.MoveNext
        If OrdenActa = 26 And Otorgadas.EOF = False Then
            .lblTotalAlumnos = OrdenActa - 1
            frmImprimeActas.PrintForm
            Unload frmImprimeActas
            'ingreso el detalle del acta en la tabla EquivalenciasResumen            ImprimioActa = "Si"
            Conexion.Execute ("INSERT INTO EquivalenciasResumen ( Asignatura, AnoSolicitud, Solicitantes, Otorgadas, Impresa, FechaImpresion, FechaActa, Profesor ) VALUES (" & dtcMaterias.BoundText & ", " & txtAño & ", " & adoEquivalencias.Recordset.RecordCount & ", " & OrdenActa - 1 & ", True, #" & Format(Date, "mm/dd/yyyy") & "#, #" & Format(FechaDeLaMesa, "mm/dd/yyyy") & "#, " & ProfesorCodigo & ") ")
            TotalActas = TotalActas + 1
            OrdenActa = 1
            ArmarEncabezado = "No"
        End If
    Wend
    If ImprimioActa = "No" Then
       .lblTotalAlumnos = OrdenActa - 1
       frmImprimeActas.PrintForm
       Unload frmImprimeActas
       Conexion.Execute ("INSERT INTO EquivalenciasResumen ( Asignatura, AnoSolicitud, Solicitantes, Otorgadas, Impresa, FechaImpresion, FechaActa, Profesor ) VALUES (" & dtcMaterias.BoundText & ", " & txtAño & ", " & adoEquivalencias.Recordset.RecordCount & ", " & OrdenActa - 1 & ", True, '" & DateValue(Date) & "', '" & DateValue(FechaDeLaMesa) & "', " & ProfesorCodigo & ") ")
    End If
    Conexion.Close
    End With
End Sub

Private Sub cmdIngresarActa_Click()
    Conexion.Open
    'controlo que el acta no este ingresada
    Set Resultado = Conexion.Execute("Select * FROM EquivalenciasResumen WHERE Asignatura =" & dtcMaterias.BoundText & " AND AnoSolicitud = " & txtAño)
    If Resultado.EOF = True Then MsgBox ("El acta no se encuentra impresa o disponible para ingresar"): Conexion.Close: Exit Sub
    If Resultado!Ingresada = True Then Respuesta = MsgBox("El acta ya fue ingresada con folio Nº " & Resultado!Folio, vbOKOnly, "Imposible ingresar"): Conexion.Close: Exit Sub
    frmEquivalenciaIngresarActa.lblCarrera = "Carrera: " & dtcCarreras
    frmEquivalenciaIngresarActa.lblCursoYMateria = "Materia: " & cbCurso & "º-" & dtcMaterias
    frmEquivalenciaIngresarActa.lblProfesor = "Profesor: " & dtcProfesor
    Set Resultado = Conexion.Execute("SELECT EquivalenciasResumen.AnoSolicitud, EquivalenciasResumen.Otorgadas, EquivalenciasResumen.Profesor, EquivalenciasResumen.FechaActa FROM EquivalenciasResumen INNER JOIN Personal ON EquivalenciasResumen.Profesor = Personal.Codigo WHERE (((EquivalenciasResumen.Asignatura)=" & dtcMaterias.BoundText & ") AND ((EquivalenciasResumen.AnoSolicitud)=" & txtAño & "))")
    frmEquivalenciaIngresarActa.lblFecha = "Fecha del acta: " & Resultado!FechaActa
    frmEquivalenciaIngresarActa.lblTotal = "Total de Equivalencias: " & Resultado!Otorgadas
    frmEquivalenciaIngresarActa.adoOtorgadas.RecordSource = "SELECT Equivalencias.Alumno, Alumnos.Nombre, Equivalencias.Nota, Equivalencias.FechaAprobacion, Equivalencias.Institucion, Equivalencias.Profesor FROM Alumnos INNER JOIN Equivalencias ON Alumnos.Permiso = Equivalencias.Alumno Where (((Equivalencias.Otorgada) = True) And ((Equivalencias.MateriaAReconocer) = " & dtcMaterias.BoundText & ") And ((Equivalencias.AnoSolicitud) = " & txtAño & "))ORDER BY Alumnos.Nombre"
    frmEquivalenciaIngresarActa.adoOtorgadas.Refresh
    Conexion.Close
    frmEquivalenciaIngresarActa.Show 1
End Sub

Private Sub cmdInstituciones_Click()
    frmInstituciones.Show 1
    adoInstituciones.Refresh
End Sub

Private Sub cmdModificar_Click()
 '   On Error GoTo RefErr
    Conexion.Open
    'verifico que el acta no haya sido impresa
    Set Resultado = Conexion.Execute("SELECT Impresa FROM EquivalenciasResumen WHERE Asignatura = " & dtcMaterias.BoundText & " AND AnoSolicitud = " & txtAño & "")
    If Resultado.EOF = False Then
        If Resultado!Impresa = True Then Respuesta = MsgBox("El acta de equivalencias para esta asignatura ya se imprimió", vbOKOnly, "Imposible realizar modificaciones"): Conexion.Close: Exit Sub
    End If
    Conexion.Close
    If frmIdentificacion.dtcUsuarios.BoundText <> adoEquivalencias.Recordset!Usuario Then
        Respuesta = MsgBox("Solo tiene permiso el usuario que inicio el trámite", vbOKOnly, "Imposible realizar modificaciones"): Exit Sub
    End If
    AlumnoActual = adoEquivalencias.Recordset!Permiso
    cmdAgregar.Enabled = False
    cmdModificar.Enabled = False
    cmdGuardar.Enabled = True
    cmdCancelar.Enabled = True
    frCarrera.Enabled = False
    frMaterias.Enabled = False
    HabilitarControles
    Exit Sub
    
'RefErr:
'    If Err.Number = 5 Then
'        Resume Next
'    Else
'        MsgBox Err.Description
'    End If
End Sub

Private Sub cmdMostrar_Click()
    cmdMostrar.Enabled = False
    adoMaterias.RecordSource = "SELECT Materias.Codigo, Materias.Nombre FROM Materias Where Materias.Curso = " & cbCurso.Text & " And Materias.Carrera = " & dtcCarreras.BoundText & " and (Materias.Detalle = 1 or Materias.Detalle = 3 or Materias.Detalle = 5) ORDER BY Materias.Curso"
    adoMaterias.Refresh
    If adoMaterias.Recordset.RecordCount > 0 Then
        dtcMaterias.BoundText = adoMaterias.Recordset!Codigo
        frMaterias.Enabled = True
    Else
        frMaterias.Enabled = False
        MsgBox ("Este curso no tiene materias con condicion de cursada")
    End If
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub dtcCarreras_Change()
    NuevasMaterias
    adoCarreras.Recordset.MoveFirst
    adoCarreras.Recordset.Find ("Codigo=" & dtcCarreras.BoundText)
    lblCurso = adoCarreras.Recordset!Medida
    cbCurso.Clear
    For i = 0 To adoCarreras.Recordset!Años - 1
        cbCurso.List(i) = i + 1
    Next i
    cbCurso.Text = cbCurso.List(0)
End Sub

Private Sub dtcMaterias_Change()
    VerEquivalencias
End Sub

Private Sub dtgEquivalencias_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    MostrarEquivalencias
End Sub


Private Sub Form_Load()
    Conexion.ConnectionString = ("DSN=Instituto")
    dtcCarreras.BoundText = adoCarreras.Recordset!Codigo
    txtAño = Format(Date, "yyyy")
    FechaDeLaMesa = Format("08/11/1978", "dd/mm/yyyy")
    letras = "abcde"
End Sub

Private Sub txtAño_Change()
    NuevasMaterias
End Sub

Private Sub UpDown1_DownClick()
    txtAño = Val(txtAño) - 1
End Sub

Private Sub UpDown1_UpClick()
    txtAño = Val(txtAño) + 1
End Sub
Private Function NuevasMaterias()
    frMaterias.Enabled = False
    cmdMostrar.Enabled = True
End Function

Private Sub VerEquivalencias()
    adoEquivalencias.RecordSource = "SELECT [Alumnos].[Permiso], [Alumnos].[Nombre], [Equivalencias].[Asignatura], [Equivalencias].[Institucion], [Instituciones].[Institucion] AS NombreInstitucion, [Usuarios].[Usuario], [Usuarios].[Identificacion], [Equivalencias].[Nota], [Equivalencias].[FechaAprobacion], [Equivalencias].[FechaSolicitud], [Equivalencias].[Profesor],[Equivalencias].[Otorgada], [Equivalencias].[Fundamentacion], [Equivalencias].[FechaResolucion], [Equivalencias].[FechaCierre], [Equivalencias].[Libro], [Equivalencias].[Folio], [Equivalencias].[MateriaAReconocer], [Equivalencias].[AnoSolicitud], [Equivalencias].[Observacion] FROM ((((Equivalencias INNER JOIN Instituciones ON [Equivalencias].[Institucion]=[Instituciones].[Codigo]) INNER JOIN Usuarios ON [Equivalencias].[Usuario]=[Usuarios].[Usuario]) INNER JOIN Personal ON [Equivalencias].[Profesor]=[Personal].[Codigo]) INNER JOIN Materias ON [Equivalencias].[MateriaAReconocer]=[Materias].[Codigo]) " _
    & " INNER JOIN Alumnos ON [Equivalencias].[Alumno]=[Alumnos].[Permiso] Where ((([Equivalencias].[MateriaAReconocer]) = " & dtcMaterias.BoundText & ") And (([Equivalencias].[AnoSolicitud]) = " & txtAño & ")) ORDER BY [Alumnos].[Nombre]"
    adoEquivalencias.Refresh
    If adoEquivalencias.Recordset.RecordCount > 0 Then
        cmdAgregar.Enabled = True
        cmdModificar.Enabled = True
        cmdEliminarEquivalencia.Enabled = True
        cmdImprimir.Enabled = True
        If frmIdentificacion.Permisos!CerrarEquivalencias = True Then
            cmdImprimir.Enabled = True
            cmdIngresarActa.Enabled = True
        Else
            cmdImprimir.Enabled = False
            cmdIngresarActa.Enabled = False
        End If
    Else
        cmdAgregar.Enabled = True
        cmdModificar.Enabled = False
        cmdEliminarEquivalencia.Enabled = False
        cmdImprimir.Enabled = False
        cmdIngresarActa.Enabled = False
    End If
End Sub

Private Sub MostrarEquivalencias()
    With adoEquivalencias.Recordset
        dtpFechaInicio = !FechaSolicitud
        txtUsuario = !Identificacion
        txtNota = Replace(Format(!Nota, "0.00"), ",", ".")
        dtpFechaAcreditacion = !FechaAprobacion
        dtcInstituciones.BoundText = !Institucion
        dtcProfesor.BoundText = !Profesor
        If !Otorgada = True Then
            chkOtorgada = 1
        Else
            chkOtorgada = 0
        End If
        If IsNull(!Libro) Then
            txtLibro = 0
        Else
            txtLibro = !Libro
        End If
        txtFolio = !Folio
        If !FechaResolucion <> Nulo Then
            dtpFechaResolucion = !FechaResolucion
        End If
        If !Fundamentacion <> Nulo Then
            txtFundamentacion = !Fundamentacion
        Else
            txtFundamentacion = ""
        End If
        If !Observacion <> Nulo Then
            txtObservacion = !Observacion
        Else
            txtObservacion = ""
        End If
    End With
End Sub

Private Sub HabilitarControles()
    dtpFechaInicio.Enabled = True
    txtNota.Enabled = True
    dtcInstituciones.Enabled = True
    dtpFechaAcreditacion.Enabled = True
    chkOtorgada.Enabled = True
    dtpFechaResolucion.Enabled = True
    dtcProfesor.Enabled = True
    txtFundamentacion.Enabled = True
    txtObservacion.Enabled = True
End Sub


Private Sub DeshabilitarControles()
    dtpFechaInicio.Enabled = False
    txtNota.Enabled = False
    dtcInstituciones.Enabled = False
    dtpFechaAcreditacion.Enabled = False
    chkOtorgada.Enabled = False
    dtpFechaResolucion.Enabled = False
    dtcProfesor.Enabled = False
    txtFundamentacion.Enabled = False
    txtObservacion.Enabled = False
End Sub

Private Sub ArmaEncabezado()
With frmImprimeActas
    .lblTituloDeActa = "Acta de Equivalencias"
    .lblCarrera = dtcCarreras.Text
    .lblMateria = dtcMaterias.BoundText & " " & dtcMaterias.Text
    .lblMesa = NumeroDeMesa
    .lblActa = i + 1
    .lblFecha = FechaDeLaMesa
    .lblHora = "18:00"
    .lblCurso = cbCurso.Text
    .lblCursada = txtAño
    .lblTitular = NombreProfesor
    .lblIntegrante1 = "Directivo"
    .lblIntegrante2 = "Secretario"
End With
End Sub
