VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmMesasArmado 
   Caption         =   "Armado de Mesas de Exámenes Finales"
   ClientHeight    =   10590
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   14880
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   10590
   ScaleWidth      =   14880
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame frMesas 
      Caption         =   "Mesas"
      Height          =   9255
      Left            =   0
      TabIndex        =   11
      Top             =   1200
      Width           =   14775
      Begin VB.CommandButton cmdSuplentes 
         Caption         =   "Suplentes"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   13560
         TabIndex        =   44
         Top             =   600
         Width           =   1095
      End
      Begin MSAdodcLib.Adodc adoInfoMesas 
         Height          =   330
         Left            =   1080
         Top             =   960
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
         RecordSource    =   $"frmMesasArmado.frx":0000
         Caption         =   "InfoMesas"
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
      Begin VB.Frame frDatos 
         Caption         =   "Datos"
         Enabled         =   0   'False
         Height          =   2175
         Left            =   120
         TabIndex        =   20
         Top             =   6960
         Width           =   13335
         Begin VB.TextBox txtLimiteInscriptos 
            Enabled         =   0   'False
            Height          =   315
            Left            =   6000
            TabIndex        =   46
            Top             =   1680
            Width           =   615
         End
         Begin VB.TextBox Text1 
            DataField       =   "Numero"
            DataSource      =   "adoMesa"
            Height          =   375
            Left            =   3360
            TabIndex        =   43
            Text            =   "Text1"
            Top             =   240
            Visible         =   0   'False
            Width           =   375
         End
         Begin VB.TextBox txtMinutos 
            Enabled         =   0   'False
            Height          =   315
            Left            =   12360
            Locked          =   -1  'True
            TabIndex        =   41
            Text            =   "00"
            Top             =   600
            Width           =   375
         End
         Begin MSComCtl2.UpDown upHora 
            Height          =   375
            Left            =   12000
            TabIndex        =   39
            Top             =   570
            Width           =   240
            _ExtentX        =   450
            _ExtentY        =   661
            _Version        =   393216
            Max             =   24
            Min             =   -1
            Enabled         =   -1  'True
         End
         Begin VB.TextBox txtLugar 
            Enabled         =   0   'False
            Height          =   315
            Left            =   1560
            MaxLength       =   30
            TabIndex        =   38
            Top             =   1650
            Width           =   2895
         End
         Begin VB.TextBox txtHora 
            Enabled         =   0   'False
            Height          =   315
            Left            =   11640
            Locked          =   -1  'True
            TabIndex        =   36
            Text            =   "00"
            Top             =   600
            Width           =   375
         End
         Begin MSAdodcLib.Adodc adoMesa 
            Height          =   330
            Left            =   1320
            Top             =   240
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
            RecordSource    =   "SELECT * FROM Mesas Where Numero = 0"
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
         Begin MSAdodcLib.Adodc adoPersonal 
            Height          =   330
            Left            =   7560
            Top             =   120
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
            RecordSource    =   "SELECT Codigo, Nombre FROM Personal WHERE Eliminado = 0 AND TrabajaActualmente = 1 ORDER BY Nombre"
            Caption         =   "Personal"
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
         Begin MSDataListLib.DataCombo dtcTitular 
            Bindings        =   "frmMesasArmado.frx":0171
            Height          =   315
            Left            =   120
            TabIndex        =   33
            Top             =   1200
            Width           =   4215
            _ExtentX        =   7435
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            MatchEntry      =   -1  'True
            Style           =   2
            ListField       =   "Nombre"
            BoundColumn     =   "Codigo"
            Text            =   ""
         End
         Begin VB.TextBox txtMesa 
            Enabled         =   0   'False
            Height          =   315
            Left            =   8160
            TabIndex        =   32
            Top             =   1680
            Width           =   615
         End
         Begin MSComCtl2.DTPicker dtpFecha 
            Height          =   315
            Left            =   9960
            TabIndex        =   31
            Top             =   600
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            Format          =   123207681
            CurrentDate     =   37846
         End
         Begin VB.TextBox txtDivision 
            Enabled         =   0   'False
            Height          =   315
            Left            =   9240
            TabIndex        =   30
            Top             =   600
            Width           =   495
         End
         Begin VB.TextBox txtMateria 
            Enabled         =   0   'False
            Height          =   315
            Left            =   120
            TabIndex        =   29
            Top             =   600
            Width           =   9015
         End
         Begin MSDataListLib.DataCombo dtcIntegrante2 
            Bindings        =   "frmMesasArmado.frx":018B
            Height          =   315
            Left            =   8880
            TabIndex        =   34
            Top             =   1200
            Width           =   4215
            _ExtentX        =   7435
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            MatchEntry      =   -1  'True
            Style           =   2
            ListField       =   "Nombre"
            BoundColumn     =   "Codigo"
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo dtcIntegrante1 
            Bindings        =   "frmMesasArmado.frx":01A5
            Height          =   315
            Left            =   4560
            TabIndex        =   35
            Top             =   1200
            Width           =   3975
            _ExtentX        =   7011
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            MatchEntry      =   -1  'True
            Style           =   2
            ListField       =   "Nombre"
            BoundColumn     =   "Codigo"
            Text            =   ""
         End
         Begin MSComCtl2.UpDown upMinutos 
            Height          =   375
            Left            =   12720
            TabIndex        =   42
            Top             =   570
            Width           =   240
            _ExtentX        =   450
            _ExtentY        =   661
            _Version        =   393216
            Increment       =   10
            Max             =   60
            Min             =   -10
            Enabled         =   -1  'True
         End
         Begin MSAdodcLib.Adodc adoControlSuperposicion 
            Height          =   330
            Left            =   4320
            Top             =   240
            Visible         =   0   'False
            Width           =   2895
            _ExtentX        =   5106
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
            RecordSource    =   $"frmMesasArmado.frx":01BF
            Caption         =   "ControlSuperposicion"
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
         Begin VB.Label Label14 
            Caption         =   "Límite de Inscriptos:"
            Height          =   255
            Left            =   4560
            TabIndex        =   47
            Top             =   1680
            Width           =   1575
         End
         Begin VB.Label Label13 
            Caption         =   "Minutos:"
            Height          =   255
            Left            =   12360
            TabIndex        =   40
            Top             =   360
            Width           =   615
         End
         Begin VB.Label Label12 
            Caption         =   "Lugar del exámen:"
            Height          =   255
            Left            =   120
            TabIndex        =   37
            Top             =   1680
            Width           =   1335
         End
         Begin VB.Label Label11 
            Caption         =   "Div:"
            Height          =   255
            Left            =   9240
            TabIndex        =   28
            Top             =   360
            Width           =   495
         End
         Begin VB.Label Label10 
            Caption         =   "Fecha:"
            Height          =   255
            Left            =   9960
            TabIndex        =   27
            Top             =   360
            Width           =   615
         End
         Begin VB.Label Label9 
            Caption         =   "Hora:"
            Height          =   255
            Left            =   11640
            TabIndex        =   26
            Top             =   360
            Width           =   615
         End
         Begin VB.Label Label8 
            Caption         =   "Nº Mesa:"
            Height          =   255
            Left            =   7440
            TabIndex        =   25
            Top             =   1680
            Width           =   735
         End
         Begin VB.Label Label7 
            Caption         =   "Titular:"
            Height          =   255
            Left            =   120
            TabIndex        =   24
            Top             =   960
            Width           =   615
         End
         Begin VB.Label Label6 
            Caption         =   "1º Integrante:"
            Height          =   255
            Left            =   4560
            TabIndex        =   23
            Top             =   960
            Width           =   1095
         End
         Begin VB.Label Label5 
            Caption         =   "2º Integrante:"
            Height          =   255
            Left            =   8880
            TabIndex        =   22
            Top             =   960
            Width           =   1095
         End
         Begin VB.Label Label2 
            Caption         =   "Materia:"
            Height          =   255
            Left            =   120
            TabIndex        =   21
            Top             =   360
            Width           =   735
         End
      End
      Begin VB.Frame frBotones 
         BackColor       =   &H8000000C&
         BorderStyle     =   0  'None
         Height          =   5295
         Left            =   13680
         TabIndex        =   13
         Top             =   2160
         Width           =   855
         Begin VB.CommandButton cmdSalir 
            Height          =   600
            Left            =   120
            Picture         =   "frmMesasArmado.frx":0295
            Style           =   1  'Graphical
            TabIndex        =   19
            ToolTipText     =   "Salir"
            Top             =   4440
            Width           =   600
         End
         Begin VB.CommandButton cmdAgregar 
            Enabled         =   0   'False
            Height          =   600
            Left            =   120
            Picture         =   "frmMesasArmado.frx":06D7
            Style           =   1  'Graphical
            TabIndex        =   18
            ToolTipText     =   "Agregar"
            Top             =   240
            Width           =   600
         End
         Begin VB.CommandButton cmdModificar 
            Enabled         =   0   'False
            Height          =   600
            Left            =   120
            Picture         =   "frmMesasArmado.frx":0B19
            Style           =   1  'Graphical
            TabIndex        =   17
            ToolTipText     =   "Modificar"
            Top             =   1080
            Width           =   600
         End
         Begin VB.CommandButton cmdGuardar 
            Enabled         =   0   'False
            Height          =   600
            Left            =   120
            Picture         =   "frmMesasArmado.frx":0F5B
            Style           =   1  'Graphical
            TabIndex        =   16
            ToolTipText     =   "Guardar"
            Top             =   1920
            Width           =   600
         End
         Begin VB.CommandButton cmdEliminar 
            Enabled         =   0   'False
            Height          =   600
            Left            =   120
            Picture         =   "frmMesasArmado.frx":139D
            Style           =   1  'Graphical
            TabIndex        =   15
            ToolTipText     =   "Borrar"
            Top             =   2760
            Width           =   600
         End
         Begin VB.CommandButton cmdCancelar 
            Enabled         =   0   'False
            Height          =   600
            Left            =   120
            Picture         =   "frmMesasArmado.frx":17DF
            Style           =   1  'Graphical
            TabIndex        =   14
            ToolTipText     =   "Cancelar"
            Top             =   3600
            Width           =   600
         End
      End
      Begin MSDataGridLib.DataGrid dtgInfoMesas 
         Bindings        =   "frmMesasArmado.frx":1C21
         Height          =   6615
         Left            =   120
         TabIndex        =   12
         Top             =   240
         Width           =   13335
         _ExtentX        =   23521
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
         ColumnCount     =   6
         BeginProperty Column00 
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
         BeginProperty Column01 
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
         BeginProperty Column02 
            DataField       =   "Division"
            Caption         =   "Division"
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
            DataField       =   "Fecha"
            Caption         =   "Fecha"
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
            DataField       =   "Hora"
            Caption         =   "Hora"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "H:mm"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   11274
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column05 
            DataField       =   "Nombre"
            Caption         =   "Titular"
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
               ColumnWidth     =   884,976
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   7395,024
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   734,74
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   1049,953
            EndProperty
            BeginProperty Column04 
               ColumnWidth     =   675,213
            EndProperty
            BeginProperty Column05 
               ColumnWidth     =   1860,095
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame frTurno 
      Height          =   1095
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   14775
      Begin VB.TextBox Text3 
         DataField       =   "TurnoLlamado"
         DataSource      =   "adoParametros"
         Height          =   375
         Left            =   2880
         TabIndex        =   45
         Text            =   "Text3"
         Top             =   720
         Visible         =   0   'False
         Width           =   375
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
         Left            =   13560
         TabIndex        =   10
         Top             =   360
         Width           =   1095
      End
      Begin VB.ComboBox cbCurso 
         Height          =   315
         Left            =   12840
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   480
         Width           =   615
      End
      Begin MSAdodcLib.Adodc adoCarreras 
         Height          =   330
         Left            =   3720
         Top             =   600
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
         RecordSource    =   $"frmMesasArmado.frx":1C3C
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
         Bindings        =   "frmMesasArmado.frx":1D40
         Height          =   315
         Left            =   2280
         TabIndex        =   8
         Top             =   480
         Width           =   10455
         _ExtentX        =   18441
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "Nombre"
         BoundColumn     =   "Codigo"
         Text            =   ""
      End
      Begin VB.TextBox txtAño 
         Height          =   315
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   480
         Width           =   495
      End
      Begin MSDataListLib.DataCombo dtcMeses 
         Bindings        =   "frmMesasArmado.frx":1D5A
         Height          =   315
         Left            =   120
         TabIndex        =   5
         Top             =   480
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "Nombre"
         BoundColumn     =   "Numero"
         Text            =   ""
      End
      Begin MSAdodcLib.Adodc adoMeses 
         Height          =   330
         Left            =   0
         Top             =   720
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
         RecordSource    =   "Meses"
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
      Begin MSComCtl2.UpDown UpDown1 
         Height          =   375
         Left            =   1920
         TabIndex        =   7
         Top             =   450
         Width           =   240
         _ExtentX        =   450
         _ExtentY        =   661
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSAdodcLib.Adodc adoParametros 
         Height          =   330
         Left            =   960
         Top             =   720
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
      Begin VB.Label Label1 
         Caption         =   "Turno:"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label4 
         Caption         =   "Año:"
         Height          =   255
         Left            =   1440
         TabIndex        =   4
         Top             =   240
         Width           =   495
      End
      Begin VB.Label Label3 
         Caption         =   "Carreras Vigentes:"
         Height          =   255
         Left            =   2280
         TabIndex        =   3
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label lblCurso 
         Caption         =   "Curso:"
         Height          =   255
         Left            =   12840
         TabIndex        =   2
         Top             =   240
         Width           =   1095
      End
   End
End
Attribute VB_Name = "frmMesasArmado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RegistroActual As String
Dim Estado As String
Public MateriaAgregar As Double
Dim Profesor As String
Dim ProfesorNombre As String
Public Mostrar As String
Dim Conexion As New Connection
Dim Resultado As New Recordset

Private Sub cbCurso_Click()
    cmdAgregar.Enabled = False
    dtgInfoMesas.Enabled = False
    cmdMostrar.Enabled = True
End Sub

Private Sub cmdAgregar_Click()
    Estado = "Agregando"
    cmdAgregar.Enabled = False
    cmdModificar.Enabled = False
    cmdGuardar.Enabled = True
    cmdEliminar.Enabled = False
    cmdCancelar.Enabled = True
    cmdSalir.Enabled = False
    cmdSuplentes.Enabled = False
    frTurno.Enabled = False
    frDatos.Enabled = True
    dtgInfoMesas.Enabled = False
    LimpiarMesa
    txtLugar = "Sede"
    HabilitarMesa
    With frmAgregarMesa
    .lblTurno = dtcMeses.Text
    .lblAño = txtAño
    .lblCarrera = dtcCarreras.Text
    .lblCurso = lblCurso & ": " & cbCurso
    .adoMaterias.RecordSource = "SELECT Materias.Codigo, Materias.Nombre From Materias Where Materias.Curso = " & cbCurso.Text & " And Materias.Carrera = " & dtcCarreras.BoundText & " And (Materias.Detalle = 1 Or Materias.Detalle = 3 Or Materias.Detalle = 5) AND Eliminada = 0 ORDER BY Materias.Codigo"
    .adoMaterias.Refresh
    .Show 1
    End With
    txtDivision = 1
End Sub

Public Sub cmdCancelar_Click()
    cmdAgregar.Enabled = True
    cmdModificar.Enabled = True
    cmdGuardar.Enabled = False
    cmdEliminar.Enabled = False
    cmdCancelar.Enabled = False
    cmdSalir.Enabled = True
    cmdSuplentes.Enabled = True
    DeshabilitarMesa
    frTurno.Enabled = True
    frDatos.Enabled = True
    dtgInfoMesas.Enabled = True
    Mostrar = "Si"
    MostrarMesa
End Sub

Private Sub cmdEliminar_Click()
    If adoMesa.Recordset!Impresas = True Then Respuesta = MsgBox("Las actas ya fueron impresas", 0, "Imposible eliminar"): Exit Sub
    Conexion.Open
    Set Resultado = Conexion.Execute("SELECT * FROM Inscripciones WHERE Mesa=" & adoMesa.Recordset!Numero)
    If Resultado.EOF = False Then
        Respuesta = MsgBox("Existen alumnos inscriptos en la mesa de " & adoInfoMesas.Recordset!Abreviatura & Chr(13) & "¿Desea quitar de todas formas la mesa y las incripciones correspondientes?", vbYesNo, "Atención")
        If Respuesta = vbYes Then
            Conexion.Execute ("DELETE * FROM Mesas WHERE Numero =" & adoInfoMesas.Recordset!Numero)
            Conexion.Execute ("DELETE * FROM Inscripciones WHERE Mesa =" & adoInfoMesas.Recordset!Numero)
        End If
    Else
        Respuesta = MsgBox("¿Está seguro de quitar la mesa de " & adoInfoMesas.Recordset!Abreviatura & "?", vbYesNo, "Atención")
        If Respuesta = vbYes Then
            Conexion.Execute ("DELETE * FROM Mesas WHERE Numero =" & adoInfoMesas.Recordset!Numero)
        End If
    End If
    Conexion.Close
    adoInfoMesas.Refresh
    If adoInfoMesas.Recordset.RecordCount = 0 Then cmdEliminar.Enabled = False: cmdModificar.Enabled = False
End Sub

Private Sub cmdGuardar_Click()
    If txtDivision < 1 Or txtDivision > 10 Then
        MsgBox ("Error de división")
        Exit Sub
    End If
    If (Not IsNumeric(txtLimiteInscriptos.Text)) Then
        MsgBox ("El límite de inscriptos no es correcto")
        Exit Sub
    End If
        
    If dtcTitular.Text = "" Or dtcIntegrante1.Text = "" Or dtcIntegrante2.Text = "" Then MsgBox ("Debe completar los tres integrantes de la mesa"): Exit Sub
    'controlo que la mesa en esa division no este creada
    'If Estado = "Agregando" Then
    '    Conexion.Open
    '    Set Resultado = Conexion.Execute("SELECT Mesas.Materia, Mesas.Division, Mesas.Turno, Mesas.Ano From Mesas WHERE (((Mesas.Materia)=" & MateriaAgregar & ") AND ((Mesas.Division)=" & txtDivision & ") AND ((Mesas.Turno)=" & dtcMeses.BoundText & ") AND ((Mesas.Ano)=" & txtAño & "))")
    '    If Resultado.EOF = False Then
    '        MsgBox ("La mesa para esta división ya fue armada"): Conexion.Close: Exit Sub
    '    End If
    '    Conexion.Close
    'Else
    '    Conexion.Open
    '    Set Resultado = Conexion.Execute("SELECT Mesas.Numero, Mesas.Materia, Mesas.Division, Mesas.Turno, Mesas.Ano From Mesas WHERE (((Mesas.Materia)=" & adoInfoMesas.Recordset!Codigo & ") AND ((Mesas.Division)=" & txtDivision & ") AND ((Mesas.Turno)=" & dtcMeses.BoundText & ") AND ((Mesas.Ano)=" & txtAño & "))")
    '    If Resultado.EOF = False Then
    '        If Resultado.EOF = False And Resultado!Numero <> adoInfoMesas.Recordset!Numero Then
    '            MsgBox ("La mesa para esta división ya fue armada"): Conexion.Close: Exit Sub
    '        End If
    '    End If
    '    Conexion.Close
    'End If
    cmdAgregar.Enabled = True
    cmdModificar.Enabled = True
    cmdGuardar.Enabled = False
    cmdEliminar.Enabled = False
    cmdCancelar.Enabled = False
    cmdSalir.Enabled = True
    cmdSuplentes.Enabled = True
    frTurno.Enabled = True
    frDatos.Enabled = True
    dtgInfoMesas.Enabled = True
    Mostrar = "Si"
    PasarMesa
    Mostrar = "No"
    dtgInfoMesas.SetFocus
End Sub

Private Sub cmdModificar_Click()
    If adoMesa.Recordset!Impresas = True Then Respuesta = MsgBox("Las actas ya fueron impresas", 0, "Imposible modificar"): Exit Sub
    cmdAgregar.Enabled = False
    cmdModificar.Enabled = False
    cmdGuardar.Enabled = True
    cmdEliminar.Enabled = False
    cmdCancelar.Enabled = True
    cmdSalir.Enabled = False
    cmdSuplentes.Enabled = False
    frTurno.Enabled = False
    frDatos.Enabled = True
    dtgInfoMesas.Enabled = False
    HabilitarMesa
    Mostrar = "No" 'para que haga el control de superposicion
End Sub

Private Sub cmdMostrar_Click()
    adoInfoMesas.RecordSource = "SELECT Materias.Codigo, Materias.Abreviatura, Mesas.Numero, Mesas.Division, Mesas.Fecha, Mesas.Hora, Mesas.Impresas, Personal.Nombre FROM (Mesas INNER JOIN Materias ON Mesas.Materia = Materias.Codigo) INNER JOIN Personal ON Mesas.Titular = Personal.Codigo WHERE Mesas.Turno = " & dtcMeses.BoundText & " AND Mesas.Ano = " & txtAño & " AND Materias.Carrera = " & dtcCarreras.BoundText & " AND Materias.Curso = " & cbCurso.Text & " ORDER BY Materias.Codigo"
    adoInfoMesas.Refresh
    Mostrar = "Si"
    MostrarMesa
    Mostrar = "No"
    cmdMostrar.Enabled = False
    dtgInfoMesas.Enabled = True
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub cmdSuplentes_Click()
    With frmSuplentes
    .lblTurno = dtcMeses.Text
    .lblAño = txtAño
    End With
    If frmIdentificacion.Permisos!ModificarSuplentes = False Then
        frmSuplentes.cmdAgregar.Enabled = False
        frmSuplentes.cmdEliminar.Enabled = False
    End If
    frmSuplentes.Show 1
End Sub

Private Sub dtcCarreras_Change()
    adoCarreras.Recordset.MoveFirst
    adoCarreras.Recordset.Find ("Codigo=" & dtcCarreras.BoundText)
    lblCurso = adoCarreras.Recordset!Medida
    cbCurso.Clear
    For i = 0 To adoCarreras.Recordset!Años - 1
        cbCurso.List(i) = i + 1
    Next i
    cbCurso.Text = cbCurso.List(0)
    cmdAgregar.Enabled = False
    dtgInfoMesas.Enabled = False
    cmdMostrar.Enabled = True
End Sub

Private Sub dtcIntegrante1_Change()
If dtcIntegrante1.Text <> "" Then
    Profesor = dtcIntegrante1.BoundText
    ProfesorNombre = dtcIntegrante1.Text
    ControlSuperposicion
End If
End Sub

Private Sub dtcIntegrante2_Change()
If dtcIntegrante2.Text <> "" Then
    Profesor = dtcIntegrante2.BoundText
    ProfesorNombre = dtcIntegrante2.Text
    ControlSuperposicion
End If
End Sub

Private Sub dtcMeses_Change()
    cmdAgregar.Enabled = False
    dtgInfoMesas.Enabled = False
    cmdMostrar.Enabled = True
End Sub

Private Sub dtcTitular_Change()
If dtcTitular.Text <> "" Then
    Profesor = dtcTitular.BoundText
    ProfesorNombre = dtcTitular.Text
    ControlSuperposicion
End If
End Sub

Private Sub dtgInfoMesas_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    Mostrar = "Si"
    MostrarMesa
    Mostrar = "No"
End Sub

Private Sub Form_Load()
    dtpFecha.Value = Date
    'If dtpFecha.Month > 8 And dtpFecha.Month < 12 Then
    '    dtcMeses.BoundText = 12
    '    txtAño = Format(Date, "yyyy")
    'ElseIf dtpFecha.Month > 3 And dtpFecha.Month <= 8 Then
    '    dtcMeses.BoundText = 8
    '    txtAño = Format(Date, "yyyy")
    'ElseIf dtpFecha.Month = 12 Then
    '    dtcMeses.BoundText = 3
    '    txtAño = Format(Date, "yyyy") + 1
    'Else
    '    dtcMeses.BoundText = 3
    '    txtAño = Format(Date, "yyyy")
    'End If
    dtcMeses.BoundText = adoParametros.Recordset!TurnoLlamado
    txtAño = adoParametros.Recordset!AñoLlamado
    dtcCarreras.BoundText = adoCarreras.Recordset!Codigo
    cmdMostrar_Click
    Conexion.ConnectionString = ("DSN=Instituto")
End Sub

Private Sub txtAño_Change()
    cmdAgregar.Enabled = False
    dtgInfoMesas.Enabled = False
    cmdMostrar.Enabled = True
End Sub

Private Sub txtDivision_Validate(Cancel As Boolean)
    If txtDivision = "" Then
        MsgBox ("Desde especificarse una divisíon")
        txtDivision.SetFocus
    End If
End Sub

Private Sub UpDown1_DownClick()
    txtAño = Val(txtAño) - 1
End Sub

Private Sub UpDown1_UpClick()
    txtAño = Val(txtAño) + 1
End Sub

Private Function HabilitarMesa()
    txtDivision.Enabled = True
    dtpFecha.Enabled = True
    txtHora.Enabled = True
    txtMinutos.Enabled = True
    dtcTitular.Enabled = True
    dtcIntegrante1.Enabled = True
    dtcIntegrante2.Enabled = True
    txtLugar.Enabled = True
    txtLimiteInscriptos.Enabled = True
End Function
Private Function DeshabilitarMesa()
    txtDivision.Enabled = False
    dtpFecha.Enabled = False
    txtHora.Enabled = False
    txtMinutos.Enabled = False
    dtcTitular.Enabled = False
    dtcIntegrante1.Enabled = False
    dtcIntegrante2.Enabled = False
    txtLugar.Enabled = False
    txtLimiteInscriptos.Enabled = False
End Function
Private Function LimpiarMesa()
    txtMateria = ""
    txtDivision = ""
    dtpFecha = Date
    upHora = 0
    upMinutos = 0
    dtcTitular = ""
    dtcIntegrante1 = ""
    dtcIntegrante2 = ""
    txtLugar = ""
    txtMesa = ""
    txtLimiteInscriptos = 999
End Function
Private Function MostrarMesa()
    If adoInfoMesas.Recordset.RecordCount > 0 Then
        adoMesa.RecordSource = "SELECT * FROM Mesas WHERE Numero = " & adoInfoMesas.Recordset!Numero
        adoMesa.Refresh
        txtMateria = adoInfoMesas.Recordset!Abreviatura
        With adoMesa.Recordset
        txtDivision = !division
        dtpFecha.Value = !Fecha
        upHora.Value = Format(!Hora, "hh")
        upMinutos.Value = Mid(Format(!Hora, "hh:mm"), Len(Format(!Hora, "hh:mm")) - 1, 2)
        txtMesa = !Numero
        If !Lugar <> "" Then
            txtLugar = !Lugar
        Else
            txtLugar = ""
        End If
        dtcTitular.BoundText = !Titular
        dtcIntegrante1.BoundText = !Integrante1
        dtcIntegrante2.BoundText = !Integrante2
        txtLimiteInscriptos.Text = !LimiteInscriptos
        End With
        If frmIdentificacion.Permisos!ModificarMesas = True Then
            cmdAgregar.Enabled = True
            cmdModificar.Enabled = True
            cmdEliminar.Enabled = True
        Else
            cmdAgregar.Enabled = False
            cmdModificar.Enabled = False
            cmdEliminar.Enabled = False
        End If
    Else
        If frmIdentificacion.Permisos!ModificarMesas = True Then
            cmdAgregar.Enabled = True
        Else
            cmdAgregar.Enabled = False
        End If
        LimpiarMesa
        cmdModificar.Enabled = False
        cmdEliminar.Enabled = False
    End If
End Function
Private Function PasarMesa()
    With adoMesa.Recordset
    horajuntada = Format(txtHora & ":" & txtMinutos, "hh:mm:ss")
    If Estado = "Agregando" Then
        Conexion.Open
        Conexion.Execute ("INSERT INTO Mesas ( Materia, Division, Turno, Ano, Fecha, Hora, Lugar, Titular, Integrante1, Integrante2, LimiteInscriptos ) VALUES (" & MateriaAgregar & "," & txtDivision & "," & dtcMeses.BoundText & "," & txtAño & ",'" & DateValue(dtpFecha.Value) & "', '" & TimeValue(horajuntada) & "','" & txtLugar & "'," & dtcTitular.BoundText & "," & dtcIntegrante1.BoundText & "," & dtcIntegrante2.BoundText & "," & txtLimiteInscriptos & ")")
        Conexion.Close
    Else
        Conexion.Open
        Conexion.Execute ("UPDATE Mesas SET Division = " & txtDivision & ",Fecha = '" & DateValue(dtpFecha.Value) & "', Hora = '" & TimeValue(horajuntada) & "',Lugar = '" & txtLugar & "',Titular = " & dtcTitular.BoundText & ",Integrante1 = " & dtcIntegrante1.BoundText & ",Integrante2 = " & dtcIntegrante2.BoundText & ", LimiteInscriptos=" & txtLimiteInscriptos.Text & " WHERE Numero = " & adoInfoMesas.Recordset!Numero)
        Conexion.Close
    End If
    adoInfoMesas.Refresh
    If Estado <> "Agregando" Then
        RegistroActual = txtMesa
        adoInfoMesas.Recordset.Find ("Numero=" & RegistroActual)
    End If
    Estado = ""
    frDatos.Enabled = False
    End With
    MostrarMesa
End Function

Private Sub upHora_Change()
    If upHora.Value = 24 Then upHora.Value = 0
    If upHora.Value = -1 Then upHora.Value = 23
    txtHora = Format(upHora.Value, "00")
End Sub

Private Sub upMinutos_Change()
    If upMinutos.Value = 60 Then
        upMinutos.Value = 0
        upHora.Value = upHora.Value + 1
    ElseIf upMinutos.Value = -10 Then
        upMinutos.Value = 50
        upHora.Value = upHora.Value - 1
    Else
        txtMinutos = Format(upMinutos.Value, "00")
    End If
End Sub

Private Function ControlSuperposicion()
    If Mostrar = "No" Then
    
    'esta anda en SQL Server
    'adoControlSuperposicion.RecordSource = "SELECT Mesas.Hora, Carreras.Abreviatura AS Carrera, Materias.Abreviatura AS Materia, Materias.Curso FROM Mesas INNER JOIN (Materias INNER JOIN Carreras ON Materias.Carrera = Carreras.Codigo) ON Mesas.Materia = Materias.Codigo Where (((Mesas.Fecha) = '" & DateValue(dtpFecha.Value) & "') And ((Mesas.Titular) = " & Profesor & ")) Or (((Mesas.Fecha) = '" & DateValue(dtpFecha.Value) & "') And ((Mesas.Integrante1) = " & Profesor & ")) Or (((Mesas.Fecha) = '" & DateValue(dtpFecha.Value) & "') And ((Mesas.Integrante2) = " & Profesor & "))ORDER BY Mesas.Fecha, Mesas.Hora"
    
    'esta anda en Acces
    adoControlSuperposicion.RecordSource = "SELECT Mesas.Fecha, Mesas.Numero, Mesas.Titular, Mesas.Hora, Carreras.Abreviatura AS Carrera, Materias.Abreviatura AS Materia, Materias.Curso FROM Mesas INNER JOIN (Materias INNER JOIN Carreras ON Materias.Carrera = Carreras.Codigo) ON Mesas.Materia = Materias.Codigo Where (((Mesas.Fecha) = #" & Format(dtpFecha.Value, "mm/dd/yyyy") & "#) And ((Mesas.Titular) = " & Profesor & ")) Or (((Mesas.Fecha) = #" & Format(dtpFecha.Value, "mm/dd/yyyy") & "#) And ((Mesas.Integrante1) = " & Profesor & ")) Or (((Mesas.Fecha) = #" & Format(dtpFecha.Value, "mm/dd/yyyy") & "#) And ((Mesas.Integrante2) = " & Profesor & "))ORDER BY Mesas.Fecha, Mesas.Hora"
    
    adoControlSuperposicion.Refresh
    With adoControlSuperposicion.Recordset
    If .RecordCount > 0 Then
        StringMesas = ""
        For i = 1 To .RecordCount
            StringMesas = StringMesas & "HORA:" & Format(!Hora, "hh:mm") & " CARRERA: " & !Carrera & " Materia: " & !Curso & "º-" & !Materia & Chr(13) & Chr(13)
            .MoveNext
        Next i
        MsgBox ("El profesor " & ProfesorNombre & " está designado en esta fecha para las siguientes mesas:" & Chr(13) & Chr(13) & StringMesas)
    End If
    End With
    End If
End Function
