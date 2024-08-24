VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmCursadas 
   Caption         =   "Armado de Horarios de Cursadas"
   ClientHeight    =   9600
   ClientLeft      =   135
   ClientTop       =   705
   ClientWidth     =   14880
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   9600
   ScaleWidth      =   14880
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame frCursadas 
      Height          =   7815
      Left            =   0
      TabIndex        =   8
      Top             =   1680
      Width           =   14655
      Begin VB.Frame frMaterias 
         Caption         =   "Materias"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   7455
         Left            =   120
         TabIndex        =   23
         Top             =   120
         Width           =   8535
         Begin MSAdodcLib.Adodc adoMaterias 
            Height          =   330
            Left            =   360
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
            RecordSource    =   $"frmCursadas.frx":0000
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
         Begin MSDataGridLib.DataGrid dtgMaterias 
            Bindings        =   "frmCursadas.frx":0121
            Height          =   6975
            Left            =   120
            TabIndex        =   24
            Top             =   240
            Width           =   8295
            _ExtentX        =   14631
            _ExtentY        =   12303
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
               DataField       =   "Modalidad"
               Caption         =   "Modalidad"
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
                  ColumnWidth     =   840,189
               EndProperty
               BeginProperty Column01 
                  ColumnWidth     =   5850,142
               EndProperty
               BeginProperty Column02 
                  ColumnWidth     =   959,811
               EndProperty
            EndProperty
         End
      End
      Begin VB.Frame frDivisiones 
         Caption         =   "Divisiones"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3735
         Left            =   8760
         TabIndex        =   17
         Top             =   120
         Width           =   5775
         Begin VB.TextBox txtLimiteMatriculados 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            Height          =   315
            Left            =   120
            MaxLength       =   15
            TabIndex        =   52
            Top             =   3000
            Width           =   1455
         End
         Begin VB.CheckBox chkLibre 
            Alignment       =   1  'Right Justify
            Caption         =   "Libre"
            Enabled         =   0   'False
            Height          =   375
            Left            =   2280
            TabIndex        =   50
            Top             =   1080
            Width           =   735
         End
         Begin VB.CommandButton Command7 
            Caption         =   "Ver Plano"
            Height          =   315
            Left            =   2400
            TabIndex        =   34
            Top             =   2280
            Width           =   1335
         End
         Begin MSAdodcLib.Adodc adoDivisiones 
            Height          =   330
            Left            =   2640
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
            RecordSource    =   $"frmCursadas.frx":013B
            Caption         =   "Divisiones"
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
         Begin VB.TextBox txtSalon 
            Enabled         =   0   'False
            Height          =   315
            Left            =   120
            MaxLength       =   15
            TabIndex        =   28
            Top             =   2280
            Width           =   2055
         End
         Begin VB.TextBox txtMateria 
            DataField       =   "Abreviatura"
            DataSource      =   "adoMaterias"
            Enabled         =   0   'False
            Height          =   375
            Left            =   120
            TabIndex        =   26
            Top             =   600
            Width           =   5535
         End
         Begin VB.Frame frBotones 
            BackColor       =   &H8000000C&
            BorderStyle     =   0  'None
            Height          =   1970
            Left            =   4200
            TabIndex        =   18
            Top             =   1080
            Width           =   1395
            Begin VB.CommandButton cmdAgregarDivision 
               Enabled         =   0   'False
               Height          =   550
               Left            =   120
               Picture         =   "frmCursadas.frx":01E1
               Style           =   1  'Graphical
               TabIndex        =   35
               ToolTipText     =   "Agregar"
               Top             =   120
               Width           =   550
            End
            Begin VB.CommandButton cmdModificarDivision 
               Enabled         =   0   'False
               Height          =   550
               Left            =   720
               Picture         =   "frmCursadas.frx":0623
               Style           =   1  'Graphical
               TabIndex        =   22
               ToolTipText     =   "Modificar"
               Top             =   120
               Width           =   550
            End
            Begin VB.CommandButton cmdGuardarDivision 
               Enabled         =   0   'False
               Height          =   550
               Left            =   120
               Picture         =   "frmCursadas.frx":0A65
               Style           =   1  'Graphical
               TabIndex        =   21
               ToolTipText     =   "Guardar"
               Top             =   720
               Width           =   550
            End
            Begin VB.CommandButton cmdEliminarDivision 
               Enabled         =   0   'False
               Height          =   550
               Left            =   720
               Picture         =   "frmCursadas.frx":0EA7
               Style           =   1  'Graphical
               TabIndex        =   20
               ToolTipText     =   "Borrar"
               Top             =   720
               Width           =   550
            End
            Begin VB.CommandButton cmdCancelarDivision 
               Enabled         =   0   'False
               Height          =   550
               Left            =   120
               Picture         =   "frmCursadas.frx":12E9
               Style           =   1  'Graphical
               TabIndex        =   19
               ToolTipText     =   "Cancelar"
               Top             =   1320
               Width           =   550
            End
         End
         Begin MSDataListLib.DataCombo dtcProfesor 
            Bindings        =   "frmCursadas.frx":172B
            Height          =   315
            Left            =   120
            TabIndex        =   27
            Top             =   1680
            Width           =   3855
            _ExtentX        =   6800
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            Style           =   2
            ListField       =   "Nombre"
            BoundColumn     =   "Codigo"
            Text            =   ""
         End
         Begin MSAdodcLib.Adodc adoPersonal 
            Height          =   330
            Left            =   3600
            Top             =   0
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
         Begin MSDataListLib.DataCombo dtcDivision 
            Bindings        =   "frmCursadas.frx":1745
            Height          =   315
            Left            =   1200
            TabIndex        =   33
            Top             =   1080
            Width           =   735
            _ExtentX        =   1296
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            ListField       =   "Division"
            BoundColumn     =   "Division"
            Text            =   ""
         End
         Begin VB.Label Label10 
            Caption         =   "Limite matriculados:"
            Height          =   255
            Left            =   120
            TabIndex        =   51
            Top             =   2760
            Width           =   2055
         End
         Begin VB.Label Label6 
            Caption         =   "Profesor:"
            Height          =   255
            Left            =   120
            TabIndex        =   32
            Top             =   1440
            Width           =   1095
         End
         Begin VB.Label Label5 
            Caption         =   "Salón:"
            Height          =   255
            Left            =   120
            TabIndex        =   31
            Top             =   2040
            Width           =   1095
         End
         Begin VB.Label Label4 
            Caption         =   "Materia:"
            Height          =   255
            Left            =   120
            TabIndex        =   30
            Top             =   360
            Width           =   1095
         End
         Begin VB.Label Label3 
            Caption         =   "Nº de división:"
            Height          =   255
            Left            =   120
            TabIndex        =   29
            Top             =   1080
            Width           =   1095
         End
      End
      Begin VB.Frame frEncuentros 
         Caption         =   "Encuentros"
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
         Height          =   3495
         Left            =   8760
         TabIndex        =   9
         Top             =   4080
         Width           =   5775
         Begin MSAdodcLib.Adodc adoDias 
            Height          =   330
            Left            =   2400
            Top             =   2280
            Visible         =   0   'False
            Width           =   1575
            _ExtentX        =   2778
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
            RecordSource    =   "SELECT * FROM Dias ORDER BY Numero"
            Caption         =   "Dias"
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
         Begin VB.TextBox txtHoraSalida 
            Enabled         =   0   'False
            Height          =   315
            Left            =   2160
            Locked          =   -1  'True
            TabIndex        =   42
            Text            =   "00"
            Top             =   3030
            Width           =   375
         End
         Begin VB.TextBox txtMinutosSalida 
            Enabled         =   0   'False
            Height          =   315
            Left            =   2880
            Locked          =   -1  'True
            TabIndex        =   40
            Text            =   "00"
            Top             =   3030
            Width           =   375
         End
         Begin VB.TextBox txtHoraEntrada 
            Enabled         =   0   'False
            Height          =   315
            Left            =   240
            Locked          =   -1  'True
            TabIndex        =   38
            Text            =   "00"
            Top             =   3030
            Width           =   375
         End
         Begin VB.TextBox txtMinutosEntrada 
            Enabled         =   0   'False
            Height          =   315
            Left            =   960
            Locked          =   -1  'True
            TabIndex        =   36
            Text            =   "00"
            Top             =   3030
            Width           =   375
         End
         Begin MSAdodcLib.Adodc adoEncuentro 
            Height          =   330
            Left            =   720
            Top             =   1440
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
            RecordSource    =   "SELECT * FROM Encuentros WHERE Numero=0"
            Caption         =   "Encuentro"
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
         Begin MSAdodcLib.Adodc adoEncuentros 
            Height          =   330
            Left            =   720
            Top             =   1080
            Visible         =   0   'False
            Width           =   2175
            _ExtentX        =   3836
            _ExtentY        =   582
            ConnectMode     =   0
            CursorLocation  =   3
            IsolationLevel  =   -1
            ConnectionTimeout=   15
            CommandTimeout  =   30
            CursorType      =   2
            LockType        =   3
            CommandType     =   8
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
            RecordSource    =   $"frmCursadas.frx":1761
            Caption         =   "Encuentros"
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
         Begin VB.Frame Frame1 
            BackColor       =   &H8000000C&
            BorderStyle     =   0  'None
            Height          =   1965
            Left            =   4320
            TabIndex        =   10
            Top             =   1440
            Width           =   1395
            Begin VB.CommandButton cmdCancelarEncuentro 
               Enabled         =   0   'False
               Height          =   550
               Left            =   120
               Picture         =   "frmCursadas.frx":185B
               Style           =   1  'Graphical
               TabIndex        =   15
               ToolTipText     =   "Cancelar"
               Top             =   1320
               Width           =   550
            End
            Begin VB.CommandButton cmdEliminarEncuentro 
               Enabled         =   0   'False
               Height          =   550
               Left            =   720
               Picture         =   "frmCursadas.frx":1C9D
               Style           =   1  'Graphical
               TabIndex        =   14
               ToolTipText     =   "Borrar"
               Top             =   720
               Width           =   550
            End
            Begin VB.CommandButton cmdGuardarEncuentro 
               Enabled         =   0   'False
               Height          =   550
               Left            =   120
               Picture         =   "frmCursadas.frx":20DF
               Style           =   1  'Graphical
               TabIndex        =   13
               ToolTipText     =   "Guardar"
               Top             =   720
               Width           =   550
            End
            Begin VB.CommandButton cmdModificarEncuentro 
               Enabled         =   0   'False
               Height          =   550
               Left            =   720
               Picture         =   "frmCursadas.frx":2521
               Style           =   1  'Graphical
               TabIndex        =   12
               ToolTipText     =   "Modificar"
               Top             =   120
               Width           =   550
            End
            Begin VB.CommandButton cmdAgregarEncuentro 
               Height          =   550
               Left            =   120
               Picture         =   "frmCursadas.frx":2963
               Style           =   1  'Graphical
               TabIndex        =   11
               ToolTipText     =   "Agregar"
               Top             =   120
               Width           =   550
            End
         End
         Begin MSDataGridLib.DataGrid dtgEncuentros 
            Bindings        =   "frmCursadas.frx":2DA5
            Height          =   1575
            Left            =   240
            TabIndex        =   16
            Top             =   360
            Width           =   3975
            _ExtentX        =   7011
            _ExtentY        =   2778
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
               DataField       =   "Nombre"
               Caption         =   "Dia"
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
               DataField       =   "Entrada"
               Caption         =   "Entrada"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   1
                  Format          =   "HH:mm"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   11274
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column02 
               DataField       =   "Salida"
               Caption         =   "Salida"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   1
                  Format          =   "HH:mm"
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
               AllowRowSizing  =   0   'False
               RecordSelectors =   0   'False
               BeginProperty Column00 
               EndProperty
               BeginProperty Column01 
                  Alignment       =   2
               EndProperty
               BeginProperty Column02 
                  Alignment       =   2
               EndProperty
            EndProperty
         End
         Begin MSComCtl2.UpDown upHoraEntrada 
            Height          =   375
            Left            =   600
            TabIndex        =   37
            Top             =   3000
            Width           =   240
            _ExtentX        =   450
            _ExtentY        =   661
            _Version        =   393216
            Max             =   24
            Min             =   -1
            Enabled         =   0   'False
         End
         Begin MSComCtl2.UpDown upMinutosEntrada 
            Height          =   375
            Left            =   1320
            TabIndex        =   39
            Top             =   3000
            Width           =   240
            _ExtentX        =   450
            _ExtentY        =   661
            _Version        =   393216
            Increment       =   5
            Max             =   60
            Min             =   -5
            Enabled         =   0   'False
         End
         Begin MSComCtl2.UpDown upHoraSalida 
            Height          =   375
            Left            =   2520
            TabIndex        =   41
            Top             =   3000
            Width           =   240
            _ExtentX        =   450
            _ExtentY        =   661
            _Version        =   393216
            Max             =   24
            Min             =   -1
            Enabled         =   0   'False
         End
         Begin MSComCtl2.UpDown upMinutosSalida 
            Height          =   375
            Left            =   3240
            TabIndex        =   43
            Top             =   3000
            Width           =   240
            _ExtentX        =   450
            _ExtentY        =   661
            _Version        =   393216
            Increment       =   5
            Max             =   60
            Min             =   -5
            Enabled         =   0   'False
         End
         Begin MSDataListLib.DataCombo dtcDias 
            Bindings        =   "frmCursadas.frx":2DC1
            Height          =   315
            Left            =   240
            TabIndex        =   47
            Top             =   2280
            Width           =   2655
            _ExtentX        =   4683
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            Style           =   2
            ListField       =   "Nombre"
            BoundColumn     =   "Numero"
            Text            =   ""
         End
         Begin VB.Label Label7 
            Caption         =   "Label7"
            Height          =   135
            Left            =   3000
            TabIndex        =   49
            Top             =   2400
            Width           =   15
         End
         Begin VB.Label Label9 
            Caption         =   "Hora de Salida:"
            Height          =   255
            Left            =   2160
            TabIndex        =   46
            Top             =   2760
            Width           =   1335
         End
         Begin VB.Label Label8 
            Caption         =   "Hora de Entrada:"
            Height          =   255
            Left            =   240
            TabIndex        =   45
            Top             =   2760
            Width           =   1335
         End
         Begin VB.Label lblDia 
            Caption         =   "Día:"
            Height          =   255
            Left            =   240
            TabIndex        =   44
            Top             =   2040
            Width           =   2655
         End
      End
   End
   Begin VB.Frame frCarrera 
      Height          =   1575
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   14655
      Begin VB.CommandButton cmdSalir 
         Height          =   550
         Left            =   13800
         Picture         =   "frmCursadas.frx":2DD7
         Style           =   1  'Graphical
         TabIndex        =   48
         ToolTipText     =   "Salir"
         Top             =   840
         Width           =   550
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
         Left            =   3960
         TabIndex        =   25
         Top             =   900
         Width           =   1095
      End
      Begin VB.TextBox txtAño 
         Height          =   285
         Left            =   2520
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   960
         Width           =   615
      End
      Begin MSComCtl2.UpDown UpDown1 
         Height          =   375
         Left            =   3120
         TabIndex        =   6
         Top             =   930
         Width           =   240
         _ExtentX        =   450
         _ExtentY        =   661
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin VB.ComboBox cbCurso 
         Height          =   315
         Left            =   720
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   960
         Width           =   735
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
         RecordSource    =   $"frmCursadas.frx":3219
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
         Bindings        =   "frmCursadas.frx":331D
         Height          =   315
         Left            =   240
         TabIndex        =   4
         Top             =   480
         Width           =   14175
         _ExtentX        =   25003
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "Nombre"
         BoundColumn     =   "Codigo"
         Text            =   ""
      End
      Begin VB.Label lblCurso 
         Caption         =   "Curso:"
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   990
         Width           =   495
      End
      Begin VB.Label Label2 
         Caption         =   "Año:"
         Height          =   255
         Left            =   2160
         TabIndex        =   2
         Top             =   990
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "Carrera:"
         Height          =   255
         Left            =   240
         TabIndex        =   1
         Top             =   240
         Width           =   615
      End
   End
End
Attribute VB_Name = "frmCursadas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Conexion As New Connection
Dim Auxiliar As Recordset
Dim maxDivision As Recordset
Dim Matriculados As New Recordset
Dim Estado As String
Dim DivisionActual As String
Dim EncuentroActual As String

Private Sub cbCurso_Click()
    cmdMostrar.Enabled = True
    frMaterias.Enabled = False
    frDivisiones.Enabled = False
End Sub

Private Sub cmdAgregarDivision_Click()
    cmdAgregarDivision.Enabled = False
    cmdModificarDivision.Enabled = False
    cmdEliminarDivision.Enabled = False
    cmdCancelarDivision.Enabled = True
    cmdGuardarDivision.Enabled = True
    dtcDivision.Enabled = False
    frCarrera.Enabled = False
    frMaterias.Enabled = False
    frEncuentros.Enabled = False
    dtcProfesor.BoundText = 0
    dtcProfesor.Enabled = True
    txtSalon = ""
    txtSalon.Enabled = True
    chkLibre = 0
    chkLibre.Enabled = False
    txtLimiteMatriculados.Text = "50"
    txtLimiteMatriculados.Enabled = True
    Estado = "Agregando"
End Sub

Private Sub cmdAgregarEncuentro_Click()
    cmdAgregarEncuentro.Enabled = False
    cmdModificarEncuentro.Enabled = False
    cmdCancelarEncuentro.Enabled = True
    cmdGuardarEncuentro.Enabled = True
    cmdSalir.Enabled = False
    frCarrera.Enabled = False
    frMaterias.Enabled = False
    frDivisiones.Enabled = False
    dtgEncuentros.Enabled = False
    lblDia = "ELIJA EL DIA ENCUENTRO"
    dtcDias = ""
    dtcDias.Enabled = True
    upHoraEntrada = 0
    upHoraEntrada.Enabled = True
    upMinutos = 0
    upMinutosEntrada.Enabled = True
    upHoraSalida = 0
    upHoraSalida.Enabled = True
    upMinutosSalida = 0
    upMinutosSalida.Enabled = True
    Estado = "Agregando"
End Sub

Private Sub cmdCancelarDivision_Click()
    cmdAgregarDivision.Enabled = True
    If adoDivisiones.Recordset.RecordCount > 0 Then cmdModificarDivision.Enabled = True
    cmdGuardarDivision.Enabled = False
    cmdCancelarDivision.Enabled = False
    dtcProfesor.Enabled = False
    txtSalon.Enabled = False
    chkLibre.Enabled = False
    txtLimiteMatriculados.Enabled = False
    frCarrera.Enabled = True
    frMaterias.Enabled = True
    frEncuentros.Enabled = True
    MostrarDivision
    Estado = ""
End Sub

Private Sub cmdCancelarEncuentro_Click()
    cmdAgregarEncuentro.Enabled = True
    If adoEncuentros.Recordset.RecordCount > 0 Then cmdModificarEncuentro.Enabled = True
    cmdCancelarEncuentro.Enabled = False
    cmdGuardarEncuentro.Enabled = False
    cmdSalir.Enabled = True
    frCarrera.Enabled = True
    frMaterias.Enabled = True
    frDivisiones.Enabled = True
    dtgEncuentros.Enabled = True
    dtcDias.Enabled = False
    txtHoraEntrada = "00"
    upHoraEntrada.Enabled = False
    txtMinutosEntrada = "00"
    upMinutosEntrada.Enabled = False
    txtHoraSalida = "00"
    upHoraSalida.Enabled = False
    txtMinutosSalida = "00"
    upMinutosSalida.Enabled = False
    Estado = ""
    MostrarEncuentro
End Sub

Private Sub cmdEliminarDivision_Click()
    Me.MousePointer = 11
    Conexion.ConnectionString = ("DSN=Instituto")
    Conexion.Open
    Set Matriculados = Conexion.Execute("SELECT count(Ano) as Total FROM Finales WHERE Ano=" & txtAño.Text & " AND Materia = " & adoMaterias.Recordset!Codigo & " AND Division = " & dtcDivision.Text)
    totalMatriculados = Matriculados!Total
    If totalMatriculados > 0 Then Me.MousePointer = 0: respuesta = MsgBox("Existen " & totalMatriculados & " alumnos matriculados para esta división", vbCritical + vbOKOnly, "Imposible eliminar"): Conexion.Close: Exit Sub
    

    Conexion.Close
    
    respuesta = MsgBox("A continuación se eliminará la división " & dtcDivision.Text & " para la asignatura " & adoMaterias.Recordset!Abreviatura & ". ¿Continúa?", vbYesNo, "Eliminar división")
    If respuesta = vbNo Then Exit Sub
    Me.MousePointer = 11
    Conexion.Open
    Conexion.Execute ("DELETE FROM Divisiones WHERE Materia=" & adoMaterias.Recordset!Codigo & " and Ano=" & txtAño.Text & " and Division = " & dtcDivision.Text)
    Me.MousePointer = 0
    Conexion.Close
    VerDivisiones
End Sub

Private Sub cmdGuardarDivision_Click()
    If (Not IsNumeric(txtLimiteMatriculados.Text)) Then
        MsgBox ("El límite de matriculados no es correcto")
        Exit Sub
    End If
    cmdAgregarDivision.Enabled = True
    cmdModificarDivision.Enabled = True
    cmdGuardarDivision.Enabled = False
    cmdCancelarDivision.Enabled = False
    If Estado = "Agregando" Then

        adoDivisiones.Recordset.AddNew
        Conexion.Open
        'Set maxDivision = Conexion.Execute("SELECT MAX(Divisiones.Numero)+1 AS nuevoId From Divisiones")
        
        'If maxDivision.RecordCount = -1 Then
        '    adoDivisiones.Recordset!Numero = 0
        'Else
        '    adoDivisiones.Recordset!Numero = maxDivision!nuevoId
        'End If
        Set Auxiliar = Conexion.Execute("SELECT MAX(Divisiones.Division)+1 AS NuevaDivision From Divisiones WHERE (((Divisiones.Materia)=" & adoMaterias.Recordset!Codigo & ") AND ((Divisiones.Ano)=" & txtAño & " ))")
        adoDivisiones.Recordset!Materia = adoMaterias.Recordset!Codigo
        adoDivisiones.Recordset!Ano = txtAño
        
        If Auxiliar!NuevaDivision > 1 Then
            adoDivisiones.Recordset!Division = Auxiliar!NuevaDivision
            DivisionActual = Auxiliar!NuevaDivision
        Else
            adoDivisiones.Recordset!Division = 1
            DivisionActual = 1
        End If
        Conexion.Close
    End If
    If dtcProfesor = "" Then
        Profesor = 0
    Else
        Profesor = dtcProfesor.BoundText
    End If
    adoDivisiones.Recordset!Profesor = Profesor
    Salon = txtSalon 'aplique esta variable suplementaria porque directamente con el txt tiraba error
    adoDivisiones.Recordset!Salon = Salon
    adoDivisiones.Recordset!Libre = chkLibre
    adoDivisiones.Recordset!LimiteMatriculados = txtLimiteMatriculados.Text
    adoDivisiones.Recordset.Update
    adoDivisiones.Recordset.Requery
    dtcProfesor.Enabled = False
    txtSalon.Enabled = False
    chkLibre.Enabled = False
    txtLimiteMatriculados.Enabled = False
    frCarrera.Enabled = True
    frMaterias.Enabled = True
    frEncuentros.Enabled = True
    VerDivisiones
End Sub

Private Sub cmdGuardarEncuentro_Click()
    If dtcDias.Text = "" Then
        MsgBox ("Debe especificar un dia")
        Exit Sub
    End If
    cmdAgregarEncuentro.Enabled = True
    If adoEncuentros.Recordset.RecordCount > 0 Then cmdModificarEncuentro.Enabled = True
    cmdCancelarEncuentro.Enabled = False
    cmdGuardarEncuentro.Enabled = False
    cmdSalir.Enabled = True
    frCarrera.Enabled = True
    frMaterias.Enabled = True
    frDivisiones.Enabled = True
    dtgEncuentros.Enabled = True
    If Estado = "Agregando" Then
        adoEncuentro.Recordset.AddNew
        Conexion.Open
        Set Auxiliar = Conexion.Execute("SELECT MAX(Encuentros.Encuentro)+1 AS NuevoEncuentro From Encuentros WHERE Encuentros.Numero=" & adoDivisiones.Recordset!Numero)
        If Auxiliar!NuevoEncuentro >= 1 Then
            EncuentroActual = Auxiliar!NuevoEncuentro
        Else
            adoEncuentro.Recordset!Encuentro = 1
            EncuentroActual = 1
        End If
        adoEncuentro.Recordset!Numero = adoDivisiones.Recordset!Numero
        adoEncuentro.Recordset!Encuentro = EncuentroActual
        Conexion.Close
    Else
        EncuentroActual = adoEncuentro.Recordset!Encuentro
    End If
    adoEncuentro.Recordset!Dia = dtcDias.BoundText
    adoEncuentro.Recordset!Entrada = txtHoraEntrada & ":" & txtMinutosEntrada
    adoEncuentro.Recordset!Salida = txtHoraSalida & ":" & txtMinutosSalida
    adoEncuentro.Recordset.Update
    adoEncuentro.Recordset.Requery
    dtcProfesor.Enabled = False
    dtcDias.Enabled = False
    upHoraEntrada.Enabled = False
    upMinutosEntrada.Enabled = False
    upHoraSalida.Enabled = False
    upMinutosSalida.Enabled = False
    VerEncuentros
End Sub

Private Sub cmdModificarDivision_Click()
    cmdAgregarDivision.Enabled = False
    cmdModificarDivision.Enabled = False
    cmdEliminarDivision.Enabled = False
    cmdCancelarDivision.Enabled = True
    cmdGuardarDivision.Enabled = True
    dtcDivision.Enabled = False
    dtcProfesor.Enabled = True
    txtSalon.Enabled = True
    chkLibre.Enabled = True
    txtLimiteMatriculados.Enabled = True
    frCarrera.Enabled = False
    frMaterias.Enabled = False
    frEncuentros.Enabled = False
    DivisionActual = adoDivisiones.Recordset!Division
    Estado = "Modificando"
End Sub

Private Sub cmdModificarEncuentro_Click()
    cmdAgregarEncuentro.Enabled = False
    cmdModificarEncuentro.Enabled = False
    cmdCancelarEncuentro.Enabled = True
    cmdGuardarEncuentro.Enabled = True
    cmdSalir.Enabled = False
    frCarrera.Enabled = False
    frMaterias.Enabled = False
    frDivisiones.Enabled = False
    dtgEncuentros.Enabled = False
    dtcDias.Enabled = True
    upHoraEntrada.Enabled = True
    upMinutosEntrada.Enabled = True
    upHoraSalida.Enabled = True
    upMinutosSalida.Enabled = True
    Estado = "Modificando"
End Sub

Private Sub cmdMostrar_Click()
    adoMaterias.RecordSource = "SELECT Materias.Codigo, Materias.Abreviatura, Modalidad.Modalidad FROM Materias INNER JOIN Modalidad ON Materias.Modalidad = Modalidad.Codigo Where Materias.Carrera = " & dtcCarreras.BoundText & "  And Materias.Curso = " & cbCurso & " And (Materias.Detalle <> 2 And Materias.Detalle <> 4 ) And Materias.Eliminada = 0 ORDER BY Materias.Codigo"
    adoMaterias.Refresh
    cmdMostrar.Enabled = False
    frMaterias.Enabled = True
    frDivisiones.Enabled = True
    VerDivisiones
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub dtcCarreras_Change()
    cmdMostrar.Enabled = True
    frMaterias.Enabled = False
    frDivisiones.Enabled = False
    adoCarreras.Recordset.MoveFirst
    adoCarreras.Recordset.Find ("Codigo=" & dtcCarreras.BoundText)
    lblCurso = adoCarreras.Recordset!Medida
    cbCurso.Clear
    For i = 0 To adoCarreras.Recordset!Años - 1
        cbCurso.List(i) = i + 1
    Next i
    cbCurso.Text = cbCurso.List(0)
End Sub

Private Sub dtcDivision_Change()
    MostrarDivision
End Sub

Private Sub dtgEncuentros_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    MostrarEncuentro
End Sub

Private Sub dtgMaterias_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    VerDivisiones
End Sub

Private Sub Form_Load()
    Conexion.ConnectionString = ("DSN=Instituto")
    dtcCarreras.BoundText = adoCarreras.Recordset!Codigo
    txtAño = Format(Date, "yyyy")
End Sub

Private Sub txtAño_Change()
    cmdMostrar.Enabled = True
    frMaterias.Enabled = False
    frDivisiones.Enabled = False
End Sub

Private Sub UpDown1_DownClick()
    txtAño = Val(txtAño) - 1
End Sub

Private Sub UpDown1_UpClick()
    txtAño = Val(txtAño) + 1
End Sub

Private Function VerDivisiones()
    adoDivisiones.RecordSource = "SELECT * FROM Divisiones WHERE (((Divisiones.Materia)=" & adoMaterias.Recordset!Codigo & ") AND ((Divisiones.Ano)=" & txtAño & ")) ORDER BY Divisiones.Division"
    adoDivisiones.Refresh
    If adoDivisiones.Recordset.RecordCount > 0 Then 'hay alguna/s division creada
        If Estado = "" Then
            DivisionActual = adoDivisiones.Recordset!Division
        Else
            Estado = ""
        End If
        If dtcDivision.BoundText = DivisionActual Then
            MostrarDivision
        Else
            dtcDivision.BoundText = DivisionActual
        End If
        If adoDivisiones.Recordset.RecordCount <= 1 Then
            dtcDivision.Enabled = False
        Else
            dtcDivision.Enabled = True
        End If
        If adoDivisiones.Recordset.RecordCount >= 1 And frmIdentificacion.Permisos!ModificarDivisiones = True Then
           cmdAgregarDivision.Enabled = True
           cmdModificarDivision.Enabled = True
           cmdEliminarDivision.Enabled = True
        Else
           cmdAgregarDivision.Enabled = False
           cmdModificarDivision.Enabled = False
           cmdEliminarDivision.Enabled = False
        End If
        frEncuentros.Enabled = True
    Else
        dtcDivision = ""
        dtcProfesor = ""
        txtSalon = ""
        txtLimiteMatriculados = ""
        cmdAgregarDivision.Enabled = True
        cmdModificarDivision.Enabled = False
        cmdEliminarDivision.Enabled = False
        frEncuentros.Enabled = False
        If frmIdentificacion.Permisos!ModificarDivisiones = False Then
            cmdAgregarDivision.Enabled = False
        End If
    End If
End Function

Private Function MostrarDivision()
    If adoDivisiones.Recordset.RecordCount > 0 Then
        adoDivisiones.Recordset.MoveFirst
        adoDivisiones.Recordset.Find ("Division=" & dtcDivision)
        dtcProfesor.BoundText = adoDivisiones.Recordset!Profesor
        If adoDivisiones.Recordset!Salon <> "" Then
            txtSalon = adoDivisiones.Recordset!Salon
        Else
            txtSalon = ""
        End If
        If adoDivisiones.Recordset!Libre = True Then
           chkLibre = 1
        Else
           chkLibre = 0
        End If
        If adoDivisiones.Recordset!LimiteMatriculados <> "" Then
            txtLimiteMatriculados = adoDivisiones.Recordset!LimiteMatriculados
        Else
            txtLimiteMatriculados = ""
        End If
        
    End If
    VerEncuentros
End Function

Private Function VerEncuentros()
    If adoDivisiones.Recordset.RecordCount = 0 Then
        adoEncuentros.RecordSource = "SELECT Dias.Nombre,Encuentros.Numero,Encuentros.Dia,Encuentros.Entrada, Encuentros.Salida, Encuentros.Encuentro FROM (Encuentros INNER JOIN Divisiones ON Encuentros.Numero = Divisiones.Numero) INNER JOIN Dias ON Encuentros.Dia = Dias.Numero WHERE (((Encuentros.Numero)=0)) ORDER BY Encuentros.Dia"
        adoEncuentros.Refresh
    Else
        adoEncuentro.RecordSource = "SELECT * FROM Encuentros WHERE Numero =" & adoDivisiones.Recordset!Numero & " ORDER BY Dia"
        adoEncuentro.Refresh
        adoEncuentros.RecordSource = "SELECT Dias.Nombre,Encuentros.Numero,Encuentros.Dia,Encuentros.Entrada, Encuentros.Salida, Encuentros.Encuentro FROM (Encuentros INNER JOIN Divisiones ON Encuentros.Numero = Divisiones.Numero) INNER JOIN Dias ON Encuentros.Dia = Dias.Numero WHERE (((Encuentros.Numero)=" & adoDivisiones.Recordset!Numero & ")) ORDER BY Encuentros.Dia"
        adoEncuentros.Refresh
    End If
    If Estado = "" And adoEncuentros.Recordset.RecordCount > 0 Then
        EncuentroActual = adoEncuentros.Recordset!Encuentro
    Else
        Estado = ""
    End If
    If adoEncuentros.Recordset.RecordCount > 0 Then
        adoEncuentros.Recordset.Find ("Encuentro=" & EncuentroActual)
        If frmIdentificacion.Permisos!ModificarEncuentros = True Then
            cmdAgregarEncuentro.Enabled = True
            cmdModificarEncuentro.Enabled = True
            cmdEliminarEncuentro.Enabled = True
        Else
            'cmdAgregarEncuentro.Enabled = False
            cmdModificarEncuentro.Enabled = False
            cmdEliminarEncuentro.Enabled = False
        End If
        MostrarEncuentro
    Else
        cmdModificarEncuentro.Enabled = False
        cmdEliminarEncuentro.Enabled = False
        dtcDias = ""
        upHoraEntrada = 0
        upMinutosEntrada = 0
        upHoraSalida = 0
        upMinutosSalida = 0
    End If
End Function

Private Function MostrarEncuentro()
    If adoEncuentro.Recordset.RecordCount > 0 Then
        adoEncuentro.Recordset.MoveFirst
        adoEncuentro.Recordset.Find ("Encuentro=" & adoEncuentros.Recordset!Encuentro)
        dtcDias.BoundText = adoEncuentro.Recordset!Dia
        upHoraEntrada = Format(adoEncuentro.Recordset!Entrada, "hh")
        upMinutosEntrada = Mid(Format(adoEncuentro.Recordset!Entrada, "hh:mm"), Len(Format(adoEncuentro.Recordset!Entrada, "hh:mm")) - 1, 2)
        upHoraSalida = Format(adoEncuentro.Recordset!Salida, "hh")
        upMinutosSalida = Mid(Format(adoEncuentro.Recordset!Salida, "hh:mm"), Len(Format(adoEncuentro.Recordset!Salida, "hh:mm")) - 1, 2)
    Else
        dtcDias = ""
        upHoraEntrada = 0
        upMinutosEntrada = 0
        upHoraSalida = 0
        upMinutosSalida = 0
    End If
    lblDia = "Dia:"
End Function

Private Sub upHoraEntrada_Change()
    If upHoraEntrada.Value = 24 Then upHoraEntrada.Value = 0
    If upHoraEntrada.Value = -1 Then upHoraEntrada.Value = 23
    txtHoraEntrada = Format(upHoraEntrada.Value, "00")
End Sub

Private Sub upHoraSalida_Change()
    If upHoraSalida.Value = 24 Then upHoraSalida.Value = 0
    If upHoraSalida.Value = -1 Then upHoraSalida.Value = 23
    txtHoraSalida = Format(upHoraSalida.Value, "00")
End Sub

Private Sub upMinutosEntrada_Change()
    If upMinutosEntrada.Value = 60 Then
        upMinutosEntrada.Value = 0
        upHoraEntrada.Value = upHoraEntrada.Value + 1
    ElseIf upMinutosEntrada.Value = -5 Then
        upMinutosEntrada.Value = 55
        upHoraEntrada.Value = upHoraEntrada.Value - 1
    Else
        txtMinutosEntrada = Format(upMinutosEntrada.Value, "00")
    End If
End Sub

Private Sub upMinutosSalida_Change()
    If upMinutosSalida.Value = 60 Then
        upMinutosSalida.Value = 0
        upHoraSalida.Value = upHoraSalida.Value + 1
    ElseIf upMinutosSalida.Value = -5 Then
        upMinutosSalida.Value = 55
        upHoraSalida.Value = upHoraSalida.Value - 1
    Else
        txtMinutosSalida = Format(upMinutosSalida.Value, "00")
    End If
End Sub
