VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmActasImpresion 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Impresión de Actas Volantes para Exámenes Finales"
   ClientHeight    =   10050
   ClientLeft      =   -1575
   ClientTop       =   -1155
   ClientWidth     =   15015
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10050
   ScaleWidth      =   15015
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame frActas 
      Height          =   8775
      Left            =   0
      TabIndex        =   12
      Top             =   1080
      Width           =   14775
      Begin VB.CommandButton cmdPrevisualizar 
         Caption         =   "Previsualizar"
         DragIcon        =   "frmActasImpresion.frx":0000
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1080
         Left            =   10800
         Picture         =   "frmActasImpresion.frx":08CA
         Style           =   1  'Graphical
         TabIndex        =   39
         Tag             =   "Genera, imprime y cierra la mesa."
         Top             =   7320
         Width           =   1575
      End
      Begin VB.CommandButton cmdCerrarMesa 
         Caption         =   "Cerrar Mesa"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1080
         Left            =   12720
         Picture         =   "frmActasImpresion.frx":1194
         Style           =   1  'Graphical
         TabIndex        =   38
         Tag             =   "Genera, imprime y cierra la mesa."
         Top             =   7320
         Width           =   1575
      End
      Begin VB.CommandButton cmdImprimir 
         Caption         =   "Imprimir y Cerrar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1080
         Left            =   8760
         Picture         =   "frmActasImpresion.frx":17FE
         Style           =   1  'Graphical
         TabIndex        =   37
         Tag             =   "Genera, imprime y cierra la mesa."
         Top             =   7320
         Width           =   1575
      End
      Begin VB.CommandButton cmdImprimirViejo 
         Caption         =   "Imprimir"
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
         Height          =   615
         Left            =   8520
         Style           =   1  'Graphical
         TabIndex        =   36
         Top             =   7920
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.CommandButton cmdCerrarMesaViejo 
         Caption         =   "Cerrar Mesa"
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
         Height          =   495
         Left            =   12480
         TabIndex        =   35
         Top             =   8280
         Visible         =   0   'False
         Width           =   2055
      End
      Begin VB.CommandButton cmdRecuperatorioDeAsistencia 
         Caption         =   "Recuperatorio de Asistencia"
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
         Height          =   495
         Left            =   9720
         TabIndex        =   34
         ToolTipText     =   "Imprime el acta con los alumnos que no aprobaron asistencia"
         Top             =   8040
         Visible         =   0   'False
         Width           =   3495
      End
      Begin MSAdodcLib.Adodc adoMesas 
         Height          =   330
         Left            =   360
         Top             =   3240
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
         RecordSource    =   $"frmActasImpresion.frx":1E68
         Caption         =   "Mesas"
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
      Begin VB.Frame frInscripciones 
         Caption         =   "Inscripciones"
         Height          =   6975
         Left            =   8520
         TabIndex        =   18
         Top             =   120
         Width           =   6135
         Begin VB.TextBox Text2 
            DataField       =   "Mesa"
            DataSource      =   "adoActas"
            Height          =   375
            Left            =   360
            TabIndex        =   32
            Text            =   "Text2"
            Top             =   5400
            Visible         =   0   'False
            Width           =   375
         End
         Begin MSAdodcLib.Adodc adoActas 
            Height          =   330
            Left            =   840
            Top             =   5400
            Visible         =   0   'False
            Width           =   2295
            _ExtentX        =   4048
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
            RecordSource    =   "Select * from Actas WHERE Acta = 0"
            Caption         =   "Actas"
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
         Begin MSAdodcLib.Adodc adoInscriptos 
            Height          =   330
            Left            =   840
            Top             =   4440
            Visible         =   0   'False
            Width           =   2415
            _ExtentX        =   4260
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
            RecordSource    =   $"frmActasImpresion.frx":20DA
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
         Begin VB.TextBox Text1 
            DataField       =   "Acta"
            DataSource      =   "adoInscripciones"
            Height          =   375
            Left            =   3120
            TabIndex        =   30
            Text            =   "Text1"
            Top             =   5040
            Visible         =   0   'False
            Width           =   495
         End
         Begin MSAdodcLib.Adodc adoInscripciones 
            Height          =   330
            Left            =   720
            Top             =   5040
            Visible         =   0   'False
            Width           =   2415
            _ExtentX        =   4260
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
            RecordSource    =   "SELECT * FROM Inscripciones WHERE Mesa = 0"
            Caption         =   "Inscripciones"
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
         Begin MSDataGridLib.DataGrid dtgInscriptos 
            Bindings        =   "frmActasImpresion.frx":2244
            Height          =   5415
            Left            =   120
            TabIndex        =   25
            Top             =   960
            Width           =   5895
            _ExtentX        =   10398
            _ExtentY        =   9551
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
            Caption         =   "Inscriptos"
            ColumnCount     =   6
            BeginProperty Column00 
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
            BeginProperty Column01 
               DataField       =   "Documento"
               Caption         =   "DNI"
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
               Caption         =   "Año"
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
            BeginProperty Column03 
               DataField       =   "FechaInscripto"
               Caption         =   "Fecha"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   1
                  Format          =   "dd/MM"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   3082
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column04 
               DataField       =   "HoraInscripto"
               Caption         =   "Hora"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   1
                  Format          =   "HH:mm"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   3082
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column05 
               DataField       =   "Libre"
               Caption         =   "Libre"
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
               AllowRowSizing  =   0   'False
               AllowSizing     =   0   'False
               RecordSelectors =   0   'False
               BeginProperty Column00 
                  ColumnWidth     =   2355,024
               EndProperty
               BeginProperty Column01 
                  Alignment       =   1
                  ColumnWidth     =   900,284
               EndProperty
               BeginProperty Column02 
                  Alignment       =   1
                  ColumnWidth     =   615,118
               EndProperty
               BeginProperty Column03 
                  Alignment       =   1
                  ColumnWidth     =   705,26
               EndProperty
               BeginProperty Column04 
                  Alignment       =   1
                  ColumnWidth     =   599,811
               EndProperty
               BeginProperty Column05 
                  Alignment       =   2
                  ColumnWidth     =   480,189
               EndProperty
            EndProperty
         End
         Begin VB.TextBox txtDivisión 
            Enabled         =   0   'False
            Height          =   315
            Left            =   3480
            TabIndex        =   24
            Top             =   480
            Width           =   375
         End
         Begin VB.TextBox txtAñoCursada 
            Enabled         =   0   'False
            Height          =   315
            Left            =   1800
            TabIndex        =   23
            Top             =   480
            Width           =   615
         End
         Begin VB.ComboBox cbNumeroActa 
            Height          =   315
            Left            =   240
            Style           =   2  'Dropdown List
            TabIndex        =   19
            Top             =   480
            Width           =   735
         End
         Begin VB.Label lblTotalDelActa 
            Caption         =   "Total del Acta:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   240
            TabIndex        =   31
            Top             =   6600
            Width           =   2415
         End
         Begin VB.Label Label9 
            Caption         =   "Año Cursada:"
            Height          =   255
            Left            =   1680
            TabIndex        =   22
            Top             =   240
            Width           =   1215
         End
         Begin VB.Label Label8 
            Caption         =   "División:"
            Height          =   255
            Left            =   3360
            TabIndex        =   21
            Top             =   240
            Width           =   735
         End
         Begin VB.Label Label2 
            Caption         =   "Acta Nº:"
            Height          =   255
            Left            =   240
            TabIndex        =   20
            Top             =   240
            Width           =   735
         End
      End
      Begin VB.Frame frTribunal 
         Caption         =   "Tribunal"
         Height          =   1455
         Left            =   240
         TabIndex        =   14
         Top             =   7200
         Width           =   8175
         Begin VB.TextBox txtIntegrante1 
            DataField       =   "Integrante1"
            DataSource      =   "adoMesas"
            Height          =   315
            Left            =   1200
            Locked          =   -1  'True
            TabIndex        =   29
            Top             =   600
            Width           =   6015
         End
         Begin VB.TextBox txtIntegrante2 
            DataField       =   "Integrante2"
            DataSource      =   "adoMesas"
            Height          =   315
            Left            =   1200
            Locked          =   -1  'True
            TabIndex        =   28
            Top             =   960
            Width           =   6015
         End
         Begin VB.TextBox txtTitular 
            DataField       =   "Titular"
            DataSource      =   "adoMesas"
            Height          =   315
            Left            =   1200
            Locked          =   -1  'True
            TabIndex        =   27
            Top             =   240
            Width           =   6015
         End
         Begin MSAdodcLib.Adodc adoPersonal 
            Height          =   330
            Left            =   5280
            Top             =   360
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
         Begin VB.Label Label5 
            Caption         =   "2º Integrante:"
            Height          =   255
            Left            =   120
            TabIndex        =   17
            Top             =   960
            Width           =   1095
         End
         Begin VB.Label Label6 
            Caption         =   "1º Integrante:"
            Height          =   255
            Left            =   120
            TabIndex        =   16
            Top             =   600
            Width           =   1095
         End
         Begin VB.Label Label7 
            Caption         =   "Titular:"
            Height          =   255
            Left            =   600
            TabIndex        =   15
            Top             =   240
            Width           =   615
         End
      End
      Begin MSDataGridLib.DataGrid dtgMesas 
         Bindings        =   "frmActasImpresion.frx":2260
         Height          =   6855
         Left            =   120
         TabIndex        =   13
         Top             =   240
         Width           =   8295
         _ExtentX        =   14631
         _ExtentY        =   12091
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
         Caption         =   "Mesas disponibles"
         ColumnCount     =   7
         BeginProperty Column00 
            DataField       =   "Materia"
            Caption         =   "Materia"
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
            Caption         =   "Div"
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
               Type            =   1
               Format          =   "dd/mm"
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
               Format          =   "HH:mm"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   11274
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column05 
            DataField       =   "Numero"
            Caption         =   "Numero"
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
            DataField       =   "Actas"
            Caption         =   "Actas"
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
            AllowRowSizing  =   0   'False
            RecordSelectors =   0   'False
            BeginProperty Column00 
               ColumnWidth     =   810,142
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   4004,788
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   360
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   615,118
            EndProperty
            BeginProperty Column04 
               ColumnWidth     =   645,165
            EndProperty
            BeginProperty Column05 
               ColumnWidth     =   794,835
            EndProperty
            BeginProperty Column06 
               ColumnWidth     =   555,024
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
         Left            =   6600
         TabIndex        =   33
         Text            =   "Text3"
         Top             =   120
         Visible         =   0   'False
         Width           =   375
      End
      Begin MSAdodcLib.Adodc adoParametros 
         Height          =   330
         Left            =   4680
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
      Begin VB.CommandButton cmdSalir 
         Height          =   550
         Left            =   14040
         Picture         =   "frmActasImpresion.frx":2277
         Style           =   1  'Graphical
         TabIndex        =   26
         ToolTipText     =   "Salir"
         Top             =   360
         Width           =   550
      End
      Begin MSComCtl2.DTPicker dtpFecha 
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   120
         Visible         =   0   'False
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   450
         _Version        =   393216
         Format          =   84606977
         CurrentDate     =   37585
      End
      Begin VB.TextBox txtAño 
         Height          =   315
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   480
         Width           =   615
      End
      Begin VB.ComboBox cbCurso 
         Height          =   315
         Left            =   12120
         Style           =   2  'Dropdown List
         TabIndex        =   2
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
         Left            =   12840
         TabIndex        =   1
         Top             =   360
         Width           =   1095
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
         RecordSource    =   $"frmActasImpresion.frx":26B9
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
         Bindings        =   "frmActasImpresion.frx":27BD
         Height          =   315
         Left            =   2280
         TabIndex        =   3
         Top             =   480
         Width           =   9735
         _ExtentX        =   17171
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "Nombre"
         BoundColumn     =   "Codigo"
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo dtcMeses 
         Bindings        =   "frmActasImpresion.frx":27D7
         Height          =   315
         Left            =   120
         TabIndex        =   5
         Top             =   480
         Width           =   1095
         _ExtentX        =   1931
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
         TabIndex        =   6
         Top             =   450
         Width           =   240
         _ExtentX        =   450
         _ExtentY        =   661
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin VB.Label lblCurso 
         Caption         =   "Curso:"
         Height          =   255
         Left            =   12120
         TabIndex        =   10
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label3 
         Caption         =   "Carreras Vigentes:"
         Height          =   255
         Left            =   2280
         TabIndex        =   9
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label4 
         Caption         =   "Año:"
         Height          =   255
         Left            =   1320
         TabIndex        =   8
         Top             =   240
         Width           =   495
      End
      Begin VB.Label Label1 
         Caption         =   "Turno:"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   735
      End
   End
End
Attribute VB_Name = "frmActasImpresion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Conexion As New Connection
Dim RecuperanAsistencia As New Recordset
Dim Actualizar As Recordset
Dim Genera As String
Private Sub cbCurso_Click()
    cmdMostrar.Enabled = True
End Sub

Private Sub cbNumeroActa_Click()
    adoInscriptos.RecordSource = "SELECT Alumnos.Permiso, Alumnos.Nombre, Inscripciones.FechaInscripto, Inscripciones.HoraInscripto, Inscripciones.Cursada, Alumnos.Documento, Inscripciones.Libre FROM Inscripciones INNER JOIN Alumnos ON Inscripciones.Alumno = Alumnos.Permiso Where (((Inscripciones.FechaBorrado) Is Null) And ((Inscripciones.Mesa) = " & adoMesas.Recordset!Numero & ") And ((Inscripciones.Acta) = " & cbNumeroActa.Text & "))ORDER BY Alumnos.Nombre"
    adoInscriptos.Refresh
    txtAñoCursada = adoInscriptos.Recordset!Cursada
    txtDivisión = adoMesas.Recordset!division
    lblTotalDelActa = "Inscriptos: " & adoInscriptos.Recordset.RecordCount
End Sub

Private Sub cmdCerrarMesa_Click()
    If adoInscriptos.Recordset.RecordCount > 0 Then
        Respuesta = MsgBox("No se puede cerrar la mesa ya que hay inscriptos en la misma", vbOKOnly, "Atención"): Exit Sub
    End If
    Respuesta = MsgBox("Si cierra la mesa los alumnos ya no podrán inscribirse o borarrse a las mismas ni tampoco imprimir las actas." & Chr(13) & "¿Desea continuar?", vbYesNo, "Atención")
    If Respuesta = vbNo Then Exit Sub
    Conexion.Open
    Set Actualizar = Conexion.Execute("UPDATE Mesas SET Mesas.Impresas = True WHERE Mesas.Numero = " & adoMesas.Recordset!Numero)
    Conexion.Close
    If adoMesas.Recordset!Impresas = False Then
        adoMesas.Refresh
        Genera = "Si" ' para que no levante otra vez las inscripciones
        adoMesas.Recordset.Move MesaActual - 1
        Genera = "No"
        cmdImprimir.Caption = "Reimprimir"
        cmdCerrarMesa.Enabled = False
    End If

End Sub

Private Sub cmdImprimir_Click()
    If frmIdentificacion.Permisos!ImprimirActas = False Then MsgBox ("Usted no tiene permiso para imprimir Actas"): Exit Sub
    If adoMesas.Recordset!Impresas = True Then
        Respuesta = MsgBox("Está seguro de reimprimir las actas", vbYesNo)
        If Respuesta = vbNo Then Exit Sub
    Else
        Respuesta = MsgBox("Está seguro de imprimir las actas", vbYesNo)
        If Respuesta = vbNo Then Exit Sub
    End If
    MesaActual = adoMesas.Recordset.Bookmark
    For i = 0 To adoMesas.Recordset!Actas - 1
        adoInscriptos.RecordSource = "SELECT Alumnos.Permiso, Alumnos.Nombre, Inscripciones.FechaInscripto, Inscripciones.HoraInscripto, Inscripciones.Cursada, [Alumnos].[Documento] AS Documento, Inscripciones.Libre  FROM Inscripciones INNER JOIN Alumnos ON Inscripciones.Alumno = Alumnos.Permiso Where (((Inscripciones.FechaBorrado) Is Null) And ((Inscripciones.Mesa) = " & adoMesas.Recordset!Numero & ") And ((Inscripciones.Acta) = " & i + 1 & "))ORDER BY Alumnos.Nombre"
        adoInscriptos.Refresh
        MsgBox ("Imprimir acta Nº " & i + 1 & " de " & adoMesas.Recordset!Actas)
        With frmImprimeActas
        adoInscriptos.Recordset.MoveFirst
        .lblEstablecimiento = adoParametros.Recordset!NombreInstitucion
        .lblTituloDeActa = "Acta Volante de Exámenes Finales"
        .lblCarrera = dtcCarreras.Text
        .lblMateria = adoMesas.Recordset!Materia & " " & adoMesas.Recordset!Nombre
        .lblMesa = adoMesas.Recordset!Numero
        .lblActa = i + 1
        .lblFecha = Format(adoMesas.Recordset!Fecha, "dd/mm/yyyy")
        .lblHora = adoMesas.Recordset!Hora
        .lblLugar = adoMesas.Recordset!Lugar
        .lblCurso = cbCurso.Text
        .lblCursada = adoInscriptos.Recordset!Cursada
        .lblDivision = txtDivisión.Text
        .lblTitular = adoMesas.Recordset!Titular
        .lblIntegrante1 = adoMesas.Recordset!Integrante1
        .lblIntegrante2 = adoMesas.Recordset!Integrante2
        .lblTotalAlumnos = adoInscriptos.Recordset.RecordCount
        .lblLocalidad = adoParametros.Recordset!Localidad
        
        For j = 1 To adoInscriptos.Recordset.RecordCount
            .lblPermiso(j) = adoInscriptos.Recordset!Permiso
            .lblAlumno(j) = adoInscriptos.Recordset!Nombre
            .lblDocumento(j) = Format(adoInscriptos.Recordset!documento, "##,###,###")
            '.lblAnioCursada(j) = adoInscriptos.Recordset!Cursada
            If adoInscriptos.Recordset!Libre = True Then
                '.lblLibre(j) = "L"
            End If
            .lblOrden(j).Visible = True
            .lblPermiso(j).Visible = True
            .lblEscritoNota(j).Visible = True
            .lnlEscritoLetras(j).Visible = True
            .lblOralNota(j).Visible = True
            .lnlOralLetras(j).Visible = True
            .lblFinalNota(j).Visible = True
            .lnlFinalLetras(j).Visible = True
            .lblAlumno(j).Visible = True
            .lblDocumento(j).Visible = True
            '.lblAnioCursada(j).Visible = True
            '.lblLibre(j).Visible = True
            adoInscriptos.Recordset.MoveNext
        Next j
        End With
        '******impresion por formulario**********
        'frmImprimeActas.PrintForm
        'Unload frmImprimeActas
        
        '****** impresiòn por reporte ************
        cn.Open
        numero_mesa = adoMesas.Recordset!Numero
        numero_acta = i + 1
        rptActaFinales.WindowState = 2
        rptActaFinales.PrintReport
        Unload rptActaFinales
        cn.Close
            'Actualizo la cantidad de actas de la mesa correspondiente
        If adoMesas.Recordset!Impresas = False Then
           Conexion.Open
           Set Actualizar = Conexion.Execute("UPDATE Mesas SET Mesas.Actas = " & adoMesas.Recordset!Actas & " WHERE (((Mesas.Numero)=" & adoMesas.Recordset!Numero & "))")
           adoInscriptos.Recordset.MoveFirst 'para sacar el año de la cursada
           adoActas.Recordset.AddNew
           adoActas.Recordset!Mesa = adoMesas.Recordset!Numero
           adoActas.Recordset!Acta = i + 1
           adoActas.Recordset!Total = adoInscriptos.Recordset.RecordCount
           adoActas.Recordset!Ano = adoInscriptos.Recordset!Cursada
           adoActas.Recordset!division = adoMesas.Recordset!division
           adoActas.Recordset.Update
               'Actualizo como que las actas ya estan impresas
           Set Actualizar = Conexion.Execute("UPDATE Mesas SET Mesas.Impresas = True WHERE Mesas.Numero = " & adoMesas.Recordset!Numero)
           Conexion.Close
        End If
    Next i
    If adoMesas.Recordset!Impresas = False Then
        adoMesas.Refresh
        Genera = "Si" ' para que no levante otra vez las inscripciones
        adoMesas.Recordset.Move MesaActual - 1
        Genera = "No"
        cmdImprimir.Caption = "Reimprimir"
        cmdCerrarMesa.Enabled = False
    End If
End Sub

Private Sub cmdMostrar_Click()
    Me.MousePointer = 11
    cmdMostrar.Enabled = False
    Genera = "Si"
    adoMesas.RecordSource = "SELECT Mesas.Materia, Materias.Nombre, Mesas.Division, Mesas.Fecha, Mesas.Hora, Mesas.Numero, Mesas.Actas, Mesas.Impresas, Mesas.Lugar, Personal.Nombre AS Titular, Personal_1.Nombre AS Integrante1, Personal_2.Nombre AS Integrante2 FROM (((Mesas INNER JOIN Materias ON Mesas.Materia = Materias.Codigo) INNER JOIN Personal ON Mesas.Titular = Personal.Codigo) INNER JOIN Personal AS Personal_1 ON Mesas.Integrante1 = Personal_1.Codigo) INNER JOIN Personal AS Personal_2 ON Mesas.Integrante2 = Personal_2.Codigo Where (((Mesas.Turno) = " & dtcMeses.BoundText & ") And ((Mesas.Ano) = " & txtAño & ") And ((Materias.Carrera) = " & dtcCarreras.BoundText & ") And ((Materias.Curso) = " & cbCurso & "))ORDER BY Mesas.Materia, Mesas.Division"
    adoMesas.Refresh
    Genera = "No"
    If adoMesas.Recordset.RecordCount > 0 Then
        Genera = "Si"
        GenerarActas
        Genera = "No"
        adoMesas.Recordset.Requery
        adoMesas.Refresh
    End If
    Me.MousePointer = 1
End Sub

Private Sub cmdPrevisualizar_Click()
    For i = 0 To adoMesas.Recordset!Actas - 1
        MsgBox ("Previsualizar acta Nº " & i + 1 & " de " & adoMesas.Recordset!Actas)
        
        '****** impresiòn por reporte ************
        cn.Open
        numero_mesa = adoMesas.Recordset!Numero
        numero_acta = i + 1
        rptActaFinales.WindowState = 2
        rptActaFinales.Show 1
        Unload rptActaFinales
        cn.Close
    Next i
End Sub

Private Sub cmdRecuperatorioDeAsistencia_Click()
    If frmIdentificacion.Permisos!ImprimirActas = False Then MsgBox ("Usted no tiene permiso para imprimir Actas"): Exit Sub
    Respuesta = MsgBox("Está seguro de imprimir las actas", vbYesNo)
    If Respuesta = vbNo Then Exit Sub
    Dim CantidadActas As Integer
    Dim Lugar As Integer
    CantidadActas = 1
    Lugar = 1
    Conexion.Open
    Set RecuperanAsistencia = Conexion.Execute("SELECT Materias.Carrera, Carreras.Nombre, Finales.Materia, Finales.Division, Finales.Cursada, Materias.Nombre, Alumnos.Permiso, Alumnos.Documento, Alumnos.Nombre, Finales.Cursada, Finales.Asistencia, Finales.Ano, Finales.AsistenciaPorcentaje FROM ((Finales INNER JOIN Alumnos ON Finales.Alumno = Alumnos.Permiso) INNER JOIN Materias ON Finales.Materia = Materias.Codigo) INNER JOIN Carreras ON Materias.Carrera = Carreras.Codigo Where (((Finales.Materia) = " & adoMesas.Recordset!Materia & ") And ((Finales.Division) = " & adoMesas.Recordset!division & ") And ((Finales.Asistencia) = False)) and Finales.Ano = " & txtAño & " ORDER BY Carreras.Nombre, Materias.Nombre")
    If RecuperanAsistencia.EOF = True Then Conexion.Close: MsgBox ("Ningún alumno desaprobó la asistencia"): Exit Sub
    'recorro los que rinden recuperatorio
    With frmImprimeActasAsistencia
    MsgBox ("Imprimir acta Nº " & CantidadActas)
    While RecuperanAsistencia.EOF = False
        If Lugar > 26 Then
            .lblTituloDeActa = "Recuperatorio de Asistencia"
            .lblCarrera = dtcCarreras.Text
            .lblMateria = adoMesas.Recordset!Materia & " " & adoMesas.Recordset!Nombre
            .lblActa = CantidadActas
            .lblFecha = Format(adoMesas.Recordset!Fecha, "dd/mm/yyyy")
            .lblHora = adoMesas.Recordset!Hora
            .lblCurso = cbCurso.Text
            .lblDivision = adoMesas.Recordset!division
            .lblTitular = adoMesas.Recordset!Titular
            .lblTotalAlumnos = 26
            .PrintForm
            'Unload frmImprimeActasAsistencia
            For h = 1 To 26
                .lblAlumno(h).Visible = False
                .lblDocumento(h).Visible = False
                .lblPermiso(h).Visible = False
                .lblOrden(h).Visible = False
                .lblPermiso(h).Visible = False
                .lblCursada(h).Visible = False
                .lblPorcentaje(h).Visible = False
                .lblCursadaAprobada(h).Visible = False
                .lblFinalNota(h).Visible = False
                .lblFinalLetras(h).Visible = False
                .lblAprobo(h).Visible = False
            Next h
            CantidadActas = CantidadActas + 1
            MsgBox ("Imprimir acta Nº " & CantidadActas)
            Lugar = 1
        End If
        .lblPermiso(Lugar) = RecuperanAsistencia!Permiso
        .lblAlumno(Lugar) = RecuperanAsistencia!Nombre
        .lblDocumento(Lugar) = Format(RecuperanAsistencia!documento, "##,###,###")
        .lblCursada(Lugar) = RecuperanAsistencia!Ano
        .lblPorcentaje(Lugar) = RecuperanAsistencia!AsistenciaPorcentaje
        If RecuperanAsistencia!Cursada = True Then
           .lblCursadaAprobada(Lugar) = "Si"
        Else
            .lblCursadaAprobada(Lugar) = "No"
        End If
        .lblAlumno(Lugar).Visible = True
        .lblDocumento(Lugar).Visible = True
        .lblPermiso(Lugar).Visible = True
        .lblOrden(Lugar).Visible = True
        .lblPermiso(Lugar).Visible = True
        .lblCursada(Lugar).Visible = True
        .lblPorcentaje(Lugar).Visible = True
        .lblCursadaAprobada(Lugar).Visible = True
        .lblFinalNota(Lugar).Visible = True
        .lblFinalLetras(Lugar).Visible = True
        .lblAprobo(Lugar).Visible = True
        Lugar = Lugar + 1
        RecuperanAsistencia.MoveNext
    Wend
    If Lugar > 26 Then
        .lblTotalAlumnos = Lugar
        .PrintForm
        Unload frmImprimeActasAsistencia
    End If
    End With
    Conexion.Close
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Command1_Click()

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
    cmdMostrar.Enabled = True
    LimpiarInscripciones
End Sub

Private Sub dtcMeses_Change()
    cmdMostrar.Enabled = True
    LimpiarInscripciones
End Sub

Private Sub dtgMesas_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
 If Genera <> "Si" Then
    If adoMesas.Recordset.EOF = False Then
        If adoMesas.Recordset!Actas > 0 Then
            cbNumeroActa.Clear
            For i = 0 To adoMesas.Recordset!Actas - 1
                cbNumeroActa.List(i) = i + 1
            Next i
            cbNumeroActa.Text = cbCurso.List(0)
            frInscripciones.Enabled = True
            If adoMesas.Recordset!Impresas = True Then
                cmdImprimir.Caption = "Reimprimir"
                cmdCerrarMesa.Enabled = False
            Else
                cmdImprimir.Caption = "Imprimir y Cerrar"
                cmdCerrarMesa.Enabled = True
            End If
            cmdImprimir.Enabled = True
        Else
            cmdImprimir.Enabled = False
            cmdCerrarMesa.Enabled = True
            LimpiarInscripciones
        End If
    End If
   If adoMesas.Recordset!Impresas = False Then
      cmdCerrarMesa.Enabled = True
      cmdPrevisualizar.Enabled = False
   Else
      cmdCerrarMesa.Enabled = False
      cmdPrevisualizar.Enabled = True
   End If
End If
End Sub

Private Sub Form_Load()
    Conexion.ConnectionString = ("DSN=Instituto")
    dtpFecha.Value = Date
    dtcMeses.BoundText = adoParametros.Recordset!TurnoLlamado
    txtAño = adoParametros.Recordset!AñoLlamado
    dtcCarreras.BoundText = adoCarreras.Recordset!Codigo
End Sub

Private Sub UpDown1_Change()
    cmdMostrar.Enabled = True
    LimpiarInscripciones
End Sub

Private Sub UpDown1_DownClick()
    txtAño = Val(txtAño) - 1
End Sub

Private Sub UpDown1_UpClick()
    txtAño = Val(txtAño) + 1
End Sub

Private Function GenerarActas()
    Conexion.Open
    For i = 1 To adoMesas.Recordset.RecordCount
        adoInscripciones.RecordSource = "SELECT Inscripciones.Acta, Inscripciones.Cursada, Inscripciones.Alumno FROM Inscripciones INNER JOIN Alumnos ON Inscripciones.Alumno = Alumnos.Permiso Where (((Inscripciones.Mesa) = " & adoMesas.Recordset!Numero & ") And ((Inscripciones.FechaBorrado) Is Null))ORDER BY Alumnos.Nombre"
        adoInscripciones.Refresh
        If adoInscripciones.Recordset.RecordCount > 0 And adoMesas.Recordset!Impresas = falso Then
            'si hay inscriptos y no esta impresa la mesa
            OrdenActa = 1
            TotalActas = 1
            For j = 1 To adoInscripciones.Recordset.RecordCount
                Set Actualizar = Conexion.Execute("UPDATE Inscripciones SET Inscripciones.Acta = " & TotalActas & " WHERE (((Inscripciones.Mesa)=" & adoMesas.Recordset!Numero & ") AND ((Inscripciones.Alumno)=" & adoInscripciones.Recordset!Alumno & ") AND ((Inscripciones.FechaBorrado) Is Null))")
                AñoCursada = adoInscripciones.Recordset!Cursada
                OrdenActa = OrdenActa + 1
                adoInscripciones.Recordset.MoveNext
                If adoInscripciones.Recordset.EOF = True Then Exit For
                If OrdenActa = 26 Then
                    TotalActas = TotalActas + 1
                    OrdenActa = 1
                End If
            Next j
            Set Actualizar = Conexion.Execute("UPDATE Mesas SET Mesas.Actas = " & TotalActas & " WHERE (((Mesas.Numero)=" & adoMesas.Recordset!Numero & "));")
        Else
            'por si se borraron todos
            If adoMesas.Recordset!Impresas = False Then Set Actualizar = Conexion.Execute("UPDATE Mesas SET Mesas.Actas = 0 WHERE (((Mesas.Numero)=" & adoMesas.Recordset!Numero & "));")
        End If
        adoMesas.Recordset.MoveNext
    Next i
    Conexion.Close
End Function

Private Function LimpiarInscripciones()
    adoInscriptos.RecordSource = "SELECT Alumnos.Nombre, Inscripciones.FechaInscripto, Inscripciones.HoraInscripto, ([Alumnos].[Tipo] & [Alumnos].[Documento]) AS Documento FROM Inscripciones INNER JOIN Alumnos ON Inscripciones.Alumno = Alumnos.Permiso Where (((Inscripciones.FechaBorrado) Is Null) And ((Inscripciones.Mesa) = 0) And ((Inscripciones.Acta) = 0 ))ORDER BY Alumnos.Nombre"
    adoInscriptos.Refresh
    cbNumeroActa.Clear
    txtAñoCursada = ""
    txtDivisión = ""
    lblTotalDelActa = "Total del Acta:"
    frInscripciones.Enabled = False
End Function
