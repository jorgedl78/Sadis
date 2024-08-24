VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmCorrelativas 
   Caption         =   "Correlativas"
   ClientHeight    =   8310
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10290
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   Moveable        =   0   'False
   ScaleHeight     =   8310
   ScaleWidth      =   10290
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Height          =   6015
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   9975
      Begin MSDataGridLib.DataGrid DataGrid1 
         Bindings        =   "frmCorrelativas.frx":0000
         Height          =   4455
         Left            =   120
         TabIndex        =   6
         Top             =   1440
         Width           =   9735
         _ExtentX        =   17171
         _ExtentY        =   7858
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
         SplitCount      =   1
         BeginProperty Split0 
            MarqueeStyle    =   4
            BeginProperty Column00 
               ColumnWidth     =   540,284
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   1019,906
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   7560
            EndProperty
         EndProperty
      End
      Begin VB.Label Label1 
         Caption         =   "Carrera:"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   360
         Width           =   615
      End
      Begin VB.Label lblCarrera 
         Height          =   255
         Left            =   960
         TabIndex        =   9
         Top             =   360
         Width           =   9015
      End
      Begin VB.Label Label2 
         Caption         =   "Materia:"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   840
         Width           =   615
      End
      Begin VB.Label lblMateria 
         Height          =   255
         Left            =   960
         TabIndex        =   7
         Top             =   840
         Width           =   9015
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1455
      Left            =   120
      TabIndex        =   1
      Top             =   6360
      Width           =   9975
      Begin VB.CommandButton cmdAgregarCorrelativa 
         Height          =   615
         Left            =   5640
         Picture         =   "frmCorrelativas.frx":001A
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Aceptar"
         Top             =   480
         Width           =   615
      End
      Begin VB.CommandButton cmdCancelarCorrelativa 
         Height          =   615
         Left            =   7200
         Picture         =   "frmCorrelativas.frx":045C
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Cancelar"
         Top             =   480
         Width           =   615
      End
      Begin VB.CheckBox chkPorFinal 
         Alignment       =   1  'Right Justify
         Caption         =   "La correlativa es por final aprobado"
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
         TabIndex        =   2
         Top             =   600
         Width           =   3855
      End
   End
   Begin VB.TextBox Text1 
      DataField       =   "Carrera"
      DataSource      =   "adoCorrelativas"
      Height          =   375
      Left            =   3960
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   7920
      Visible         =   0   'False
      Width           =   615
   End
   Begin MSAdodcLib.Adodc adoCorrelativas 
      Height          =   330
      Left            =   6720
      Top             =   7920
      Visible         =   0   'False
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   2
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
      RecordSource    =   "Correlativas"
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
   Begin MSAdodcLib.Adodc adoMaterias 
      Height          =   330
      Left            =   1680
      Top             =   8040
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
      RecordSource    =   "SELECT Curso, Codigo,  Nombre FROM Materias WHERE Codigo = 0"
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
End
Attribute VB_Name = "frmCorrelativas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAgregarCorrelativa_Click()
    adoCorrelativas.Recordset.AddNew
    adoCorrelativas.Recordset!Carrera = frmPlanes.adoCarreras.Recordset!Codigo
    adoCorrelativas.Recordset!Principal = frmPlanes.adoMaterias.Recordset!Codigo
    adoCorrelativas.Recordset!Correlativa = adoMaterias.Recordset!Codigo
    adoCorrelativas.Recordset!PorFinal = chkPorFinal
    adoCorrelativas.Recordset.Update
    Unload Me
    frmPlanes.MostrarMaterias
End Sub

Private Sub cmdCancelarCorrelativa_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    lblCarrera = frmPlanes.adoCarreras.Recordset!Abreviatura
    lblMateria = frmPlanes.adoMaterias.Recordset!Abreviatura
    adoMaterias.RecordSource = "SELECT Curso, Codigo, Nombre FROM Materias WHERE Carrera = " & frmPlanes.adoMaterias.Recordset!Carrera & " AND Detalle <>4 AND Curso <= " & frmPlanes.adoMaterias.Recordset!Curso & " AND Eliminada = 0 ORDER BY Codigo"
    adoMaterias.Refresh
    If adoMaterias.Recordset.RecordCount = 0 Then frmCorrelativas.cmdAgregarCorrelativa.Enabled = False
End Sub
