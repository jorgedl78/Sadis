VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmAgregarMesa 
   Caption         =   "Agregar Mesa de Exámen"
   ClientHeight    =   8760
   ClientLeft      =   4245
   ClientTop       =   1200
   ClientWidth     =   10245
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   Moveable        =   0   'False
   ScaleHeight     =   8760
   ScaleWidth      =   10245
   Begin VB.CommandButton cmdVerHabilitados 
      Caption         =   "Ver Habilitados"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   9240
      TabIndex        =   12
      ToolTipText     =   "Cancelar"
      Top             =   4080
      Width           =   1095
   End
   Begin MSAdodcLib.Adodc adoCorrelativas 
      Height          =   375
      Left            =   2880
      Top             =   7920
      Visible         =   0   'False
      Width           =   2175
      _ExtentX        =   3836
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
      RecordSource    =   $"frmAgregarMesa.frx":0000
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
   Begin MSDataGridLib.DataGrid dtgCorrelativas 
      Bindings        =   "frmAgregarMesa.frx":00D5
      Height          =   1335
      Left            =   240
      TabIndex        =   10
      Top             =   7200
      Width           =   8895
      _ExtentX        =   15690
      _ExtentY        =   2355
      _Version        =   393216
      AllowUpdate     =   0   'False
      ColumnHeaders   =   0   'False
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
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   "Correlativa"
         Caption         =   "Correlativa"
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
      SplitCount      =   1
      BeginProperty Split0 
         MarqueeStyle    =   5
         AllowFocus      =   0   'False
         AllowRowSizing  =   0   'False
         AllowSizing     =   0   'False
         RecordSelectors =   0   'False
         BeginProperty Column00 
            ColumnWidth     =   915.024
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1739.906
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc adoMaterias 
      Height          =   330
      Left            =   480
      Top             =   3360
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
      RecordSource    =   $"frmAgregarMesa.frx":00F3
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
   Begin VB.CommandButton cmdAgregarMesa 
      Height          =   615
      Left            =   9480
      Picture         =   "frmAgregarMesa.frx":01C9
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Aceptar"
      Top             =   2040
      Width           =   615
   End
   Begin VB.CommandButton cmdCancelarMesa 
      Height          =   615
      Left            =   9480
      Picture         =   "frmAgregarMesa.frx":060B
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Cancelar"
      Top             =   3000
      Width           =   615
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frmAgregarMesa.frx":0A4D
      Height          =   4455
      Left            =   120
      TabIndex        =   2
      Top             =   2040
      Width           =   8895
      _ExtentX        =   15690
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
      ColumnCount     =   2
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
            ColumnWidth     =   1005.165
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   7244.788
         EndProperty
      EndProperty
   End
   Begin VB.Label Label2 
      Caption         =   "Correlativas:"
      Height          =   255
      Left            =   240
      TabIndex        =   11
      Top             =   6840
      Width           =   1935
   End
   Begin VB.Label lblTurno 
      Height          =   255
      Left            =   840
      TabIndex        =   9
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label lblAño 
      Height          =   255
      Left            =   2760
      TabIndex        =   8
      Top             =   120
      Width           =   615
   End
   Begin VB.Label Label4 
      Caption         =   "Año:"
      Height          =   255
      Left            =   2280
      TabIndex        =   7
      Top             =   120
      Width           =   375
   End
   Begin VB.Label Label3 
      Caption         =   "Turno:"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   495
   End
   Begin VB.Label Label1 
      Caption         =   "Carrera:"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   720
      Width           =   615
   End
   Begin VB.Label lblCarrera 
      Height          =   615
      Left            =   840
      TabIndex        =   4
      Top             =   720
      Width           =   5655
   End
   Begin VB.Label lblCurso 
      Caption         =   "Curso:"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   1560
      Width           =   1935
   End
End
Attribute VB_Name = "frmAgregarMesa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAgregarMesa_Click()
    With frmMesasArmado
    .txtMateria = adoMaterias.Recordset!Nombre
    .MateriaAgregar = adoMaterias.Recordset!Codigo
    .Mostrar = "No"
    End With
    Unload Me
End Sub

Private Sub cmdCancelarMesa_Click()
    frmMesasArmado.cmdCancelar_Click
    frmMesasArmado.Mostrar = "Si"
    Unload Me
End Sub

Private Sub Command1_Click()

End Sub

Private Sub cmdVerHabilitados_Click()
    With frmVerHabilitadosParaRendir
    .adoAlumnos.RecordSource = "SELECT Alumnos.Permiso, Alumnos.Tipo, Alumnos.Documento, Alumnos.Nombre AS Alumno, Finales.Ano, Personal.Nombre AS Profesor, Finales.PerdioTurno, Finales.Asistencia, Finales.AsistenciaPorcentaje FROM (Alumnos INNER JOIN Finales ON Alumnos.Permiso = Finales.Alumno) INNER JOIN Personal ON Finales.Profesor = Personal.Codigo WHERE (((Finales.Cursada)=True) AND ((Finales.Aprobada)=False) AND ((Finales.Materia)=" & adoMaterias.Recordset!Codigo & ")) ORDER BY Alumnos.Nombre"
    .adoAlumnos.Refresh
    If .adoAlumnos.Recordset.RecordCount = 0 Then MsgBox ("No existen alumnos habilitados para rendir"): Exit Sub
    .Caption = "Habilitados para rendir " & adoMaterias.Recordset!Nombre
    .lblHabilitados = .adoAlumnos.Recordset.RecordCount & " alumnos habilitados"
    .Show 1
    End With
End Sub

Private Sub DataGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    BuscarCorrelativas
End Sub

Private Sub Form_Resize()
    BuscarCorrelativas
End Sub

Private Function BuscarCorrelativas()
    adoCorrelativas.RecordSource = "SELECT Correlativas.Correlativa, Materias.Nombre FROM Correlativas INNER JOIN Materias ON Correlativas.Correlativa = Materias.Codigo Where (((Correlativas.Principal) = " & adoMaterias.Recordset!Codigo & " )) ORDER BY Correlativas.Correlativa"
    adoCorrelativas.Refresh
End Function
