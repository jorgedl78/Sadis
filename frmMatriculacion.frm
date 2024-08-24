VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "msadodc.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Begin VB.Form frmMatriculacion 
   ClientHeight    =   8880
   ClientLeft      =   210
   ClientTop       =   525
   ClientWidth     =   11880
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MousePointer    =   1  'Arrow
   Moveable        =   0   'False
   ScaleHeight     =   8880
   ScaleWidth      =   11880
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdPresencial 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Presencial"
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
      Left            =   6720
      MouseIcon       =   "frmMatriculacion.frx":0000
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   4320
      Width           =   1815
   End
   Begin VB.Frame frInscripcionFinales 
      BackColor       =   &H00C0FFFF&
      Height          =   7815
      Left            =   0
      TabIndex        =   0
      Top             =   1200
      Width           =   11895
      Begin VB.Frame Frame1 
         BackColor       =   &H00C0FFFF&
         Height          =   1335
         Left            =   120
         TabIndex        =   13
         Top             =   2400
         Width           =   11655
         Begin VB.CommandButton cmdCursarLibre 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Cursar Libre"
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
            Left            =   6600
            MouseIcon       =   "frmMatriculacion.frx":030A
            MousePointer    =   99  'Custom
            Style           =   1  'Graphical
            TabIndex        =   17
            Top             =   240
            Width           =   1815
         End
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
            Height          =   975
            Left            =   9840
            MouseIcon       =   "frmMatriculacion.frx":0614
            MousePointer    =   99  'Custom
            Picture         =   "frmMatriculacion.frx":091E
            Style           =   1  'Graphical
            TabIndex        =   16
            Top             =   240
            Width           =   1215
         End
         Begin VB.CommandButton cmdAgregar 
            BackColor       =   &H00C0FFC0&
            Caption         =   "Click aquí para matricularse"
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
            Height          =   855
            Left            =   1200
            MaskColor       =   &H000000FF&
            MouseIcon       =   "frmMatriculacion.frx":0D60
            MousePointer    =   99  'Custom
            Style           =   1  'Graphical
            TabIndex        =   15
            Top             =   240
            Width           =   2415
         End
         Begin VB.CommandButton cmdBorrarme 
            BackColor       =   &H008080FF&
            Caption         =   "Borrarme"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   855
            Left            =   3840
            MouseIcon       =   "frmMatriculacion.frx":106A
            MousePointer    =   99  'Custom
            Style           =   1  'Graphical
            TabIndex        =   14
            Top             =   240
            Width           =   2415
         End
      End
      Begin VB.TextBox Text1 
         DataField       =   "Materia"
         DataSource      =   "adoMatriculacion"
         Height          =   285
         Left            =   2640
         TabIndex        =   10
         Text            =   "Text1"
         Top             =   4920
         Visible         =   0   'False
         Width           =   255
      End
      Begin MSAdodcLib.Adodc adoMatriculacion 
         Height          =   330
         Left            =   360
         Top             =   4920
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
         RecordSource    =   $"frmMatriculacion.frx":1374
         Caption         =   "Matriculacion"
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
         Top             =   480
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
         RecordSource    =   $"frmMatriculacion.frx":1595
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
      Begin MSDataGridLib.DataGrid dtgInscripto 
         Bindings        =   "frmMatriculacion.frx":172C
         Height          =   3135
         Left            =   120
         TabIndex        =   7
         Top             =   4320
         Width           =   11655
         _ExtentX        =   20558
         _ExtentY        =   5530
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
            DataField       =   "Curso"
            Caption         =   "Curso"
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
            DataField       =   "Codigo"
            Caption         =   "Codigo"
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
            DataField       =   "Materia"
            Caption         =   "Materia"
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
            DataField       =   "Profesor"
            Caption         =   "Profesor"
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
            DataField       =   "Salon"
            Caption         =   "Salon"
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
         BeginProperty Column05 
            DataField       =   "Condicion"
            Caption         =   "Condicion"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   5
               Format          =   ""
               HaveTrueFalseNull=   1
               TrueValue       =   "Libre"
               FalseValue      =   "Presencial"
               NullValue       =   ""
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   7
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            MarqueeStyle    =   4
            BeginProperty Column00 
               ColumnWidth     =   555.024
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   1019.906
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   5910.236
            EndProperty
            BeginProperty Column03 
            EndProperty
            BeginProperty Column04 
               ColumnWidth     =   794.835
            EndProperty
            BeginProperty Column05 
               ColumnWidth     =   1335.118
            EndProperty
         EndProperty
      End
      Begin MSDataListLib.DataCombo dtcCarrera 
         Bindings        =   "frmMatriculacion.frx":174B
         Height          =   360
         Left            =   120
         TabIndex        =   6
         Top             =   1440
         Width           =   11655
         _ExtentX        =   20558
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
      Begin VB.Label lblDivision 
         BackStyle       =   0  'Transparent
         Caption         =   "Div."
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
         Left            =   1200
         TabIndex        =   12
         Top             =   1920
         Width           =   615
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "División:"
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
         TabIndex        =   11
         Top             =   1920
         Width           =   1095
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
         TabIndex        =   9
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Materias en las que se matriculó para este año:"
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
         Top             =   3960
         Width           =   5775
      End
      Begin VB.Label lblTitulo 
         BackStyle       =   0  'Transparent
         Caption         =   "Matriculación"
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
         Height          =   615
         Left            =   3600
         TabIndex        =   5
         Top             =   360
         Width           =   6375
      End
   End
   Begin VB.Frame frDatosAlumno 
      BackColor       =   &H0080FFFF&
      Height          =   1215
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   11895
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
         TabIndex        =   4
         Top             =   360
         Width           =   7575
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
         TabIndex        =   3
         Top             =   360
         Width           =   1695
      End
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
         TabIndex        =   2
         Top             =   360
         Width           =   855
      End
   End
End
Attribute VB_Name = "frmMatriculacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public CargarMaterias As String 'para saber si ya se cargaron las materias en el formulario de AgregarMatriculacion
Dim Conexion As New Connection
Dim MateriasPorAño As New Recordset

Private Sub cmdAgregar_Click()
If CargarMaterias = "Si" Then
    Me.MousePointer = 11
    With frmConexionAlumnos
    frmAgregarMatriculacion.adoMaterias.RecordSource = "SELECT Materias.Codigo, Materias.Curso, Materias.Nombre AS Materia, Personal.Nombre AS Profesor, Divisiones.Profesor AS CodigoProfesor, Divisiones.Salon FROM (Divisiones INNER JOIN Materias ON Divisiones.Materia = Materias.Codigo) INNER JOIN Personal ON Divisiones.Profesor = Personal.Codigo Where Divisiones.Ano = " & frmConexionAlumnos.adoParametros.Recordset!AñoMatriculacion & " And Materias.Carrera = " & dtcCarrera.BoundText & " And Divisiones.Division = " & adoCarreras.Recordset!Division & " ORDER BY Materias.Curso, Materias.Codigo"
    frmAgregarMatriculacion.adoMaterias.Refresh
    End With
    Me.MousePointer = 0
    CargarMaterias = "No"
End If
    If frmAgregarMatriculacion.adoMaterias.Recordset.RecordCount = 0 Then
        MsgBox ("Por el momento no existen materias disponibles para matricularse")
    Else
        frmAgregarMatriculacion.Show 1
    End If
End Sub

Private Sub cmdBorrarme_Click()
    Respuesta = MsgBox("Esta seguro de borrarse de la cursada de: " & Chr(13) & adoMatriculacion.Recordset!Materia, vbYesNo, "Borrase")
    If Respuesta = vbYes Then
        Me.MousePointer = 11
        Conexion.Open
        Conexion.Execute ("DELETE * FROM Finales WHERE Alumno = " & frmConexionAlumnos.adoAlumnos.Recordset!Permiso & " AND Materia = " & adoMatriculacion.Recordset!Codigo & " AND Ano = " & frmConexionAlumnos.adoParametros.Recordset!AñoMatriculacion)
        Conexion.Close
        dtcCarrera_Change
        frmMatriculacion.MousePointer = 0
    End If
End Sub

Private Sub cmdCursarLibre_Click()
    If adoMatriculacion.Recordset!Libre = False Then
       MsgBox ("Esta asignatura no se encuentra habilitada para cursarla libre")
       Exit Sub
    End If
    Conexion.Open
    Set MateriasPorAño = Conexion.Execute("SELECT Count(Materias.Curso) AS Total, (SELECT Count(Materias.Curso) AS Expr1 FROM Finales INNER JOIN Materias ON Finales.Materia = Materias.Codigo WHERE (((Finales.Alumno)=" & frmConexionAlumnos.adoAlumnos.Recordset!Permiso & ") AND ((Finales.Ano)=" & frmConexionAlumnos.adoParametros.Recordset!AñoMatriculacion & ") AND ((Materias.Curso)=" & adoMatriculacion.Recordset!Curso & ") AND ((Finales.Libre)=True))) AS Inscripto FROM Materias WHERE (((Materias.Curso)=" & adoMatriculacion.Recordset!Curso & ") AND ((Materias.Carrera)=" & dtcCarrera.BoundText & ") AND ((Materias.Detalle)=1 Or (Materias.Detalle)=2) AND ((Materias.Eliminada)=0))")
    If (MateriasPorAño!Inscripto + 1) / MateriasPorAño!Total > 0.3 Then
       MsgBox ("Ha superado el 30 % de asignaturas libres para este año")
       Conexion.Close
       Exit Sub
    End If
    Conexion.Close
    Respuesta = MsgBox("Esta seguro cursar libre esta asignatura?: " & Chr(13) & adoMatriculacion.Recordset!Materia, vbYesNo, "Atención")
    If Respuesta = vbYes Then
        Me.MousePointer = 11
        Conexion.Open
        Conexion.Execute ("UPDATE Finales set Libre = 1 WHERE Alumno = " & frmConexionAlumnos.adoAlumnos.Recordset!Permiso & " AND Materia = " & adoMatriculacion.Recordset!Codigo & " AND Ano = " & frmConexionAlumnos.adoParametros.Recordset!AñoMatriculacion)
        Conexion.Close
        dtcCarrera_Change
        frmMatriculacion.MousePointer = 0
    End If
End Sub

Private Sub cmdPresencial_Click()
    Respuesta = MsgBox("Esta seguro cursar presencial esta asignatura?: " & Chr(13) & adoMatriculacion.Recordset!Materia, vbYesNo, "Atención")
    If Respuesta = vbYes Then
        Me.MousePointer = 11
        Conexion.Open
        Conexion.Execute ("UPDATE Finales set Libre = 0 WHERE Alumno = " & frmConexionAlumnos.adoAlumnos.Recordset!Permiso & " AND Materia = " & adoMatriculacion.Recordset!Codigo & " AND Ano = " & frmConexionAlumnos.adoParametros.Recordset!AñoMatriculacion)
        Conexion.Close
        dtcCarrera_Change
        frmMatriculacion.MousePointer = 0
    End If

End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Public Sub dtcCarrera_Change()
    If adoCarreras.Recordset.RecordCount > 1 Then
        adoCarreras.Recordset.MoveFirst
        adoCarreras.Recordset.Find ("Carrera=" & dtcCarrera.BoundText)
    End If
    lblDivision = adoCarreras.Recordset!Division
    frmMatriculacion.MousePointer = 11
    adoMatriculacion.RecordSource = "SELECT Materias.Curso, Materias.Codigo, Materias.Nombre AS Materia, Personal.Nombre AS Profesor, Divisiones.Salon, finales.libre As Condicion, Divisiones.Libre as Libre FROM ((Finales INNER JOIN Divisiones ON (Finales.Division = Divisiones.Division) AND (Finales.Ano = Divisiones.Ano) AND (Finales.Materia = Divisiones.Materia)) INNER JOIN Personal ON Divisiones.Profesor = Personal.Codigo) INNER JOIN Materias ON Divisiones.Materia = Materias.Codigo Where (((Finales.Alumno) = " & frmConexionAlumnos.adoAlumnos.Recordset!Permiso & ") And ((Divisiones.Ano) = " & frmConexionAlumnos.adoParametros.Recordset!AñoMatriculacion & ") And ((Materias.Carrera) = " & dtcCarrera.BoundText & ")) ORDER BY Materias.Curso, Materias.Codigo"
    adoMatriculacion.Refresh
    cmdAgregar.Enabled = True
    If adoMatriculacion.Recordset.RecordCount > 0 Then
        cmdBorrarme.Visible = True
    Else
        cmdBorrarme.Visible = False
    End If
    CargarMaterias = "Si"
    frmMatriculacion.MousePointer = 0
End Sub

Private Sub Form_GotFocus()
    dtgInscripto.SetFocus
End Sub

Private Sub Form_Load()
    With frmConexionAlumnos
    lblNombre = .lblNombre
    lblTipo = .lblTipo
    lblDocumento = .lblDocumento
    adoCarreras.RecordSource = "SELECT CarrerasHechas.Permiso, CarrerasHechas.Carrera, CarrerasHechas.Ingreso, CarrerasHechas.Division, Condicion.Condicion, CarrerasHechas.Fecha, CarrerasHechas.Libro, CarrerasHechas.Folio, Carreras.Nombre FROM (CarrerasHechas INNER JOIN Carreras ON CarrerasHechas.Carrera = Carreras.Codigo) INNER JOIN Condicion ON CarrerasHechas.Condición = Condicion.Codigo WHERE CarrerasHechas.Permiso=" & .adoAlumnos.Recordset!Permiso & " AND ((CarrerasHechas.Condición)=1 Or (CarrerasHechas.Condición)=4 Or (CarrerasHechas.Condición)=6)"
    adoCarreras.Refresh
    dtcCarrera.BoundText = adoCarreras.Recordset!Carrera
    If adoCarreras.Recordset.RecordCount > 1 Then dtcCarrera.Enabled = True
    End With
    Conexion.ConnectionString = ("DSN=Instituto")
    CargarMaterias = "Si"
End Sub

