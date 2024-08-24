VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmInscripcionFinales 
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
   Begin VB.Frame frInscripcionFinales 
      BackColor       =   &H00C0FFFF&
      Height          =   7815
      Left            =   0
      TabIndex        =   0
      Top             =   1200
      Width           =   11895
      Begin VB.Frame Frame1 
         BackColor       =   &H00C0FFFF&
         Height          =   1695
         Left            =   120
         TabIndex        =   12
         Top             =   1920
         Width           =   11655
         Begin VB.CommandButton cmdverTodasLasMaterias 
            BackColor       =   &H00C0C000&
            Caption         =   "Ver Todas las Mesas"
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
            Left            =   6600
            MouseIcon       =   "frmInscripciónFinales.frx":0000
            MousePointer    =   99  'Custom
            Style           =   1  'Graphical
            TabIndex        =   16
            ToolTipText     =   "Elija esta opción para borrarse de una mesa"
            Top             =   360
            Width           =   2295
         End
         Begin VB.CommandButton cmdAgregar 
            BackColor       =   &H00C0FFC0&
            Caption         =   "Click aquí para INSCRIBIRSE en una materia"
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
            Height          =   975
            Left            =   600
            MouseIcon       =   "frmInscripciónFinales.frx":030A
            MousePointer    =   99  'Custom
            Style           =   1  'Graphical
            TabIndex        =   15
            ToolTipText     =   "Elija esta opción para buscar las materias a rendir"
            Top             =   360
            UseMaskColor    =   -1  'True
            Width           =   2295
         End
         Begin VB.CommandButton cmdBorrarme 
            BackColor       =   &H00C0C0FF&
            Caption         =   "Click aquí para BORRARSE en la materia"
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
            Height          =   975
            Left            =   3600
            MouseIcon       =   "frmInscripciónFinales.frx":0614
            MousePointer    =   99  'Custom
            Style           =   1  'Graphical
            TabIndex        =   14
            ToolTipText     =   "Elija esta opción para borrarse de una mesa"
            Top             =   360
            Width           =   2295
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
            Height          =   960
            Left            =   9840
            MouseIcon       =   "frmInscripciónFinales.frx":091E
            MousePointer    =   99  'Custom
            Picture         =   "frmInscripciónFinales.frx":0C28
            Style           =   1  'Graphical
            TabIndex        =   13
            ToolTipText     =   "Salir"
            Top             =   360
            Width           =   1200
         End
      End
      Begin VB.TextBox Text1 
         DataField       =   "Mesa"
         DataSource      =   "adoInscripcion"
         Height          =   285
         Left            =   2640
         TabIndex        =   10
         Text            =   "Text1"
         Top             =   5160
         Visible         =   0   'False
         Width           =   255
      End
      Begin MSAdodcLib.Adodc adoInscripciones 
         Height          =   330
         Left            =   360
         Top             =   5160
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
         RecordSource    =   $"frmInscripciónFinales.frx":106A
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
         RecordSource    =   $"frmInscripciónFinales.frx":134F
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
         Bindings        =   "frmInscripciónFinales.frx":14CD
         Height          =   3375
         Left            =   120
         TabIndex        =   7
         Top             =   4080
         Width           =   11655
         _ExtentX        =   20558
         _ExtentY        =   5953
         _Version        =   393216
         AllowUpdate     =   0   'False
         HeadLines       =   1
         RowHeight       =   15
         RowDividerStyle =   0
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
         ColumnCount     =   8
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
            DataField       =   "Abreviatura"
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
         BeginProperty Column02 
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
         BeginProperty Column03 
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
         BeginProperty Column04 
            DataField       =   "Lugar"
            Caption         =   "Lugar"
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
            DataField       =   "Titular"
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
         BeginProperty Column06 
            DataField       =   "Integrante1"
            Caption         =   "1º Integrante"
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
         BeginProperty Column07 
            DataField       =   "Integrante2"
            Caption         =   "2º Integrante"
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
            RecordSelectors =   0   'False
            BeginProperty Column00 
               DividerStyle    =   0
               ColumnWidth     =   524,976
            EndProperty
            BeginProperty Column01 
               DividerStyle    =   0
               ColumnWidth     =   4155,024
            EndProperty
            BeginProperty Column02 
               DividerStyle    =   0
               ColumnWidth     =   900,284
            EndProperty
            BeginProperty Column03 
               DividerStyle    =   0
               ColumnWidth     =   555,024
            EndProperty
            BeginProperty Column04 
               DividerStyle    =   0
               ColumnWidth     =   854,929
            EndProperty
            BeginProperty Column05 
               DividerStyle    =   0
               ColumnWidth     =   1454,74
            EndProperty
            BeginProperty Column06 
               DividerStyle    =   0
               ColumnWidth     =   1425,26
            EndProperty
            BeginProperty Column07 
               DividerStyle    =   0
            EndProperty
         EndProperty
      End
      Begin MSDataListLib.DataCombo dtcCarrera 
         Bindings        =   "frmInscripciónFinales.frx":14EC
         Height          =   360
         Left            =   120
         TabIndex        =   6
         Top             =   1440
         Width           =   11535
         _ExtentX        =   20346
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
         Caption         =   "Lista de materias en las que se encuentra inscripto:"
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
         Top             =   3720
         Width           =   6375
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Inscripción a Exámenes Finales"
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
         Height          =   375
         Left            =   3120
         TabIndex        =   5
         Top             =   360
         Width           =   4695
      End
   End
   Begin MSComCtl2.MonthView Calendario 
      Height          =   2370
      Left            =   7440
      TabIndex        =   11
      Top             =   1440
      Visible         =   0   'False
      Width           =   2595
      _ExtentX        =   4577
      _ExtentY        =   4180
      _Version        =   393216
      ForeColor       =   -2147483630
      BackColor       =   12632256
      Appearance      =   1
      StartOfWeek     =   103088130
      CurrentDate     =   37600
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
Attribute VB_Name = "frmInscripcionFinales"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public CargarMesas As String 'para saber si ya se cargaron las mesas en el formulario de AgregarInscripcion
Dim Conexion As New Connection

Private Sub Calendario_DateClick(ByVal DateClicked As Date)
    Label4 = Calendario.DayOfWeek
End Sub

Private Sub cmdAgregar_Click()
If CargarMesas = "Si" Then
    Me.MousePointer = 11
    With frmConexionAlumnos
    frmAgregarInscripcion.adoMesas.RecordSource = "SELECT Materias.Codigo, Materias.Curso,Finales.Ano, Finales.Asistencia, Finales.PerdioTurno, Finales.CantidadMesas, Materias.Abreviatura, Mesas.Fecha, Mesas.Numero, Mesas.Hora, Mesas.Lugar, Mesas.Impresas, Personal.Nombre AS Titular, Personal_1.Nombre AS Integrante1, Personal_2.Nombre AS Integrante2, Finales.Libre FROM (Materias INNER JOIN (((Mesas INNER JOIN Personal ON Mesas.Titular = Personal.Codigo) INNER JOIN Personal AS Personal_1 ON Mesas.Integrante1 = Personal_1.Codigo) INNER JOIN Personal AS Personal_2 ON Mesas.Integrante2 = Personal_2.Codigo) ON Materias.Codigo = Mesas.Materia) INNER JOIN Finales  ON (Mesas.Division = Finales.Division) AND (Materias.Codigo = Finales.Materia) Where (((Mesas.Turno) = " & .adoParametros.Recordset!TurnoLlamado & ") And ((Mesas.Ano) = " & .adoParametros.Recordset!AñoLlamado & ") And ((Finales.Alumno) = " & .adoAlumnos.Recordset!Permiso & ") And ((Finales.Cursada) = True) And ((Finales.Aprobada) = False)" _
    & " And ((Finales.Promocion) = False) And ((Materias.Carrera) = " & dtcCarrera.BoundText & "))" & " AND ((Mesas.Fecha) >=#" & Format(Date, "mm/dd/yyyy") & "#) ORDER BY Materias.Curso, Mesas.Fecha"
    frmAgregarInscripcion.adoMesas.Refresh
    End With
    Me.MousePointer = 0
    CargarMesas = "No"
End If
If frmAgregarInscripcion.adoMesas.Recordset.RecordCount = 0 Then
    MsgBox ("Por el momento no existen mesas disponibles para inscribirse")
Else
    frmAgregarInscripcion.Show 1
End If
End Sub

Private Sub cmdBorrarme_Click()
    If adoInscripciones.Recordset!Impresas = True Then Respuesta = MsgBox("Ya se imprimieron las actas", 0, "Imposible Borrarse"): Exit Sub
    If adoInscripciones.Recordset!Fecha < Date Then
        Respuesta = MsgBox("No puede borrarse despues de la fecha de la mesa", 0, "Imposible Borrarse")
        Exit Sub
    'ElseIf adoInscripciones.Recordset!Fecha = Date Then
    '    If Calendario.DayOfWeek = mvwMonday Then 'si hoy es Lunes
    '        If adoInscripciones.Recordset!Hora <= Time Then Respuesta = MsgBox("Se sobrepasó la hora de la mesa", 0, "Imposible Borrarse"): Exit Sub
    '    Else
    '        Respuesta = MsgBox("Debe borrarse 24 hs. antes de la mesa a excepción del día Lunes", 0, "Imposible Borrarse")
    '        Exit Sub
    '    End If
    End If
    Mensaje = "¿Está seguro de borrarse de la mesa " & adoInscripciones.Recordset!Abreviatura & " ?"
    Estilo = vbYesNo + vbDefaultButton2   ' Define los botones.
    Título = "Atención"   ' Define el título.
    Respuesta = MsgBox(Mensaje, Estilo, Título)
    If Respuesta = vbYes Then
        'if adoInscripciones.Recordset!Dia
        Me.MousePointer = 11
        Conexion.Open
        Conexion.Execute ("UPDATE Inscripciones SET Inscripciones.FechaBorrado = #" & Format(Date, "mm/dd/yyyy") & "#, Inscripciones.HoraBorrado = #" & Format(Time, "hh:mm:ss") & "#, Inscripciones.Acta = 0, Inscripciones.Medio = 1 WHERE (((Inscripciones.Mesa)=" & adoInscripciones.Recordset!Numero & ") AND ((Inscripciones.Alumno)=" & frmConexionAlumnos.adoAlumnos.Recordset!Permiso & "))")
        Conexion.Close
        dtcCarrera_Change
    End If
    dtgInscripto.SetFocus
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Command1_Click()

End Sub

Private Sub cmdverTodasLasMaterias_Click()
    With frmConexionAlumnos
    frrVerTodasLasMesas.adoMesas.RecordSource = "SELECT Materias.Curso, Materias.Abreviatura, Mesas.Fecha, Mesas.Hora, Mesas.Lugar, Personal.Nombre AS Titular, Personal_1.Nombre AS Integrante1, Personal_2.Nombre AS Integrante2 FROM Materias INNER JOIN (((Mesas INNER JOIN Personal ON Mesas.Titular = Personal.Codigo) INNER JOIN Personal AS Personal_1 ON Mesas.Integrante1 = Personal_1.Codigo) INNER JOIN Personal AS Personal_2 ON Mesas.Integrante2 = Personal_2.Codigo) ON Materias.Codigo = Mesas.Materia " _
    & " WHERE (((Mesas.Turno)=" & .adoParametros.Recordset!TurnoLlamado & ") AND ((Mesas.Ano)=" & .adoParametros.Recordset!AñoLlamado & ") AND ((Materias.Carrera)=" & dtcCarrera.BoundText & ")) ORDER BY Materias.Curso, Materias.Abreviatura"
    frrVerTodasLasMesas.adoMesas.Refresh
    End With
    frrVerTodasLasMesas.Show 1
End Sub

Public Sub dtcCarrera_Change()
    frmInscripcionFinales.MousePointer = 11
    adoInscripciones.RecordSource = "SELECT Materias.Curso, Materias.Abreviatura, Mesas.Numero, Mesas.Fecha, Mesas.Hora, Mesas.Lugar, Mesas.Impresas, Personal.Nombre AS Titular, Personal_1.Nombre AS Integrante1, Personal_2.Nombre AS Integrante2 FROM ((((Inscripciones INNER JOIN Mesas ON Inscripciones.Mesa = Mesas.Numero) INNER JOIN Materias ON Mesas.Materia = Materias.Codigo) INNER JOIN Personal ON Mesas.Titular = Personal.Codigo) INNER JOIN Personal AS Personal_1 ON Mesas.Integrante1 = Personal_1.Codigo) INNER JOIN Personal AS Personal_2 ON Mesas.Integrante2 = Personal_2.Codigo Where (((Materias.Carrera) = " & dtcCarrera.BoundText & ") And ((Mesas.Turno) = " & frmConexionAlumnos.adoParametros.Recordset!TurnoLlamado & ") And ((Mesas.Ano) = " & frmConexionAlumnos.adoParametros.Recordset!AñoLlamado & ") And ((Inscripciones.FechaBorrado) Is Null) And ((Inscripciones.HoraBorrado) Is Null) And ((Inscripciones.Alumno) = " & frmConexionAlumnos.adoAlumnos.Recordset!Permiso & "))" _
    & " AND Mesas.Fecha >=#" & Format(Date, "mm/dd/yyyy") & "# ORDER BY Materias.Curso, Mesas.Fecha"
    adoInscripciones.Refresh
    cmdAgregar.Enabled = True
    If adoInscripciones.Recordset.RecordCount > 0 Then
        cmdBorrarme.Visible = True
        cmdBorrarme.Enabled = True
    Else
        cmdBorrarme.Visible = False
        cmdBorrarme.Enabled = False
    End If
    CargarMesas = "Si"
    Calendario.Value = Date
    frmInscripcionFinales.MousePointer = 0
End Sub

Private Sub Form_Activate()
    dtgInscripto.SetFocus
End Sub

Private Sub Form_Load()
    Conexion.ConnectionString = ("DSN=Instituto")
    With frmConexionAlumnos
    lblNombre = .lblNombre
    lblTipo = .lblTipo
    lblDocumento = .lblDocumento
    adoCarreras.RecordSource = "SELECT CarrerasHechas.Permiso, CarrerasHechas.Carrera, CarrerasHechas.Ingreso, Condicion.Condicion, CarrerasHechas.Fecha, CarrerasHechas.Libro, CarrerasHechas.Folio, Carreras.Nombre FROM (CarrerasHechas INNER JOIN Carreras ON CarrerasHechas.Carrera = Carreras.Codigo) INNER JOIN Condicion ON CarrerasHechas.Condición = Condicion.Codigo WHERE CarrerasHechas.Permiso=" & .adoAlumnos.Recordset!Permiso & " AND ((CarrerasHechas.Condición)=1 Or (CarrerasHechas.Condición)=4 Or (CarrerasHechas.Condición)=6)"
    adoCarreras.Refresh
    dtcCarrera.BoundText = adoCarreras.Recordset!Carrera
    If adoCarreras.Recordset.RecordCount > 1 Then dtcCarrera.Enabled = True
    End With
    CargarMesas = "Si"
End Sub

