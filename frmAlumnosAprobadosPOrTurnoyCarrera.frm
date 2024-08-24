VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "msadodc.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Begin VB.Form frmAlumnosAprobadosPOrTurnoyCarrera 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Alumnos Aprobados por Turno y Carrera"
   ClientHeight    =   2850
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8595
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2850
   ScaleWidth      =   8595
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   2655
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8055
      Begin VB.CommandButton cmdSalir 
         Height          =   720
         Left            =   6840
         Picture         =   "frmAlumnosAprobadosPOrTurnoyCarrera.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Salir"
         Top             =   1800
         Width           =   720
      End
      Begin VB.CommandButton cmdMostrar 
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
         Left            =   3360
         Picture         =   "frmAlumnosAprobadosPOrTurnoyCarrera.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   1800
         Width           =   975
      End
      Begin VB.TextBox Text3 
         DataField       =   "TurnoLlamado"
         DataSource      =   "adoParametros"
         Height          =   375
         Left            =   4560
         TabIndex        =   2
         Text            =   "Text3"
         Top             =   120
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.TextBox txtAño 
         Height          =   315
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   1
         Top             =   480
         Width           =   735
      End
      Begin Crystal.CrystalReport rptAlumnosAprobados 
         Left            =   5040
         Top             =   1680
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   262150
         ReportFileName  =   "aprobados.rpt"
         WindowTitle     =   "Mesas de Exámenes por Carrera"
         WindowControlBox=   -1  'True
         WindowMaxButton =   -1  'True
         WindowMinButton =   -1  'True
      End
      Begin MSAdodcLib.Adodc adoCarreras 
         Height          =   330
         Left            =   4440
         Top             =   840
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
         RecordSource    =   $"frmAlumnosAprobadosPOrTurnoyCarrera.frx":0AAC
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
         Bindings        =   "frmAlumnosAprobadosPOrTurnoyCarrera.frx":0B4D
         Height          =   315
         Left            =   120
         TabIndex        =   5
         Top             =   1200
         Width           =   7815
         _ExtentX        =   13785
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "Nombre"
         BoundColumn     =   "Codigo"
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo dtcMeses 
         Bindings        =   "frmAlumnosAprobadosPOrTurnoyCarrera.frx":0B67
         Height          =   315
         Left            =   120
         TabIndex        =   6
         Top             =   480
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "Nombre"
         BoundColumn     =   "Numero"
         Text            =   ""
      End
      Begin MSAdodcLib.Adodc adoParametros 
         Height          =   330
         Left            =   2640
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
      Begin MSAdodcLib.Adodc adoMeses 
         Height          =   330
         Left            =   2760
         Top             =   480
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
         Left            =   2400
         TabIndex        =   7
         Top             =   360
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   661
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin VB.Label Label1 
         Caption         =   "Carreras Vigentes:"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   960
         Width           =   1455
      End
      Begin VB.Label Label2 
         Caption         =   "Turno:"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label4 
         Caption         =   "Año:"
         Height          =   255
         Left            =   1680
         TabIndex        =   8
         Top             =   240
         Width           =   495
      End
   End
End
Attribute VB_Name = "frmAlumnosAprobadosPOrTurnoyCarrera"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Conexion As New Connection
Dim Resultado As New Recordset

Private Sub cbCurso_Click()
    adoDivision.RecordSource = "SELECT DISTINCT Mesas.Division FROM Mesas INNER JOIN Materias ON Mesas.Materia = Materias.Codigo Where Materias.Carrera = " & dtcCarreras.BoundText & " And Mesas.Turno = " & dtcMeses.BoundText & " And Mesas.Ano = " & txtAño & " And Materias.Curso = " & cbCurso & " ORDER BY Mesas.Division"
    adoDivision.Refresh
    If adoDivision.Recordset.RecordCount > o Then
        dtcDivision.Enabled = True
        dtcDivision = adoDivision.Recordset!Division
        cmdMostrar.Enabled = True
    Else
        MsgBox ("No se armaron mesas")
        cmdMostrar.Enabled = False
        dtcDivision.Enabled = False
    End If
End Sub

Private Sub cmdMostrar_Click()
    Conexion.Open
    Conexion.Execute ("DELETE * FROM rptAlumnosAprobadosPorTurnoYCarrera")
    Conexion.Execute ("INSERT INTO rptAlumnosAprobadosPorTurnoYCarrera ( Permiso, Alumno, Tipo, Documento, Libro, Folio, Turno, Ano, Carrera ) SELECT DISTINCT Alumnos.Permiso, Alumnos.Nombre, Alumnos.Tipo, Alumnos.Documento, CarrerasHechas.Libro, CarrerasHechas.Folio, Meses.Nombre, Mesas.Ano, Carreras.Nombre FROM (((((Alumnos INNER JOIN Inscripciones ON Alumnos.Permiso = Inscripciones.Alumno) INNER JOIN Mesas ON Inscripciones.Mesa = Mesas.Numero) INNER JOIN Materias ON Mesas.Materia = Materias.Codigo) INNER JOIN CarrerasHechas ON (Materias.Carrera = CarrerasHechas.Carrera) AND (Alumnos.Permiso = CarrerasHechas.Permiso)) INNER JOIN Meses ON Mesas.Turno = Meses.Numero) INNER JOIN Carreras ON CarrerasHechas.Carrera = Carreras.Codigo" _
    & " WHERE (((Inscripciones.FechaBorrado) Is Null) AND ((Mesas.Turno)=" & dtcMeses.BoundText & " ) AND ((Mesas.Ano)=" & txtAño & ") AND ((Materias.Carrera)=" & dtcCarreras.BoundText & ") AND ((Inscripciones.Nota)>=" & adoParametros.Recordset!NotaAprobacionFinal & ")) ORDER BY Alumnos.Nombre")
    Conexion.Close
    rptAlumnosAprobados.PrintReport
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub dtcCarreras_Change()
    adoCarreras.Recordset.MoveFirst
    adoCarreras.Recordset.Find ("Codigo=" & dtcCarreras.BoundText)
End Sub

Private Sub Form_Activate()
    Conexion.ConnectionString = ("DSN=Instituto")
    dtcMeses.BoundText = adoParametros.Recordset!TurnoLlamado
    txtAño = adoParametros.Recordset!AñoLlamado
    dtcCarreras.BoundText = adoCarreras.Recordset!Codigo
End Sub


Private Sub UpDown1_DownClick()
    txtAño = Val(txtAño) - 1
End Sub

Private Sub UpDown1_UpClick()
    txtAño = Val(txtAño) + 1
End Sub

