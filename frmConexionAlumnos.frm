VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmConexionAlumnos 
   BackColor       =   &H80000008&
   ClientHeight    =   8880
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   11880
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   8880
   ScaleWidth      =   11880
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame frMenu 
      BackColor       =   &H00C0FFFF&
      Height          =   7695
      Left            =   0
      TabIndex        =   5
      Top             =   1200
      Visible         =   0   'False
      Width           =   11895
      Begin VB.CommandButton cmdSituacionAcademica 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Situación Académica"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   6480
         MouseIcon       =   "frmConexionAlumnos.frx":0000
         MousePointer    =   99  'Custom
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   1200
         Width           =   4335
      End
      Begin VB.CommandButton cmdMensajesAlumnos 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Mensajes"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   960
         MouseIcon       =   "frmConexionAlumnos.frx":030A
         MousePointer    =   99  'Custom
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   6000
         Width           =   4335
      End
      Begin VB.CommandButton cmdInscripcionFinales 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Inscripción a Exámenes Finales"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   960
         MouseIcon       =   "frmConexionAlumnos.frx":0614
         MousePointer    =   99  'Custom
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   2400
         Width           =   4335
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
         Height          =   1215
         Left            =   8760
         MouseIcon       =   "frmConexionAlumnos.frx":091E
         MousePointer    =   99  'Custom
         Picture         =   "frmConexionAlumnos.frx":0C28
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   6120
         Width           =   1695
      End
      Begin VB.CommandButton cmdCambiarContraseña 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Cambiar Contraseña"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   960
         MouseIcon       =   "frmConexionAlumnos.frx":106A
         MousePointer    =   99  'Custom
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   4800
         Width           =   4335
      End
      Begin VB.CommandButton cmdMatriculacion 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Matriculación"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   960
         MouseIcon       =   "frmConexionAlumnos.frx":1374
         MousePointer    =   99  'Custom
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   3600
         Width           =   4335
      End
      Begin VB.CommandButton cmdCursadasyAsistencias 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFC0&
         Caption         =   "Cursadas y Asistencias"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   960
         MaskColor       =   &H00FFFFFF&
         MouseIcon       =   "frmConexionAlumnos.frx":167E
         MousePointer    =   99  'Custom
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   1200
         UseMaskColor    =   -1  'True
         Width           =   4335
      End
      Begin VB.Label lblNotificacion 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   615
         Left            =   480
         TabIndex        =   13
         Top             =   360
         Width           =   10815
      End
   End
   Begin MSAdodcLib.Adodc adoAlumnos 
      Height          =   330
      Left            =   2760
      Top             =   360
      Visible         =   0   'False
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   2
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
      RecordSource    =   "SELECT * FROM Alumnos WHERE Permiso = 0"
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
   Begin VB.Frame frDatosAlumno 
      BackColor       =   &H0080FFFF&
      Height          =   1215
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   11895
      Begin VB.TextBox Text1 
         DataField       =   "TurnoLlamado"
         DataSource      =   "adoParametros"
         Height          =   285
         Left            =   7200
         TabIndex        =   4
         Text            =   "Text1"
         Top             =   720
         Visible         =   0   'False
         Width           =   375
      End
      Begin MSAdodcLib.Adodc adoParametros 
         Height          =   330
         Left            =   5040
         Top             =   720
         Visible         =   0   'False
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   582
         ConnectMode     =   1
         CursorLocation  =   3
         IsolationLevel  =   -1
         ConnectionTimeout=   15
         CommandTimeout  =   30
         CursorType      =   2
         LockType        =   1
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
         TabIndex        =   3
         Top             =   360
         Width           =   855
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
         TabIndex        =   1
         Top             =   360
         Width           =   1695
      End
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
         TabIndex        =   2
         Top             =   360
         Width           =   7575
      End
   End
End
Attribute VB_Name = "frmConexionAlumnos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Control As New Connection
Dim Obligadas As New Recordset
Dim Cancelada As New Recordset
Public SalirPorCompleto As String
Dim AportesNoCancelados As String

Private Sub cmdCambiarContraseña_Click()
    frmCambiarContraseñaAlumnos.Show 1
End Sub

Private Sub cmdCursadasyAsistencias_Click()
    Control.Open
    Set Obligadas = Control.Execute("SELECT [Control Final].Codigo, [Cooperadora Conceptos].Concepto, [Control Final].Año FROM [Control Final] INNER JOIN [Cooperadora Conceptos] ON [Control Final].Codigo = [Cooperadora Conceptos].Codigo ORDER BY  [Control Final].Año,  [Control Final].Codigo")
    If Obligadas.EOF = False Then 'hay que controlar
        While Obligadas.EOF = False
            Set Cancelada = Control.Execute("SELECT Cancelado From [Cooperadora Pagos] Where Alumno= " & adoAlumnos.Recordset!Permiso & " And Año= " & Obligadas!Año & " And Concepto = " & Obligadas!Codigo & " And Cancelado = True")
            If Cancelada.EOF = True Then 'el aporte no fue cancelado
                AportesNoCancelados = AportesNoCancelados & Obligadas!Concepto & " " & Obligadas!Año & Chr(13)
            End If
            Obligadas.MoveNext
        Wend
        If AportesNoCancelados <> "" Then Respuesta = MsgBox("No se registran los siguientes aportes a Cooperadora" & Chr(13) & Chr(13) & AportesNoCancelados & Chr(13) & Chr(13) & "Las inscripciones que realice quedarán en forma condicional hasta tanto regularice su situación", , "Aportes a Cooperadora")
        AportesNoCancelados = ""
    End If
    Control.Close
    frmCursadasyAsistencias.Show 1
End Sub

Private Sub cmdInscripcionFinales_Click()
    Me.MousePointer = 11
    Control.Open
    Set Obligadas = Control.Execute("SELECT [Control Final].Codigo, [Cooperadora Conceptos].Concepto, [Control Final].Año FROM [Control Final] INNER JOIN [Cooperadora Conceptos] ON [Control Final].Codigo = [Cooperadora Conceptos].Codigo ORDER BY  [Control Final].Año,  [Control Final].Codigo")
    If Obligadas.EOF = False Then 'hay que controlar
        While Obligadas.EOF = False
            Set Cancelada = Control.Execute("SELECT Cancelado From [Cooperadora Pagos] Where Alumno= " & adoAlumnos.Recordset!Permiso & " And Año= " & Obligadas!Año & " And Concepto = " & Obligadas!Codigo & " And Cancelado = True")
            If Cancelada.EOF = True Then 'el aporte no fue cancelado
                AportesNoCancelados = AportesNoCancelados & Obligadas!Concepto & " " & Obligadas!Año & Chr(13)
            End If
            Obligadas.MoveNext
        Wend
        If AportesNoCancelados <> "" Then Respuesta = MsgBox("No se registran los siguientes aportes a Cooperadora" & Chr(13) & Chr(13) & AportesNoCancelados & Chr(13) & Chr(13) & "Las inscripciones que realice quedarán en forma condicional hasta tanto regularice su situación", , "Aportes a Cooperadora")
        AportesNoCancelados = ""
    End If
    Control.Close
    frmInscripcionFinales.Show 1
    Me.MousePointer = 0
End Sub

Private Sub cmdMatriculacion_Click()
'    MsgBox ("El período de matriculación ha finalizado")
 '   Exit Sub
    If frmConexionAlumnos.adoParametros.Recordset!Habilitar_matriculacion = 0 Then
       MsgBox ("Este trámite se encuentra deshabilitado")
       Exit Sub
    End If
    Respuesta = MsgBox("Si es ingresante a 1º año no debe realizar este trámite de matriculación, la institución lo hará de forma automática", vbOKOnly, "ATENCIÓN")
    frmMatriculacion.lblTitulo = "Matriculación año " & adoParametros.Recordset!AñoMatriculacion
    frmMatriculacion.Show 1
End Sub

Private Sub cmdMensajesAlumnos_Click()
   frmMensajesAlumnos.Show 1
End Sub

Private Sub cmdSalir_Click()
    frDatosAlumno.Visible = False
    frMenu.Visible = False
    frmLoginAlumno.Show 1
End Sub

Private Sub cmdSituacionAcademica_Click()
    frmSituacionAcademica.Show 1
End Sub

Private Sub Form_Activate()
    If SalirPorCompleto = "Si" Then Unload Me
    Control.ConnectionString = ("DSN=Instituto")
End Sub

Private Sub Form_Load()
    DisableKeys True
End Sub

Private Sub Form_Resize()
    If SalirPorCompleto = "No" Then frmLoginAlumno.Show 1
End Sub

Public Function ControlaCooperadora()
    Control.Open
    Set Obligadas = Control.Execute("SELECT [Control Final].Codigo, [Cooperadora Conceptos].Concepto, [Control Final].Año FROM [Control Final] INNER JOIN [Cooperadora Conceptos] ON [Control Final].Codigo = [Cooperadora Conceptos].Codigo ORDER BY  [Control Final].Año,  [Control Final].Codigo")
    If Obligadas.EOF = False Then 'hay que controlar
        While Obligadas.EOF = False
            Set Cancelada = Control.Execute("SELECT Cancelado From [Cooperadora Pagos] Where Alumno= " & adoAlumnos.Recordset!Permiso & " And Año= " & Obligadas!Año & " And Concepto = " & Obligadas!Codigo & " And Cancelado = True")
            If Cancelada.EOF = True Then 'el aporte no fue cancelado
                AportesNoCancelados = AportesNoCancelados & Obligadas!Concepto & " " & Obligadas!Año & Chr(13)
            End If
            Obligadas.MoveNext
        Wend
        If AportesNoCancelados <> "" Then Respuesta = MsgBox("No se registran los siguientes aportes a Cooperadora" & Chr(13) & Chr(13) & AportesNoCancelados & Chr(13) & Chr(13) & "Las inscripciones que realice quedarán en forma condicional hasta tanto regularice su situación", , "Aportes a Cooperadora")
        AportesNoCancelados = ""
    End If
    Control.Close
End Function

Private Sub Form_Unload(Cancel As Integer)
    DisableKeys False
End Sub

