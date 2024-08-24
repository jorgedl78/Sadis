VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmAgregarInscripcion 
   BackColor       =   &H0080FFFF&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   5790
   ClientLeft      =   390
   ClientTop       =   1665
   ClientWidth     =   11175
   ControlBox      =   0   'False
   FillColor       =   &H00404040&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5790
   ScaleWidth      =   11175
   ShowInTaskbar   =   0   'False
   Begin MSAdodcLib.Adodc adoFinal 
      Height          =   330
      Left            =   3000
      Top             =   3240
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
      RecordSource    =   "SELECT Aprobada FROM Finales WHERE Alumno = 0 AND Materia=0 AND Cursada=1"
      Caption         =   "Final"
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
   Begin MSAdodcLib.Adodc adoCorrelativas 
      Height          =   330
      Left            =   720
      Top             =   3240
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
      RecordSource    =   $"frmAgregarInscripcion.frx":0000
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
   Begin MSDataGridLib.DataGrid dtgMesas 
      Bindings        =   "frmAgregarInscripcion.frx":0074
      Height          =   4095
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   10935
      _ExtentX        =   19288
      _ExtentY        =   7223
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
            Format          =   "HH:mm"
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
         Caption         =   "Integrante1"
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
         Caption         =   "Integrante2"
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
            ColumnAllowSizing=   -1  'True
            ColumnWidth     =   585,071
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   3240
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1049,953
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   569,764
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   915,024
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   1635,024
         EndProperty
         BeginProperty Column06 
         EndProperty
         BeginProperty Column07 
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdCancelar 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Cancelar"
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
      Left            =   6120
      MouseIcon       =   "frmAgregarInscripcion.frx":008B
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   5040
      Width           =   1935
   End
   Begin VB.CommandButton cmdAceptar 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Aceptar"
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
      Left            =   1440
      MouseIcon       =   "frmAgregarInscripcion.frx":0395
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   5040
      Width           =   1935
   End
   Begin MSAdodcLib.Adodc adoMesas 
      Height          =   330
      Left            =   3600
      Top             =   5160
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
      RecordSource    =   $"frmAgregarInscripcion.frx":069F
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
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Elija la materia en la que se quiere inscribir y presione ""Aceptar"""
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
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   240
      Width           =   10335
   End
End
Attribute VB_Name = "frmAgregarInscripcion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim DebeCorrelativa As String
Dim NombreCorrelativa(30) As String
Dim TotalCorrelativas As Integer
Dim TotalQueDebe As Integer
Dim Conexion As New Connection
Dim TurnoAnterior As New Recordset
Dim Auxiliar As New Recordset
Dim documento As New Recordset

Private Sub cmdAceptar_Click()
'    If adoMesas.Recordset!codigo = 540172 Or adoMesas.Recordset!codigo = 540263 Or adoMesas.Recordset!codigo = 540363 Then
'       Conexion.Open
'       Set documento = Conexion.Execute("SELECT Documento FROM Alumnos WHERE Permiso= " & frmConexionAlumnos.adoAlumnos.Recordset!Permiso)
'       If documento!documento = 33479764 Or documento!documento = 30928396 Or documento!documento = 34146671 Or documento!documento = 34542783 Or documento!documento = 33951670 Or documento!documento = 32923975 Or documento!documento = 32923713 Or documento!documento = 32920423 Or documento!documento = 30277465 Or documento!documento = 33416165 Or documento!documento = 33335306 Or documento!documento = 33096742 Or documento!documento = 32363758 Or documento!documento = 32781208 Then
'           Respuesta = MsgBox("Consulte al personal administrativo", vbOKOnly, "Imposible inscribirse en esta materia")
'           Conexion.Close
'           Exit Sub
'       End If
    '
    '   If documento!documento = 30928487 Or documento!documento = 30573464 Or documento!documento = 32772598 Or documento!documento = 31526741 Or documento!documento = 31813759 Or documento!documento = 32772612 Or documento!documento = 29652084 Or documento!documento = 30573238 Or documento!documento = 31062070 Or documento!documento = 32988393 Or documento!documento = 32195521 Or documento!documento = 32456531 Or documento!documento = 31730022 Or documento!documento = 30928291 Or documento!documento = 32527486 Or documento!documento = 31062300 Or documento!documento = 32209746 Or documento!documento = 32209707 Or documento!documento = 31114903 Or documento!documento = 32773029 Or documento!documento = 31941334 Or documento!documento = 32066360 Then
    '          Respuesta = MsgBox("Consulte al personal administrativo", vbOKOnly, "Imposible inscribirse en esta materia")
    '          Conexion.Close
    '          Exit Sub
    '   End If
    '   Conexion.Close
    'End If
    If adoMesas.Recordset!Impresas = True Then Respuesta = MsgBox("Ya se imprimieron las actas", 0, "Imposible Inscribirse"): Exit Sub
    Me.MousePointer = 11
    If adoMesas.Recordset!Asistencia = "Falso" Then
        Respuesta = MsgBox("No ha aprobado la asistencia para esta materia" & Chr(13) & "Debe rendir recuperatorio de faltas", vbOKOnly, "Imposible inscribirse")
        Me.MousePointer = 0
        Exit Sub
    End If
    If adoMesas.Recordset!PerdioTurno = "Verdadero" Then
        Respuesta = MsgBox("Ha perdido un turno por estar ausente en la mesa anterior", vbOKOnly, "Imposible inscribirse")
        Me.MousePointer = 0
        Exit Sub
    End If
    If frmInscripcionFinales.adoInscripciones.Recordset.RecordCount > 0 Then 'si ya se inscribió en alguna materia
        frmInscripcionFinales.adoInscripciones.Recordset.MoveFirst
        frmInscripcionFinales.adoInscripciones.Recordset.Find ("Numero=" & adoMesas.Recordset!Numero)
        If frmInscripcionFinales.adoInscripciones.Recordset.EOF = True Or frmInscripcionFinales.adoInscripciones.Recordset.BOF = True Then
            'no encontro la mesa en las inscripciones
        ElseIf frmInscripcionFinales.adoInscripciones.Recordset!Numero = adoMesas.Recordset!Numero Then
            MsgBox ("Ya se inscribió en esta materia")
            Me.MousePointer = 0
            Exit Sub
        End If
        
        'evito inscripcion en misma fecha y hora
        'With frmInscripcionFinales.adoInscripciones.Recordset
        '.MoveFirst
        'For i = 1 To frmInscripcionFinales.adoInscripciones.Recordset.RecordCount
        '    If !Fecha = adoMesas.Recordset!Fecha And !Hora = adoMesas.Recordset!Hora Then MsgBox ("Ya se inscribio en una mesa en esta misma fecha y horario"): Me.MousePointer = 0: Exit Sub
        '    .MoveNext
        'Next i
        'End With
    
    End If
    
    'controlo que haya limite disponible de inscriptos
    Conexion.Open
    'MsgBox ("select LimiteInscriptos, (SELECT count(Inscripciones.Alumno) FROM Inscripciones WHERE Inscripciones.Mesa=" & adoMesas.Recordset!Numero & " AND Inscripciones.FechaBorrado Is Null) as Inscriptos from mesas where numero=" & adoMesas.Recordset!Numero)
    'MsgBox ("select LimiteInscriptos, (SELECT count(Inscripciones.Alumno) FROM Inscripciones WHERE Inscripciones.Mesa=" & adoMesas.Recordset!Numero & " AND Inscripciones.FechaBorrado Is Null) as Inscriptos from mesas where numero=" & adoMesas.Recordset!Numero)
    Set Auxiliar = Conexion.Execute("select LimiteInscriptos, (SELECT count(Inscripciones.Alumno) FROM Inscripciones WHERE Inscripciones.Mesa=" & adoMesas.Recordset!Numero & " AND Inscripciones.FechaBorrado Is Null) as Inscriptos from mesas where numero=" & adoMesas.Recordset!Numero)
    'Auxiliar.MoveNext
    If (Auxiliar!Inscriptos >= Auxiliar!LimiteInscriptos) Then
        Respuesta = MsgBox("Por el momento se llegó al límite máximo de inscriptos permitidos", vbOKOnly, "Imposible inscribirse"): Conexion.Close: Me.MousePointer = 0: Exit Sub
    End If
    Conexion.Close
    
    
    'controlo que no se haya anotado en la misma mesa de este turno en otra fecha
    Conexion.Open
    Set Auxiliar = Conexion.Execute("SELECT DISTINCT Mesas.Materia FROM Inscripciones INNER JOIN Mesas ON Inscripciones.Mesa = Mesas.Numero WHERE Mesas.Materia=" & adoMesas.Recordset!Codigo & " AND Mesas.Turno=" & frmConexionAlumnos.adoParametros.Recordset!TurnoLlamado & " AND Mesas.Ano=" & frmConexionAlumnos.adoParametros.Recordset!AñoLlamado & "  AND Inscripciones.Alumno=" & frmConexionAlumnos.adoAlumnos.Recordset!Permiso & " AND Inscripciones.FechaBorrado Is Null")
    If Auxiliar.EOF = False Then Respuesta = MsgBox("Ya se inscribio en este turno", vbOKOnly, "Imposible inscribirse"): Conexion.Close: Me.MousePointer = 0: Exit Sub
    Conexion.Close
    
    
    'controlo correlativas
    adoCorrelativas.RecordSource = "SELECT Correlativas.Principal, Correlativas.Correlativa, Materias.Nombre, Materias.Curso FROM Correlativas INNER JOIN Materias ON Correlativas.Correlativa = Materias.Codigo Where (((Correlativas.Principal) = " & adoMesas.Recordset!Codigo & "))ORDER BY Correlativas.Correlativa    "
    adoCorrelativas.Refresh
    If adoCorrelativas.Recordset.RecordCount > 0 Then 'la materia elegida tiene correlativas
       'levanto las materias que son correlativas pero tiene aprobadas
        AdoFinal.RecordSource = "SELECT Correlativas.Correlativa FROM Correlativas INNER JOIN Finales ON Correlativas.Correlativa = Finales.Materia Where (((Correlativas.Principal) = " & adoMesas.Recordset!Codigo & ") And ((Finales.Alumno) = " & frmConexionAlumnos.adoAlumnos.Recordset!Permiso & ") And ((Finales.Aprobada) = True))"
        AdoFinal.Refresh
        TotalQueDebe = adoCorrelativas.Recordset.RecordCount - AdoFinal.Recordset.RecordCount
        If TotalQueDebe > 0 Then 'Debe alguna correlativa
            TotalCorrelativas = 0
            If AdoFinal.Recordset.RecordCount = 0 Then 'debe todas las correlativas
                adoCorrelativas.Recordset.MoveFirst
                For i = 1 To TotalQueDebe
                    NombreCorrelativa(i) = adoCorrelativas.Recordset!Curso & "°- " & adoCorrelativas.Recordset!Nombre
                    adoCorrelativas.Recordset.MoveNext
                Next i
            Else ' debe solo alguna/s correlativa/s
                adoCorrelativas.Recordset.MoveFirst
                For i = 1 To adoCorrelativas.Recordset.RecordCount
                    AdoFinal.Recordset.MoveFirst
                    AdoFinal.Recordset.Find ("Correlativa=" & adoCorrelativas.Recordset!Correlativa)
                    If AdoFinal.Recordset.BOF = True Or AdoFinal.Recordset.EOF = True Then 'no encontro la materia en los finales aprobados (entonces la debe)
                        TotalCorrelativas = TotalCorrelativas + 1
                        NombreCorrelativa(TotalCorrelativas) = adoCorrelativas.Recordset!Curso & "°- " & adoCorrelativas.Recordset!Nombre
                        AdoFinal.Recordset.MoveFirst
                    End If
                    adoCorrelativas.Recordset.MoveNext
                Next i
            End If
            StrinCorrelativa = ""
            For i = 1 To TotalQueDebe
                StrinCorrelativas = StrinCorrelativas & NombreCorrelativa(i) & Chr(13) & Chr(13)
            Next i
            MsgBox ("No se puede inscribir: debe las siguientes correlativas:" & Chr(13) & Chr(13) & StrinCorrelativas)
            Me.MousePointer = 0
            Exit Sub
        End If
    End If
    
    'chequeo que no se haya inscripto en el turno anterior segun parametro
    If frmConexionAlumnos.adoParametros.Recordset!ControlaInscripcionAnterior = True Then
        Conexion.Open
        Set TurnoAnterior = Conexion.Execute("SELECT Inscripciones.Acta FROM Inscripciones INNER JOIN Mesas ON Inscripciones.Mesa = Mesas.Numero WHERE Inscripciones.Alumno=" & frmConexionAlumnos.adoAlumnos.Recordset!Permiso & " AND Mesas.Materia=" & adoMesas.Recordset!Codigo & " AND Mesas.Turno=" & frmConexionAlumnos.adoParametros.Recordset!TurnoControl & " AND Mesas.Ano=" & frmConexionAlumnos.adoParametros.Recordset!AñoControl & " AND Inscripciones.FechaBorrado Is Null")
        If TurnoAnterior.EOF = False Then Respuesta = MsgBox("Ya se inscribio en el turno anterior", vbOKOnly, "Imposible inscribirse"): Conexion.Close: Me.MousePointer = 0: Exit Sub
        Conexion.Close
    End If
    
    'esta todo correcto y agrego la inscripción
    Conexion.Open
    If adoMesas.Recordset!Libre = 0 Then 'no es cursada libre
        Conexion.Execute ("INSERT INTO Inscripciones ( Mesa, Alumno, Cursada, FechaInscripto, HoraInscripto, Medio, Libre ) VALUES ( " & adoMesas.Recordset!Numero & ", " & frmConexionAlumnos.adoAlumnos.Recordset!Permiso & ", " & adoMesas.Recordset!Ano & ", '" & DateValue(Date) & "','" & TimeValue(Time) & "', 1,0)")
    Else
        Conexion.Execute ("INSERT INTO Inscripciones ( Mesa, Alumno, Cursada, FechaInscripto, HoraInscripto, Medio, Libre ) VALUES ( " & adoMesas.Recordset!Numero & ", " & frmConexionAlumnos.adoAlumnos.Recordset!Permiso & ", " & adoMesas.Recordset!Ano & ", '" & DateValue(Date) & "','" & TimeValue(Time) & "', 1,1)")
    End If
    Conexion.Close
    '.AddNew
    '!Mesa = adoMesas.Recordset!Numero
    '!Alumno = frmConexionAlumnos.adoAlumnos.Recordset!Permiso
    '!Cursada = adoMesas.Recordset!Ano
    '!FechaInscripto = Date
    '!HoraInscripto = Time
    '.Update
    '.Requery

    frmInscripcionFinales.dtcCarrera_Change
    Me.MousePointer = 0
    frmInscripcionFinales.CargarMesas = "No"
    Me.Hide
End Sub

Private Sub cmdCancelar_Click()
    Me.Hide
End Sub

Private Sub dtgMesas_GotFocus()
    cmdAceptar.Enabled = True
End Sub

Private Sub Form_Activate()
    dtgMesas.SetFocus
End Sub

Private Sub Form_Load()
    Conexion.ConnectionString = ("DSN=Instituto")
End Sub

