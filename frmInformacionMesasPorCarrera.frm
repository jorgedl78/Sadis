VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Begin VB.Form frmInformacionMesasPorCarrera 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Mesas de Exámenes Parciales por Carrera"
   ClientHeight    =   4170
   ClientLeft      =   1935
   ClientTop       =   3090
   ClientWidth     =   9660
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   4170
   ScaleWidth      =   9660
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   4095
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9615
      Begin VB.Frame Frame4 
         Caption         =   "Listado de Suplentes"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2415
         Left            =   6480
         TabIndex        =   21
         Top             =   1560
         Width           =   3015
         Begin VB.CommandButton cmdSuplentes 
            Caption         =   "Listado de suplentes"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   975
            Left            =   600
            Picture         =   "frmInformacionMesasPorCarrera.frx":0000
            Style           =   1  'Graphical
            TabIndex        =   22
            Top             =   1200
            Width           =   1575
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Listado General"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2415
         Left            =   3360
         TabIndex        =   17
         Top             =   1560
         Width           =   3015
         Begin VB.CommandButton cmdOrdenadoPorFecha 
            Caption         =   "Ordenado por Fecha"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   975
            Left            =   600
            Picture         =   "frmInformacionMesasPorCarrera.frx":066A
            Style           =   1  'Graphical
            TabIndex        =   20
            Top             =   1200
            Width           =   1575
         End
         Begin VB.OptionButton optCarrera 
            Caption         =   "Todas las Carreras"
            Height          =   255
            Index           =   0
            Left            =   480
            TabIndex        =   19
            Top             =   360
            Value           =   -1  'True
            Width           =   1815
         End
         Begin VB.OptionButton optCarrera 
            Caption         =   "Carrera Seleccionada"
            Height          =   255
            Index           =   1
            Left            =   480
            TabIndex        =   18
            Top             =   720
            Width           =   1935
         End
      End
      Begin Crystal.CrystalReport rptMesasPorCarrera 
         Left            =   2040
         Top             =   2760
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   262150
         ReportFileName  =   "mesascar.rpt"
         WindowTitle     =   "Mesas de Exámenes por Carrera"
         WindowControlBox=   -1  'True
         WindowMaxButton =   -1  'True
         WindowMinButton =   -1  'True
      End
      Begin VB.Frame Frame2 
         Caption         =   "Con Filtros"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2415
         Left            =   120
         TabIndex        =   10
         Top             =   1560
         Width           =   3015
         Begin VB.CommandButton cmdMostrar 
            Caption         =   "Con Filtros"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   975
            Left            =   480
            Picture         =   "frmInformacionMesasPorCarrera.frx":0CD4
            Style           =   1  'Graphical
            TabIndex        =   16
            Top             =   1200
            Width           =   1455
         End
         Begin MSAdodcLib.Adodc adoDivision 
            Height          =   330
            Left            =   960
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
            RecordSource    =   $"frmInformacionMesasPorCarrera.frx":133E
            Caption         =   "Division"
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
            Bindings        =   "frmInformacionMesasPorCarrera.frx":1406
            Height          =   315
            Left            =   1920
            TabIndex        =   14
            Top             =   600
            Width           =   735
            _ExtentX        =   1296
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            ListField       =   "Division"
            Text            =   "DataCombo1"
         End
         Begin VB.ComboBox cbCurso 
            Enabled         =   0   'False
            Height          =   315
            Left            =   960
            Style           =   2  'Dropdown List
            TabIndex        =   13
            Top             =   600
            Width           =   735
         End
         Begin VB.OptionButton optCurso 
            Caption         =   "Curso"
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   12
            Top             =   600
            Width           =   1095
         End
         Begin VB.OptionButton optCurso 
            Caption         =   "Todos"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   11
            Top             =   240
            Value           =   -1  'True
            Width           =   1095
         End
         Begin VB.Label Label3 
            Caption         =   "División:"
            Height          =   255
            Left            =   1920
            TabIndex        =   15
            Top             =   360
            Width           =   735
         End
      End
      Begin VB.TextBox txtAño 
         Height          =   315
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   480
         Width           =   735
      End
      Begin VB.TextBox Text3 
         DataField       =   "TurnoLlamado"
         DataSource      =   "adoParametros"
         Height          =   375
         Left            =   4560
         TabIndex        =   4
         Text            =   "Text3"
         Top             =   120
         Visible         =   0   'False
         Width           =   375
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
         RecordSource    =   $"frmInformacionMesasPorCarrera.frx":1420
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
      Begin VB.CommandButton cmdSalir 
         Height          =   960
         Left            =   8400
         Picture         =   "frmInformacionMesasPorCarrera.frx":14C1
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Salir"
         Top             =   240
         Width           =   960
      End
      Begin MSDataListLib.DataCombo dtcCarreras 
         Bindings        =   "frmInformacionMesasPorCarrera.frx":1903
         Height          =   315
         Left            =   120
         TabIndex        =   2
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
         Bindings        =   "frmInformacionMesasPorCarrera.frx":191D
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
         TabIndex        =   9
         Top             =   360
         Width           =   240
         _ExtentX        =   450
         _ExtentY        =   661
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin VB.Label Label4 
         Caption         =   "Año:"
         Height          =   255
         Left            =   1680
         TabIndex        =   8
         Top             =   240
         Width           =   495
      End
      Begin VB.Label Label2 
         Caption         =   "Turno:"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Carreras Vigentes:"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   960
         Width           =   1455
      End
   End
End
Attribute VB_Name = "frmInformacionMesasPorCarrera"
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
    Conexion.Execute ("DELETE * FROM [rpt Mesas Por Carrera]")
    If optCurso(0).Value = True Then 'se imprimen todos los cursos
        Conexion.Execute ("INSERT INTO [rpt Mesas Por Carrera] ( Carreras_Nombre, Curso, Abreviatura, Fecha, Hora, Lugar, Division, Turno, Ano, Titular, Integrante1, Integrante2, Carrera ) SELECT Carreras.Nombre, Materias.Curso, Materias.Abreviatura, Mesas.Fecha, Mesas.Hora, Mesas.Lugar, Mesas.Division, Meses.Nombre, Mesas.Ano, Personal.Nombre, Personal_1.Nombre, Personal_2.Nombre, Materias.Carrera FROM (((((Mesas INNER JOIN Materias ON Mesas.Materia = Materias.Codigo) INNER JOIN Carreras ON Materias.Carrera = Carreras.Codigo) INNER JOIN Personal ON Mesas.Titular = Personal.Codigo) INNER JOIN Personal AS Personal_1 ON Mesas.Integrante1 = Personal_1.Codigo) INNER JOIN Personal AS Personal_2 ON Mesas.Integrante2 = Personal_2.Codigo) INNER JOIN Meses ON Mesas.Turno = Meses.Numero Where Mesas.Ano = " & txtAño & " And Materias.Carrera = " & dtcCarreras.BoundText & " And Mesas.Turno = " & dtcMeses.BoundText & " AND Mesas.Division = " & dtcDivision & "  ORDER BY Materias.Curso")
    Else 'se imprime un curso elegido
        Conexion.Execute ("INSERT INTO [rpt Mesas Por Carrera] ( Carreras_Nombre, Curso, Abreviatura, Fecha, Hora, Lugar, Division, Turno, Ano, Titular, Integrante1, Integrante2, Carrera ) SELECT Carreras.Nombre, Materias.Curso, Materias.Abreviatura, Mesas.Fecha, Mesas.Hora, Mesas.Lugar, Mesas.Division, Meses.Nombre, Mesas.Ano, Personal.Nombre, Personal_1.Nombre, Personal_2.Nombre, Materias.Carrera FROM (((((Mesas INNER JOIN Materias ON Mesas.Materia = Materias.Codigo) INNER JOIN Carreras ON Materias.Carrera = Carreras.Codigo) INNER JOIN Personal ON Mesas.Titular = Personal.Codigo) INNER JOIN Personal AS Personal_1 ON Mesas.Integrante1 = Personal_1.Codigo) INNER JOIN Personal AS Personal_2 ON Mesas.Integrante2 = Personal_2.Codigo) INNER JOIN Meses ON Mesas.Turno = Meses.Numero Where Mesas.Ano = " & txtAño & " And Materias.Carrera = " & dtcCarreras.BoundText & " And Materias.Curso = " & cbCurso & " And Mesas.Turno = " & dtcMeses.BoundText & " AND Mesas.Division = " & dtcDivision & " ORDER BY Materias.Curso")
    End If
    Conexion.Execute ("UPDATE [rpt Mesas Por Carrera], Parametros SET [rpt Mesas Por Carrera].Institucion = [Parametros].[nombreinstitucion]")
    Set Resultado = Conexion.Execute("SELECT Curso FROM [rpt Mesas Por Carrera]")
    If Resultado.EOF = True Then MsgBox ("No se armaron mesas para este turno"): Conexion.Close: Exit Sub
    Conexion.Close
    rptMesasPorCarrera.PrintReport
End Sub

Private Sub cmdOrdenadoPorFecha_Click()

    If optCarrera(1).Value = True Then
        desdeCarrera = dtcCarreras.BoundText
        hastaCarrera = dtcCarreras.BoundText
        nombreCarrera = dtcCarreras.Text
    Else
        desdeCarrera = 0
        hastaCarrera = 99999
        nombreCarrera = "Todas las Carreras"
    End If
    cn.Open
    Set rs = cn.Execute("SELECT Carreras.Abreviatura as Carrera, Materias.Curso, Materias.Abreviatura as Asignatura, Mesas.Fecha, Mesas.Hora, Mesas.Lugar, Mesas.Division, Meses.Nombre, Mesas.Ano, Personal.Nombre as Titular, Personal_1.Nombre as Secundario, Personal_2.Nombre, Materias.Carrera FROM (((((Mesas INNER JOIN Materias ON Mesas.Materia = Materias.Codigo) INNER JOIN Carreras ON Materias.Carrera = Carreras.Codigo) INNER JOIN Personal ON Mesas.Titular = Personal.Codigo) INNER JOIN Personal AS Personal_1 ON Mesas.Integrante1 = Personal_1.Codigo) INNER JOIN Personal AS Personal_2 ON Mesas.Integrante2 = Personal_2.Codigo) INNER JOIN Meses ON Mesas.Turno = Meses.Numero Where Mesas.Ano = " & txtAño & " And Materias.Carrera between  " & desdeCarrera & " and " & hastaCarrera & " And Mesas.Turno = " & dtcMeses.BoundText & " ORDER BY Mesas.Fecha, Mesas.Hora, Carreras.Abreviatura")
    Set rptMesasOrdenadasPorFecha.DataSource = rs
    With rptMesasOrdenadasPorFecha.Sections("Sección4")
        .Controls("lblCarrera").Caption = nombreCarrera
        .Controls("lblTurno").Caption = dtcMeses.Text & " - " & txtAño.Text
        '.Controls("lblTelefonos").Caption = rs!Telefono
        'On Error Resume Next
        'Set .Controls("imgLogo").Picture = LoadPicture(rs!PathLogoReporte)
        'Set .Controls("imgSello").Picture = LoadPicture("sello.jpg")
        
        '.Controls("lblTextoCompleto").Caption = "Se deja constancia de que, a la fecha, " & txtAlumnoNombre & ", DNI: " & txtAlumnoDocumento & " es alumno/a regular " & rs!NombreINstitucion & " DISTRITO " & rs!Localidad & ", y cursa la carrera " & dtcCarreras.Text & Chr(13) & Chr(10) & Chr(13) & Chr(10) & "A pedido del interesado/a y para ser presentado ante las autoridades que correspondan, se extiende la presente en la ciudad de " & rs!Localidad & " Prov. de " & rs!Provincia & " a los " & dtFecha.Day & " dias del mes de " & adoMeses.Recordset!Nombre & " del año " & dtFecha.Year & ".-"
        
        '.Controls("lblTextoCompleto").Caption = "LA DIRECCIÓN DEL NIVEL SUPERIOR del " & rs!NombreINstitucion & " DISTRITO " & rs!Localidad & " CERTIFICA que " & txtAlumnoNombre & ", DNI: " & txtAlumnoDocumento & " es alumno regular de la carrera " & dtcCarreras.Text & Chr(13) & Chr(10) & Chr(13) & Chr(10) & "A pedido del interesado y para ser presentado ante las autoridades que correspondan, se extiende la presente en " & rs!Localidad & " Prov. de " & rs!Provincia & " a los " & dtFecha.Day & " dias del mes de " & adoMeses.Recordset!Nombre & " del año " & dtFecha.Year & ".-"
    End With
    
    
    rptMesasOrdenadasPorFecha.WindowState = 2
    rptMesasOrdenadasPorFecha.Show 1
    cn.Close
    Me.Refresh

End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Command1_Click()

End Sub

Private Sub cmdSuplentes_Click()
    cn.Open
    Set rs = cn.Execute("SELECT Mesas.Fecha, Mesas.Hora, Materias.Nombre as Asignatura, Personal.Nombre as Titular FROM (Mesas INNER JOIN Materias ON Mesas.Materia = Materias.Codigo) INNER JOIN Personal ON Mesas.Titular = Personal.Codigo WHERE (((Mesas.Turno)=" & dtcMeses.BoundText & ") AND ((Mesas.Ano)=" & txtAño & ") AND ((Materias.Carrera)=16)) ORDER BY Mesas.Fecha, Mesas.Hora")
    
    Set rptListadoSuplentesPorTurno.DataSource = rs
    With rptListadoSuplentesPorTurno.Sections("Sección4")
        .Controls("lblTurno").Caption = dtcMeses.Text & " - " & txtAño.Text
    End With
    
    
    rptListadoSuplentesPorTurno.WindowState = 2
    rptListadoSuplentesPorTurno.Show 1
    cn.Close
    Me.Refresh

End Sub

Private Sub dtcCarreras_Change()
    adoCarreras.Recordset.MoveFirst
    adoCarreras.Recordset.Find ("Codigo=" & dtcCarreras.BoundText)
    optCurso(0).Value = True
    cbCurso.Clear
    For i = 0 To adoCarreras.Recordset!Años - 1
        cbCurso.List(i) = i + 1
    Next i
    cbCurso.Text = cbCurso.List(0)
End Sub

Private Sub Form_Activate()
    Conexion.ConnectionString = ("DSN=Instituto")
    dtcMeses.BoundText = adoParametros.Recordset!TurnoLlamado
    txtAño = adoParametros.Recordset!AñoLlamado
    dtcCarreras.BoundText = adoCarreras.Recordset!Codigo
End Sub

Private Sub optCurso_Click(Index As Integer)
    If Index = 0 Then 'todos los cursos
        cbCurso.Enabled = False
        'dtcDivision.Enabled = False
    Else 'un curso en particular
        cbCurso.Enabled = True
        'dtcDivision.Enabled = True
    End If
End Sub

Private Sub UpDown1_DownClick()
    txtAño = Val(txtAño) - 1
End Sub

Private Sub UpDown1_UpClick()
    txtAño = Val(txtAño) + 1
End Sub
