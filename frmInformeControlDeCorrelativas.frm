VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Begin VB.Form frmInformeControlDeCorrelativas 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Control de Correlativas por Alumnos"
   ClientHeight    =   3480
   ClientLeft      =   3120
   ClientTop       =   4590
   ClientWidth     =   8895
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   3480
   ScaleWidth      =   8895
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   2055
      Left            =   0
      TabIndex        =   0
      Top             =   1200
      Width           =   8895
      Begin VB.CommandButton cmdControlaSoloPorCursadaVencida 
         Caption         =   "Controlar por cursada"
         Height          =   735
         Left            =   2400
         Picture         =   "frmInformeControlDeCorrelativas.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   1200
         Width           =   2295
      End
      Begin VB.TextBox txtAño 
         Height          =   315
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   1440
         Width           =   735
      End
      Begin VB.CommandButton cmdSalir 
         Height          =   720
         Left            =   7560
         Picture         =   "frmInformeControlDeCorrelativas.frx":066A
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Salir"
         Top             =   1200
         Width           =   720
      End
      Begin VB.CommandButton cmdMostrar 
         Caption         =   "Controlar por Final"
         Height          =   735
         Left            =   4800
         Picture         =   "frmInformeControlDeCorrelativas.frx":0AAC
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   1200
         Width           =   2295
      End
      Begin Crystal.CrystalReport rptInformeControlCorrelativas 
         Left            =   1560
         Top             =   1320
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   262150
         ReportFileName  =   "contcorr.rpt"
         WindowTitle     =   "Control de Correlativas por cursada"
         WindowControlBox=   -1  'True
         WindowMaxButton =   -1  'True
         WindowMinButton =   -1  'True
      End
      Begin MSAdodcLib.Adodc adoCarreras 
         Height          =   330
         Left            =   2280
         Top             =   480
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
         RecordSource    =   $"frmInformeControlDeCorrelativas.frx":1116
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
         Bindings        =   "frmInformeControlDeCorrelativas.frx":11B7
         Height          =   315
         Left            =   120
         TabIndex        =   3
         Top             =   720
         Width           =   7815
         _ExtentX        =   13785
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "Nombre"
         BoundColumn     =   "Codigo"
         Text            =   ""
      End
      Begin MSComCtl2.UpDown UpDown1 
         Height          =   375
         Left            =   840
         TabIndex        =   6
         Top             =   1440
         Width           =   240
         _ExtentX        =   450
         _ExtentY        =   661
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin VB.Label Label4 
         Caption         =   "Año:"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   1200
         Width           =   495
      End
      Begin VB.Label Label1 
         Caption         =   "Carreras Vigentes:"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   360
         Width           =   1455
      End
   End
   Begin VB.Label Label2 
      Caption         =   "El presente informa releva los alumnos que se encuentren matriculados por materia y aún no han aprobado alguna de sus correlativas"
      Height          =   495
      Left            =   480
      TabIndex        =   9
      Top             =   360
      Width           =   7935
   End
End
Attribute VB_Name = "frmInformeControlDeCorrelativas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Conexion As New Connection
Dim Materias As New Recordset
Dim Correlativas As New Recordset
Dim Matriculados As New Recordset
Dim Final As New Recordset

Private Sub cmdControlaSoloPorCursadaVencida_Click()
Respuesta = MsgBox("Este proceso puede tardar algunos minutos" & Chr(13) & "¿Desea continuar?", vbYesNo, "Atención")
If Respuesta = vbNo Then Exit Sub
Me.MousePointer = 11
Conexion.Open
Conexion.Execute ("DELETE * FROM rptControlCorrelativas")
Set Materias = Conexion.Execute("SELECT Materias.Codigo, Materias.Curso, Materias.Abreviatura From Materias Where (((Materias.Eliminada) = False) And ((Materias.Detalle) = 1 Or (Materias.Detalle) = 3 Or (Materias.Detalle) = 5) And ((Materias.Carrera) = " & dtcCarreras.BoundText & ")) ORDER BY Materias.Curso, Materias.Codigo")
While Materias.EOF = False  'recorro las materias una por una
    Set Correlativas = Conexion.Execute("SELECT Correlativas.Correlativa, Materias.Abreviatura, Materias.Curso FROM Materias INNER JOIN Correlativas ON Materias.Codigo = Correlativas.Correlativa Where Correlativas.Principal = " & Materias!Codigo & " ORDER BY Correlativas.Correlativa")
    If Correlativas.EOF = False Then 'si la materia tiene correlativas
        Set Matriculados = Conexion.Execute("SELECT Alumnos.Permiso, Alumnos.Nombre, Finales.Division FROM (((Finales INNER JOIN Alumnos ON Finales.Alumno = Alumnos.Permiso) INNER JOIN Materias ON Finales.Materia = Materias.Codigo) INNER JOIN CarrerasHechas ON (Materias.Carrera = CarrerasHechas.Carrera) AND (Alumnos.Permiso = CarrerasHechas.Permiso)) INNER JOIN Condicion ON CarrerasHechas.Condición = Condicion.Codigo Where (((Finales.Materia) = " & Materias!Codigo & ") And ((Finales.Ano) = " & txtAño & ")) ORDER BY Finales.Division, Alumnos.Nombre")
        While Matriculados.EOF = False
            Correlativas.MoveFirst
            While Correlativas.EOF = False
                Set Final = Conexion.Execute("SELECT Finales.Alumno, Finales.Materia, Finales.Aprobada From Finales WHERE (((Finales.Alumno)=" & Matriculados!Permiso & ") AND ((Finales.Materia)=" & Correlativas!Correlativa & ") AND ((Finales.Cursada)=True))")
                If Final.EOF = True Then
                    'debe la correlativa
                    Conexion.Execute ("INSERT INTO rptControlCorrelativas ( Carrera, Curso, Materia, Division, Permiso, Alumno, CursoCorrelativa, Correlativa ) VALUES ('" & dtcCarreras.Text & "'," & Materias!Curso & ",'" & Materias!Abreviatura & "'," & Matriculados!Division & "," & Matriculados!Permiso & ",'" & Matriculados!Nombre & "'," & Correlativas!Curso & ",'" & Correlativas!Abreviatura & "')")
                End If
                Correlativas.MoveNext
            Wend
        Matriculados.MoveNext
        Wend
    End If
    Materias.MoveNext
Wend
Conexion.Close
Me.MousePointer = 0
rptInformeControlCorrelativas.PrintReport
End Sub

Private Sub cmdMostrar_Click()
Respuesta = MsgBox("Este proceso puede tardar algunos minutos" & Chr(13) & "¿Desea continuar?", vbYesNo, "Atención")
If Respuesta = vbNo Then Exit Sub
Me.MousePointer = 11
Conexion.Open
Conexion.Execute ("DELETE * FROM rptControlCorrelativas")
Set Materias = Conexion.Execute("SELECT Materias.Codigo, Materias.Curso, Materias.Abreviatura From Materias Where (((Materias.Eliminada) = False) And ((Materias.Detalle) = 1 Or (Materias.Detalle) = 3 Or (Materias.Detalle) = 5) And ((Materias.Carrera) = " & dtcCarreras.BoundText & ")) ORDER BY Materias.Curso, Materias.Codigo")
While Materias.EOF = False  'recorro las materias una por una
    Set Correlativas = Conexion.Execute("SELECT Correlativas.Correlativa, Materias.Abreviatura, Materias.Curso FROM Materias INNER JOIN Correlativas ON Materias.Codigo = Correlativas.Correlativa Where Correlativas.Principal = " & Materias!Codigo & " ORDER BY Correlativas.Correlativa")
    If Correlativas.EOF = False Then 'si la materia tiene correlativas
        Set Matriculados = Conexion.Execute("SELECT Alumnos.Permiso, Alumnos.Nombre, Finales.Division FROM (((Finales INNER JOIN Alumnos ON Finales.Alumno = Alumnos.Permiso) INNER JOIN Materias ON Finales.Materia = Materias.Codigo) INNER JOIN CarrerasHechas ON (Materias.Carrera = CarrerasHechas.Carrera) AND (Alumnos.Permiso = CarrerasHechas.Permiso)) INNER JOIN Condicion ON CarrerasHechas.Condición = Condicion.Codigo Where (((Finales.Materia) = " & Materias!Codigo & ") And ((Finales.Ano) = " & txtAño & ")) ORDER BY Finales.Division, Alumnos.Nombre")
        While Matriculados.EOF = False
            Correlativas.MoveFirst
            While Correlativas.EOF = False
                Set Final = Conexion.Execute("SELECT Finales.Alumno, Finales.Materia, Finales.Aprobada From Finales WHERE (((Finales.Alumno)=" & Matriculados!Permiso & ") AND ((Finales.Materia)=" & Correlativas!Correlativa & ") AND ((Finales.Aprobada)=True))")
                If Final.EOF = True Then
                    'debe la correlativa
                    Conexion.Execute ("INSERT INTO rptControlCorrelativas ( Carrera, Curso, Materia, Division, Permiso, Alumno, CursoCorrelativa, Correlativa ) VALUES ('" & dtcCarreras.Text & "'," & Materias!Curso & ",'" & Materias!Abreviatura & "'," & Matriculados!Division & "," & Matriculados!Permiso & ",'" & Matriculados!Nombre & "'," & Correlativas!Curso & ",'" & Correlativas!Abreviatura & "')")
                End If
                Correlativas.MoveNext
            Wend
        Matriculados.MoveNext
        Wend
    End If
    Materias.MoveNext
Wend
Conexion.Close
Me.MousePointer = 0
rptInformeControlCorrelativas.PrintReport
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Conexion.ConnectionString = ("DSN=Instituto")
    txtAño = Format(Date, "yyyy")
    dtcCarreras.BoundText = adoCarreras.Recordset!Codigo
End Sub

Private Sub UpDown1_DownClick()
    txtAño = Val(txtAño) - 1
End Sub

Private Sub UpDown1_UpClick()
    txtAño = Val(txtAño) + 1
End Sub

