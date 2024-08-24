VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Begin VB.Form frmBajaCursadasVencidas 
   Caption         =   "Baja de Cursadas Vencidas"
   ClientHeight    =   4215
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   9405
   LinkTopic       =   "Form1"
   ScaleHeight     =   4215
   ScaleWidth      =   9405
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   2895
      Left            =   240
      TabIndex        =   0
      Top             =   1200
      Width           =   9015
      Begin VB.CheckBox chkIncluirCorrelativas 
         Alignment       =   1  'Right Justify
         Caption         =   "Incluir las asignaturas correlativas"
         Height          =   255
         Left            =   3120
         TabIndex        =   8
         Top             =   720
         Width           =   2895
      End
      Begin MSComctlLib.ProgressBar pbProcesando 
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   2520
         Visible         =   0   'False
         Width           =   8415
         _ExtentX        =   14843
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   1
      End
      Begin Crystal.CrystalReport rptPrevisualizar 
         Left            =   2160
         Top             =   1560
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   262150
         ReportFileName  =   "bajacurs.rpt"
         WindowTitle     =   "Cursadas a dar de baja"
         WindowControlBox=   -1  'True
         WindowMaxButton =   -1  'True
         WindowMinButton =   -1  'True
      End
      Begin VB.TextBox txtAnio 
         Height          =   285
         Left            =   5520
         TabIndex        =   1
         Top             =   360
         Width           =   735
      End
      Begin VB.CommandButton cmdPrevisualizar 
         Caption         =   "Previsualizar"
         Height          =   975
         Left            =   600
         Picture         =   "frmBajaCursadasVencidas.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   1200
         Width           =   1455
      End
      Begin VB.CommandButton cmdSalir 
         Height          =   720
         Left            =   7680
         Picture         =   "frmBajaCursadasVencidas.frx":066A
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Salir"
         Top             =   1320
         Width           =   720
      End
      Begin VB.CommandButton cmdProcesar 
         Caption         =   "Procesar Baja"
         Height          =   975
         Left            =   3960
         Picture         =   "frmBajaCursadasVencidas.frx":0AAC
         TabIndex        =   3
         Top             =   1200
         Width           =   1455
      End
      Begin VB.Label Label2 
         Caption         =   "Especifique desde que año hacia atrás  se darán de baja las cursadas:"
         Height          =   375
         Left            =   360
         TabIndex        =   6
         Top             =   360
         Width           =   5175
      End
   End
   Begin VB.Label Label1 
      Caption         =   $"frmBajaCursadasVencidas.frx":1116
      Height          =   615
      Left            =   240
      TabIndex        =   2
      Top             =   240
      Width           =   9015
   End
End
Attribute VB_Name = "frmBajaCursadasVencidas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Conexion As New Connection
Dim CursadasVencidas As New Recordset
Dim CorrelativasVencidas As New Recordset
Dim CorrelativaPorAlumnoVencida As New Recordset
Dim TotalaBajar As New Recordset
Dim Final As New Recordset



Private Sub cmdPrevisualizar_Click()
    If txtAnio = "" Then MsgBox ("Debe especificar un año"): Exit Sub
    Respuesta = MsgBox("¿Desea previsualizar las asignaturas a dar de baja? ", vbYesNo, "Previsualizar")
    If Respuesta = vbYes Then
        Conexion.Open
        Set CursadasVencidas = Conexion.Execute("DELETE * FROM rptCursadasAVencer")
        Set CursadasVencidas = Conexion.Execute("INSERT INTO rptCursadasAVencer ( Permiso, CodigoMateria, Alumno, Resolucion, Carrera, Materia, Curso, Anio, Cursada, Aprobada) SELECT Alumnos.Permiso, Finales.Materia, Alumnos.Nombre, Carreras.Resolucion, Carreras.Nombre, Materias.Nombre, Materias.Curso, Finales.Ano, Finales.Cursada, Finales.Aprobada FROM ((Finales INNER JOIN Alumnos ON Finales.Alumno = Alumnos.Permiso) INNER JOIN Materias ON Finales.Materia = Materias.Codigo) INNER JOIN Carreras ON Materias.Carrera = Carreras.Codigo Where (((Finales.Ano) <= " & txtAnio.Text & ") And ((Finales.Cursada) = True) And ((Finales.Aprobada) = False)) ORDER BY Carreras.Nombre, Materias.Curso, Materias.Nombre, Alumnos.Nombre")
        Set CursadasVencidas = Conexion.Execute("UPDATE rptCursadasAVencer SET aniobase=" & txtAnio)
        Conexion.Close
        rptPrevisualizar.PrintReport
     End If
End Sub

Private Sub cmdProcesar_Click()
   If txtAnio = "" Then MsgBox ("Debe especificar un año"): Exit Sub
   Respuesta = MsgBox("¿Desea procesar las bajas de las cursadas? ", vbYesNo, "Procesar")
   If Respuesta = vbNo Then Exit Sub
   
   If chkIncluirCorrelativas.Value = 1 Then
       Respuesta = MsgBox("Tambien seleccionó marcar como vencidas las asignaturas correlativas." + Chr(13) + "¿Continúa?", vbYesNo, "Atención")
       If Respuesta = vbNo Then Exit Sub
   End If
   
   Me.MousePointer = 11
   
   TotalVencidas = 0
   TotalVencidasCorrelativas = 0
   Conexion.Open
   'levanto las cursadas que estan para marcar como vencidas
   Set TotalaBajar = Conexion.Execute("SELECT count(Alumnos.Permiso) as total FROM ((Finales INNER JOIN Alumnos ON Finales.Alumno = Alumnos.Permiso) INNER JOIN Materias ON Finales.Materia = Materias.Codigo) INNER JOIN Carreras ON Materias.Carrera = Carreras.Codigo Where (((Finales.Ano) <= " & txtAnio & ") And ((Finales.Cursada) = True) And ((Finales.Aprobada) = False))")
   TotalaBajar.MoveFirst
   pbProcesando.Max = TotalaBajar!Total
   pbProcesando.Visible = True
   Set CursadasVencidas = Conexion.Execute("SELECT Alumnos.Permiso,Finales.Materia, Alumnos.Nombre, Carreras.Resolucion, Carreras.Nombre, Materias.Nombre, Materias.Curso, Finales.Ano, Finales.Cursada, finales.Aprobada FROM ((Finales INNER JOIN Alumnos ON Finales.Alumno = Alumnos.Permiso) INNER JOIN Materias ON Finales.Materia = Materias.Codigo) INNER JOIN Carreras ON Materias.Carrera = Carreras.Codigo Where (((Finales.Ano) <= " & txtAnio & ") And ((Finales.Cursada) = True) And ((Finales.Aprobada) = False)) ORDER BY Carreras.Nombre, Materias.Curso, Materias.Nombre, Alumnos.Nombre")
   CursadasVencidas.MoveFirst
   While CursadasVencidas.EOF = False
      Conexion.Execute ("UPDATE Finales SET Finales.PerdioCursada = True, Finales.Cursada = False WHERE (((Finales.Alumno)=" & CursadasVencidas!Permiso & ") AND ((Finales.Materia)=" & CursadasVencidas!Materia & ") AND ((Finales.Ano)=" & CursadasVencidas!Ano & "))")
      Conexion.Execute ("INSERT INTO Cursadas_Vencidas ( Permiso, Materia, Ano_Cursada, Fecha_Proceso, Detalle ) values(" & CursadasVencidas!Permiso & "," & CursadasVencidas!Materia & "," & CursadasVencidas!Ano & ",#" & DateValue(Date) & "#,1 )")
      Conexion.Execute ("insert into Mensajes_alumnos (Permiso, Fecha,Asunto,Detalle) values (" & CursadasVencidas!Permiso & ",#" & DateValue(Date) & "#,'Cursada Vencida','Se vencio la cursada de la materia " & CursadasVencidas!Nombre & "')")
        If chkIncluirCorrelativas.Value = 1 Then
            'busco las materias correlativas en cadena
            Set CorrelativasVencidas = Conexion.Execute("SELECT Correlativas.Principal From Correlativas WHERE Correlativas.Correlativa=" & CursadasVencidas!Materia)
            While CorrelativasVencidas.EOF = False
               'me fijo cada correlativa si el alumno la tiene aprobada
               Set CorrelativaPorAlumnoVencida = Conexion.Execute("SELECT Finales.Ano, Materias.Nombre FROM Finales INNER JOIN Materias ON Finales.Materia = Materias.Codigo WHERE Finales.Alumno=" & CursadasVencidas!Permiso & " AND Finales.Materia=" & CorrelativasVencidas!Principal & " AND Finales.Cursada=True AND Finales.Aprobada=False")
               While CorrelativaPorAlumnoVencida.EOF = False
                  Conexion.Execute ("UPDATE Finales SET Finales.PerdioCursada = True, Finales.Cursada = False WHERE (((Finales.Alumno)=" & CursadasVencidas!Permiso & ") AND ((Finales.Materia)=" & CorrelativasVencidas!Principal & ") AND ((Finales.Ano)=" & CorrelativaPorAlumnoVencida!Ano & "))")
                  Conexion.Execute ("INSERT INTO Cursadas_Vencidas ( Permiso, Materia, Ano_Cursada, Fecha_Proceso, Detalle ) values(" & CursadasVencidas!Permiso & "," & CorrelativasVencidas!Principal & "," & CorrelativaPorAlumnoVencida!Ano & ",#" & DateValue(Date) & "#,2 )")
                  Conexion.Execute ("insert into Mensajes_alumnos (Permiso, Fecha,Asunto,Detalle) values (" & CursadasVencidas!Permiso & ",#" & DateValue(Date) & "#,'Cursada Vencida','Se vencio la cursada de la materia " & CorrelativaPorAlumnoVencida!Nombre & "')")
                  TotalVencidasCorrelativas = TotalVencidasCorrelativas + 1
                  CorrelativaPorAlumnoVencida.MoveNext
               Wend
               CorrelativasVencidas.MoveNext
            Wend
        End If
      CursadasVencidas.MoveNext
      TotalVencidas = TotalVencidas + 1
      pbProcesando.Value = TotalVencidas
   Wend
   Conexion.Close
   Me.MousePointer = 0
   MsgBox ("Se vencieron y actualizaron " & TotalVencidas & " cursadas")
   If TotalVencidasCorrelativas > 0 Then
        MsgBox ("Se vencieron y actualizaron " & TotalVencidasCorrelativas & " cursadas correlativas")
   End If
   pbProcesando.Visible = False

End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Conexion.ConnectionString = ("DSN=Instituto")
    txtAño = Format(Date, "yyyy")
On Error GoTo hErr
   Conexion.Open
   Conexion.Execute ("CREATE TABLE rptCursadasAVencer (Permiso int, CodigoMateria int, Alumno char(50), Resolucion char(80), Carrera char(120), Materia char(123), Curso int, Anio int, Cursada bit, aprobada bit, aniobase int)")
   Conexion.Close
   Exit Sub
hErr:
   'MsgBox Err.Number & " " & Err.Description
   Conexion.Close
   Exit Sub
End Sub



