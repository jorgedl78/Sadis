VERSION 5.00
Begin VB.Form frmInformeRecursadasPorAlumno 
   Caption         =   "Recursadas por Alumno"
   ClientHeight    =   5265
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4845
   LinkTopic       =   "Form1"
   ScaleHeight     =   5265
   ScaleWidth      =   4845
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdProcesar 
      Caption         =   "Procesar"
      Height          =   615
      Left            =   1080
      TabIndex        =   0
      Top             =   3720
      Width           =   2535
   End
   Begin VB.Label Label2 
      Caption         =   "Solo para I.S.F.D. y T. Nº 20"
      Height          =   255
      Left            =   1200
      TabIndex        =   2
      Top             =   360
      Width           =   2175
   End
   Begin VB.Label Label1 
      Caption         =   $"frmInformeRecursadasPorAlumno.frx":0000
      Height          =   1215
      Left            =   240
      TabIndex        =   1
      Top             =   1920
      Width           =   4095
   End
End
Attribute VB_Name = "frmInformeRecursadasPorAlumno"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Conexion As New Connection
Dim Materias As New Recordset
Dim Materia As New Recordset
Dim Alumnos As New Recordset

Private Sub cmdProcesar_Click()
Conexion.Open
Set Materias = Conexion.Execute("SELECT Materias.Codigo As Codigo, Materias.Nombre, Materias.Detalle FROM Materias WHERE (((Materias.Carrera)=54) AND ((Materias.Curso)=1) AND ((Materias.Detalle)=1 Or (Materias.Detalle)=5)) ORDER BY Materias.Codigo ")
Set Alumnos = Conexion.Execute("SELECT Alumnos.Permiso AS Permiso, Alumnos.Documento, Alumnos.Nombre FROM Alumnos INNER JOIN CarrerasHechas ON Alumnos.Permiso = CarrerasHechas.Permiso WHERE CarrerasHechas.Carrera=54 AND CarrerasHechas.Ingreso<=2008 ORDER BY Alumnos.Nombre")
While Materias.EOF = False  'recorro las materias una por una
   Alumnos.MoveFirst
   While Alumnos.EOF = False
       Set Materia = Conexion.Execute("SELECT Finales.Alumno, Finales.Materia, Finales.Ano, Finales.Asistencia, Finales.PerdioCursada, Finales.Aprobada FROM Finales WHERE Finales.Alumno=" & Alumnos!Permiso & " AND Finales.Materia=" & Materias!Codigo & " ORDER BY Finales.Ano desc")
       If Materia.EOF = False Then 'al menos la curso una vez
          If Materia!Aprobada = False Then 'no esta aprobado el final
             If Materia!Asistencia = 0 Then 'abandono
                Conexion.Execute ("INSERT INTO Informe_alumnos_adeudan_materia ( permiso, alumno, materia, anio, noaprobo, vencida, nocurso, abandono ) values (" & Alumnos!Permiso & ",'" & Alumnos!Nombre & "'," & Materias!Codigo & "," & Materia!Ano & ",0,0,0,1)")
             ElseIf Materia!PerdioCursada = True Then
                Conexion.Execute ("INSERT INTO Informe_alumnos_adeudan_materia ( permiso, alumno, materia, anio, noaprobo, vencida, nocurso, abandono ) values (" & Alumnos!Permiso & ",'" & Alumnos!Nombre & "'," & Materias!Codigo & "," & Materia!Ano & ",0,1,0,0)")
             Else
                Conexion.Execute ("INSERT INTO Informe_alumnos_adeudan_materia ( permiso, alumno, materia, anio,  noaprobo, vencida, nocurso, abandono ) values (" & Alumnos!Permiso & ",'" & Alumnos!Nombre & "'," & Materias!Codigo & "," & Materia!Ano & ",1,0,0,0)")
             End If
          End If
        Else 'nunca la curso
                Conexion.Execute ("INSERT INTO Informe_alumnos_adeudan_materia ( permiso, alumno, materia, anio, noaprobo, vencida, nocurso, abandono ) values (" & Alumnos!Permiso & ",'" & Alumnos!Nombre & "'," & Materias!Codigo & ",0,0,0,1,0)")
        End If
       Alumnos.MoveNext
   Wend
   Materias.MoveNext
Wend
Conexion.Close
End Sub

Private Sub Form_Load()
    Conexion.ConnectionString = ("DSN=Instituto")
End Sub
