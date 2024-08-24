VERSION 5.00
Begin VB.Form frmInformesyReportes 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Informes, Reportes y Procesos"
   ClientHeight    =   6480
   ClientLeft      =   4470
   ClientTop       =   4335
   ClientWidth     =   9795
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   6480
   ScaleWidth      =   9795
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Height          =   1455
      Left            =   120
      TabIndex        =   7
      Top             =   4800
      Width           =   9495
      Begin VB.CommandButton cmdSalir 
         Height          =   840
         Left            =   3720
         Picture         =   "frmInformesyReportes.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Salir"
         Top             =   360
         Width           =   1440
      End
   End
   Begin VB.Frame Frame1 
      Height          =   4455
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   9495
      Begin VB.CommandButton cmdControlDeCorrelativasPorAlumno 
         Caption         =   "Control de Correlativas por Alumno"
         Height          =   975
         Left            =   360
         TabIndex        =   6
         Top             =   360
         Width           =   3255
      End
      Begin VB.CommandButton cmdRecursantesPorCarrera 
         Caption         =   "Recursantes por Carrera"
         Enabled         =   0   'False
         Height          =   975
         Left            =   5400
         TabIndex        =   5
         Top             =   3000
         Visible         =   0   'False
         Width           =   3255
      End
      Begin VB.CommandButton cmdPorcentajeAvance 
         Caption         =   "Alumnos con Porcentaje de Avance"
         Height          =   975
         Left            =   360
         TabIndex        =   4
         Top             =   1680
         Width           =   3255
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Procesar Cursadas Vencidas"
         Height          =   975
         Left            =   360
         TabIndex        =   3
         Top             =   3000
         Width           =   3255
      End
      Begin VB.CommandButton cmdAprobarLibres 
         Caption         =   "Aprobar Cursada para Libres"
         Height          =   975
         Left            =   5400
         TabIndex        =   2
         Top             =   360
         Width           =   3255
      End
      Begin VB.CommandButton cmdDesaprobarLibres 
         Caption         =   "Desaprobar Cursadas Libres sin Aprobación de Final"
         Height          =   975
         Left            =   5400
         TabIndex        =   1
         Top             =   1680
         Width           =   3255
      End
   End
End
Attribute VB_Name = "frmInformesyReportes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Conexion As New Connection
Dim Parametros As New Recordset
Dim Libres As New Recordset


Private Sub cmdAprobarLibres_Click()
    Conexion.ConnectionString = ("DSN=Instituto")
    Respuesta = MsgBox("A continuación se calculara cuantas cursadas estan afectadas. ¿Continúa?", vbYesNo, "Aprobación de cursadas libres")
    If Respuesta = vbNo Then Exit Sub
    Me.MousePointer = 11
    Conexion.Open
    Set Parametros = Conexion.Execute("SELECT AñoMatriculacion FROM Parametros")
    anioMatriculacion = Parametros!AñoMatriculacion
    Set Libres = Conexion.Execute("SELECT count(Ano) as Total FROM Finales WHERE Ano=" & anioMatriculacion & " AND Libre = True AND Cursada=False")
    totalLibres = Libres!Total
    Me.MousePointer = 0
    Respuesta = MsgBox("¿Confirma la Aprobación de todas las cursadas libres del año " & anioMatriculacion & " para " & totalLibres & " cursadas?", vbYesNo, "Atención")
    If Respuesta = vbYes Then
        Me.MousePointer = 11
        Conexion.Execute ("UPDATE Finales set CURSADA=True, Asistencia=True where Libre= True and Ano = " & anioMatriculacion)
        Me.MousePointer = 0
        MsgBox ("El proceso ha finalizado")
    End If
    Conexion.Close
End Sub

Private Sub cmdControlDeCorrelativasPorAlumno_Click()
    frmInformeControlDeCorrelativas.Show 1
End Sub

Private Sub cmdDesaprobarLibres_Click()
    Conexion.ConnectionString = ("DSN=Instituto")
    Respuesta = MsgBox("A continuación se calculara cuantas cursadas estan afectadas. ¿Continúa?", vbYesNo, "Desaprobación de cursadas libres")
    If Respuesta = vbNo Then Exit Sub
    Me.MousePointer = 11
    Conexion.Open
    Set Libres = Conexion.Execute("SELECT count(Ano) as Total FROM Finales WHERE Libre = True AND Aprobada = False AND Cursada=True")
    totalLibres = Libres!Total
    Me.MousePointer = 0
    Respuesta = MsgBox("¿Confirma Desaprobar todas las cursadas libres que no han aprobado el exámen final a la fecha de hoy para " & totalLibres & " cursadas?", vbYesNo, "Atención")
    If Respuesta = vbYes Then
        Me.MousePointer = 11
        Conexion.Execute ("UPDATE Finales set CURSADA=False, Asistencia=False WHERE Libre = True AND Aprobada = False AND Cursada=True")
        Me.MousePointer = 0
        MsgBox ("El proceso ha finalizado")
    End If
    Me.MousePointer = 0
    Conexion.Close
End Sub

Private Sub cmdPorcentajeAvance_Click()
   frmInformePorcentajeDeAvance.Show 1
End Sub

Private Sub cmdRecursantesPorCarrera_Click()
    frmInformeRecursadasPorAlumno.Show 1
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Command1_Click()
    frmBajaCursadasVencidas.Show 1
End Sub

Private Sub Command2_Click()

End Sub

Private Sub Command3_Click()

End Sub
