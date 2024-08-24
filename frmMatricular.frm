VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmMatricular 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Matricular"
   ClientHeight    =   5670
   ClientLeft      =   4965
   ClientTop       =   4110
   ClientWidth     =   5250
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5670
   ScaleWidth      =   5250
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame3 
      Height          =   1095
      Left            =   120
      TabIndex        =   9
      Top             =   4440
      Width           =   4935
      Begin VB.CommandButton cmdAceptar 
         Enabled         =   0   'False
         Height          =   615
         Left            =   1320
         Picture         =   "frmMatricular.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Aceptar"
         Top             =   240
         Width           =   615
      End
      Begin VB.CommandButton cmdCancelarCorrelativa 
         Height          =   615
         Left            =   2880
         Picture         =   "frmMatricular.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Cancelar"
         Top             =   240
         Width           =   615
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Alumno"
      Height          =   2775
      Left            =   120
      TabIndex        =   5
      Top             =   1560
      Width           =   4935
      Begin VB.Label lblAlumno 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         TabIndex        =   8
         Top             =   1560
         Width           =   4215
      End
      Begin VB.Label lblDni 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         TabIndex        =   7
         Top             =   840
         Width           =   4215
      End
      Begin VB.Label lblPermiso 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         TabIndex        =   6
         Top             =   360
         Width           =   4215
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Buscar..."
      Height          =   1335
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4935
      Begin VB.TextBox txtPermiso 
         Height          =   375
         Left            =   960
         MaxLength       =   5
         TabIndex        =   2
         Top             =   240
         Width           =   735
      End
      Begin VB.TextBox txtDNI 
         Height          =   375
         Left            =   960
         MaxLength       =   9
         TabIndex        =   1
         Top             =   720
         Width           =   1575
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Permiso:"
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "DNI:"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   840
         Width           =   735
      End
   End
   Begin MSAdodcLib.Adodc adoCorrelativas 
      Height          =   330
      Left            =   2400
      Top             =   0
      Visible         =   0   'False
      Width           =   2175
      _ExtentX        =   3836
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
      RecordSource    =   $"frmMatricular.frx":0884
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
   Begin MSAdodcLib.Adodc adoFinal 
      Height          =   330
      Left            =   2400
      Top             =   360
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
      RecordSource    =   "SELECT Aprobada FROM Finales WHERE Alumno = 0 AND Materia=0 AND Cursada=Verdadero"
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
End
Attribute VB_Name = "frmMatricular"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim DebeCorrelativa As String
Dim NombreCorrelativa(30) As String
Dim TotalCorrelativas As Integer
Dim TotalQueDebe As Integer
Dim Conexion As New Connection
Dim Resultado As New Recordset

Private Sub cmdAceptar_Click()
    Conexion.Open
    Set Resultado = Conexion.Execute("SELECT Cursada,Materia FROM Finales WHERE Alumno = " & txtPermiso & " AND Materia = " & frmParciales.dtcMaterias.BoundText & " AND Cursada = True")
    If Resultado.EOF = False Or Resultado.BOF = False Then
        'ya tiene aprobada la cursada
        MsgBox ("Ya ha aprobado esta cursada"): Conexion.Close: Unload Me: Exit Sub
    End If
    
    Set Resultado = Conexion.Execute("SELECT Alumnos.Permiso FROM Finales INNER JOIN Alumnos ON Finales.Alumno = Alumnos.Permiso Where Finales.Materia = " & frmParciales.dtcMaterias.BoundText & " And Finales.Ano = " & frmParciales.txtAño & " And Finales.Division = " & frmParciales.cbDivision & " AND Alumnos.Permiso=" & txtPermiso)
    If Resultado.EOF = False Then 'el alumno ya está inscripto para este año
        Respuesta = MsgBox("El alumno ya está matriculado", , "Imposible matricular"): cmdAceptar.Enabled = False: txtPermiso = "": txtPermiso.SetFocus: lblAlumno = "": Conexion.Close: Exit Sub
    End If
    Conexion.Close
    
    'comienza el control si debe correlativas
    adoCorrelativas.RecordSource = "SELECT Correlativas.Principal, Correlativas.Correlativa, Materias.Nombre, Materias.Curso FROM Correlativas INNER JOIN Materias ON Correlativas.Correlativa = Materias.Codigo Where Correlativas.Principal = " & frmParciales.dtcMaterias.BoundText & " AND Materias.Curso <> " & frmParciales.cbCurso & " ORDER BY Correlativas.Correlativa"
    adoCorrelativas.Refresh
    If adoCorrelativas.Recordset.RecordCount > 0 Then
        'levanto las materias que son correlativas pero tiene aprobadas
        AdoFinal.RecordSource = "SELECT Correlativas.Correlativa FROM Correlativas INNER JOIN Finales ON Correlativas.Correlativa = Finales.Materia Where (((Correlativas.Principal) = " & frmParciales.dtcMaterias.BoundText & ") And ((Finales.Alumno) = " & txtPermiso & ") And ((Finales.Cursada) = True))"
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
            Else 'debe solo alguna/s correlativa/s
                adoCorrelativas.Recordset.MoveFirst
                For i = 1 To adoCorrelativas.Recordset.RecordCount
                   AdoFinal.Recordset.MoveFirst
                   AdoFinal.Recordset.Find ("Correlativa=" & adoCorrelativas.Recordset!Correlativa)
                   If AdoFinal.Recordset.BOF = True Or AdoFinal.Recordset.EOF = True Then 'no encontro la materia en los cursadas aprobadas (entonces la debe)
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
            Respuesta = MsgBox("Debe las siguientes correlativas:" & Chr(13) & Chr(13) & StrinCorrelativas & Chr(13) & "¿Desea matricularlo igual?", vbYesNo, "Debe Correlativas")
            If Respuesta = vbNo Then Exit Sub
        End If
    End If
    Conexion.Open
    Conexion.Execute ("INSERT INTO Finales ( Alumno, Materia, Ano, Division, Habilitada, Profesor) VALUES (" & txtPermiso & "," & frmParciales.dtcMaterias.BoundText & "," & frmParciales.txtAño & "," & frmParciales.cbDivision & ", True, " & frmParciales.lblCodigoProfesor & " )")
    Conexion.Close
    frmParciales.adoMatriculados.Refresh
    Unload Me
End Sub

Private Sub cmdCancelarCorrelativa_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
    txtPermiso.SetFocus
End Sub

Private Sub Form_Load()
    Conexion.ConnectionString = ("DSN=Instituto")
End Sub

Private Sub Text1_Change()

End Sub


Private Sub txtDNI_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        'busco el alumno si tiene esa carrera
        Conexion.Open
        Set Resultado = Conexion.Execute("SELECT Alumnos.Permiso, Alumnos.Nombre, Alumnos.Documento, CarrerasHechas.Carrera, CarrerasHechas.Condición FROM Alumnos INNER JOIN CarrerasHechas ON Alumnos.Permiso = CarrerasHechas.Permiso WHERE (((Alumnos.Documento)=" & txtDNI & ") AND ((CarrerasHechas.Carrera)=" & frmParciales.dtcCarreras.BoundText & ") AND ((CarrerasHechas.Condición)=1 Or (CarrerasHechas.Condición)=4 Or (CarrerasHechas.Condición)=6))")
        If Resultado.EOF = True Then 'el alumno no tiene asignada esta carrera o la condicion no es = a 1,4 o 6
            Respuesta = MsgBox("El alumno no pertenece a esta carrera o su condición no corresponde", , "Imposible matricular"): lblAlumno = "": txtDNI = "": cmdAceptar.Enabled = False: txtPermiso.SetFocus
        Else
            txtPermiso = Resultado!Permiso
            lblPermiso = "Permiso: " & Resultado!Permiso
            lblDni = "DNI: " & Resultado!documento
            lblAlumno = Resultado!Nombre
            cmdAceptar.Enabled = True
            cmdAceptar.SetFocus
        End If
        Conexion.Close
    End If
    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 Then KeyAscii = 0
End Sub

Private Sub txtPermiso_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        'busco el alumno si tiene esa carrera
        Conexion.Open
        Set Resultado = Conexion.Execute("SELECT Alumnos.Permiso, Alumnos.Nombre, Alumnos.Documento, CarrerasHechas.Carrera, CarrerasHechas.Condición FROM Alumnos INNER JOIN CarrerasHechas ON Alumnos.Permiso = CarrerasHechas.Permiso WHERE (((Alumnos.Permiso)=" & txtPermiso & ") AND ((CarrerasHechas.Carrera)=" & frmParciales.dtcCarreras.BoundText & ") AND ((CarrerasHechas.Condición)=1 Or (CarrerasHechas.Condición)=4 Or (CarrerasHechas.Condición)=6))")
        If Resultado.EOF = True Then 'el alumno no tiene asignada esta carrera o la condicion no es = a 1,4 o 6
            Respuesta = MsgBox("El alumno no pertenece a esta carrera o su condición no corresponde", , "Imposible matricular"): lblAlumno = "": txtPermiso = "": cmdAceptar.Enabled = False: txtPermiso.SetFocus
        Else
            lblPermiso = "Permiso: " & Resultado!Permiso
            lblDni = "DNI: " & Resultado!documento
            lblAlumno = Resultado!Nombre
            cmdAceptar.Enabled = True
            cmdAceptar.SetFocus
        End If
        Conexion.Close
    End If
    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 Then KeyAscii = 0
End Sub
