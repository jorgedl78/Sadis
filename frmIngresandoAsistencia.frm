VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmIngresandoAsistencia 
   BackColor       =   &H00808080&
   ClientHeight    =   4170
   ClientLeft      =   2520
   ClientTop       =   1770
   ClientWidth     =   9330
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   4170
   ScaleWidth      =   9330
   Begin VB.Frame Frame2 
      Height          =   1575
      Left            =   120
      TabIndex        =   5
      Top             =   2520
      Width           =   9135
      Begin MSComctlLib.ProgressBar pbActualizando 
         Height          =   255
         Left            =   2400
         TabIndex        =   10
         Top             =   1275
         Width           =   6735
         _ExtentX        =   11880
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   1
      End
      Begin MSComctlLib.StatusBar StatusBar1 
         Height          =   375
         Left            =   0
         TabIndex        =   9
         Top             =   1200
         Width           =   9135
         _ExtentX        =   16113
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
            NumPanels       =   1
            BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
               AutoSize        =   1
               Bevel           =   2
               Object.Width           =   16060
            EndProperty
         EndProperty
      End
      Begin VB.TextBox txtPresente 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   360
         MaxLength       =   1
         TabIndex        =   7
         Top             =   360
         Width           =   855
      End
      Begin VB.Label lblAlumno 
         Caption         =   "lblAlumno"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1440
         TabIndex        =   6
         Top             =   480
         Width           =   7575
      End
   End
   Begin VB.Frame Frame1 
      Height          =   2415
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   9135
      Begin VB.Label lblFecha 
         Caption         =   "lblFecha"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   2040
         Width           =   7575
      End
      Begin VB.Label lblProfesor 
         Caption         =   "lblProfesor"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1560
         TabIndex        =   4
         Top             =   1560
         Width           =   6375
      End
      Begin VB.Label lblDivision 
         Caption         =   "lblDivision"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   1560
         Width           =   1095
      End
      Begin VB.Label lblMateria 
         Caption         =   "lblMateria"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   2
         Top             =   960
         Width           =   8895
      End
      Begin VB.Label lblCarrera 
         Caption         =   "lblCarrera"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   8895
      End
   End
End
Attribute VB_Name = "frmIngresandoAsistencia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Conectar As New Connection
Dim Presentes As New Recordset
Dim Horas As New Recordset
Dim NumeroCursada As New Recordset
Dim CursadaNumero As Single
Dim TotalPresentes As Integer
Dim TotalHoras As Integer
Dim HastaFecha As Date
Dim Auxiliar As New Recordset
Dim PorcentajeAsistencia As New Recordset
Dim PorcentajeAsistencias As Integer
Dim Permisos(200) As Double
Dim Asistencia(200) As String
Dim ActualizaAsistencia As New Recordset

Private Sub Form_Activate()
    lblCarrera = "Carrera: " & frmParciales.dtcCarreras
    lblMateria = "Materia: " & frmParciales.dtcMaterias
    lblProfesor = "Profesor: " & frmParciales.lblProfesor
    lblDivision = "División: " & frmParciales.cbDivision
    lblFecha = "Fecha: " & Format(frmIngresarAsistencia.dtpFecha, "dddd, d MMMM yyyy")
    frmParciales.adoMatriculados.Recordset.MoveFirst
    lblAlumno = frmParciales.adoMatriculados.Recordset!Nombre
    Permisos(frmParciales.adoMatriculados.Recordset.Bookmark) = frmParciales.adoMatriculados.Recordset!Permiso
    Conectar.ConnectionString = ("DSN=Instituto")
    Conectar.Open
    Set NumeroCursada = Conectar.Execute("SELECT Divisiones.Numero From Divisiones WHERE Divisiones.Materia=" & frmParciales.dtcMaterias.BoundText & " AND Divisiones.Ano=" & frmParciales.txtAño & " AND Divisiones.Division=" & frmParciales.cbDivision)
    CursadaNumero = NumeroCursada!Numero
    Set PorcentajeAsistencia = Conectar.Execute("SELECT Parametros.PorcentajeAsistencia FROM Parametros")
    PorcentajeAsistencias = PorcentajeAsistencia!PorcentajeAsistencia
    Conectar.Close
    Conectar.Open
    Set Auxiliar = Conectar.Execute("SELECT Count([Agente]) AS CantidadDeAsistencias From Asistencias WHERE Asistencias.Fecha=#" & Format(frmIngresarAsistencia.dtpFecha, "mm/dd/yyyy") & "# AND Asistencias.Agente=" & frmParciales.adoMatriculados.Recordset!Permiso & " AND Asistencias.Numero=" & CursadaNumero)
    If Auxiliar!CantidadDeAsistencias > 0 Then 'ya se ingreso alguna asistencia ese dia
        Respuesta = MsgBox("Ya se ingresaron " & Auxiliar!CantidadDeAsistencias & " horas para este día", vbOKCancel, "Atención")
        If Respuesta = vbCancel Then Conectar.Close: Unload Me: Exit Sub
    End If
    Conectar.Close
End Sub

Private Sub txtPresente_Change()
    If txtPresente <> "" Then
        Asistencia(frmParciales.adoMatriculados.Recordset.Bookmark) = txtPresente
        frmParciales.adoMatriculados.Recordset.MoveNext
      'para saltear los desertantes o retirados con pase y asignarles A(Ausente)
        'If frmParciales.adoMatriculados.Recordset.EOF = False Then
        '    If frmParciales.adoMatriculados.Recordset!Condicion = "Deserción" Or frmParciales.adoMatriculados.Recordset!Condicion = "Retiro con Pase" And frmParciales.adoMatriculados.Recordset.EOF = False Then
        '        Asistencia(frmParciales.adoMatriculados.Recordset.Bookmark) = "a"
        '        Permisos(frmParciales.adoMatriculados.Recordset.Bookmark) = frmParciales.adoMatriculados.Recordset!Permiso
        '        frmParciales.adoMatriculados.Recordset.MoveNext
        '    End If
        'End If
        
        If frmParciales.adoMatriculados.Recordset.EOF = True Then
            Respuesta = MsgBox("Confirma el ingreso de asistencias", vbYesNo, "Confirmar")
            If Respuesta = vbYes Then
                Me.MousePointer = 11: Registrar: Me.MousePointer = 0
            Else
                txtPresente = ""
                Me.Hide
                Exit Sub
            End If
            
            'pregunta si repite el ingreso por otra hora
            Do While a = 0
                Respuesta = MsgBox("¿Desea repetir el ingreso por otra hora?", vbYesNo, "Repetir")
                If Respuesta = vbYes Then
                    Me.MousePointer = 11: Registrar: Me.MousePointer = 0
                Else
                    DeseaActualizar = MsgBox("¿Desea actualizar los porcentajes?", vbYesNo + vbQuestion, "Actualizar porcentajes")
                    If DeseaActualizar = vbYes Then Actualizar
                    pbActualizando.Value = 0
                    frmParciales.VerMatriculados
                    Me.Hide
                    Exit Do
                End If
            Loop
        Else
            lblAlumno = frmParciales.adoMatriculados.Recordset!Nombre
            Permisos(frmParciales.adoMatriculados.Recordset.Bookmark) = frmParciales.adoMatriculados.Recordset!Permiso
        End If
    End If
    txtPresente = ""
End Sub

Private Sub txtPresente_KeyPress(KeyAscii As Integer)
    If KeyAscii <> 80 And KeyAscii <> 112 And KeyAscii <> 65 And KeyAscii <> 97 Then KeyAscii = 0: Exit Sub
End Sub

Private Function Registrar()
    Me.MousePointer = 11
    Conectar.Open
    For i = 1 To frmParciales.adoMatriculados.Recordset.RecordCount
        If Asistencia(i) = "p" Or Asistencia(i) = "P" Then
            Set Auxiliar = Conectar.Execute("INSERT INTO Asistencias (Numero, Agente, Fecha, Dia, Clasificacion, Presente) VALUES (" & CursadaNumero & "," & Permisos(i) & ",'" & DateValue(frmIngresarAsistencia.dtpFecha) & "'," & frmIngresarAsistencia.dtpFecha.DayOfWeek & ",1, True)")
        Else
            Set Auxiliar = Conectar.Execute("INSERT INTO Asistencias (Numero, Agente, Fecha, Dia, Clasificacion, Presente) VALUES (" & CursadaNumero & "," & Permisos(i) & ",'" & DateValue(frmIngresarAsistencia.dtpFecha) & "'," & frmIngresarAsistencia.dtpFecha.DayOfWeek & ",1, False)")
        End If
    Next i
    Conectar.Close
    Me.MousePointer = 0
End Function

Private Function Actualizar()
    Me.MousePointer = 11
    pbActualizando.Max = frmParciales.adoMatriculados.Recordset.RecordCount
    StatusBar1.Panels(1) = "Actualizando Asistencia..."
    Conectar.Open
    Set Auxiliar = Conectar.Execute("SELECT max(Fecha) as UltimaFecha from Asistencias WHERE Numero=" & CursadaNumero)
    HastaFecha = Auxiliar!UltimaFecha
    Conectar.Execute ("INSERT INTO AsistenciaTemporal ( Numero, Agente, Presente ) SELECT Asistencias.Numero, Asistencias.Agente, Asistencias.Presente From Asistencias WHERE Asistencias.Numero=" & CursadaNumero)
    Set Horas = Conectar.Execute("SELECT Count(Asistencias.Presente) AS cantidad, Asistencias.Agente FROM Asistencias WHERE (((Asistencias.Numero)=" & CursadaNumero & ")) GROUP BY Asistencias.Agente ORDER BY Count(Asistencias.Presente) desc")
    TotalHoras = Horas!Cantidad
    ActualizaAsistencia.Open "SELECT Finales.Alumno, Finales.Asistencia, Finales.AsistenciaPorcentaje, AsistenciaHasta FROM Finales WHERE Finales.Ano=" & frmParciales.txtAño & " AND Finales.Materia=" & frmParciales.dtcMaterias.BoundText & " and Finales.Libre=0 and Finales.Division=" & frmParciales.cbDivision, Conectar, adOpenDynamic, adLockPessimistic
    a = 0
    While ActualizaAsistencia.EOF = False
        StatusBar1.Panels(1) = "Actualizando Asistencia"
        Set Presentes = Conectar.Execute("SELECT sum(AsistenciaTemporal.Presente)*-1 AS Presentes FROM AsistenciaTemporal WHERE ((AsistenciaTemporal.Numero)=" & CursadaNumero & ") AND ((AsistenciaTemporal.Agente)=" & ActualizaAsistencia!Alumno & ")")
        TotalPresentes = Presentes!Presentes
        StatusBar1.Panels(1) = "Actualizando Asistencia..."
        porcentaje = Format(((TotalPresentes * 100) / TotalHoras), "00")
        ActualizaAsistencia!AsistenciaPorcentaje = porcentaje
        ActualizaAsistencia!AsistenciaHasta = DateValue(HastaFecha)
        If porcentaje >= PorcentajeAsistencias Then
            ActualizaAsistencia!Asistencia = True
            'Set Auxiliar = Conectar.Execute("UPDATE Finales SET Finales.[AsistenciaPorcentaje] = " & porcentaje & ", Finales.Asistencia = True, Finales.AsistenciaHasta = '" & DateValue(HastaFecha) & "' WHERE (((Finales.Alumno)=" & Permisos(i) & ") AND ((Finales.Materia)=" & frmParciales.dtcMaterias.BoundText & ") AND ((Finales.Ano)=" & frmParciales.txtAño & ") AND ((Finales.Division)=" & frmParciales.cbDivision & "))")
        Else
            ActualizaAsistencia!Asistencia = False
            'Set Auxiliar = Conectar.Execute("UPDATE Finales SET Finales.[AsistenciaPorcentaje] = " & porcentaje & ", Finales.Asistencia = False, Finales.AsistenciaHasta = '" & DateValue(HastaFecha) & "' WHERE (((Finales.Alumno)=" & Permisos(i) & ") AND ((Finales.Materia)=" & frmParciales.dtcMaterias.BoundText & ") AND ((Finales.Ano)=" & frmParciales.txtAño & ") AND ((Finales.Division)=" & frmParciales.cbDivision & "))")
        End If
      a = a + 1
      pbActualizando.Value = a
      ActualizaAsistencia.Update
      ActualizaAsistencia.MoveNext
   Wend
   Conectar.Close
    
    
    
 '   For i = 1 To frmParciales.adoMatriculados.Recordset.RecordCount
 '       StatusBar1.Panels(1) = "Actualizando Asistencia"
 '       Set Presentes = Conectar.Execute("SELECT sum(AsistenciaTemporal.Presente)*-1 AS Presentes FROM AsistenciaTemporal WHERE ((AsistenciaTemporal.Numero)=" & CursadaNumero & ") AND ((AsistenciaTemporal.Agente)=" & Permisos(i) & ")")
 '       TotalPresentes = Presentes!Presentes
 '       StatusBar1.Panels(1) = "Actualizando Asistencia..."
 '       porcentaje = Format(((TotalPresentes * 100) / TotalHoras), "00")
 '       If porcentaje >= PorcentajeAsistencias Then
 '           Set Auxiliar = Conectar.Execute("UPDATE Finales SET Finales.[AsistenciaPorcentaje] = " & porcentaje & ", Finales.Asistencia = True, Finales.AsistenciaHasta = '" & DateValue(HastaFecha) & "' WHERE (((Finales.Alumno)=" & Permisos(i) & ") AND ((Finales.Materia)=" & frmParciales.dtcMaterias.BoundText & ") AND ((Finales.Ano)=" & frmParciales.txtAño & ") AND ((Finales.Division)=" & frmParciales.cbDivision & "))")
 '       Else
 '           Set Auxiliar = Conectar.Execute("UPDATE Finales SET Finales.[AsistenciaPorcentaje] = " & porcentaje & ", Finales.Asistencia = False, Finales.AsistenciaHasta = '" & DateValue(HastaFecha) & "' WHERE (((Finales.Alumno)=" & Permisos(i) & ") AND ((Finales.Materia)=" & frmParciales.dtcMaterias.BoundText & ") AND ((Finales.Ano)=" & frmParciales.txtAño & ") AND ((Finales.Division)=" & frmParciales.cbDivision & "))")
 '       End If
 '   pbActualizando.Value = i
 '   Next i
    Conectar.Open
    Conectar.Execute ("delete * from AsistenciaTemporal WHERE Numero=" & CursadaNumero)
    Conectar.Close
    StatusBar1.Panels(1) = ""
    Me.MousePointer = 0
End Function

