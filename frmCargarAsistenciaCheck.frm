VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmCargarAsistenciaCheck 
   Caption         =   "Cargar Asistencia"
   ClientHeight    =   9660
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   14430
   LinkTopic       =   "Form1"
   ScaleHeight     =   9660
   ScaleWidth      =   14430
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.Frame Frame5 
      Caption         =   "Todos"
      Height          =   1095
      Left            =   5280
      TabIndex        =   25
      Top             =   5280
      Width           =   3015
      Begin VB.CheckBox chkTodos 
         Caption         =   "Marcar todos"
         Height          =   375
         Left            =   360
         TabIndex        =   26
         Top             =   360
         Width           =   1575
      End
   End
   Begin VB.Frame Frame4 
      Height          =   3135
      Left            =   12000
      TabIndex        =   20
      Top             =   5400
      Width           =   2175
      Begin VB.CommandButton cmdGuardar 
         Height          =   600
         Left            =   840
         Picture         =   "frmCargarAsistenciaCheck.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   22
         ToolTipText     =   "Guardar"
         Top             =   360
         Width           =   600
      End
      Begin VB.CommandButton cmdCancelar 
         Height          =   600
         Left            =   840
         Picture         =   "frmCargarAsistenciaCheck.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   21
         ToolTipText     =   "Cancelar"
         Top             =   2160
         Width           =   600
      End
   End
   Begin VB.Frame Frame3 
      Height          =   1815
      Left            =   5160
      TabIndex        =   13
      Top             =   3360
      Width           =   9015
      Begin VB.TextBox txtTotal 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   19
         Top             =   1320
         Width           =   855
      End
      Begin VB.TextBox txtAusentes 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   18
         Top             =   720
         Width           =   855
      End
      Begin VB.TextBox txtPresentes 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   17
         Top             =   240
         Width           =   855
      End
      Begin VB.Line Line1 
         X1              =   120
         X2              =   2040
         Y1              =   1200
         Y2              =   1200
      End
      Begin VB.Shape Shape8 
         Height          =   375
         Left            =   120
         Shape           =   4  'Rounded Rectangle
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label8 
         Caption         =   "Presentes:"
         Height          =   255
         Left            =   240
         TabIndex        =   16
         Top             =   315
         Width           =   735
      End
      Begin VB.Shape Shape7 
         Height          =   375
         Left            =   120
         Shape           =   4  'Rounded Rectangle
         Top             =   720
         Width           =   975
      End
      Begin VB.Label Label7 
         Caption         =   "Ausentes:"
         Height          =   255
         Left            =   240
         TabIndex        =   15
         Top             =   795
         Width           =   855
      End
      Begin VB.Shape Shape6 
         Height          =   375
         Left            =   120
         Shape           =   4  'Rounded Rectangle
         Top             =   1320
         Width           =   975
      End
      Begin VB.Label Label6 
         Caption         =   "Total:"
         Height          =   255
         Left            =   480
         TabIndex        =   14
         Top             =   1395
         Width           =   495
      End
   End
   Begin VB.Frame Frame2 
      Height          =   1815
      Left            =   5160
      TabIndex        =   6
      Top             =   1320
      Width           =   9015
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
         Left            =   1080
         TabIndex        =   12
         Top             =   1320
         Width           =   3975
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
         Left            =   1080
         TabIndex        =   11
         Top             =   840
         Width           =   7815
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
         Left            =   1080
         TabIndex        =   10
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label5 
         Caption         =   "Fecha:"
         Height          =   255
         Left            =   360
         TabIndex        =   9
         Top             =   1275
         Width           =   495
      End
      Begin VB.Shape Shape5 
         Height          =   375
         Left            =   120
         Shape           =   4  'Rounded Rectangle
         Top             =   1200
         Width           =   855
      End
      Begin VB.Label Label4 
         Caption         =   "Profesor:"
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   795
         Width           =   615
      End
      Begin VB.Shape Shape4 
         Height          =   375
         Left            =   120
         Shape           =   4  'Rounded Rectangle
         Top             =   720
         Width           =   855
      End
      Begin VB.Label Label3 
         Caption         =   "División:"
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   315
         Width           =   615
      End
      Begin VB.Shape Shape3 
         Height          =   375
         Left            =   120
         Shape           =   4  'Rounded Rectangle
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1215
      Left            =   240
      TabIndex        =   1
      Top             =   120
      Width           =   13935
      Begin VB.Label Label2 
         Caption         =   "Materia:"
         Height          =   255
         Left            =   360
         TabIndex        =   5
         Top             =   800
         Width           =   615
      End
      Begin VB.Shape Shape2 
         Height          =   375
         Left            =   240
         Shape           =   4  'Rounded Rectangle
         Top             =   720
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Carrera:"
         Height          =   255
         Left            =   360
         TabIndex        =   4
         Top             =   310
         Width           =   615
      End
      Begin VB.Shape Shape1 
         Height          =   375
         Left            =   240
         Shape           =   4  'Rounded Rectangle
         Top             =   240
         Width           =   855
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
         Height          =   255
         Left            =   1200
         TabIndex        =   3
         Top             =   360
         Width           =   12615
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
         Height          =   255
         Left            =   1200
         TabIndex        =   2
         Top             =   795
         Width           =   12615
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid gralumnos 
      Height          =   8055
      Left            =   240
      TabIndex        =   0
      Top             =   1440
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   14208
      _Version        =   393216
      Rows            =   1
      FixedRows       =   0
      FixedCols       =   0
      GridColorFixed  =   255
      TextStyleFixed  =   3
      HighLight       =   2
      ScrollBars      =   2
      SelectionMode   =   1
      GridLineWidthFixed=   1
      FontWidthFixed  =   0
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin MSComctlLib.ProgressBar pbActualizando 
      Height          =   255
      Left            =   7680
      TabIndex        =   23
      Top             =   9210
      Visible         =   0   'False
      Width           =   6495
      _ExtentX        =   11456
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Height          =   375
      Left            =   5280
      TabIndex        =   24
      Top             =   9120
      Width           =   9015
      _ExtentX        =   15901
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Bevel           =   2
            Object.Width           =   15849
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmCargarAsistenciaCheck"
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


Private Function CalcularTotales()

    txtPresentes = 0
    txtAusentes = 0
    For i = 1 To gralumnos.Rows - 1
        If gralumnos.TextMatrix(i, 2) = Chr(254) Then
            txtPresentes = txtPresentes + 1
        Else
            txtAusentes = txtAusentes + 1
        End If
    Next i
    txtTotal = Val(txtPresentes) + Val(txtAusentes)
End Function
Private Sub chkTodos_Click()
    For i = 1 To gralumnos.Rows - 1
        If chkTodos.Value = 1 Then
            gralumnos.TextMatrix(i, 2) = Chr(254)
        Else
            gralumnos.TextMatrix(i, 2) = Chr(168)
        End If
    Next i
    CalcularTotales
End Sub

Private Sub cmdCancelar_Click()
    Unload Me
End Sub

Private Sub cmdGuardar_Click()
            Respuesta = MsgBox("Confirma el ingreso de asistencias", vbYesNo, "Confirmar")
            If Respuesta = vbYes Then
                Me.MousePointer = 11: Registrar: Me.MousePointer = 0
            Else
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
                    Unload Me
                    Exit Do
                End If
            Loop
End Sub

Private Sub Form_Load()
    lblCarrera = frmParciales.dtcCarreras
    lblMateria = frmParciales.dtcMaterias
    lblProfesor = frmParciales.lblProfesor
    lblDivision = frmParciales.cbDivision
    lblFecha = Format(frmIngresarAsistencia.dtpFecha, "dddd, d MMMM yyyy")
    
    gralumnos.Cols = 3
    gralumnos.FixedCols = 0
    gralumnos.TextArray(0) = "PERMISO"
    gralumnos.TextArray(1) = "ALUMNO"
    gralumnos.TextArray(2) = "P/A"
    'gralumnos.FixedRows = 1
    gralumnos.SelectionMode = flexSelectionFree
    gralumnos.ColWidth(0) = 700
    gralumnos.ColWidth(1) = 3000
    gralumnos.ColWidth(2) = 600
    
    
    With frmParciales.adoMatriculados.Recordset
    .MoveFirst
    For i = 1 To .RecordCount
        gralumnos.Rows = gralumnos.Rows + 1
        gralumnos.TextMatrix(gralumnos.Rows - 1, 0) = !Permiso
        gralumnos.TextMatrix(gralumnos.Rows - 1, 1) = !Nombre
        gralumnos.Row = gralumnos.Rows - 1
        gralumnos.Col = 2 ' se para en la columna
        gralumnos.CellFontName = "Wingdings" ' cambia la fuente para esta celda
        gralumnos.CellFontSize = 14
        gralumnos.CellAlignment = flexAlignCenterCenter
        gralumnos.TextMatrix(gralumnos.Rows - 1, 2) = Chr(254)
        .MoveNext
    Next i
    gralumnos.FixedRows = 1
    'txtPresentes = .RecordCount
    'txtAusentes = 0
    'txtTotal = .RecordCount
    chkTodos.Value = 1
    CalcularTotales
    
    End With
   
    Conectar.ConnectionString = ("DSN=Instituto")
    Conectar.Open
    Set NumeroCursada = Conectar.Execute("SELECT Divisiones.Numero From Divisiones WHERE Divisiones.Materia=" & frmParciales.dtcMaterias.BoundText & " AND Divisiones.Ano=" & frmParciales.txtAño & " AND Divisiones.Division=" & frmParciales.cbDivision)
    CursadaNumero = NumeroCursada!Numero
    Set PorcentajeAsistencia = Conectar.Execute("SELECT Parametros.PorcentajeAsistencia FROM Parametros")
    PorcentajeAsistencias = PorcentajeAsistencia!PorcentajeAsistencia
    Set Auxiliar = Conectar.Execute("SELECT Count([Agente]) AS CantidadDeAsistencias From Asistencias WHERE Asistencias.Fecha=#" & Format(frmIngresarAsistencia.dtpFecha, "mm/dd/yyyy") & "# AND Asistencias.Agente=" & gralumnos.TextMatrix(1, 0) & " AND Asistencias.Numero=" & CursadaNumero)
    If Auxiliar!CantidadDeAsistencias > 0 Then 'ya se ingreso alguna asistencia ese dia
        Respuesta = MsgBox("Ya se ingresaron " & Auxiliar!CantidadDeAsistencias & " horas para este día", vbOKCancel, "Atención")
        If Respuesta = vbCancel Then Conectar.Close: Unload Me: Exit Sub
    End If
    Conectar.Close
    
    
    
End Sub

Private Sub gralumnos_Click()
    With gralumnos
    If .TextMatrix(.Row, 2) = Chr(168) Then
        .TextMatrix(.Row, 2) = Chr(254)
        txtPresentes = txtPresentes + 1
        txtAusentes = txtAusentes - 1
    Else
        .TextMatrix(.Row, 2) = Chr(168)
        txtPresentes = txtPresentes - 1
        txtAusentes = txtAusentes + 1
    End If
    End With
    CalcularTotales
End Sub
Private Function Registrar()
    Me.MousePointer = 11
    Conectar.Open
    For i = 1 To gralumnos.Rows - 1
        If gralumnos.TextMatrix(i, 2) = Chr(254) Then
            Set Auxiliar = Conectar.Execute("INSERT INTO Asistencias (Numero, Agente, Fecha, Dia, Clasificacion, Presente) VALUES (" & CursadaNumero & "," & gralumnos.TextMatrix(i, 0) & ",'" & DateValue(frmIngresarAsistencia.dtpFecha) & "'," & frmIngresarAsistencia.dtpFecha.DayOfWeek & ",1, True)")
        Else
            Set Auxiliar = Conectar.Execute("INSERT INTO Asistencias (Numero, Agente, Fecha, Dia, Clasificacion, Presente) VALUES (" & CursadaNumero & "," & gralumnos.TextMatrix(i, 0) & ",'" & DateValue(frmIngresarAsistencia.dtpFecha) & "'," & frmIngresarAsistencia.dtpFecha.DayOfWeek & ",1, False)")
        End If
    Next i
    Conectar.Close
    Me.MousePointer = 0
End Function

Private Function Actualizar()
    Me.MousePointer = 11
    pbActualizando.Max = frmParciales.adoMatriculados.Recordset.RecordCount
    StatusBar1.Panels(1) = "Actualizando Asistencia..."
    pbActualizando.Visible = True
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
