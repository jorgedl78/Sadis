VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmAgregarMatriculacion 
   BackColor       =   &H0080FFFF&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   6000
   ClientLeft      =   435
   ClientTop       =   1665
   ClientWidth     =   11025
   ControlBox      =   0   'False
   FillColor       =   &H00404040&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6000
   ScaleWidth      =   11025
   ShowInTaskbar   =   0   'False
   Begin MSAdodcLib.Adodc adoFinal 
      Height          =   330
      Left            =   2760
      Top             =   4920
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
   Begin MSAdodcLib.Adodc adoCorrelativas 
      Height          =   330
      Left            =   240
      Top             =   4920
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
      RecordSource    =   $"frmAgregarMatriculacion.frx":0000
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
   Begin VB.CommandButton cmdCancelar 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Cancelar"
      Height          =   615
      Left            =   6600
      MouseIcon       =   "frmAgregarMatriculacion.frx":0077
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   5160
      Width           =   2175
   End
   Begin VB.CommandButton cmdAceptar 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Aceptar"
      Enabled         =   0   'False
      Height          =   615
      Left            =   2520
      MouseIcon       =   "frmAgregarMatriculacion.frx":0381
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   5160
      Width           =   2175
   End
   Begin MSAdodcLib.Adodc adoMaterias 
      Height          =   330
      Left            =   600
      Top             =   1920
      Visible         =   0   'False
      Width           =   1935
      _ExtentX        =   3413
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
      RecordSource    =   $"frmAgregarMatriculacion.frx":068B
      Caption         =   "Materias"
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
   Begin MSDataGridLib.DataGrid dtgMaterias 
      Bindings        =   "frmAgregarMatriculacion.frx":0826
      Height          =   3975
      Left            =   120
      TabIndex        =   0
      Top             =   840
      Width           =   10815
      _ExtentX        =   19076
      _ExtentY        =   7011
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
      ColumnCount     =   4
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
         DataField       =   "Materia"
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
         DataField       =   "Profesor"
         Caption         =   "Profesor"
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
      BeginProperty Column03 
         DataField       =   "Salon"
         Caption         =   "Salon"
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
            ColumnWidth     =   585,071
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   6584,882
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   2145,26
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   975,118
         EndProperty
      EndProperty
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Elija la materia en la que se quiere matricular y presione ""Aceptar"""
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
      Width           =   10455
   End
End
Attribute VB_Name = "frmAgregarMatriculacion"
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
    Me.MousePointer = 11
    Conexion.Open
    Set Resultado = Conexion.Execute("SELECT Cursada,Materia FROM Finales WHERE Alumno = " & frmConexionAlumnos.adoAlumnos.Recordset!Permiso & " AND Materia = " & adoMaterias.Recordset!Codigo & " AND Cursada = True")
    If Resultado.EOF = False Or Resultado.BOF = False Then
        'ya tiene aprobada la cursada
        MsgBox ("Ya ha aprobado esta cursada"): Conexion.Close: Me.MousePointer = 0: Exit Sub
    End If
    Conexion.Close
    
    'controlo que haya limite disponible de matriculados
    Conexion.Open
    Set Auxiliar = Conexion.Execute("SELECT LimiteMatriculados, (SELECT count(Alumno) FROM Finales WHERE Materia=" & adoMaterias.Recordset!Codigo & " AND Ano=" & frmConexionAlumnos.adoParametros.Recordset!AñoMatriculacion & " AND Libre=False AND Division=" & frmMatriculacion.adoCarreras.Recordset!Division & ")  as Matriculados FROM Divisiones WHERE Materia=" & adoMaterias.Recordset!Codigo & " AND Ano=" & frmConexionAlumnos.adoParametros.Recordset!AñoMatriculacion & " AND Division=" & frmMatriculacion.adoCarreras.Recordset!Division & "")
    'Set Auxiliar = Conexion.Execute("SELECT divisiones.limitematriculados, (SELECT Count(Finales.Alumno) AS total FROM Finales WHERE Finales.Materia=" & adoMaterias.Recordset!Codigo & " AND Finales.Ano=" & frmConexionAlumnos.adoParametros.Recordset!AñoMatriculacion & " AND Finales.Division=" & frmMatriculacion.adoCarreras.Recordset!Division & " AND  Finales.Libre=0) AS Matriculados FROM divisiones WHERE divisiones.[numero]=585")
    'Auxiliar.MoveNext
    If (Auxiliar!Matriculados >= Auxiliar!LimiteMatriculados) Then
        Respuesta = MsgBox("Por el momento se llegó al límite máximo de matriculados para esta división", vbOKOnly, "Imposible matricularse"): Conexion.Close: Me.MousePointer = 0: Exit Sub
    End If
    Conexion.Close

    
    With frmMatriculacion.adoMatriculacion.Recordset
    If .RecordCount > 0 Then 'si ya se matriculó en alguna materia
        .MoveFirst
        .Find ("Codigo=" & adoMaterias.Recordset!Codigo)
        If .EOF = True Or .BOF = True Then
            'no encontro la materia en la matriculacion
        ElseIf !Codigo = adoMaterias.Recordset!Codigo Then
            MsgBox ("Ya se matriculó en esta materia")
            Me.MousePointer = 0
            Exit Sub
        End If
    End If
    'levanto las correlativas de la materia elegida a excepcion de la que corresponde al mismo curso
    adoCorrelativas.RecordSource = "SELECT Correlativas.Principal, Correlativas.Correlativa, Materias.Nombre, Materias.Curso FROM Correlativas INNER JOIN Materias ON Correlativas.Correlativa = Materias.Codigo Where Correlativas.Principal = " & adoMaterias.Recordset!Codigo & " AND Materias.Curso <> " & adoMaterias.Recordset!Curso & " ORDER BY Correlativas.Correlativa"
    adoCorrelativas.Refresh
    If adoCorrelativas.Recordset.RecordCount > 0 Then
        'levanto las materias que son correlativas pero tiene aprobadas
        adoFinal.RecordSource = "SELECT Correlativas.Correlativa FROM Correlativas INNER JOIN Finales ON Correlativas.Correlativa = Finales.Materia Where (((Correlativas.Principal) = " & adoMaterias.Recordset!Codigo & ") And ((Finales.Alumno) = " & frmConexionAlumnos.adoAlumnos.Recordset!Permiso & ") And ((Finales.Cursada) = True))"
        adoFinal.Refresh
        TotalQueDebe = adoCorrelativas.Recordset.RecordCount - adoFinal.Recordset.RecordCount
        If TotalQueDebe > 0 Then 'Debe alguna correlativa
            TotalCorrelativas = 0
            If adoFinal.Recordset.RecordCount = 0 Then 'debe todas las correlativas
                adoCorrelativas.Recordset.MoveFirst
                For i = 1 To TotalQueDebe
                    NombreCorrelativa(i) = adoCorrelativas.Recordset!Curso & "°- " & adoCorrelativas.Recordset!Nombre
                    adoCorrelativas.Recordset.MoveNext
               Next i
            Else 'debe solo alguna/s correlativa/s
                adoCorrelativas.Recordset.MoveFirst
                For i = 1 To adoCorrelativas.Recordset.RecordCount
                   adoFinal.Recordset.MoveFirst
                   adoFinal.Recordset.Find ("Correlativa=" & adoCorrelativas.Recordset!Correlativa)
                   If adoFinal.Recordset.BOF = True Or adoFinal.Recordset.EOF = True Then 'no encontro la materia en los cursadas aprobadas (entonces la debe)
                        TotalCorrelativas = TotalCorrelativas + 1
                        NombreCorrelativa(TotalCorrelativas) = adoCorrelativas.Recordset!Curso & "°- " & adoCorrelativas.Recordset!Nombre
                        adoFinal.Recordset.MoveFirst
                    End If
                    adoCorrelativas.Recordset.MoveNext
                Next i
            End If
            StrinCorrelativa = ""
            For i = 1 To TotalQueDebe
                StrinCorrelativas = StrinCorrelativas & NombreCorrelativa(i) & Chr(13) & Chr(13)
            Next i
            MsgBox ("No se puede matricular en esta materia: debe las siguientes correlativas:" & Chr(13) & Chr(13) & StrinCorrelativas)
            Me.MousePointer = 0
            Exit Sub
        End If
    End If
    End With
    Conexion.Open
    Conexion.Execute ("INSERT INTO Finales ( Alumno, Materia, Ano, Division, Habilitada, Profesor) VALUES (" & frmConexionAlumnos.adoAlumnos.Recordset!Permiso & "," & adoMaterias.Recordset!Codigo & "," & frmConexionAlumnos.adoParametros.Recordset!AñoMatriculacion & "," & frmMatriculacion.adoCarreras.Recordset!Division & ", True, " & adoMaterias.Recordset!CodigoProfesor & ")")
    Conexion.Close
    frmMatriculacion.dtcCarrera_Change
    Me.MousePointer = 0
    frmMatriculacion.CargarMaterias = "No"
    Me.Hide
End Sub

Private Sub cmdCancelar_Click()
    Me.Hide
End Sub

Private Sub dtgMaterias_GotFocus()
    cmdAceptar.Enabled = True
End Sub

Private Sub Form_Load()
    Conexion.ConnectionString = ("DSN=Instituto")
End Sub

