VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "msadodc.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Begin VB.Form frmInformacionMesasParaProfesores 
   Caption         =   "Mesas para entregar a profesores"
   ClientHeight    =   2235
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5940
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   Moveable        =   0   'False
   ScaleHeight     =   2235
   ScaleWidth      =   5940
   StartUpPosition =   2  'CenterScreen
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   1560
      Top             =   840
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "mesaspro.rpt"
   End
   Begin MSAdodcLib.Adodc MesasParaProfesores 
      Height          =   330
      Left            =   600
      Top             =   1920
      Visible         =   0   'False
      Width           =   3495
      _ExtentX        =   6165
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
      RecordSource    =   "SELECT * FROM [Mesas de examenes por profesor]"
      Caption         =   "Mesas por profesores"
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
      Height          =   840
      Left            =   4680
      Picture         =   "frmInformacionMesasParaProfesores.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Salir"
      Top             =   1080
      Width           =   960
   End
   Begin VB.TextBox Text3 
      DataField       =   "TurnoLlamado"
      DataSource      =   "adoParametros"
      Height          =   375
      Left            =   4680
      TabIndex        =   6
      Text            =   "Text3"
      Top             =   0
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox txtAño 
      Height          =   315
      Left            =   1800
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   360
      Width           =   735
   End
   Begin VB.CommandButton cmdMostrar 
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
      Left            =   4560
      Picture         =   "frmInformacionMesasParaProfesores.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   240
      Width           =   1215
   End
   Begin MSDataListLib.DataCombo dtcMeses 
      Bindings        =   "frmInformacionMesasParaProfesores.frx":0AAC
      Height          =   315
      Left            =   240
      TabIndex        =   2
      Top             =   360
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   556
      _Version        =   393216
      Style           =   2
      ListField       =   "Nombre"
      BoundColumn     =   "Numero"
      Text            =   ""
   End
   Begin MSComCtl2.UpDown UpDown1 
      Height          =   375
      Left            =   2520
      TabIndex        =   5
      Top             =   360
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   661
      _Version        =   393216
      Enabled         =   -1  'True
   End
   Begin MSAdodcLib.Adodc adoParametros 
      Height          =   330
      Left            =   2760
      Top             =   0
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
      Left            =   2880
      Top             =   360
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
   Begin MSDataListLib.DataCombo dtcProfesor 
      Bindings        =   "frmInformacionMesasParaProfesores.frx":0AC3
      Height          =   315
      Left            =   240
      TabIndex        =   8
      Top             =   1560
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   556
      _Version        =   393216
      MatchEntry      =   -1  'True
      Style           =   2
      ListField       =   "Nombre"
      BoundColumn     =   "Codigo"
      Text            =   ""
   End
   Begin MSAdodcLib.Adodc adoPersonal 
      Height          =   330
      Left            =   2280
      Top             =   960
      Visible         =   0   'False
      Width           =   2055
      _ExtentX        =   3625
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
      RecordSource    =   "SELECT Codigo, Nombre FROM Personal WHERE Eliminado = 0 AND TrabajaActualmente = 1 ORDER BY Nombre"
      Caption         =   "Personal"
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
   Begin VB.Label Label6 
      Caption         =   "Profesor:"
      Height          =   255
      Left            =   240
      TabIndex        =   9
      Top             =   1200
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Turno:"
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   120
      Width           =   735
   End
   Begin VB.Label Label4 
      Caption         =   "Año:"
      Height          =   255
      Left            =   1800
      TabIndex        =   3
      Top             =   120
      Width           =   495
   End
End
Attribute VB_Name = "frmInformacionMesasParaProfesores"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Conexion As New Connection
Dim Resultado As Recordset

Private Sub cmdMostrar_Click()
    Conexion.Open
    Set Resultado = Conexion.Execute("DELETE [Mesas de examenes por profesor].* FROM [Mesas de examenes por profesor]")
    Set Resultado = Conexion.Execute("INSERT INTO [Mesas de examenes por profesor] ( Fecha, Hora, Carrera, Materia, Curso, Division, Lugar, Titular, Integrante1, Integrante2, Profesor, Turno, Año )SELECT [Mesas].[Fecha], [Mesas].[Hora], [Carreras].[Abreviatura] AS Carrera, [Materias].[Abreviatura] AS Materia, [Materias].[Curso], [Mesas].[Division], [Mesas].[Lugar], [Personal].[Nombre] AS Titular, Personal_1.Nombre AS Integrante1, Personal_2.Nombre AS Integrante2, [Personal].[Nombre] AS Profesor, [Meses].[Nombre], [Mesas].[Ano]FROM (((((Mesas INNER JOIN Personal ON [Mesas].[Titular]=[Personal].[Codigo]) INNER JOIN Personal AS Personal_1 ON [Mesas].[Integrante1]=Personal_1.Codigo) INNER JOIN Personal AS Personal_2 ON [Mesas].[Integrante2]=Personal_2.Codigo) INNER JOIN Materias ON [Mesas].[Materia]=[Materias].[Codigo]) INNER JOIN Carreras ON [Materias].[Carrera]=[Carreras].[Codigo]) INNER JOIN Meses ON [Mesas].[Turno]=[Meses].[Numero]" & _
        "Where (Mesas.Turno = " & dtcMeses.BoundText & " And Mesas.Ano = " & txtAño & ") AND (Mesas.Titular = " & dtcProfesor.BoundText & " Or Mesas.Integrante1 = " & dtcProfesor.BoundText & " Or Mesas.Integrante2 = " & dtcProfesor.BoundText & ")ORDER BY [Mesas].[Fecha], [Mesas].[Hora]")
    Set Resultado = Conexion.Execute("UPDATE [Mesas de examenes por profesor] SET [Mesas de examenes por profesor].Profesor = '" & dtcProfesor.Text & "'")
    Conexion.Execute ("UPDATE [Mesas de examenes por profesor], Parametros SET [Mesas de examenes por profesor].Institucion = [Parametros].[nombreinstitucion]")
    Conexion.Close
    MesasParaProfesores.Refresh
    If MesasParaProfesores.Recordset.RecordCount < 1 Then MsgBox ("No se encontraron mesas"): Exit Sub
    CrystalReport1.PrintReport
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Conexion.ConnectionString = ("DSN=Instituto")
    dtcMeses.BoundText = adoParametros.Recordset!TurnoLlamado
    txtAño = adoParametros.Recordset!AñoLlamado
    dtcProfesor.BoundText = adoPersonal.Recordset!Codigo
End Sub

Private Sub UpDown1_DownClick()
    txtAño = Val(txtAño) - 1
End Sub

Private Sub UpDown1_UpClick()
    txtAño = Val(txtAño) + 1
End Sub

