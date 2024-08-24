VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "msadodc.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Begin VB.Form frmInformacionMesasActasSinIngresar 
   Caption         =   "Actas sin Ingresar"
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
      Left            =   720
      Top             =   1200
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "actasin.rpt"
      WindowTitle     =   "Actas sin Ingresar"
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
   End
   Begin VB.CommandButton cmdSalir 
      Height          =   840
      Left            =   4680
      Picture         =   "frmInformacionMesasActasSinIngresar.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Salir"
      Top             =   1080
      Width           =   960
   End
   Begin VB.TextBox Text3 
      DataField       =   "TurnoLlamado"
      DataSource      =   "adoParametros"
      Height          =   375
      Left            =   4680
      TabIndex        =   4
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
      Picture         =   "frmInformacionMesasActasSinIngresar.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   240
      Width           =   1215
   End
   Begin MSAdodcLib.Adodc adoParametros 
      Height          =   330
      Left            =   1800
      Top             =   1080
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
      Left            =   1680
      Top             =   1680
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
   Begin MSDataListLib.DataCombo dtcMeses 
      Bindings        =   "frmInformacionMesasActasSinIngresar.frx":0AAC
      Height          =   315
      Left            =   240
      TabIndex        =   6
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
      TabIndex        =   7
      Top             =   360
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   661
      _Version        =   393216
      Enabled         =   -1  'True
   End
   Begin VB.Label Label1 
      Caption         =   "Turno:"
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   120
      Width           =   735
   End
   Begin VB.Label Label4 
      Caption         =   "Año:"
      Height          =   255
      Left            =   1800
      TabIndex        =   2
      Top             =   120
      Width           =   495
   End
End
Attribute VB_Name = "frmInformacionMesasActasSinIngresar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Conexion As New Connection
Dim Resultado As Recordset

Private Sub cmdMostrar_Click()
    Conexion.Open
    Set Resultado = Conexion.Execute("DELETE rptActasSinIngresar.* FROM rptActasSinIngresar")
    Set Resultado = Conexion.Execute("INSERT INTO rptActasSinIngresar ( Mesa, Carreras_Nombre, Curso, Materias_Nombre, Fecha, Actas, Acta, Division, Personal_Nombre, Turno, Ano ) SELECT Actas.Mesa, Carreras.Nombre, Materias.Curso, Materias.Nombre, Mesas.Fecha, Mesas.Actas, Actas.Acta, Mesas.Division, Personal.Nombre, '" & dtcMeses.Text & "' , " & txtAño & " FROM (((Actas INNER JOIN Mesas ON Actas.Mesa = Mesas.Numero) INNER JOIN Materias ON Mesas.Materia = Materias.Codigo) INNER JOIN Carreras ON Materias.Carrera = Carreras.Codigo) INNER JOIN Personal ON Mesas.Titular = Personal.Codigo Where (((Mesas.Turno) = " & dtcMeses.BoundText & ") And ((Mesas.Ano) = " & txtAño & ") And ((Actas.Ingresada) = False)) ORDER BY Actas.Mesa, Actas.Acta")
    Conexion.Execute ("UPDATE rptActasSinIngresar, Parametros SET rptActasSinIngresar.Institucion = [Parametros].[nombreinstitucion]")
    Set Resultado = Conexion.Execute("SELECT Count(mesa) as total FROM rptActasSinIngresar")
    If Resultado!total = 0 Then MsgBox ("No se encontraron actas sin ingresar para este turno"): Conexion.Close: Exit Sub
    Conexion.Close
    CrystalReport1.PrintReport
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Conexion.ConnectionString = ("DSN=Instituto")
    dtcMeses.BoundText = adoParametros.Recordset!TurnoLlamado
    txtAño = adoParametros.Recordset!AñoLlamado
End Sub

Private Sub UpDown1_DownClick()
    txtAño = Val(txtAño) - 1
End Sub

Private Sub UpDown1_UpClick()
    txtAño = Val(txtAño) + 1
End Sub

