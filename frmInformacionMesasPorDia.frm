VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Begin VB.Form frmInformacionMesasPorDia 
   Caption         =   "Mesas Por Día Para Firmar"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4830
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   Moveable        =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   4830
   StartUpPosition =   2  'CenterScreen
   Begin MSAdodcLib.Adodc adoMesas 
      Height          =   330
      Left            =   2520
      Top             =   120
      Visible         =   0   'False
      Width           =   2160
      _ExtentX        =   3810
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
      RecordSource    =   "SELECT * FROM [Mesas por dia para firmar]"
      Caption         =   "Mesas"
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
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   3480
      Top             =   1320
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "mesasdia.rpt"
      WindowTitle     =   "Mesas por día para firmar"
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
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
      Left            =   3240
      Picture         =   "frmInformacionMesasPorDia.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   720
      Width           =   1215
   End
   Begin VB.CommandButton cmdSalir 
      Height          =   840
      Left            =   3360
      Picture         =   "frmInformacionMesasPorDia.frx":066A
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Salir"
      Top             =   1920
      Width           =   960
   End
   Begin MSComCtl2.MonthView Calendario 
      Height          =   2370
      Left            =   240
      TabIndex        =   0
      Top             =   480
      Width           =   2595
      _ExtentX        =   4577
      _ExtentY        =   4180
      _Version        =   393216
      ForeColor       =   -2147483630
      BackColor       =   -2147483633
      Appearance      =   1
      StartOfWeek     =   46792705
      CurrentDate     =   37629
   End
End
Attribute VB_Name = "frmInformacionMesasPorDia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Conexion As New Connection
Dim Auxiliar As New Recordset

Private Sub Calendario_DateClick(ByVal DateClicked As Date)
    cmdMostrar.SetFocus
End Sub

Private Sub cmdMostrar_Click()
    Conexion.Open
    
    'para SQL Server
    'Set Auxiliar = Conexion.Execute("DELETE [Mesas por dia para firmar]")
    'Set Auxiliar = Conexion.Execute("INSERT INTO [Mesas por dia para firmar] ( Carrera, Materia, Numero,Curso, Division, Fecha, Hora, Lugar, Titular, Integrante1, Integrante2, Actas ) SELECT Carreras.Abreviatura AS Carrera, Materias.Abreviatura AS Materia, Mesas.Numero, Materias.Curso, Mesas.Division, Mesas.Fecha, Mesas.Hora, Mesas.Lugar, Personal.Nombre AS Titular, Personal_1.Nombre AS Integrante1, Personal_2.Nombre AS Integrante2, Mesas.Actas FROM ((((Mesas INNER JOIN Materias ON Mesas.Materia = Materias.Codigo) INNER JOIN Carreras ON Materias.Carrera = Carreras.Codigo) INNER JOIN Personal ON Mesas.Titular = Personal.Codigo) INNER JOIN Personal AS Personal_1 ON Mesas.Integrante1 = Personal_1.Codigo) INNER JOIN Personal AS Personal_2 ON Mesas.Integrante2 = Personal_2.Codigo Where (((Mesas.Fecha) = '" & DateValue(Calendario.Value) & "')) ORDER BY Carreras.Abreviatura")
    
    'para Acces
    Set Auxiliar = Conexion.Execute("DELETE * FROM [Mesas por dia para firmar]")
    Set Auxiliar = Conexion.Execute("INSERT INTO [Mesas por dia para firmar] ( Carrera, Materia, Numero,Curso, Division, Fecha, Hora, Lugar, Titular, Integrante1, Integrante2, Actas ) SELECT Carreras.Abreviatura AS Carrera, Materias.Abreviatura AS Materia, Mesas.Numero, Materias.Curso, Mesas.Division, Mesas.Fecha, Mesas.Hora, Mesas.Lugar, Personal.Nombre AS Titular, Personal_1.Nombre AS Integrante1, Personal_2.Nombre AS Integrante2, Mesas.Actas FROM ((((Mesas INNER JOIN Materias ON Mesas.Materia = Materias.Codigo) INNER JOIN Carreras ON Materias.Carrera = Carreras.Codigo) INNER JOIN Personal ON Mesas.Titular = Personal.Codigo) INNER JOIN Personal AS Personal_1 ON Mesas.Integrante1 = Personal_1.Codigo) INNER JOIN Personal AS Personal_2 ON Mesas.Integrante2 = Personal_2.Codigo Where (((Mesas.Fecha) = #" & Format(Calendario.Value, "mm/dd/yyyy") & "#)) ORDER BY Carreras.Abreviatura")
    
    Conexion.Close
    adoMesas.Refresh
    If adoMesas.Recordset.RecordCount < 1 Then MsgBox ("No se armaron mesas para esta fecha"): Exit Sub
    CrystalReport1.PrintReport
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Conexion.ConnectionString = ("DSN=Instituto")
    Calendario.Value = Date
End Sub
