VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmGenerarRecibos 
   Caption         =   "Generar Recibos de Cooperadora"
   ClientHeight    =   8430
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   LinkTopic       =   "Form1"
   ScaleHeight     =   8430
   ScaleWidth      =   11880
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdAsignarCobrador 
      Caption         =   "Asignar Cobrador"
      Height          =   615
      Left            =   600
      TabIndex        =   6
      Top             =   7560
      Width           =   2175
   End
   Begin VB.CommandButton cmdGenerar 
      Caption         =   "Generar"
      Height          =   615
      Left            =   3480
      TabIndex        =   5
      Top             =   7560
      Width           =   2175
   End
   Begin VB.CommandButton cmdImprimir 
      Caption         =   "Imprimir"
      Height          =   615
      Left            =   6720
      TabIndex        =   4
      Top             =   7560
      Width           =   2175
   End
   Begin VB.TextBox Text1 
      DataField       =   "Periodo"
      DataSource      =   "adoPeriodos"
      Height          =   285
      Left            =   9240
      TabIndex        =   3
      Text            =   "Text1"
      Top             =   120
      Visible         =   0   'False
      Width           =   255
   End
   Begin MSAdodcLib.Adodc adoPeriodos 
      Height          =   375
      Left            =   9600
      Top             =   120
      Visible         =   0   'False
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   661
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
      RecordSource    =   $"frmGenerarRecibos.frx":0000
      Caption         =   "Periodos"
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
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   6135
      Left            =   -120
      TabIndex        =   0
      Top             =   1080
      Width           =   10215
      Begin MSAdodcLib.Adodc adoAlumnos 
         Height          =   495
         Left            =   600
         Top             =   960
         Visible         =   0   'False
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   873
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
         RecordSource    =   $"frmGenerarRecibos.frx":0068
         Caption         =   "Alumnos"
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
      Begin MSDataGridLib.DataGrid dtgAlumnos 
         Bindings        =   "frmGenerarRecibos.frx":0256
         Height          =   5655
         Left            =   0
         TabIndex        =   1
         Top             =   360
         Width           =   9855
         _ExtentX        =   17383
         _ExtentY        =   9975
         _Version        =   393216
         AllowUpdate     =   0   'False
         ColumnHeaders   =   -1  'True
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
         ColumnCount     =   6
         BeginProperty Column00 
            DataField       =   "Permiso"
            Caption         =   "Permiso"
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
            DataField       =   "Nombre"
            Caption         =   "Nombre"
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
            DataField       =   "Tipo"
            Caption         =   "Tipo"
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
            DataField       =   "Documento"
            Caption         =   "Documento"
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
         BeginProperty Column04 
            DataField       =   "Domicilio"
            Caption         =   "Domicilio"
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
         BeginProperty Column05 
            DataField       =   "Localidad"
            Caption         =   "Localidad"
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
            BeginProperty Column00 
               ColumnWidth     =   915.024
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   1739.906
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   615.118
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   915.024
            EndProperty
            BeginProperty Column04 
               ColumnWidth     =   1739.906
            EndProperty
            BeginProperty Column05 
               ColumnWidth     =   1739.906
            EndProperty
         EndProperty
      End
   End
   Begin VB.Label lblTotalAlumnos 
      Caption         =   "lblTotalAlumnos"
      Height          =   375
      Left            =   5640
      TabIndex        =   2
      Top             =   360
      Width           =   2055
   End
End
Attribute VB_Name = "frmGenerarRecibos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Conexion As New Connection
Dim DatoAlumno As New Recordset
Dim NumeroRecibo As Double
Dim UltimoRecibo As New Recordset
Dim DatosDelAlumno As New Recordset
Dim TotalImpreso As Integer

Private Sub cmdAsignarCobrador_Click()
    frmCooperadoraAsignarCobradores.Show 1
End Sub

Private Sub cmdGenerar_Click()
    Conexion.Open
    While adoAlumnos.Recordset.EOF = False
        Conexion.Execute ("INSERT INTO Recibos_Cooperadora (Alumno, FechaGeneracion, Concepto, Ano, Importe ) SELECT " & adoAlumnos.Recordset!Permiso & ", #05/01/2006#,[Plan Cooperadora].Concepto, " & adoPeriodos.Recordset!Ano & ",[Plan Cooperadora].Importe FROM [Plan Cooperadora]")
        adoAlumnos.Recordset.MoveNext
    Wend
    Conexion.Close
End Sub

Private Sub cmdImprimir_Click()
    While adoAlumnos.Recordset.EOF = False
        Respuesta = MsgBox("Continua con la Hoja nº " & TotalImpreso & "?", vbYesNo, "Impresion de recibos")
        If Respuesta = vbNo Then Exit Sub
        frmCooperadoraImprimeRecibo.PrintForm
        TotalImpreso = TotalImpreso + 1
        adoAlumnos.Recordset.MoveNext
    Wend
    MsgBox ("Se imprimieron todas las hojas")
End Sub

Private Sub Form_Load()
    Conexion.ConnectionString = ("DSN=Instituto")
    lblTotalAlumnos = adoAlumnos.Recordset.RecordCount
    Conexion.Open
    Set UltimoRecibo = Conexion.Execute("SELECT MAX(Comprobante) as Ultimo from [Recibos_Cooperadora]")
    NumeroRecibo = UltimoRecibo!Ultimo + 1
    Conexion.Close
    TotalImpreso = 1
End Sub

