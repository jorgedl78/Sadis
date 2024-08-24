VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmCooperadoraGeneracionDeRecibos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   8610
   ClientLeft      =   -15
   ClientTop       =   225
   ClientWidth     =   11565
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8610
   ScaleWidth      =   11565
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdSalir 
      Caption         =   "Salir"
      Height          =   1095
      Left            =   9960
      Picture         =   "frmCooperadoraGeneracionDeRecibos.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   6600
      Width           =   1335
   End
   Begin MSComCtl2.DTPicker dtpFechaDeGeneracion 
      Height          =   375
      Left            =   7560
      TabIndex        =   4
      Top             =   6480
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      _Version        =   393216
      Format          =   80347137
      CurrentDate     =   39197
   End
   Begin MSDataGridLib.DataGrid DataGrid2 
      Bindings        =   "frmCooperadoraGeneracionDeRecibos.frx":08CA
      Height          =   2415
      Left            =   240
      TabIndex        =   3
      Top             =   5880
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   4260
      _Version        =   393216
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
         DataField       =   "Periodo"
         Caption         =   "Período"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "CodigoConcepto"
         Caption         =   "Código"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column02 
         DataField       =   "Concepto"
         Caption         =   "Concepto"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column03 
         DataField       =   "importe"
         Caption         =   "Importe"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
            ColumnWidth     =   675.213
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   705.26
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1800
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   1035.213
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc adoConceptosDlPeriodo 
      Height          =   375
      Left            =   4320
      Top             =   7320
      Visible         =   0   'False
      Width           =   2175
      _ExtentX        =   3836
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
      RecordSource    =   $"frmCooperadoraGeneracionDeRecibos.frx":08EE
      Caption         =   "adoConceptosDelPeriodo"
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
   Begin VB.CommandButton cmdGeneracionInicialDeRecibos 
      Caption         =   "Generar Recibos"
      Height          =   495
      Left            =   6480
      TabIndex        =   2
      Top             =   7320
      Width           =   1935
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frmCooperadoraGeneracionDeRecibos.frx":0A5A
      Height          =   4215
      Left            =   120
      TabIndex        =   0
      Top             =   1320
      Width           =   11295
      _ExtentX        =   19923
      _ExtentY        =   7435
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
            LCID            =   3082
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
            LCID            =   3082
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
            LCID            =   3082
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
            LCID            =   3082
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
            LCID            =   3082
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
            LCID            =   3082
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         MarqueeStyle    =   4
         BeginProperty Column00 
            ColumnWidth     =   780.095
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   3030.236
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   420.095
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   1094.74
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   1904.882
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   3044.977
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc adoAlumnosMatriculados 
      Height          =   495
      Left            =   5040
      Top             =   720
      Visible         =   0   'False
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   873
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
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
      RecordSource    =   $"frmCooperadoraGeneracionDeRecibos.frx":0A7F
      Caption         =   "adoAlumnosMatriculados"
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
      Caption         =   "Generación de Recibos"
      Height          =   2775
      Left            =   120
      TabIndex        =   8
      Top             =   5640
      Width           =   9135
      Begin VB.Label Label3 
         Caption         =   "Fecha de Generación:"
         Height          =   375
         Left            =   5760
         TabIndex        =   9
         Top             =   960
         Width           =   1815
      End
   End
   Begin VB.Label lblTotal 
      Caption         =   "0"
      Height          =   255
      Left            =   3240
      TabIndex        =   7
      Top             =   960
      Width           =   495
   End
   Begin VB.Label lblTitulo 
      Alignment       =   2  'Center
      Caption         =   "Label3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   6
      Top             =   240
      Width           =   10935
   End
   Begin VB.Label Label2 
      Caption         =   "Fecha de Generacion"
      Height          =   255
      Left            =   2280
      TabIndex        =   5
      Top             =   6000
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "Alumnos matriculados en el presente año:"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   960
      Width           =   3015
   End
End
Attribute VB_Name = "frmCooperadoraGeneracionDeRecibos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Conexion As New Connection
Dim DatoAlumno As New Recordset
Dim UltimoRecibo As New Recordset
Dim NumeroRecibo As Double
Dim Orden As Integer

Private Sub cmdGeneracionInicialDeRecibos_Click()
    Respuesta = MsgBox("A continuación se generarán los recibos para el periodo" & Chr(13) & "¿Continua?", vbYesNo, "Atención")
    If Respuesta = vbNo Then Exit Sub
    cmdGeneracionInicialDeRecibos.Enabled = False
    Me.MousePointer = 11
    Conexion.Open
    Set UltimoRecibo = Conexion.Execute("SELECT MAX(Comprobante) as Ultimo from [Recibos_Cooperadora]")
    NumeroRecibo = UltimoRecibo!Ultimo + 1
    Orden = 1
    'recorro los alumnos matriculados en este periodo
    While adoAlumnosMatriculados.Recordset.EOF = False
       adoConceptosDlPeriodo.Recordset.MoveFirst
       'recorro los conceptos de este año y voy agregando los recibos
       While adoConceptosDlPeriodo.Recordset.EOF = False
          Conexion.Execute ("INSERT INTO Recibos_Cooperadora (Alumno, FechaGeneracion, Concepto, Ano, Importe, Comprobante, Periodo, Orden )  VALUES ( " & adoAlumnosMatriculados.Recordset!Permiso & " ,'" & DateValue(dtpFechaDeGeneracion.Value) & "'," & adoConceptosDlPeriodo.Recordset!CodigoConcepto & " ," & Anio & "," & adoConceptosDlPeriodo.Recordset!Importe & ", " & NumeroRecibo & "," & adoConceptosDlPeriodo.Recordset!Periodo & "," & Orden & ")")
          NumeroRecibo = NumeroRecibo + 1
          adoConceptosDlPeriodo.Recordset.MoveNext
       Wend
       adoAlumnosMatriculados.Recordset.MoveNext
       Orden = Orden + 1
    Wend
    Conexion.Execute ("UPDATE Periodos_Cooperadora set Generado=true where Periodo=" & frmCooperadora.adoPeriodos.Recordset!Periodo)
    frmCooperadora.adoPeriodos.Refresh
    Conexion.Close
    Me.MousePointer = 0
    MsgBox ("El proceso de generación a finalizado")
End Sub


Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Conexion.ConnectionString = ("DSN=Instituto")
    Conexion.Open
    lblTitulo = "Generación de recibos para el período " & Anio
    adoAlumnosMatriculados.RecordSource = "SELECT DISTINCT Alumnos.Permiso,Alumnos.Nombre, Alumnos.Tipo, Alumnos.Documento, Alumnos.Domicilio, Alumnos.Localidad FROM Finales INNER JOIN Alumnos ON Finales.Alumno = Alumnos.Permiso Where (((Finales.Ano) = " & Anio & ")) ORDER BY Alumnos.Nombre"
    adoAlumnosMatriculados.Refresh
    lblTotal = adoAlumnosMatriculados.Recordset.RecordCount
    Conexion.Close
    dtpFechaDeGeneracion.Value = Date
End Sub

