VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmCooperadora 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cooperadora"
   ClientHeight    =   6990
   ClientLeft      =   -15
   ClientTop       =   225
   ClientWidth     =   6870
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6990
   ScaleWidth      =   6870
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdSalir 
      Caption         =   "Salir"
      Height          =   975
      Left            =   5640
      Picture         =   "frmCooperadora.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   5760
      Width           =   1095
   End
   Begin VB.CommandButton cmdGenerarRecibosParaElPeriodo 
      Caption         =   "Generar recibos para el Período"
      Height          =   1095
      Left            =   360
      Picture         =   "frmCooperadora.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   5640
      Width           =   2055
   End
   Begin VB.CommandButton cmdImprimirRecibos 
      Caption         =   "Imprimir Recibos"
      Height          =   1095
      Left            =   2880
      Picture         =   "frmCooperadora.frx":1194
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   5640
      Width           =   2055
   End
   Begin VB.Frame frPeriodo 
      Caption         =   "Periodos"
      Height          =   5175
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6615
      Begin MSAdodcLib.Adodc adoPeriodos 
         Height          =   330
         Left            =   3840
         Top             =   4920
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
         RecordSource    =   $"frmCooperadora.frx":1A5E
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
      Begin MSDataGridLib.DataGrid DataGrid1 
         Bindings        =   "frmCooperadora.frx":1AC1
         Height          =   4575
         Left            =   240
         TabIndex        =   2
         Top             =   360
         Width           =   5895
         _ExtentX        =   10398
         _ExtentY        =   8070
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
         ColumnCount     =   5
         BeginProperty Column00 
            DataField       =   "Periodo"
            Caption         =   "Periodo"
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
            DataField       =   "Ano"
            Caption         =   "Ano"
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
            DataField       =   "Generado"
            Caption         =   "Generado"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   5
               Format          =   ""
               HaveTrueFalseNull=   1
               TrueValue       =   "Si"
               FalseValue      =   "No"
               NullValue       =   ""
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   11274
               SubFormatType   =   7
            EndProperty
         EndProperty
         BeginProperty Column03 
            DataField       =   "DesdeFecha"
            Caption         =   "DesdeFecha"
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
            DataField       =   "HastaFecha"
            Caption         =   "HastaFecha"
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
            AllowRowSizing  =   -1  'True
            AllowSizing     =   -1  'True
            BeginProperty Column00 
               ColumnWidth     =   705.26
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   900.284
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   929.764
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   1319.811
            EndProperty
            BeginProperty Column04 
               ColumnWidth     =   1214.929
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Opciones"
      Height          =   1455
      Left            =   120
      TabIndex        =   4
      Top             =   5400
      Width           =   5295
   End
End
Attribute VB_Name = "frmCooperadora"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdGenerarRecibosParaElPeriodo_Click()
    If adoPeriodos.Recordset!Generado = True Then
        MsgBox ("Este período ya ha sido generado")
        Exit Sub
    End If
    Anio = adoPeriodos.Recordset!Ano
    frmCooperadoraGeneracionDeRecibos.adoConceptosDlPeriodo.RecordSource = "SELECT Conceptos_del_periodo.Periodo as Periodo, Conceptos_del_periodo.Concepto as CodigoConcepto, [Cooperadora Conceptos].Concepto as Concepto, Conceptos_del_periodo.Importe as importe FROM Conceptos_del_periodo INNER JOIN [Cooperadora Conceptos] ON Conceptos_del_periodo.Concepto = [Cooperadora Conceptos].Codigo WHERE (((Conceptos_del_periodo.Periodo)=" & adoPeriodos.Recordset!Periodo & "))"
    frmCooperadoraGeneracionDeRecibos.adoConceptosDlPeriodo.Refresh
    frmCooperadoraGeneracionDeRecibos.Show 1
 End Sub

Private Sub cmdImprimirRecibos_Click()
    If adoPeriodos.Recordset!Generado = False Then
        MsgBox ("Aún no se generaron los recibos para este período")
        Exit Sub
    End If
    Anio = adoPeriodos.Recordset!Ano
    frmCooperadoraImpresionDeRecibos.adoaAlumnosSinImprimir.RecordSource = "SELECT distinct Recibos_Cooperadora.Orden,Alumnos.Permiso, Alumnos.Nombre, Alumnos.Tipo, Alumnos.Documento, Alumnos.Domicilio, Alumnos.Localidad, Alumnos.DomicilioEnJunin, Recibos_Cooperadora.Periodo FROM Alumnos INNER JOIN Recibos_Cooperadora ON Alumnos.Permiso = Recibos_Cooperadora.Alumno WHERE (((Recibos_Cooperadora.Impreso)=0) AND ((Recibos_Cooperadora.Periodo)=" & adoPeriodos.Recordset!Periodo & ")) ORDER BY Alumnos.Nombre"
    frmCooperadoraImpresionDeRecibos.adoaAlumnosSinImprimir.Refresh
    frmCooperadoraImpresionDeRecibos.Show 1
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub
