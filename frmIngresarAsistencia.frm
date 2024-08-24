VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmIngresarAsistencia 
   ClientHeight    =   2625
   ClientLeft      =   7935
   ClientTop       =   3135
   ClientWidth     =   3570
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   Moveable        =   0   'False
   ScaleHeight     =   2625
   ScaleWidth      =   3570
   Begin VB.Frame Frame1 
      Height          =   2535
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3495
      Begin VB.CommandButton cmdCancelarCorrelativa 
         Height          =   615
         Left            =   2400
         Picture         =   "frmIngresarAsistencia.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Cancelar"
         Top             =   1800
         Width           =   615
      End
      Begin MSDataGridLib.DataGrid dtgEncuentros 
         Bindings        =   "frmIngresarAsistencia.frx":0442
         Height          =   975
         Left            =   120
         TabIndex        =   3
         Top             =   720
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   1720
         _Version        =   393216
         AllowUpdate     =   -1  'True
         BorderStyle     =   0
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
         ColumnCount     =   2
         BeginProperty Column00 
            DataField       =   "Entrada"
            Caption         =   "Entrada"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "hh:mm"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   11274
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   "Salida"
            Caption         =   "Salida"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "hh:mm"
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
               DividerStyle    =   0
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
      Begin VB.CommandButton cmdIngresarAsistencia 
         Caption         =   "Registrar"
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
         Left            =   240
         TabIndex        =   2
         Top             =   1920
         Width           =   1215
      End
      Begin MSComCtl2.DTPicker dtpFecha 
         Height          =   375
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   661
         _Version        =   393216
         Format          =   71827456
         CurrentDate     =   37858
      End
      Begin MSAdodcLib.Adodc adoEncuentros 
         Height          =   330
         Left            =   480
         Top             =   2040
         Visible         =   0   'False
         Width           =   2535
         _ExtentX        =   4471
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
         RecordSource    =   $"frmIngresarAsistencia.frx":045E
         Caption         =   "Encuentros"
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
End
Attribute VB_Name = "frmIngresarAsistencia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancelarCorrelativa_Click()
    Unload Me
End Sub

Private Sub cmdIngresarAsistencia_Click()
    If adoEncuentros.Recordset.RecordCount = 0 Then
        MsgBox ("En este día no se cursa la materia " & frmParciales.dtcMaterias.Text & Chr(13) & "Elija la fecha y el horario correspondiente")
        Exit Sub
    Else
        'frmIngresandoAsistencia.Show 1
        frmCargarAsistenciaCheck.Show 1
    End If
End Sub

Private Sub dtpFecha_Change()
    adoEncuentros.RecordSource = "SELECT Encuentros.Numero, Encuentros.Dia, Encuentros.Entrada, Encuentros.Salida FROM Encuentros INNER JOIN Divisiones ON Encuentros.Numero = Divisiones.Numero Where Encuentros.Dia = " & dtpFecha.DayOfWeek & " And Divisiones.Materia = " & frmParciales.dtcMaterias.BoundText & " And Divisiones.Ano = " & frmParciales.txtAño & " And Divisiones.Division = " & frmParciales.cbDivision & " ORDER BY Encuentros.Entrada"
    adoEncuentros.Refresh
End Sub

Private Sub Form_Activate()
    dtpFecha = Date
    dtpFecha_Change
End Sub

