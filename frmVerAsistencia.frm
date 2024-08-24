VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmVerAsistencia 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Ver Asistencia"
   ClientHeight    =   9120
   ClientLeft      =   4410
   ClientTop       =   1695
   ClientWidth     =   7320
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   9120
   ScaleWidth      =   7320
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSAdodcLib.Adodc adoVerAsistencia 
      Height          =   330
      Left            =   1920
      Top             =   2160
      Visible         =   0   'False
      Width           =   2415
      _ExtentX        =   4260
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
      RecordSource    =   $"frmVerAsistencia.frx":0000
      Caption         =   "VerAsistencia"
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
      Height          =   9015
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7215
      Begin VB.CommandButton cmdCambiar 
         Caption         =   "Cambiar"
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
         Left            =   5880
         TabIndex        =   13
         Top             =   4080
         Width           =   1095
      End
      Begin VB.CommandButton cmdCancelarCorrelativa 
         Height          =   615
         Left            =   4440
         Picture         =   "frmVerAsistencia.frx":0132
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Cancelar"
         Top             =   8040
         Width           =   615
      End
      Begin MSDataGridLib.DataGrid DataGrid1 
         Bindings        =   "frmVerAsistencia.frx":0574
         Height          =   6255
         Left            =   720
         TabIndex        =   3
         Top             =   1560
         Width           =   4935
         _ExtentX        =   8705
         _ExtentY        =   11033
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
            DataField       =   "Nombre"
            Caption         =   "Día"
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
            DataField       =   "Fecha"
            Caption         =   "Fecha"
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
            DataField       =   "Entro"
            Caption         =   "Entro"
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
            DataField       =   "Salio"
            Caption         =   "Salio"
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
            DataField       =   "Presente"
            Caption         =   "P/A"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   5
               Format          =   ""
               HaveTrueFalseNull=   1
               TrueValue       =   "P"
               FalseValue      =   "A"
               NullValue       =   ""
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
               ColumnWidth     =   975,118
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   1049,953
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   569,764
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   615,118
            EndProperty
            BeginProperty Column04 
            EndProperty
         EndProperty
      End
      Begin VB.Label lblPorcentaje 
         Alignment       =   2  'Center
         Caption         =   "Porcentaje"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2760
         TabIndex        =   12
         Top             =   8400
         Width           =   1455
      End
      Begin VB.Label Label5 
         Caption         =   "Porcentaje"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2760
         TabIndex        =   11
         Top             =   8040
         Width           =   1455
      End
      Begin VB.Label lblClases 
         Caption         =   "lblTotal"
         Height          =   255
         Left            =   1680
         TabIndex        =   10
         Top             =   8520
         Width           =   1095
      End
      Begin VB.Label lblAusentes 
         Caption         =   "Label4"
         Height          =   255
         Left            =   1680
         TabIndex        =   9
         Top             =   8280
         Width           =   1095
      End
      Begin VB.Label lblPresentes 
         Caption         =   "Label4"
         Height          =   255
         Left            =   1680
         TabIndex        =   8
         Top             =   8040
         Width           =   1095
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Clases:"
         Height          =   255
         Left            =   720
         TabIndex        =   7
         Top             =   8520
         Width           =   855
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Ausentes:"
         Height          =   255
         Left            =   720
         TabIndex        =   6
         Top             =   8280
         Width           =   855
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Presentes:"
         Height          =   255
         Left            =   720
         TabIndex        =   5
         Top             =   8040
         Width           =   855
      End
      Begin VB.Label lblMateria 
         Caption         =   "Label1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   120
         TabIndex        =   2
         Top             =   720
         Width           =   5295
      End
      Begin VB.Label lblAlumno 
         Caption         =   "Label1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   5295
      End
   End
End
Attribute VB_Name = "frmVerAsistencia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim CursadaNumero As Single
Dim Conectar As New Connection
Dim NumeroCursada As New Recordset

Private Sub cmdCambiar_Click()
    frmCambiarAsistencia.lblMateria = lblMateria
    frmCambiarAsistencia.lblAlumno = lblAlumno
    frmCambiarAsistencia.lblFecha = adoVerAsistencia.Recordset!Fecha
    frmCambiarAsistencia.lblPermiso = frmParciales.adoMatriculados.Recordset!Permiso
    frmCambiarAsistencia.lblNumero = CursadaNumero
    If adoVerAsistencia.Recordset!Presente = True Then
        frmCambiarAsistencia.chkPresente.Value = 1
    Else
        frmCambiarAsistencia.chkPresente.Value = 0
    End If
    frmCambiarAsistencia.Show 1
End Sub

Private Sub cmdCancelarCorrelativa_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
lblAlumno = frmParciales.adoMatriculados.Recordset!Nombre
lblMateria = frmParciales.dtcMaterias
Conectar.Open
Set NumeroCursada = Conectar.Execute("SELECT Divisiones.Numero From Divisiones WHERE Divisiones.Materia=" & frmParciales.dtcMaterias.BoundText & " AND Divisiones.Ano=" & frmParciales.txtAño & " AND Divisiones.Division=" & frmParciales.cbDivision)
CursadaNumero = NumeroCursada!Numero
Conectar.Close
adoVerAsistencia.RecordSource = "SELECT Asistencias.Numero, Dias.Nombre, Asistencias.Fecha, Asistencias.Entro, Asistencias.Salio, Asistencias.Presente, Asistencias.Agente FROM Asistencias INNER JOIN Dias ON Asistencias.Dia = Dias.Numero Where Asistencias.Numero = " & CursadaNumero & " And Asistencias.Agente = " & frmParciales.adoMatriculados.Recordset!Permiso & "  ORDER BY Asistencias.Fecha"
adoVerAsistencia.Refresh
lblPresentes = 0
lblAusentes = 0
lblClases = 0
lblPorcentaje = "% 0"
If adoVerAsistencia.Recordset.EOF = True Then Exit Sub
adoVerAsistencia.Recordset.MoveFirst
For i = 1 To adoVerAsistencia.Recordset.RecordCount
    If adoVerAsistencia.Recordset!Presente = True Then
        lblPresentes = lblPresentes + 1
    Else
        lblAusentes = lblAusentes + 1
    End If
    adoVerAsistencia.Recordset.MoveNext
Next i
lblClases = Val(lblPresentes) + Val(lblAusentes)
lblPorcentaje = "% " & Format((lblPresentes * 100) / lblClases, "00")
End Sub

Private Sub Form_Load()
    Conectar.ConnectionString = ("DSN=Instituto")
End Sub

