VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "msadodc.ocx"
Begin VB.Form frmMateriasPromedio 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3975
   ClientLeft      =   1455
   ClientTop       =   780
   ClientWidth     =   8400
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   3975
   ScaleWidth      =   8400
   Begin VB.CommandButton cmdEliminarComponente 
      Enabled         =   0   'False
      Height          =   615
      Left            =   4680
      Picture         =   "frmMateriasPromedio.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Eliminar"
      Top             =   3120
      Width           =   615
   End
   Begin VB.CommandButton cmdAgregarComponente 
      Height          =   615
      Left            =   2280
      Picture         =   "frmMateriasPromedio.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Agregar"
      Top             =   3120
      Width           =   615
   End
   Begin VB.CommandButton cmdSalir 
      Height          =   615
      Left            =   7440
      Picture         =   "frmMateriasPromedio.frx":0884
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Salir"
      Top             =   3240
      Width           =   615
   End
   Begin MSAdodcLib.Adodc adoMateriasPromedio 
      Height          =   330
      Left            =   600
      Top             =   2160
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
      RecordSource    =   $"frmMateriasPromedio.frx":0CC6
      Caption         =   "MateriasPromedio"
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
   Begin MSDataGridLib.DataGrid dtgMateriasPromedio 
      Bindings        =   "frmMateriasPromedio.frx":0E1E
      Height          =   1815
      Left            =   240
      TabIndex        =   0
      Top             =   1200
      Width           =   7935
      _ExtentX        =   13996
      _ExtentY        =   3201
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
         DataField       =   "Componente"
         Caption         =   "Componente"
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
      BeginProperty Column02 
         DataField       =   "Abreviatura"
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
      BeginProperty Column03 
         DataField       =   "Detalle"
         Caption         =   "Detalle"
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
         RecordSelectors =   0   'False
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   540.284
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   3509.858
         EndProperty
         BeginProperty Column03 
         EndProperty
      EndProperty
   End
   Begin VB.Label lblNombreMateria 
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
      Left            =   240
      TabIndex        =   2
      Top             =   600
      Width           =   6735
   End
   Begin VB.Label Label1 
      Caption         =   "Materias que promedian a:"
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
      Left            =   240
      TabIndex        =   1
      Top             =   120
      Width           =   3375
   End
End
Attribute VB_Name = "frmMateriasPromedio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Conexion As New Connection
Dim Resultado As New Recordset
Private Sub Label2_Click()

End Sub

Private Sub cmdAgregarComponente_Click()
    frmAgregarComponente.adoComponentes.RecordSource = "SELECT Materias.Codigo, Materias.Curso, Materias.Abreviatura, Detalles.Detalle FROM Detalles INNER JOIN Materias ON Detalles.Codigo = Materias.Detalle Where (((Materias.Carrera) = " & frmPlanes.adoCarreras.Recordset!Codigo & ") And ((Materias.Curso) = " & frmPlanes.adoMaterias.Recordset!Curso & "))ORDER BY Materias.Codigo"
    frmAgregarComponente.adoComponentes.Refresh
    frmAgregarComponente.Show 1
End Sub

Private Sub cmdEliminarComponente_Click()
    Respuesta = MsgBox("Desea eliminar la Componente:" & Chr(13) & adoMateriasPromedio.Recordset!Abreviatura, vbYesNo, "BorrarComponente")
    If Respuesta = vbYes Then
        Conexion.Open
        
        'para SQL Server
        'Conexion.Execute ("DELETE MateriasPromedio WHERE Principal=" & frmPlanes.adoMaterias.Recordset!Codigo & " AND Componente=" & adoMateriasPromedio.Recordset!Componente & "")
        
        'para Acces
        Conexion.Execute ("DELETE * FROM MateriasPromedio WHERE Principal=" & frmPlanes.adoMaterias.Recordset!Codigo & " AND Componente=" & adoMateriasPromedio.Recordset!Componente & "")
        
        Conexion.Close
        adoMateriasPromedio.Refresh
        If adoMateriasPromedio.Recordset.RecordCount > 0 Then
            cmdEliminarComponente.Enabled = True
        Else
            cmdEliminarComponente.Enabled = False
        End If
    End If
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Conexion.ConnectionString = ("DSN=Instituto")
End Sub
