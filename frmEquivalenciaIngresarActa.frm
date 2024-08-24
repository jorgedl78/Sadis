VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmEquivalenciaIngresarActa 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   9180
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   7725
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9180
   ScaleWidth      =   7725
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Height          =   4935
      Left            =   120
      TabIndex        =   11
      Top             =   3360
      Width           =   7455
      Begin MSAdodcLib.Adodc adoOtorgadas 
         Height          =   375
         Left            =   4320
         Top             =   2400
         Visible         =   0   'False
         Width           =   2415
         _ExtentX        =   4260
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
         RecordSource    =   $"frmEquivalenciaIngresarActa.frx":0000
         Caption         =   "Otorgadas"
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
      Begin MSDataGridLib.DataGrid dtgOtorgadas 
         Bindings        =   "frmEquivalenciaIngresarActa.frx":017F
         Height          =   3735
         Left            =   120
         TabIndex        =   12
         Top             =   240
         Width           =   6855
         _ExtentX        =   12091
         _ExtentY        =   6588
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
         ColumnCount     =   3
         BeginProperty Column00 
            DataField       =   "Alumno"
            Caption         =   "Alumno"
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
            DataField       =   "Nota"
            Caption         =   "Nota"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "0,00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   11274
               SubFormatType   =   1
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            MarqueeStyle    =   4
            BeginProperty Column00 
               ColumnWidth     =   884,976
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   4605,166
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   659,906
            EndProperty
         EndProperty
      End
      Begin VB.Label lblTotal 
         Alignment       =   1  'Right Justify
         Caption         =   "Total"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4320
         TabIndex        =   14
         Top             =   4200
         Width           =   2655
      End
   End
   Begin VB.CommandButton cmdAceptar 
      Height          =   615
      Left            =   2280
      Picture         =   "frmEquivalenciaIngresarActa.frx":019A
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Aceptar"
      Top             =   8400
      Width           =   615
   End
   Begin VB.CommandButton cmdCancelar 
      Height          =   615
      Left            =   4320
      Picture         =   "frmEquivalenciaIngresarActa.frx":05DC
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Cancelar"
      Top             =   8400
      Width           =   615
   End
   Begin VB.Frame Frame1 
      Height          =   3135
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   7455
      Begin VB.TextBox txtLibro 
         Height          =   285
         Left            =   4320
         TabIndex        =   0
         Top             =   2280
         Width           =   735
      End
      Begin VB.TextBox txtFolio 
         Height          =   285
         Left            =   5880
         TabIndex        =   2
         Top             =   2280
         Width           =   855
      End
      Begin VB.Label Label3 
         Caption         =   "Libro:"
         Height          =   255
         Left            =   3720
         TabIndex        =   13
         Top             =   2280
         Width           =   495
      End
      Begin VB.Label lblFecha 
         Caption         =   "Fecha"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   2280
         Width           =   2895
      End
      Begin VB.Label lblProfesor 
         Caption         =   "Profesor"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   1800
         Width           =   5775
      End
      Begin VB.Label lblCursoYMateria 
         Caption         =   "Curso y Materia"
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
         TabIndex        =   8
         Top             =   1320
         Width           =   5775
      End
      Begin VB.Label lblCarrera 
         Caption         =   "Carera:"
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
         Left            =   120
         TabIndex        =   7
         Top             =   720
         Width           =   7095
      End
      Begin VB.Label Label2 
         Caption         =   "Folio:"
         Height          =   255
         Left            =   5400
         TabIndex        =   6
         Top             =   2280
         Width           =   495
      End
      Begin VB.Label Label1 
         Caption         =   "Ingreso de Acta de Equivalencia"
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
         Left            =   1560
         TabIndex        =   3
         Top             =   240
         Width           =   4815
      End
   End
End
Attribute VB_Name = "frmEquivalenciaIngresarActa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Conexion As New Connection

Private Sub cmdAceptar_Click()
    If txtFolio = "" Then MsgBox ("Debe ingresar el Nº de folio"): txtFolio.SetFocus: Exit Sub
    Respuesta = MsgBox("A continuación se cerrará el acta de equivalencias", vbYesNo, "Cierre de acta")
    If Respuesta = vbNo Then Exit Sub
    Conexion.Open
    adoOtorgadas.Recordset.MoveFirst
    Ingresadas = 0
    While adoOtorgadas.Recordset.EOF = False
        'para ingresar en tabla finales
        Conexion.Execute ("INSERT INTO Finales ( Alumno, Materia, Ano, Division, Cursada, Asistencia, Aprobada, Nota, Fecha, Libro, Folio, Acta, Equivalencia, Establecimiento, Profesor )VALUES (" & adoOtorgadas.Recordset!Alumno & ", " & frmEquivalencias.dtcMaterias.BoundText & ", 0, 1,True,True, True, " & adoOtorgadas.Recordset!Nota & ", # " & Format(adoOtorgadas.Recordset!FechaAprobacion, "mm/dd/yyyy") & " # , " & txtLibro & ", " & txtFolio & ", 1, True, '" & adoOtorgadas.Recordset!Institucion & "', " & adoOtorgadas.Recordset!Profesor & ")")
        adoOtorgadas.Recordset.MoveNext
        Ingresadas = Ingresadas + 1
    Wend
    Conexion.Execute ("UPDATE Equivalencias SET Libro=" & txtLibro & ", Folio =" & txtFolio & " WHERE Otorgada = True AND MateriaAReconocer = " & frmEquivalencias.dtcMaterias.BoundText & " AND AnoSolicitud = " & frmEquivalencias.txtAño)
    Conexion.Execute ("UPDATE EquivalenciasResumen SET Ingresada = True, Libro=" & txtLibro & ", Folio =" & txtFolio & " WHERE Asignatura = " & frmEquivalencias.dtcMaterias.BoundText & " AND AnoSolicitud=" & frmEquivalencias.txtAño)
    Conexion.Close
    Respuesta = MsgBox("Se ingresaron " & Ingresadas & " equivalencias al sistema", vbOKOnly, "Ingreso completado")
    Unload Me
End Sub

Private Sub cmdCancelar_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
    txtLibro.SetFocus
End Sub

Private Sub Form_Load()
    Conexion.ConnectionString = ("DSN=Instituto")
End Sub
