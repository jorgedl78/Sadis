VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmMovimientos 
   BackColor       =   &H00404040&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Movimientos"
   ClientHeight    =   8070
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11310
   ClipControls    =   0   'False
   Icon            =   "Movimientos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8070
   ScaleWidth      =   11310
   Begin VB.TextBox txtCupof2 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      Height          =   285
      Left            =   6960
      TabIndex        =   19
      Top             =   1485
      Width           =   975
   End
   Begin VB.TextBox txtRevista 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      Height          =   285
      Left            =   1320
      Locked          =   -1  'True
      MaxLength       =   3
      MouseIcon       =   "Movimientos.frx":0442
      MousePointer    =   99  'Custom
      TabIndex        =   21
      Top             =   1845
      Width           =   372
   End
   Begin VB.TextBox txtTurno 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      Height          =   285
      Left            =   9020
      Locked          =   -1  'True
      MouseIcon       =   "Movimientos.frx":074C
      MousePointer    =   99  'Custom
      TabIndex        =   11
      Top             =   1125
      Width           =   372
   End
   Begin VB.TextBox txtDivision 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      Height          =   285
      Left            =   8650
      TabIndex        =   9
      Top             =   1125
      Width           =   372
   End
   Begin VB.TextBox txtCupof 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      Height          =   285
      Left            =   4920
      TabIndex        =   17
      Top             =   1485
      Width           =   975
   End
   Begin VB.TextBox txtAño 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      Height          =   285
      Left            =   8280
      TabIndex        =   7
      Top             =   1125
      Width           =   372
   End
   Begin MSAdodcLib.Adodc mdbMovimientos 
      Height          =   330
      Left            =   9840
      Top             =   7320
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   582
      ConnectMode     =   3
      CursorLocation  =   2
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   1
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
      RecordSource    =   "SELECT * FROM Movimientos ORDER BY CodPersonal"
      Caption         =   ""
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
   Begin MSDataGridLib.DataGrid dbgMovimientos 
      Bindings        =   "Movimientos.frx":0A56
      Height          =   4335
      Left            =   120
      TabIndex        =   35
      Top             =   3600
      Width           =   11055
      _ExtentX        =   19500
      _ExtentY        =   7646
      _Version        =   393216
      AllowUpdate     =   -1  'True
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
      Caption         =   "Situación de Revista"
      ColumnCount     =   8
      BeginProperty Column00 
         DataField       =   "Carrera"
         Caption         =   "Carrera"
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
         DataField       =   "Año"
         Caption         =   "Año"
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
         DataField       =   "Asignatura"
         Caption         =   "Asignatura"
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
         DataField       =   "Horas"
         Caption         =   "H/C"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   11274
            SubFormatType   =   1
         EndProperty
      EndProperty
      BeginProperty Column04 
         DataField       =   "Modulos"
         Caption         =   "Mód."
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   11274
            SubFormatType   =   1
         EndProperty
      EndProperty
      BeginProperty Column05 
         DataField       =   "Revista"
         Caption         =   "Revista"
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
      BeginProperty Column06 
         DataField       =   "Desde"
         Caption         =   "Desde"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "dd/MM/yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   11274
            SubFormatType   =   3
         EndProperty
      EndProperty
      BeginProperty Column07 
         DataField       =   "Hasta"
         Caption         =   "Hasta"
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
         BeginProperty Column00 
            ColumnWidth     =   2324,977
         EndProperty
         BeginProperty Column01 
            Alignment       =   2
            ColumnWidth     =   345,26
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   4500,284
         EndProperty
         BeginProperty Column03 
            Alignment       =   2
            ColumnWidth     =   390,047
         EndProperty
         BeginProperty Column04 
            Alignment       =   2
            ColumnWidth     =   420,095
         EndProperty
         BeginProperty Column05 
            Alignment       =   2
            ColumnWidth     =   629,858
         EndProperty
         BeginProperty Column06 
            Alignment       =   2
            ColumnWidth     =   959,811
         EndProperty
         BeginProperty Column07 
            Alignment       =   2
            ColumnWidth     =   959,811
         EndProperty
      EndProperty
   End
   Begin VB.ComboBox txtAsignatura 
      Enabled         =   0   'False
      Height          =   315
      Left            =   1320
      TabIndex        =   5
      Top             =   1125
      Width           =   6615
   End
   Begin VB.ComboBox txtCarrera 
      Enabled         =   0   'False
      Height          =   315
      Left            =   1320
      TabIndex        =   3
      Top             =   765
      Width           =   6615
   End
   Begin VB.ComboBox txtNombre 
      Height          =   315
      ItemData        =   "Movimientos.frx":0A73
      Left            =   1680
      List            =   "Movimientos.frx":0A75
      TabIndex        =   1
      Top             =   120
      Width           =   7815
   End
   Begin VB.CommandButton cmdImprimir 
      BackColor       =   &H00FFE0C0&
      Caption         =   "Imprimir"
      Height          =   375
      Left            =   8280
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   2040
      Width           =   1215
   End
   Begin VB.CommandButton cmdGuardar 
      BackColor       =   &H00FFE0C0&
      Caption         =   "Guardar"
      Enabled         =   0   'False
      Height          =   375
      Left            =   9960
      Style           =   1  'Graphical
      TabIndex        =   30
      Top             =   1680
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancelar 
      BackColor       =   &H00FFE0C0&
      Caption         =   "Cancelar"
      Enabled         =   0   'False
      Height          =   375
      Left            =   9960
      Style           =   1  'Graphical
      TabIndex        =   31
      Top             =   2160
      Width           =   1215
   End
   Begin VB.Frame frmBuscar 
      BackColor       =   &H00404040&
      Caption         =   " Buscar "
      ForeColor       =   &H00FFFFFF&
      Height          =   660
      Left            =   120
      TabIndex        =   32
      Top             =   2760
      Width           =   11055
      Begin VB.TextBox txtBuscarNombre 
         Height          =   285
         Left            =   2280
         MaxLength       =   40
         TabIndex        =   34
         Top             =   240
         Width           =   8655
      End
      Begin VB.Label Label15 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Apellido y Nombres:"
         ForeColor       =   &H00FFFFFF&
         Height          =   252
         Left            =   240
         TabIndex        =   33
         Top             =   240
         Width           =   1932
      End
   End
   Begin VB.CommandButton cmdEliminar 
      BackColor       =   &H00FFE0C0&
      Caption         =   "Eliminar"
      Height          =   375
      Left            =   9960
      Style           =   1  'Graphical
      TabIndex        =   29
      Top             =   1200
      Width           =   1215
   End
   Begin VB.CommandButton cmdModificar 
      BackColor       =   &H00FFE0C0&
      Caption         =   "Modificar"
      Height          =   375
      Left            =   9960
      Style           =   1  'Graphical
      TabIndex        =   28
      Top             =   720
      Width           =   1215
   End
   Begin VB.CommandButton cmdNuevo 
      BackColor       =   &H00FFE0C0&
      Caption         =   "Nuevo"
      Height          =   375
      Left            =   9960
      Style           =   1  'Graphical
      TabIndex        =   27
      Top             =   120
      Width           =   1215
   End
   Begin VB.TextBox txtHasta 
      Enabled         =   0   'False
      Height          =   285
      Left            =   3360
      MaxLength       =   10
      TabIndex        =   25
      Top             =   2205
      Width           =   1095
   End
   Begin VB.TextBox txtDesde 
      Enabled         =   0   'False
      Height          =   285
      Left            =   1320
      MaxLength       =   10
      TabIndex        =   23
      Top             =   2205
      Width           =   1095
   End
   Begin VB.TextBox txtModulos 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      Height          =   285
      Left            =   3120
      TabIndex        =   15
      Top             =   1485
      Width           =   495
   End
   Begin VB.TextBox txtHoras 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      Height          =   285
      Left            =   1320
      TabIndex        =   13
      Top             =   1485
      Width           =   495
   End
   Begin VB.Label lblCUPOF2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Cupof 2:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   6240
      TabIndex        =   18
      Top             =   1515
      Width           =   615
   End
   Begin VB.Label lblTurno 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Turno"
      ForeColor       =   &H00FFFFFF&
      Height          =   252
      Left            =   9000
      TabIndex        =   10
      Top             =   860
      Width           =   492
   End
   Begin VB.Label lblDiv 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Div"
      ForeColor       =   &H00FFFFFF&
      Height          =   252
      Left            =   8640
      TabIndex        =   8
      Top             =   860
      Width           =   372
   End
   Begin VB.Label lblCupof1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Cupof 1:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   4200
      TabIndex        =   16
      Top             =   1515
      Width           =   615
   End
   Begin VB.Label lblAño 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Año"
      ForeColor       =   &H00FFFFFF&
      Height          =   252
      Left            =   8280
      TabIndex        =   6
      Top             =   860
      Width           =   372
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00FFFFFF&
      X1              =   11880
      X2              =   0
      Y1              =   600
      Y2              =   600
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00FFFFFF&
      X1              =   9720
      X2              =   9720
      Y1              =   2640
      Y2              =   0
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      X1              =   0
      X2              =   11880
      Y1              =   2640
      Y2              =   2640
   End
   Begin VB.Label lblHasta 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Hasta:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2760
      TabIndex        =   24
      Top             =   2235
      Width           =   495
   End
   Begin VB.Label lblDesde 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Desde:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   600
      TabIndex        =   22
      Top             =   2235
      Width           =   615
   End
   Begin VB.Label lblRevista 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Revista:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   20
      Top             =   1875
      Width           =   1095
   End
   Begin VB.Label lblModulos 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Módulos:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2280
      TabIndex        =   14
      Top             =   1515
      Width           =   735
   End
   Begin VB.Label lhlHorasCatedra 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Horas Cátedra:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   1515
      Width           =   1095
   End
   Begin VB.Label lblAsignatura 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Asignatura:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   360
      TabIndex        =   4
      Top             =   1155
      Width           =   855
   End
   Begin VB.Label lblCarrera 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Carrera:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   795
      Width           =   975
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Apellido y Nombres:"
      ForeColor       =   &H00FFFFFF&
      Height          =   252
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   1572
   End
End
Attribute VB_Name = "frmMovimientos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rstPersonal As New Recordset
Dim rstCarreras As New Recordset
Dim rstMaterias As New Recordset
Dim Conexion As New Connection

Private Sub dbgMovimientos_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    If mdbMovimientos.Recordset.EditMode <> adEditNone Then Exit Sub
    txtCarrera.Text = ""
    txtAño.Text = ""
    txtDivision.Text = ""
    txtTurno.Text = ""
    txtAsignatura.Text = ""
    txtHoras.Text = ""
    txtModulos.Text = ""
    txtCupof.Text = ""
    txtCupof2.Text = ""
    txtRevista.Text = ""
    txtDesde.Text = ""
    txtHasta.Text = ""
    If mdbMovimientos.Recordset.EOF = False Then
        If mdbMovimientos.Recordset!Carrera <> "" Then txtCarrera.Text = mdbMovimientos.Recordset!Carrera
        If mdbMovimientos.Recordset!Asignatura <> "" Then txtAsignatura.Text = mdbMovimientos.Recordset!Asignatura
        If mdbMovimientos.Recordset!Año <> "" Then txtAño.Text = mdbMovimientos.Recordset!Año
        If mdbMovimientos.Recordset!Division <> "" Then txtDivision.Text = mdbMovimientos.Recordset!Division
        If mdbMovimientos.Recordset!Turno <> "" Then txtTurno.Text = mdbMovimientos.Recordset!Turno
        If mdbMovimientos.Recordset!Horas <> "" Then txtHoras.Text = mdbMovimientos.Recordset!Horas
        If mdbMovimientos.Recordset!Modulos <> "" Then txtModulos.Text = mdbMovimientos.Recordset!Modulos
        If mdbMovimientos.Recordset!Cupof <> "" Then txtCupof.Text = mdbMovimientos.Recordset!Cupof
        If mdbMovimientos.Recordset!Cupof2 <> "" Then txtCupof2.Text = mdbMovimientos.Recordset!Cupof2
        If mdbMovimientos.Recordset!Revista <> "" Then txtRevista.Text = mdbMovimientos.Recordset!Revista
        If mdbMovimientos.Recordset!Desde <> "" Then txtDesde.Text = mdbMovimientos.Recordset!Desde
        If mdbMovimientos.Recordset!Hasta <> "" Then txtHasta.Text = mdbMovimientos.Recordset!Hasta
    End If
End Sub

Private Sub Form_Load()
    Conexion.ConnectionString = ("DSN=Instituto")
    Conexion.Mode = adModeReadWrite
    Conexion.Open
    Set rstPersonal = New ADODB.Recordset
    rstPersonal.Open "SELECT * FROM Personal ORDER BY Nombre", Conexion, adOpenKeyset, adLockOptimistic, adCmdText
    Set rstCarreras = New ADODB.Recordset
    Set rstMaterias = New ADODB.Recordset
End Sub

Private Sub Form_Activate()
    txtNombre.Clear
    If rstPersonal.State <> adStateClosed Then rstPersonal.Close
    rstPersonal.Open "SELECT Codigo, Nombre, Tipo, Documento, Calificacion1, Calificacion2 FROM Personal ORDER BY Nombre", Conexion, adOpenKeyset, adLockOptimistic, adCmdText
    rstPersonal.Requery
    opersonal = ""
    npersonal = 0
    While rstPersonal.EOF = False
        If opersonal <> rstPersonal!Nombre Then
            txtNombre.AddItem rstPersonal!Nombre
            txtNombre.ItemData(npersonal) = rstPersonal!Codigo
            npersonal = npersonal + 1
        End If
        opersonal = rstPersonal!Nombre
        rstPersonal.MoveNext
    Wend
    rstPersonal.MoveFirst
    txtNombre.ListIndex = 0
End Sub

Private Sub cmdNuevo_Click()
    mdbMovimientos.Recordset.AddNew
    cmdNuevo.Enabled = False
    cmdModificar.Enabled = False
    cmdEliminar.Enabled = False
    cmdGuardar.Enabled = True
    cmdCancelar.Enabled = True
    dbgMovimientos.AllowAddNew = True
    dbgMovimientos.AllowUpdate = True
    frmBuscar.Enabled = False
    txtCarrera.Enabled = True
    txtAño.Enabled = True
    txtDivision.Enabled = True
    txtTurno.Enabled = True
    txtAsignatura.Enabled = True
    txtHoras.Enabled = True
    txtModulos.Enabled = True
    txtCupof.Enabled = True
    txtCupof2.Enabled = True
    txtRevista.Enabled = True
    txtDesde.Enabled = True
    txtHasta.Enabled = True
   
    txtCarrera.Clear
    rstCarreras.Open "SELECT Codigo, Abreviatura FROM Carreras ORDER BY Abreviatura", Conexion, adOpenKeyset, adLockOptimistic, adCmdText
    rstCarreras.Requery
    ncarrera = 0
    While rstCarreras.EOF = False
        Carrera = Trim(LCase(rstCarreras!Abreviatura))
        Mid(Carrera, 1, 1) = UCase(Left(Carrera, 1))
        For i = 1 To Len(Carrera)
            If Mid(Carrera, i, 1) = " " Then
                If Mid(Carrera, i + 1, 2) = "y " Then i = i + 1: cont = 1
                If Mid(Carrera, i + 1, 2) = "a " Then i = i + 1: cont = 1
                If Mid(Carrera, i + 1, 2) = "e " Then i = i + 1: cont = 1
                If Mid(Carrera, i + 1, 2) = "o " Then i = i + 1: cont = 1
                If Mid(Carrera, i + 1, 3) = "de " Then i = i + 2: cont = 1
                If Mid(Carrera, i + 1, 3) = "la " Then i = i + 2: cont = 1
                If Mid(Carrera, i + 1, 3) = "en " Then i = i + 2: cont = 1
                If Mid(Carrera, i + 1, 4) = "del " Then i = i + 3: cont = 1
                If Mid(Carrera, i + 1, 4) = "las " Then i = i + 3: cont = 1
                If Mid(Carrera, i + 1, 4) = "los " Then i = i + 3: cont = 1
                If cont = 0 Then Mid(Carrera, i + 1, 1) = UCase(Mid(Carrera, i + 1, 1)) Else cont = 0
            End If
        Next
        txtCarrera.AddItem Carrera
        txtCarrera.ItemData(ncarrera) = rstCarreras!Codigo
        ncarrera = ncarrera + 1
        rstCarreras.MoveNext
    Wend
    rstCarreras.MoveFirst
    txtCarrera.ListIndex = 0
    rstCarreras.Close
End Sub

Private Sub cmdModificar_Click()
    cmdNuevo.Enabled = False
    cmdModificar.Enabled = False
    cmdEliminar.Enabled = False
    cmdGuardar.Enabled = True
    cmdCancelar.Enabled = True
    dbgMovimientos.AllowDelete = True
    dbgMovimientos.AllowUpdate = True
    frmBuscar.Enabled = False
    txtCarrera.Enabled = True
    txtAño.Enabled = True
    txtDivision.Enabled = True
    txtTurno.Enabled = True
    txtAsignatura.Enabled = True
    txtHoras.Enabled = True
    txtModulos.Enabled = True
    txtCupof.Enabled = True
    txtCupof2.Enabled = True
    txtRevista.Enabled = True
    txtDesde.Enabled = True
    txtHasta.Enabled = True
    
    txtCarrera.Clear
    rstCarreras.Open "SELECT Codigo, Abreviatura FROM Carreras ORDER BY Abreviatura", Conexion, adOpenKeyset, adLockOptimistic, adCmdText
    rstCarreras.Requery
    ncarrera = 0
    While rstCarreras.EOF = False
        Carrera = Trim(LCase(rstCarreras!Abreviatura))
        Mid(Carrera, 1, 1) = UCase(Left(Carrera, 1))
        For i = 1 To Len(Carrera)
            If Mid(Carrera, i, 1) = " " Then
                If Mid(Carrera, i + 1, 2) = "y " Then i = i + 1: cont = 1
                If Mid(Carrera, i + 1, 2) = "a " Then i = i + 1: cont = 1
                If Mid(Carrera, i + 1, 2) = "e " Then i = i + 1: cont = 1
                If Mid(Carrera, i + 1, 2) = "o " Then i = i + 1: cont = 1
                If Mid(Carrera, i + 1, 3) = "de " Then i = i + 2: cont = 1
                If Mid(Carrera, i + 1, 3) = "la " Then i = i + 2: cont = 1
                If Mid(Carrera, i + 1, 3) = "en " Then i = i + 2: cont = 1
                If Mid(Carrera, i + 1, 4) = "del " Then i = i + 3: cont = 1
                If Mid(Carrera, i + 1, 4) = "las " Then i = i + 3: cont = 1
                If Mid(Carrera, i + 1, 4) = "los " Then i = i + 3: cont = 1
                If cont = 0 Then Mid(Carrera, i + 1, 1) = UCase(Mid(Carrera, i + 1, 1)) Else cont = 0
            End If
        Next
        txtCarrera.AddItem Carrera
        txtCarrera.ItemData(ncarrera) = rstCarreras!Codigo
        ncarrera = ncarrera + 1
        rstCarreras.MoveNext
    Wend
    rstCarreras.MoveFirst
    txtCarrera.Text = mdbMovimientos.Recordset!Carrera
    rstCarreras.Close
    If mdbMovimientos.Recordset!CodCarrera > 0 Then
        'Call txtCarrera_Click
        'txtAsignatura.Text = mdbMovimientos.Recordset!Asignatura
    End If
End Sub

Private Sub cmdGuardar_Click()
    On Error Resume Next
    If txtDesde.Text = "" Then MsgBox "Debe especificar la Fecha de toma de posesión.", , "Movimientos": Exit Sub
    If txtHasta.Text = "" Then MsgBox "Debe especificar la Fecha de cese o si continúa.", , "Movimientos": Exit Sub
    If Len(txtCarrera.Text) = 0 Then txtCarrera.Text = "------------------"
    If Len(txtAño.Text) = 0 Then txtAño.Text = "-"
    If Len(txtDivision.Text) = 0 Then txtDivision.Text = "-"
    If Len(txtTurno.Text) = 0 Then txtTurno.Text = "-"
    If Len(txtAsignatura.Text) = 0 Then txtAsignatura.Text = "------------------"
    If Len(txtRevista.Text) = 0 Then txtRevista.Text = "P"
    If mdbMovimientos.Recordset.RecordCount = 0 Then
        mdbMovimientos.Recordset.AddNew
        mdbMovimientos.Recordset!CodPersonal = txtNombre.ItemData(txtNombre.ListIndex)
        'If txtCarrera.ItemData(txtCarrera.ListIndex) = "" Then
        '    mdbMovimientos.Recordset!CodCarrera = 0
        'Else
            mdbMovimientos.Recordset!CodCarrera = txtCarrera.ItemData(txtCarrera.ListIndex)
        'End If
        mdbMovimientos.Recordset!Carrera = txtCarrera.Text
        mdbMovimientos.Recordset!Año = txtAño.Text
        mdbMovimientos.Recordset!Division = txtDivision.Text
        mdbMovimientos.Recordset!Turno = txtTurno.Text
        mdbMovimientos.Recordset!Asignatura = txtAsignatura.Text
        mdbMovimientos.Recordset!Horas = txtHoras.Text
        mdbMovimientos.Recordset!Modulos = txtModulos.Text
        If Len(txtCupof.Text) = 0 Then mdbMovimientos.Recordset!Cupof = Null Else mdbMovimientos.Recordset!Cupof = txtCupof.Text
        If Len(txtCupof2.Text) = 0 Then mdbMovimientos.Recordset!Cupof2 = Null Else mdbMovimientos.Recordset!Cupof2 = txtCupof2.Text
        mdbMovimientos.Recordset!Revista = txtRevista.Text
        mdbMovimientos.Recordset!Desde = txtDesde.Text
        mdbMovimientos.Recordset!Hasta = txtHasta.Text
        mdbMovimientos.Recordset.Update
        Exit Sub
    Else
        mdbMovimientos.Recordset!CodPersonal = txtNombre.ItemData(txtNombre.ListIndex)
        'If txtCarrera.ItemData(txtCarrera.ListIndex) = "" Then
        '    mdbMovimientos.Recordset!CodCarrera = 0
        'Else
            mdbMovimientos.Recordset!CodCarrera = txtCarrera.ItemData(txtCarrera.ListIndex)
        'End If
        mdbMovimientos.Recordset!Carrera = txtCarrera.Text
        mdbMovimientos.Recordset!Año = txtAño.Text
        mdbMovimientos.Recordset!Division = txtDivision.Text
        mdbMovimientos.Recordset!Turno = txtTurno.Text
        mdbMovimientos.Recordset!Asignatura = txtAsignatura.Text
        mdbMovimientos.Recordset!Horas = txtHoras.Text
        mdbMovimientos.Recordset!Modulos = txtModulos.Text
        If Len(txtCupof.Text) = 0 Then mdbMovimientos.Recordset!Cupof = Null Else mdbMovimientos.Recordset!Cupof = txtCupof.Text
        If Len(txtCupof2.Text) = 0 Then mdbMovimientos.Recordset!Cupof2 = Null Else mdbMovimientos.Recordset!Cupof2 = txtCupof2.Text
        mdbMovimientos.Recordset!Revista = txtRevista.Text
        mdbMovimientos.Recordset!Desde = txtDesde.Text
        mdbMovimientos.Recordset!Hasta = txtHasta.Text
        mdbMovimientos.Recordset.Update
    End If
    mdbMovimientos.Recordset.Requery
    dbgMovimientos.Refresh
    cmdNuevo.Enabled = True
    cmdModificar.Enabled = True
    cmdEliminar.Enabled = True
    cmdGuardar.Enabled = False
    cmdCancelar.Enabled = False
    dbgMovimientos.AllowAddNew = False
    dbgMovimientos.AllowDelete = False
    dbgMovimientos.AllowUpdate = False
    frmBuscar.Enabled = True
    txtCarrera.Enabled = False
    txtAño.Enabled = False
    txtDivision.Enabled = False
    txtTurno.Enabled = False
    txtAsignatura.Enabled = False
    txtHoras.Enabled = False
    txtModulos.Enabled = False
    txtCupof.Enabled = False
    txtCupof2.Enabled = False
    txtRevista.Enabled = False
    txtDesde.Enabled = False
    txtHasta.Enabled = False
End Sub

Private Sub cmdCancelar_Click()
    mdbMovimientos.Recordset.CancelUpdate
    cmdNuevo.Enabled = True
    cmdModificar.Enabled = True
    cmdEliminar.Enabled = True
    cmdGuardar.Enabled = False
    cmdCancelar.Enabled = False
    dbgMovimientos.AllowAddNew = False
    dbgMovimientos.AllowDelete = False
    dbgMovimientos.AllowUpdate = False
    frmBuscar.Enabled = True
    txtCarrera.Enabled = False
    txtAño.Enabled = False
    txtDivision.Enabled = False
    txtTurno.Enabled = False
    txtAsignatura.Enabled = False
    txtHoras.Enabled = False
    txtModulos.Enabled = False
    txtCupof.Enabled = False
    txtCupof2.Enabled = False
    txtRevista.Enabled = False
    txtDesde.Enabled = False
    txtHasta.Enabled = False
End Sub

Private Sub cmdEliminar_Click()
    If mdbMovimientos.Recordset.RecordCount > 0 Then
        mdbMovimientos.Recordset.Delete
        dbgMovimientos.ReBind
        mdbMovimientos.Recordset.Requery
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Conexion.Close
End Sub

Private Sub txtAsignatura_Click()
    'If mdbMovimientos.Recordset.EditMode <> adEditAdd And mdbMovimientos.Recordset.EditMode <> adEditInProgress Then Exit Sub
    If rstMaterias.State = adStateOpen Then rstMaterias.Close
    rstMaterias.Open "SELECT Codigo, Curso, Horas, Modulos FROM Materias WHERE Codigo = " & txtAsignatura.ItemData(txtAsignatura.ListIndex) & " ORDER BY Codigo", Conexion, adOpenKeyset, adLockOptimistic, adCmdText
    rstMaterias.Requery
    If rstMaterias!Curso <> "" Then txtAño.Text = rstMaterias!Curso
    txtDivision.Text = "A"
    If rstMaterias!Horas <> "" Then txtHoras.Text = rstMaterias!Horas
    If rstMaterias!Modulos <> "" Then txtModulos.Text = rstMaterias!Modulos
    rstMaterias.Close
End Sub

Private Sub txtBuscarNombre_Change()
    If txtBuscarNombre.Text = "" Then
        txtNombre.ListIndex = 0
        AutoLoad
        Exit Sub
    End If
    For i = 1 To txtNombre.ListCount
        If UCase(Left(txtNombre.List(i), Len(txtBuscarNombre.Text))) = UCase(txtBuscarNombre.Text) Then
            txtNombre.ListIndex = i
            AutoLoad
            Exit For
        End If
    Next
End Sub

Private Sub cmdImprimir_Click()
    If rstPersonal.State Then rstPersonal.Close
    rstPersonal.Open "SELECT Codigo, Tipo, Documento, Calificacion1, Calificacion2, Observaciones FROM Personal WHERE Codigo = " & txtNombre.ItemData(txtNombre.ListIndex), Conexion, adOpenKeyset, adLockOptimistic, adCmdText
    rstPersonal.Requery
    With DataEnvironment1
        If .rsMovimientos.State Then .rsMovimientos.Close
        .Commands!Movimientos.CommandText = "SELECT * FROM Movimientos WHERE CodPersonal = " & txtNombre.ItemData(txtNombre.ListIndex) & " ORDER BY Desde"
    End With
    rptServicios.Sections(1).Controls(1).Caption = "Certifico que el/la docente " & txtNombre.Text & ", " & rstPersonal!Tipo & ": " & rstPersonal!documento & ", se desempeña/ó en las siguientes asignaturas y/o cargos en los periodos que a continuación se detallan:"
    If rstPersonal!Calificacion1 <> "" Then
        If rstPersonal!Calificacion2 <> "" Then
            rptServicios.Sections(5).Controls(1).Caption = "Mereció en los últimos dos años de desempeño en este establecimiento las siguientes calificaciones:  " & rstPersonal!Calificacion1 & "  /  " & rstPersonal!Calificacion2
        Else
            rptServicios.Sections(5).Controls(1).Caption = "Mereció en el último año de desempeño en este establecimiento la siguiente calificación:  " & rstPersonal!Calificacion1
        End If
    Else
        If rstPersonal!Calificacion2 <> "" Then
            rptServicios.Sections(5).Controls(1).Caption = "Mereció en el último año de desempeño en este establecimiento la siguiente calificación:  " & rstPersonal!Calificacion2
        End If
    End If
    If rstPersonal!Observaciones <> "" Then
        rptServicios.Sections(5).Controls(2).Caption = rstPersonal!Observaciones
        rptServicios.Sections(5).Controls(4).Caption = "A pedido del interesado/a y para presentar ante quien corresponda, se extiende la presente en la ciudad de Junín, Provincia de Buenos Aires, el día " & Format(Date, "Long Date") & "."
    Else
        rptServicios.Sections(5).Controls(2).Caption = "A pedido del interesado/a y para presentar ante quien corresponda, se extiende la presente en la ciudad de Junín, Provincia de Buenos Aires, el día " & Format(Date, "Long Date") & "."
    End If
    rptServicios.Refresh
    rptServicios.Show 1
    EstablecerZoom rptServicios.hWnd, zoom75
End Sub

Private Sub txtCarrera_Click()
    'If mdbMovimientos.Recordset.EditMode <> adEditAdd And mdbMovimientos.Recordset.EditMode <> adEditInProgress Then Exit Sub
    txtAsignatura.Clear
    If rstMaterias.State = adStateOpen Then rstMaterias.Close
    rstMaterias.Open "SELECT Codigo, Nombre, Carrera FROM Materias WHERE Carrera = " & txtCarrera.ItemData(txtCarrera.ListIndex) & " ORDER BY Codigo, Nombre", Conexion, adOpenKeyset, adLockOptimistic, adCmdText
    rstMaterias.Requery
    nmateria = 0
    While rstMaterias.EOF = False
        If rstMaterias!Nombre <> "" Then
            Asignatura = Trim(LCase(rstMaterias!Nombre))
            Mid(Asignatura, 1, 1) = UCase(Left(Asignatura, 1))
            For i = 1 To Len(Asignatura)
                If Mid(Asignatura, i, 1) = " " Then
                    If Mid(Asignatura, i + 1, 2) = "y " Then i = i + 1: cont = 1
                    If Mid(Asignatura, i + 1, 2) = "a " Then i = i + 1: cont = 1
                    If Mid(Asignatura, i + 1, 2) = "e " Then i = i + 1: cont = 1
                    If Mid(Asignatura, i + 1, 2) = "o " Then i = i + 1: cont = 1
                    If Mid(Asignatura, i + 1, 3) = "de " Then i = i + 2: cont = 1
                    If Mid(Asignatura, i + 1, 3) = "la " Then i = i + 2: cont = 1
                    If Mid(Asignatura, i + 1, 3) = "en " Then i = i + 2: cont = 1
                    If Mid(Asignatura, i + 1, 4) = "del " Then i = i + 3: cont = 1
                    If Mid(Asignatura, i + 1, 4) = "las " Then i = i + 3: cont = 1
                    If Mid(Asignatura, i + 1, 4) = "los " Then i = i + 3: cont = 1
                    If cont = 0 Then Mid$(Asignatura, i + 1, 1) = UCase(Mid$(Asignatura, i + 1, 1)) Else cont = 0
                End If
            Next
            If Right(Asignatura, 2) = " i" Then Mid(Asignatura, Len(Asignatura) - 1, 2) = " I"
            If Right(Asignatura, 3) = " Ii" Then Mid(Asignatura, Len(Asignatura) - 2, 3) = " II"
            If Right(Asignatura, 4) = " Iii" Then Mid(Asignatura, Len(Asignatura) - 3, 4) = " III"
            If Right(Asignatura, 3) = " Iv" Then Mid(Asignatura, Len(Asignatura) - 2, 3) = " IV"
            txtAsignatura.AddItem Asignatura
            txtAsignatura.ItemData(nmateria) = rstMaterias!Codigo
            nmateria = nmateria + 1
        End If
        rstMaterias.MoveNext
    Wend
    If txtAsignatura.ListCount > 0 Then txtAsignatura.ListIndex = 0
    If rstMaterias.State = adStateOpen Then rstMaterias.Close
End Sub

Private Sub txtNombre_Click()
    AutoLoad
End Sub

Private Sub AutoLoad()
    mdbMovimientos.RecordSource = "SELECT * FROM Movimientos WHERE CodPersonal = " & txtNombre.ItemData(txtNombre.ListIndex) & " ORDER BY Desde"
    mdbMovimientos.Refresh
    dbgMovimientos.Refresh
    txtCarrera.Text = ""
    txtAño.Text = ""
    txtDivision.Text = ""
    txtTurno.Text = ""
    txtAsignatura.Text = ""
    txtHoras.Text = ""
    txtModulos.Text = ""
    txtCupof.Text = ""
    txtCupof2.Text = ""
    txtRevista.Text = ""
    txtDesde.Text = ""
    txtHasta.Text = ""
    If mdbMovimientos.Recordset.EOF = False Then
        If mdbMovimientos.Recordset!Carrera <> "" Then txtCarrera.Text = mdbMovimientos.Recordset!Carrera
        If mdbMovimientos.Recordset!Asignatura <> "" Then txtAsignatura.Text = mdbMovimientos.Recordset!Asignatura
        If mdbMovimientos.Recordset!Año <> "" Then txtAño.Text = mdbMovimientos.Recordset!Año
        If mdbMovimientos.Recordset!Division <> "" Then txtDivision.Text = mdbMovimientos.Recordset!Division
        If mdbMovimientos.Recordset!Turno <> "" Then txtTurno.Text = mdbMovimientos.Recordset!Turno
        If mdbMovimientos.Recordset!Horas <> "" Then txtHoras.Text = mdbMovimientos.Recordset!Horas
        If mdbMovimientos.Recordset!Modulos <> "" Then txtModulos.Text = mdbMovimientos.Recordset!Modulos
        If mdbMovimientos.Recordset!Cupof <> "" Then txtCupof.Text = mdbMovimientos.Recordset!Cupof
        If mdbMovimientos.Recordset!Cupof2 <> "" Then txtCupof2.Text = mdbMovimientos.Recordset!Cupof2
        If mdbMovimientos.Recordset!Revista <> "" Then txtRevista.Text = mdbMovimientos.Recordset!Revista
        If mdbMovimientos.Recordset!Desde <> "" Then txtDesde.Text = mdbMovimientos.Recordset!Desde
        If mdbMovimientos.Recordset!Hasta <> "" Then txtHasta.Text = mdbMovimientos.Recordset!Hasta
        cmdModificar.Enabled = True
        cmdEliminar.Enabled = True
        cmdImprimir.Enabled = True
    Else
        cmdModificar.Enabled = False
        cmdEliminar.Enabled = False
        cmdImprimir.Enabled = False
    End If
End Sub

Private Sub txtRevista_Click()
    If txtRevista.Text <> "T" And txtRevista.Text <> "P" And txtRevista.Text <> "S" And txtRevista.Text <> "I" And txtRevista.Text <> "T/P" Then
        txtRevista = "P"
    ElseIf txtRevista.Text = "P" Then
        txtRevista.Text = "S"
    ElseIf txtRevista.Text = "S" Then
        txtRevista.Text = "I"
    ElseIf txtRevista.Text = "I" Then
        txtRevista.Text = "T/P"
    ElseIf txtRevista.Text = "T/P" Then
        txtRevista.Text = "T"
    ElseIf txtRevista.Text = "T" Then
        txtRevista.Text = "P"
    End If
End Sub

Private Sub txtTurno_Click()
    If txtTurno.Text <> "" And txtTurno.Text <> "M" And txtTurno.Text <> "T" And txtTurno.Text <> "V" And txtTurno.Text <> "N" Then
        txtTurno = ""
    ElseIf txtTurno.Text = "" Then
        txtTurno.Text = "M"
    ElseIf txtTurno.Text = "M" Then
        txtTurno.Text = "T"
    ElseIf txtTurno.Text = "T" Then
        txtTurno.Text = "V"
    ElseIf txtTurno.Text = "V" Then
        txtTurno.Text = "N"
    ElseIf txtTurno.Text = "N" Then
        txtTurno.Text = "M"
    End If
End Sub
