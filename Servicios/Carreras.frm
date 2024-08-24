VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmCarreras 
   BackColor       =   &H00404040&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Carreras"
   ClientHeight    =   8076
   ClientLeft      =   48
   ClientTop       =   432
   ClientWidth     =   11304
   Icon            =   "Carreras.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8076
   ScaleWidth      =   11304
   Begin VB.TextBox txtCodigo 
      BackColor       =   &H00B0FFB0&
      Enabled         =   0   'False
      Height          =   285
      Left            =   1560
      MaxLength       =   5
      TabIndex        =   3
      Top             =   840
      Width           =   1215
   End
   Begin MSAdodcLib.Adodc mdbMaterias 
      Height          =   330
      Left            =   9720
      Top             =   7200
      Visible         =   0   'False
      Width           =   1215
      _ExtentX        =   2138
      _ExtentY        =   572
      ConnectMode     =   3
      CursorLocation  =   2
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
      RecordSource    =   "SELECT * FROM Materias WHERE Carrera = 9999 ORDER BY Codigo, Curso, Nombre"
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSDataGridLib.DataGrid dbgMaterias 
      Bindings        =   "Carreras.frx":0442
      Height          =   4455
      Left            =   120
      TabIndex        =   21
      Top             =   3480
      Width           =   11055
      _ExtentX        =   19495
      _ExtentY        =   7853
      _Version        =   393216
      AllowUpdate     =   -1  'True
      HeadLines       =   1
      RowHeight       =   15
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Asignaturas"
      ColumnCount     =   12
      BeginProperty Column00 
         DataField       =   "Codigo"
         Caption         =   "Código"
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
      BeginProperty Column02 
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
      BeginProperty Column03 
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
      BeginProperty Column04 
         DataField       =   "Abreviatura"
         Caption         =   "Abreviatura"
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
         DataField       =   "Horas"
         Caption         =   "H/C"
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
         DataField       =   "Modulos"
         Caption         =   "Mód."
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
      BeginProperty Column07 
         DataField       =   "CUPOF"
         Caption         =   "CUPOF"
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
      BeginProperty Column08 
         DataField       =   "AbreviaturaPof"
         Caption         =   "Contralor"
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
      BeginProperty Column09 
         DataField       =   "Modalidad"
         Caption         =   "Modalidad"
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
      BeginProperty Column10 
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
      BeginProperty Column11 
         DataField       =   "Eliminada"
         Caption         =   "Eliminada"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   5
            Format          =   ""
            HaveTrueFalseNull=   1
            TrueValue       =   "Sí"
            FalseValue      =   "No"
            NullValue       =   ""
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   11274
            SubFormatType   =   7
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         MarqueeStyle    =   4
         BeginProperty Column00 
            Alignment       =   2
            ColumnWidth     =   659,906
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   2232
         EndProperty
         BeginProperty Column02 
            Alignment       =   2
            ColumnWidth     =   599,811
         EndProperty
         BeginProperty Column03 
            Alignment       =   2
            ColumnWidth     =   515,906
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   1692,284
         EndProperty
         BeginProperty Column05 
            Alignment       =   2
            ColumnWidth     =   396,283
         EndProperty
         BeginProperty Column06 
            Alignment       =   2
            ColumnWidth     =   420,095
         EndProperty
         BeginProperty Column07 
            Alignment       =   2
            ColumnWidth     =   815,811
         EndProperty
         BeginProperty Column08 
            Alignment       =   2
            ColumnWidth     =   887,811
         EndProperty
         BeginProperty Column09 
            ColumnWidth     =   815,811
         EndProperty
         BeginProperty Column10 
            ColumnWidth     =   612,284
         EndProperty
         BeginProperty Column11 
            Alignment       =   2
            ColumnWidth     =   768,189
         EndProperty
      EndProperty
   End
   Begin VB.ComboBox txtAños 
      Enabled         =   0   'False
      Height          =   315
      ItemData        =   "Carreras.frx":045C
      Left            =   1560
      List            =   "Carreras.frx":046C
      TabIndex        =   11
      Top             =   1920
      Width           =   1695
   End
   Begin VB.CheckBox chkVigente 
      BackColor       =   &H00404040&
      Height          =   255
      Left            =   9120
      TabIndex        =   5
      Top             =   840
      Width           =   255
   End
   Begin VB.ComboBox txtModalidad 
      Enabled         =   0   'False
      Height          =   315
      ItemData        =   "Carreras.frx":047C
      Left            =   1560
      List            =   "Carreras.frx":0486
      TabIndex        =   9
      Top             =   1560
      Width           =   2535
   End
   Begin VB.TextBox txtAbreviatura 
      Enabled         =   0   'False
      Height          =   285
      Left            =   1560
      MaxLength       =   30
      TabIndex        =   13
      Top             =   2280
      Width           =   5055
   End
   Begin VB.TextBox txtAbreviaturaPof 
      Enabled         =   0   'False
      Height          =   285
      Left            =   1560
      MaxLength       =   5
      TabIndex        =   15
      Top             =   2640
      Width           =   1455
   End
   Begin VB.ComboBox txtNombre 
      Height          =   315
      ItemData        =   "Carreras.frx":04A6
      Left            =   1560
      List            =   "Carreras.frx":04A8
      TabIndex        =   1
      Top             =   90
      Width           =   7815
   End
   Begin VB.CommandButton cmdGuardar 
      BackColor       =   &H00FFE0C0&
      Caption         =   "Guardar"
      Enabled         =   0   'False
      Height          =   375
      Left            =   9960
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   2160
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancelar 
      BackColor       =   &H00FFE0C0&
      Caption         =   "Cancelar"
      Enabled         =   0   'False
      Height          =   375
      Left            =   9960
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   2640
      Width           =   1215
   End
   Begin VB.CommandButton cmdEliminar 
      BackColor       =   &H00FFE0C0&
      Caption         =   "Eliminar"
      Height          =   375
      Left            =   9960
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   1320
      Width           =   1215
   End
   Begin VB.CommandButton cmdModificar 
      BackColor       =   &H00FFE0C0&
      Caption         =   "Modificar"
      Height          =   375
      Left            =   9960
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   840
      Width           =   1215
   End
   Begin VB.CommandButton cmdNuevo 
      BackColor       =   &H00FFE0C0&
      Caption         =   "Nuevo"
      Height          =   375
      Left            =   9960
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   120
      Width           =   1215
   End
   Begin VB.TextBox txtResoluciones 
      Enabled         =   0   'False
      Height          =   285
      Left            =   1560
      MaxLength       =   60
      TabIndex        =   7
      Top             =   1200
      Width           =   6255
   End
   Begin VB.Label lblCodigo 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Código:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   1335
   End
   Begin VB.Line Line5 
      BorderColor     =   &H00FFFFFF&
      X1              =   9720
      X2              =   11280
      Y1              =   1920
      Y2              =   1920
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00FFFFFF&
      X1              =   9720
      X2              =   11280
      Y1              =   600
      Y2              =   600
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00FFFFFF&
      X1              =   9720
      X2              =   9720
      Y1              =   0
      Y2              =   3240
   End
   Begin VB.Label lblVigente 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Vigente:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   8280
      TabIndex        =   4
      Top             =   840
      Width           =   735
   End
   Begin VB.Label lblAbreviatura 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Abreviatura:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   2280
      Width           =   1335
   End
   Begin VB.Label lblAbreviaturaPof 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Código Contralor:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   2640
      Width           =   1335
   End
   Begin VB.Label lblModalidad 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Modalidad:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   1560
      Width           =   1335
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      X1              =   0
      X2              =   11280
      Y1              =   3240
      Y2              =   3240
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      X1              =   0
      X2              =   9720
      Y1              =   600
      Y2              =   600
   End
   Begin VB.Label lblAños 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Años:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   1920
      Width           =   1335
   End
   Begin VB.Label lblResoluciones 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Resoluciones:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   1200
      Width           =   1335
   End
   Begin VB.Label lblNombre 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Carrera:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1335
   End
End
Attribute VB_Name = "frmCarreras"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mdbCarreras As Database
Dim rstCarreras As New Recordset
Dim Conexion As New Connection

Private Sub Form_Load()
    Conexion.ConnectionString = ("DSN=Instituto")
    Conexion.Mode = adModeReadWrite
    Conexion.Open
    Set rstCarreras = New ADODB.Recordset
    rstCarreras.Open "SELECT * FROM Carreras ORDER BY Nombre", Conexion, adOpenKeyset, adLockOptimistic, adCmdText
End Sub

Private Sub Form_Activate()
    txtNombre.Clear
    rstCarreras.Close
    rstCarreras.Open "SELECT Nombre FROM Carreras ORDER BY Nombre", Conexion, adOpenKeyset, adLockOptimistic, adCmdText
    rstCarreras.Requery
    While rstCarreras.EOF = False
        If ocarrera <> rstCarreras!Nombre Then txtNombre.AddItem rstCarreras!Nombre
        ocarrera = rstCarreras!Nombre
        rstCarreras.MoveNext
    Wend
    rstCarreras.MoveFirst
    txtNombre.ListIndex = 0
End Sub

Private Sub cmdNuevo_Click()
    rstCarreras.Close
    rstCarreras.Open "SELECT * FROM Carreras ORDER BY Codigo", Conexion, adOpenKeyset, adLockOptimistic, adCmdText
    NuevoCodigo = 0
    If rstCarreras.RecordCount > 0 Then
        For i = 0 To rstCarreras.RecordCount - 1
            If rstCarreras!Codigo > i Then
                NuevoCodigo = i
                Exit For
            End If
            rstCarreras.MoveNext
        Next
        If rstCarreras.EOF Then NuevoCodigo = i
    End If
    txtCodigo.Text = NuevoCodigo
    rstCarreras.AddNew
    chkVigente.Value = 1
    cmdNuevo.Enabled = False
    cmdModificar.Enabled = False
    cmdEliminar.Enabled = False
    cmdGuardar.Enabled = True
    cmdCancelar.Enabled = True
    txtAños.Enabled = True
    txtResoluciones.Enabled = True
    txtModalidad.Enabled = True
    txtAbreviatura.Enabled = True
    txtAbreviaturaPof.Enabled = True
    chkVigente.Enabled = True
    txtNombre.Text = ""
    txtAños.Text = ""
    txtResoluciones.Text = ""
    txtModalidad.Text = ""
    txtAbreviatura.Text = ""
    txtAbreviaturaPof.Text = ""
End Sub

Private Sub cmdModificar_Click()
    cmdNuevo.Enabled = False
    cmdModificar.Enabled = False
    cmdEliminar.Enabled = False
    cmdGuardar.Enabled = True
    cmdCancelar.Enabled = True
    dbgMaterias.AllowAddNew = True
    dbgMaterias.AllowDelete = True
    dbgMaterias.AllowUpdate = True
    txtAños.Enabled = True
    txtResoluciones.Enabled = True
    txtModalidad.Enabled = True
    txtAbreviatura.Enabled = True
    txtAbreviaturaPof.Enabled = True
    chkVigente.Enabled = True
End Sub

Private Sub cmdGuardar_Click()
    On Error Resume Next
    If Len(txtNombre.Text) = 0 Then MsgBox "Debe especificar el Nombre de la Carrera.", , "Carreras": Exit Sub
    If txtResoluciones.Text = "" Then MsgBox "Debe especificar la/s Resolución/es de la Carrera.", , "Carreras": Exit Sub
    If txtModalidad.Text = "" Then MsgBox "Debe especificar la Modalidad de la Carrera.", , "Carreras": Exit Sub
    If Val(txtAños.Text) < "0" Or Val(txtAños.Text) > 4 Then MsgBox "Debe especificar entre 1 y 4 años de duración de la Carrera.", , "Carreras": Exit Sub
    If txtAbreviatura.Text = "" Then MsgBox "Debe especificar la Abreviatura de la Carrera.", , "Carreras": Exit Sub
    If txtAbreviaturaPof.Text = "" Then MsgBox "Debe especificar el Código de Contralor de la Carrera.", , "Carreras": Exit Sub
    If rstCarreras.RecordCount = 0 Then
        rstCarreras.AddNew
        rstCarreras!Codigo = txtCodigo.Text
        rstCarreras!Nombre = txtNombre.Text
        rstCarreras!Resolucion = txtResoluciones.Text
        rstCarreras!Modalidad = txtModalidad.ItemData(txtModalidad.ListIndex)
        rstCarreras!Años = txtAños.Text
        rstCarreras!Abreviatura = txtAbreviatura.Text
        rstCarreras!AbreviaturaPof = txtAbreviaturaPof.Text
        rstCarreras!Vigente = chkVigente.Value
        rstCarreras!Caracteristica = 0
        rstCarreras.Update
        Exit Sub
    Else
        rstCarreras!Codigo = txtCodigo.Text
        rstCarreras!Nombre = txtNombre.Text
        rstCarreras!Resolucion = txtResoluciones.Text
        rstCarreras!Modalidad = txtModalidad.ItemData(txtModalidad.ListIndex)
        rstCarreras!Años = txtAños.Text
        rstCarreras!Abreviatura = txtAbreviatura.Text
        rstCarreras!AbreviaturaPof = txtAbreviaturaPof.Text
        rstCarreras!Vigente = chkVigente.Value
        rstCarreras.Update
    End If
    
    If rstCarreras.EditMode = adEditInProgress Then rstCarreras.Update
    If rstCarreras.EditMode = adEditAdd Then rstCarreras.Update: rstCarreras.Requery
    cmdNuevo.Enabled = True
    cmdModificar.Enabled = True
    cmdEliminar.Enabled = True
    cmdGuardar.Enabled = False
    cmdCancelar.Enabled = False
    dbgMaterias.AllowAddNew = False
    dbgMaterias.AllowDelete = False
    dbgMaterias.AllowUpdate = False
    txtAños.Enabled = False
    txtResoluciones.Enabled = False
    txtModalidad.Enabled = False
    txtAbreviatura.Enabled = False
    txtAbreviaturaPof.Enabled = False
    chkVigente.Enabled = False
End Sub

Private Sub cmdCancelar_Click()
    rstCarreras.CancelUpdate
    cmdNuevo.Enabled = True
    cmdModificar.Enabled = True
    cmdEliminar.Enabled = True
    cmdGuardar.Enabled = False
    cmdCancelar.Enabled = False
    dbgMaterias.AllowAddNew = False
    dbgMaterias.AllowDelete = False
    dbgMaterias.AllowUpdate = False
    txtAños.Enabled = False
    txtResoluciones.Enabled = False
    txtModalidad.Enabled = False
    txtAbreviatura.Enabled = False
    txtAbreviaturaPof.Enabled = False
    chkVigente.Enabled = False
End Sub

Private Sub cmdEliminar_Click()
    If rstCarreras.RecordCount > 0 Then
        rstCarreras!Eliminada = True
        rstCarreras.Update
        rstCarreras.Requery
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Conexion.Close
End Sub

Private Sub txtNombre_Click()
    AutoLoad
End Sub

Private Sub AutoLoad()
    rstCarreras.Close
    rstCarreras.Open "SELECT * FROM Carreras ORDER BY Nombre", Conexion, adOpenKeyset, adLockOptimistic, adCmdText
    txtCodigo.Text = ""
    txtAños.Text = ""
    txtResoluciones.Text = ""
    txtModalidad.Text = ""
    txtAbreviatura.Text = ""
    txtAbreviaturaPof.Text = ""
    chkVigente.Value = False
    rstCarreras.Find "Nombre = '" & txtNombre.Text & "'", , adSearchForward, 1
    If rstCarreras.EOF = False Then
        If rstCarreras!Codigo <> "" Then txtCodigo.Text = rstCarreras!Codigo
        If rstCarreras!Resolucion <> "" Then txtResoluciones.Text = rstCarreras!Resolucion
        If rstCarreras!Modalidad <> "" Then txtModalidad.ListIndex = rstCarreras!Modalidad - 1
        If rstCarreras!Modalidad = 2 Then lblAños.Caption = "Cuatrimestres:" Else lblAños = "Años:"
        If rstCarreras!Años <> "" Then txtAños.Text = rstCarreras!Años
        If rstCarreras!Abreviatura <> "" Then txtAbreviatura.Text = rstCarreras!Abreviatura
        If rstCarreras!AbreviaturaPof <> "" Then txtAbreviaturaPof.Text = rstCarreras!AbreviaturaPof
        If rstCarreras!Vigente <> "" Then
            If rstCarreras!Vigente = False Then chkVigente.Value = 0
            If rstCarreras!Vigente = True Then chkVigente.Value = 1
        End If
        mdbMaterias.RecordSource = "SELECT * FROM Materias WHERE Carrera = " & rstCarreras!Codigo & " ORDER BY Codigo, Curso, Nombre"
        mdbMaterias.Refresh
        dbgMaterias.Refresh
    End If
End Sub
