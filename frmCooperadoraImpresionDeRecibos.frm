VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Begin VB.Form frmCooperadoraImpresionDeRecibos 
   Caption         =   "Impresion de Recibos"
   ClientHeight    =   8190
   ClientLeft      =   3660
   ClientTop       =   810
   ClientWidth     =   11040
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   8190
   ScaleWidth      =   11040
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdSalir 
      Caption         =   "Salir"
      Height          =   975
      Left            =   9240
      Picture         =   "frmCooperadoraImpresionDeRecibos.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   6840
      Width           =   1215
   End
   Begin VB.Frame Frame2 
      Caption         =   "Desmarcar impresos"
      Height          =   1575
      Left            =   5160
      TabIndex        =   7
      Top             =   6480
      Width           =   3495
      Begin VB.CommandButton cmdDesmarcar 
         Caption         =   "Desmarcar"
         Height          =   735
         Left            =   840
         Picture         =   "frmCooperadoraImpresionDeRecibos.frx":08CA
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   720
         Width           =   1815
      End
      Begin VB.TextBox txtReimprimir 
         Height          =   375
         Left            =   2400
         TabIndex        =   9
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label2 
         Caption         =   "Desmarcar a partir del orden:"
         Height          =   375
         Left            =   240
         TabIndex        =   8
         Top             =   360
         Width           =   2175
      End
   End
   Begin Crystal.CrystalReport rptRecibo 
      Left            =   1920
      Top             =   6720
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "recibo.rpt"
      Destination     =   1
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
   End
   Begin VB.TextBox txtHojasAImprimir 
      Height          =   375
      Left            =   1680
      TabIndex        =   2
      Top             =   7200
      Width           =   855
   End
   Begin VB.CommandButton cmdImprimir 
      Caption         =   "Imprimir"
      Height          =   975
      Left            =   2880
      Picture         =   "frmCooperadoraImpresionDeRecibos.frx":1194
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   6840
      Width           =   1575
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frmCooperadoraImpresionDeRecibos.frx":1A5E
      Height          =   5415
      Left            =   120
      TabIndex        =   0
      Top             =   960
      Width           =   10695
      _ExtentX        =   18865
      _ExtentY        =   9551
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
      ColumnCount     =   9
      BeginProperty Column00 
         DataField       =   "Orden"
         Caption         =   "Orden"
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
      BeginProperty Column02 
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
      BeginProperty Column03 
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
      BeginProperty Column04 
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
      BeginProperty Column05 
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
      BeginProperty Column06 
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
      BeginProperty Column07 
         DataField       =   "DomicilioEnJunin"
         Caption         =   "DomicilioEnJunin"
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
      SplitCount      =   1
      BeginProperty Split0 
         MarqueeStyle    =   4
         BeginProperty Column00 
            ColumnWidth     =   689.953
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   750.047
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   2009.764
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   615.118
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   915.024
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   1890.142
         EndProperty
         BeginProperty Column07 
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column08 
            ColumnWidth     =   615.118
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc adoaAlumnosSinImprimir 
      Height          =   615
      Left            =   3240
      Top             =   5880
      Visible         =   0   'False
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   1085
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
      RecordSource    =   $"frmCooperadoraImpresionDeRecibos.frx":1A83
      Caption         =   "adoAlumnosSinImprimir"
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
      Caption         =   "Impresión de recibos"
      Height          =   1575
      Left            =   120
      TabIndex        =   5
      Top             =   6480
      Width           =   4695
      Begin VB.Label Label1 
         Caption         =   "Cantidad a imprimir:"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   840
         Width           =   1575
      End
   End
   Begin VB.Label lblRestantes 
      Caption         =   "Recibos a emitir"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   720
      Width           =   4455
   End
   Begin VB.Label lblTitulo 
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   3
      Top             =   240
      Width           =   8775
   End
End
Attribute VB_Name = "frmCooperadoraImpresionDeRecibos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Conexion As New Connection
Dim ConceptosDelAlumno As New Recordset
Dim Orden As Integer

Private Sub cmdDesmarcar_Click()
    If txtReimprimir = "" Then
        MsgBox ("Especifique desde que Orden desmarca")
        Exit Sub
    End If
    Respuesta = MsgBox("A continuación desmarcará los recibos especificados" & Chr(13) & "¿Continua?", vbYesNo, "Atención")
    If Respuesta = vbNo Then Exit Sub
    Me.MousePointer = 11
    Conexion.Open
    Conexion.Execute ("UPDATE Recibos_Cooperadora SET Recibos_Cooperadora.Impreso = False WHERE (((Recibos_Cooperadora.Periodo)=" & frmCooperadora.adoPeriodos.Recordset!Periodo & ") AND ((Recibos_Cooperadora.Orden)>=" & txtReimprimir & ") AND ((Recibos_Cooperadora.Impreso)=True))")
    Conexion.Close
    adoaAlumnosSinImprimir.Refresh
    lblRestantes = "Recibos restantes para imprimir: " & adoaAlumnosSinImprimir.Recordset.RecordCount
    Me.MousePointer = 0
End Sub

Private Sub cmdImprimir_Click()
   If txtHojasAImprimir = "" Or adoaAlumnosSinImprimir.Recordset.RecordCount = 0 Then Exit Sub
   Me.MousePointer = 11
   HojasImpresas = 0
   Conexion.Open
    adoaAlumnosSinImprimir.Recordset.MoveFirst
    With frmCooperadoraImprimeRecibo
    While adoaAlumnosSinImprimir.Recordset.EOF = False
        Conexion.Execute ("delete * from rptRecibos_de_cooperadora")
        .lblAlumno = adoaAlumnosSinImprimir.Recordset!Permiso & "-" & adoaAlumnosSinImprimir.Recordset!Nombre
        .lblOrden = adoaAlumnosSinImprimir.Recordset!Orden
        .lblPermisoAporteVoluntario = adoaAlumnosSinImprimir.Recordset!Permiso & "-" & adoaAlumnosSinImprimir.Recordset!Nombre
        .lblTipoYNumeroDocumento = adoaAlumnosSinImprimir.Recordset!Tipo & ": " & adoaAlumnosSinImprimir.Recordset!documento
        .lblPermisoAbril = adoaAlumnosSinImprimir.Recordset!Permiso & "-" & adoaAlumnosSinImprimir.Recordset!Nombre
        .lblPermisoMayo = adoaAlumnosSinImprimir.Recordset!Permiso & "-" & adoaAlumnosSinImprimir.Recordset!Nombre
        .lblPermisoJunio = adoaAlumnosSinImprimir.Recordset!Permiso & "-" & adoaAlumnosSinImprimir.Recordset!Nombre
        .lblPermisoJulio = adoaAlumnosSinImprimir.Recordset!Permiso & "-" & adoaAlumnosSinImprimir.Recordset!Nombre
        .lblPermisoAgosto = adoaAlumnosSinImprimir.Recordset!Permiso & "-" & adoaAlumnosSinImprimir.Recordset!Nombre
        .lblPermisoSetiembre = adoaAlumnosSinImprimir.Recordset!Permiso & "-" & adoaAlumnosSinImprimir.Recordset!Nombre
        .lblPermisoOctubre = adoaAlumnosSinImprimir.Recordset!Permiso & "-" & adoaAlumnosSinImprimir.Recordset!Nombre
        .lblPermisoNoviembre = adoaAlumnosSinImprimir.Recordset!Permiso & "-" & adoaAlumnosSinImprimir.Recordset!Nombre
        .lblAno = Anio
        If adoaAlumnosSinImprimir.Recordset!domicilio <> vacio Then
           .lblDomicilio = adoaAlumnosSinImprimir.Recordset!domicilio
        Else
           .lblDomicilio = ""
        End If
        If adoaAlumnosSinImprimir.Recordset!DomicilioEnJunin <> vacio Then
            .lblDomicilioEnJunin = adoaAlumnosSinImprimir.Recordset!DomicilioEnJunin
        Else
           .lblDomicilioEnJunin = ""
        End If
        If adoaAlumnosSinImprimir.Recordset!Localidad <> vacio Then
            .lblLocalidad = adoaAlumnosSinImprimir.Recordset!Localidad
        Else
            .lblLocalidad = ""
        End If
        Set ConceptosDelAlumno = Conexion.Execute("SELECT DISTINCT [Cooperadora Conceptos].Concepto, Recibos_Cooperadora.Importe, Recibos_Cooperadora.Comprobante, Recibos_Cooperadora.Orden FROM Recibos_Cooperadora INNER JOIN [Cooperadora Conceptos] ON Recibos_Cooperadora.Concepto = [Cooperadora Conceptos].Codigo WHERE (((Recibos_Cooperadora.Periodo)=" & frmCooperadora.adoPeriodos.Recordset!Periodo & ") AND ((Recibos_Cooperadora.Alumno)=" & adoaAlumnosSinImprimir.Recordset!Permiso & ")) ORDER BY Recibos_Cooperadora.Comprobante")
        Orden = ConceptosDelAlumno!Orden
        While ConceptosDelAlumno.EOF = False
           .lblReciboAportevoluntario = ConceptosDelAlumno!Comprobante
           .lblReciboNumeroaporteVoluntario(0) = ConceptosDelAlumno!Comprobante
           ConceptosDelAlumno.MoveNext
           .lblReciboAbril = ConceptosDelAlumno!Comprobante
           .lblReciboNumeroAbril(0) = ConceptosDelAlumno!Comprobante
           ConceptosDelAlumno.MoveNext
           .lblReciboMayo = ConceptosDelAlumno!Comprobante
           .lblReciboNumeroMayo(0) = ConceptosDelAlumno!Comprobante
           ConceptosDelAlumno.MoveNext
           .lblReciboJunio = ConceptosDelAlumno!Comprobante
           .lblReciboNumeroJunio(0) = ConceptosDelAlumno!Comprobante
           ConceptosDelAlumno.MoveNext
           .lblReciboJulio = ConceptosDelAlumno!Comprobante
           .lblReciboNumeroJulio(0) = ConceptosDelAlumno!Comprobante
           ConceptosDelAlumno.MoveNext
           .lblReciboAgosto = ConceptosDelAlumno!Comprobante
           .lblReciboNumeroAgosto(0) = ConceptosDelAlumno!Comprobante
           ConceptosDelAlumno.MoveNext
           .lblReciboSetiembre = ConceptosDelAlumno!Comprobante
           .lblReciboNumeroSetiembre(0) = ConceptosDelAlumno!Comprobante
           ConceptosDelAlumno.MoveNext
           .lblReciboOctubre = ConceptosDelAlumno!Comprobante
           .lblReciboNumeroOctubre(0) = ConceptosDelAlumno!Comprobante
           ConceptosDelAlumno.MoveNext
           .lblReciboNoviembre = ConceptosDelAlumno!Comprobante
           .lblReciboNumeroNoviembre(0) = ConceptosDelAlumno!Comprobante
           ConceptosDelAlumno.MoveNext
        Wend
        'frmCooperadoraImprimeRecibo.PrintForm
        Conexion.Execute ("INSERT INTO rptRecibos_de_cooperadora ( Permiso, Nombre, Tipo, Documento, Domicilio, Localidad, Domicilio_en_junin, Ano, AV, Abril, Mayo, Junio, Julio, Agosto, Septiembre, Octubre, Noviembre, Orden )  values (" & adoaAlumnosSinImprimir.Recordset!Permiso & ",'" & adoaAlumnosSinImprimir.Recordset!Nombre & "','" & adoaAlumnosSinImprimir.Recordset!Tipo & "'," & adoaAlumnosSinImprimir.Recordset!documento & ",'" & .lblDomicilio & "','" & .lblLocalidad & "','" & .lblDomicilioEnJunin & "'," & .lblAno & "," & .lblReciboAportevoluntario & "," & .lblReciboAbril & "," & .lblReciboMayo & "," & .lblReciboJunio & "," & .lblReciboJulio & "," & .lblReciboAgosto & "," & .lblReciboSetiembre & "," & .lblReciboOctubre & "," & .lblReciboNoviembre & ", " & Orden & ")")
        Conexion.Close
        Conexion.Open
        rptRecibo.PrintReport
        Conexion.Execute ("UPDATE Recibos_Cooperadora SET Recibos_Cooperadora.Impreso = 1 WHERE (((Recibos_Cooperadora.Alumno)=" & adoaAlumnosSinImprimir.Recordset!Permiso & ") AND ((Recibos_Cooperadora.Ano)=" & Anio & "))")
        adoaAlumnosSinImprimir.Recordset.MoveNext
        lblRestantes = "Recibos restantes para imprimir: " & adoaAlumnosSinImprimir.Recordset.RecordCount
        HojasImpresas = HojasImpresas + 1
        If HojasImpresas = txtHojasAImprimir Then
            Me.MousePointer = 0
            MsgBox ("Ha terminado de imprimir la cantidad de hojas seleccionadas")
            adoaAlumnosSinImprimir.Refresh
            lblRestantes = "Recibos restantes para imprimir: " & adoaAlumnosSinImprimir.Recordset.RecordCount
            txtHojasAImprimir = 0
            Conexion.Close
            Exit Sub
        End If
    Wend
    Conexion.Close
    End With
End Sub

Private Sub cmdReimprimir_Click()
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
    lblTitulo = "Impresión de recibos correspondientes al período " & Anio
    lblRestantes = "Recibos restantes para imprimir: " & adoaAlumnosSinImprimir.Recordset.RecordCount
End Sub

Private Sub Form_Load()
    Conexion.ConnectionString = ("DSN=Instituto")
End Sub

Private Sub txtHojasAImprimir_KeyPress(KeyAscii As Integer)
   If InStr(1, "0123456789" & Chr(13) & Chr(8), Chr(KeyAscii)) = 0 Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtReimprimir_KeyPress(KeyAscii As Integer)
   If InStr(1, "0123456789" & Chr(13) & Chr(8), Chr(KeyAscii)) = 0 Then
        KeyAscii = 0
    End If
End Sub
