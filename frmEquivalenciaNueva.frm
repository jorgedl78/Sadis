VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmEquivalenciaNueva 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   7080
   ClientLeft      =   1005
   ClientTop       =   750
   ClientWidth     =   9915
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7080
   ScaleWidth      =   9915
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdCancelar 
      Height          =   615
      Left            =   5760
      Picture         =   "frmEquivalenciaNueva.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   29
      ToolTipText     =   "Cancelar"
      Top             =   6240
      Width           =   615
   End
   Begin VB.CommandButton cmdAceptar 
      Height          =   615
      Left            =   3720
      Picture         =   "frmEquivalenciaNueva.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   28
      ToolTipText     =   "Aceptar"
      Top             =   6240
      Width           =   615
   End
   Begin VB.Frame Frame1 
      Height          =   3495
      Left            =   0
      TabIndex        =   5
      Top             =   2520
      Width           =   9855
      Begin VB.TextBox txtObservacion 
         Height          =   975
         Left            =   120
         TabIndex        =   21
         Top             =   2400
         Width           =   9615
      End
      Begin VB.CommandButton cmdBuscarAlumno 
         Caption         =   "Buscar..."
         Height          =   255
         Left            =   1800
         TabIndex        =   19
         Top             =   240
         Width           =   855
      End
      Begin VB.TextBox txtNota 
         Height          =   285
         Left            =   720
         TabIndex        =   16
         Top             =   1800
         Width           =   855
      End
      Begin MSDataListLib.DataCombo dtcInstituciones 
         Bindings        =   "frmEquivalenciaNueva.frx":0884
         Height          =   315
         Left            =   1800
         TabIndex        =   14
         Top             =   1320
         Width           =   7935
         _ExtentX        =   13996
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "Institucion"
         BoundColumn     =   "Codigo"
         Text            =   ""
      End
      Begin VB.TextBox txtAsignatura 
         Height          =   285
         Left            =   1800
         TabIndex        =   12
         Top             =   840
         Width           =   7935
      End
      Begin VB.TextBox txtPermiso 
         Height          =   285
         Left            =   840
         TabIndex        =   7
         Top             =   240
         Width           =   855
      End
      Begin MSAdodcLib.Adodc adoInstituciones 
         Height          =   330
         Left            =   5280
         Top             =   1440
         Visible         =   0   'False
         Width           =   2655
         _ExtentX        =   4683
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
         RecordSource    =   $"frmEquivalenciaNueva.frx":08A3
         Caption         =   "Instituciones"
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
      Begin MSComCtl2.DTPicker dtpFechaAcreditacion 
         DragIcon        =   "frmEquivalenciaNueva.frx":08F8
         Height          =   285
         Left            =   3840
         TabIndex        =   18
         Top             =   1800
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   503
         _Version        =   393216
         Format          =   35192833
         CurrentDate     =   37550
         MinDate         =   -36522
      End
      Begin VB.Label lblObservacion 
         Caption         =   "Observacion"
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   2160
         Width           =   1695
      End
      Begin VB.Label Label7 
         Caption         =   "Fecha de Acreditación:"
         Height          =   255
         Left            =   2040
         TabIndex        =   17
         Top             =   1800
         Width           =   1815
      End
      Begin VB.Label Label6 
         Caption         =   "Nota:"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   1800
         Width           =   615
      End
      Begin VB.Label Label5 
         Caption         =   "Institución de origen:"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   1320
         Width           =   1575
      End
      Begin VB.Label Label4 
         Caption         =   "Asignatura Aprobada:"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   840
         Width           =   1575
      End
      Begin VB.Label lblDocumento 
         Height          =   255
         Left            =   8280
         TabIndex        =   10
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label lblTipoDocumento 
         Height          =   255
         Left            =   7560
         TabIndex        =   9
         Top             =   240
         Width           =   495
      End
      Begin VB.Label lblAlumno 
         Height          =   255
         Left            =   2880
         TabIndex        =   8
         Top             =   240
         Width           =   4575
      End
      Begin VB.Label Label3 
         Caption         =   "Alumno:"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   735
      End
      Begin VB.Line Line1 
         X1              =   0
         X2              =   9840
         Y1              =   720
         Y2              =   720
      End
   End
   Begin VB.Frame Frame2 
      Height          =   735
      Left            =   0
      TabIndex        =   23
      Top             =   1800
      Width           =   9855
      Begin MSComCtl2.DTPicker dtpFechaSolicitud 
         DragIcon        =   "frmEquivalenciaNueva.frx":0D3A
         Height          =   285
         Left            =   1560
         TabIndex        =   25
         Top             =   240
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   503
         _Version        =   393216
         Format          =   91815937
         CurrentDate     =   37550
         MinDate         =   -36522
      End
      Begin MSDataListLib.DataCombo dtcProfesor 
         Bindings        =   "frmEquivalenciaNueva.frx":117C
         Height          =   315
         Left            =   3960
         TabIndex        =   26
         Top             =   240
         Width           =   5415
         _ExtentX        =   9551
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "Nombre"
         BoundColumn     =   "Codigo"
         Text            =   ""
      End
      Begin MSAdodcLib.Adodc adoProfesor 
         Height          =   330
         Left            =   4920
         Top             =   120
         Visible         =   0   'False
         Width           =   1695
         _ExtentX        =   2990
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
         RecordSource    =   "SELECT * FROM Personal ORDER BY Nombre"
         Caption         =   "Profesor"
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
      Begin VB.Label Label20 
         Caption         =   "Profesor:"
         Height          =   255
         Left            =   3240
         TabIndex        =   27
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label8 
         Caption         =   "Fecha de solicitud:"
         Height          =   255
         Left            =   120
         TabIndex        =   24
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.Label Label1 
      Caption         =   "Carrera:"
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
      TabIndex        =   22
      Top             =   600
      Width           =   975
   End
   Begin VB.Label Label2 
      Caption         =   "º año"
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
      Left            =   600
      TabIndex        =   4
      Top             =   1320
      Width           =   615
   End
   Begin VB.Label lblCurso 
      Alignment       =   1  'Right Justify
      Caption         =   "Label2"
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
      TabIndex        =   3
      Top             =   1320
      Width           =   375
   End
   Begin VB.Label lblMateria 
      Caption         =   "Label2"
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
      Left            =   1440
      TabIndex        =   2
      Top             =   1320
      Width           =   8295
   End
   Begin VB.Label lblCarrera 
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1200
      TabIndex        =   1
      Top             =   600
      Width           =   8655
   End
   Begin VB.Label lblTitulo 
      Alignment       =   2  'Center
      Caption         =   "Ingreso de Equivalencia"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   9615
   End
End
Attribute VB_Name = "frmEquivalenciaNueva"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Conexion As New Connection
Dim Resultado As New Recordset

Private Sub cmdAceptar_Click()
    If dtcProfesor.Text = "" Then MsgBox ("Debe especificar el profesor de la asignatura"): Exit Sub
    If lblAlumno.Caption = "" Then MsgBox ("No se indicò el alumno"): txtPermiso.SetFocus: Exit Sub
    If txtAsignatura.Text = "" Then MsgBox ("Debe especificar el nombre de la asignatura que presenta como equivalente"): txtAsignatura.SetFocus: Exit Sub
    If dtcInstituciones.Text = "" Then MsgBox ("Debe especificar la instituciòn de origen. Esta debe estar previamente cargada"): Exit Sub
    If txtNota.Text = "" Then MsgBox ("Falta especificar la nota"): Exit Sub
    If Not IsNumeric(txtNota) Then MsgBox ("La nota no corresponde a un número"): txtNota.SetFocus: Exit Sub
    Conexion.Open
    Conexion.Execute ("INSERT INTO Equivalencias ( Alumno, FechaSolicitud, AnoSolicitud, Asignatura, FechaAprobacion, Nota, Institucion, MateriaAReconocer, Usuario, Profesor, Observacion ) VALUES (" & txtPermiso & ", '" & DateValue(dtpFechaSolicitud) & "', " & frmEquivalencias.txtAño & ", '" & txtAsignatura & "', '" & DateValue(dtpFechaAcreditacion) & "', " & Replace(txtNota, ",", ".") & ", " & dtcInstituciones.BoundText & ", " & frmEquivalencias.dtcMaterias.BoundText & ", " & frmIdentificacion.dtcUsuarios.BoundText & ", " & dtcProfesor.BoundText & ", '" & txtObservacion & "')")
    Conexion.Close
    frmEquivalencias.adoEquivalencias.Refresh
    Unload Me
End Sub

Private Sub cmdBuscarAlumno_Click()
    Respuesta = InputBox("Ingrese Nº de Documento", "Buscar Alumno")
    If Respuesta = "" Then Exit Sub
    Conexion.Open
    Set Resultado = Conexion.Execute("SELECT Permiso,Nombre,Tipo,Documento FROM Alumnos WHERE Documento = " & Respuesta & " AND Eliminado = False")
    If Resultado.EOF = False Then
        txtPermiso = Resultado!Permiso: lblAlumno = Resultado!Nombre: lblTipoDocumento = Resultado!Tipo: lblDocumento = Resultado!documento: txtPermiso.SetFocus
    Else
        MsgBox ("El documento no existe"): txtPermiso = ""
    End If
    Conexion.Close
End Sub

Private Sub cmdCancelar_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Conexion.ConnectionString = ("DSN=Instituto")
    lblTitulo = "Ingreso de Equivalencias del año " & Str(frmEquivalencias.txtAño)
    lblCarrera = frmEquivalencias.dtcCarreras
    lblCurso = frmEquivalencias.cbCurso
    lblMateria = frmEquivalencias.dtcMaterias
    Conexion.Open
    Set Resultado = Conexion.Execute("SELECT Profesor FROM Divisiones WHERE Materia =" & frmEquivalencias.dtcMaterias.BoundText & " and Ano = " & frmEquivalencias.txtAño)
    If Resultado.EOF = False Then 'Si se ce creo la division, predetermino el profesor que dicta la materia en este año
        dtcProfesor.BoundText = Resultado!Profesor
    End If
    Conexion.Close
End Sub

Private Sub txtPermiso_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       If txtPermiso = "" Then MsgBox ("Debe ingresar un Nº de permiso"): txtPermiso.ShowWhatsThis: Exit Sub
       Conexion.Open
       Set Resultado = Conexion.Execute("SELECT Permiso,Nombre,Tipo,Documento FROM Alumnos WHERE Permiso = " & txtPermiso & " AND Eliminado = False")
       If Resultado.EOF = False Then
            lblAlumno = Resultado!Nombre: lblTipoDocumento = Resultado!Tipo: lblDocumento = Resultado!documento
            Set Resultado = Conexion.Execute("SELECT Carrera FROM CarrerasHechas WHERE Permiso =" & txtPermiso & " AND Carrera = " & frmEquivalencias.dtcCarreras.BoundText)
            If Resultado.EOF = True Then
                MsgBox ("El alumno mo corresponde a esta carrera"): Conexion.Close: txtPermiso = "": txtPermiso.SetFocus: Exit Sub
            End If
        Else '
            MsgBox ("El alumno no existe"): Conexion.Close: txtPermiso = "": txtPermiso.SetFocus: Exit Sub
        End If
        Conexion.Close
   End If
End Sub
