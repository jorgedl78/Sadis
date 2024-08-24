VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmInstituciones 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "ABM de Instituciones"
   ClientHeight    =   7305
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7965
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   7305
   ScaleWidth      =   7965
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtBuscador 
      Height          =   285
      Left            =   840
      TabIndex        =   27
      Top             =   4200
      Width           =   1695
   End
   Begin MSAdodcLib.Adodc adoInstituciones 
      Height          =   330
      Left            =   840
      Top             =   5760
      Visible         =   0   'False
      Width           =   2175
      _ExtentX        =   3836
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
      RecordSource    =   "SELECT * FROM Instituciones  where codigo > 0 ORDER BY Institucion"
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
   Begin VB.Frame frComandos 
      Height          =   4335
      Left            =   7080
      TabIndex        =   2
      Top             =   0
      Width           =   855
      Begin VB.CommandButton cmdSalir 
         Height          =   600
         Left            =   120
         Picture         =   "frmInstituciones.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Salir"
         Top             =   3600
         Width           =   600
      End
      Begin VB.CommandButton cmdAgregar 
         Height          =   600
         Left            =   120
         Picture         =   "frmInstituciones.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Agregar"
         Top             =   240
         Width           =   600
      End
      Begin VB.CommandButton cmdModificar 
         Enabled         =   0   'False
         Height          =   600
         Left            =   120
         Picture         =   "frmInstituciones.frx":0884
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Modificar"
         Top             =   1080
         Width           =   600
      End
      Begin VB.CommandButton cmdGuardar 
         Enabled         =   0   'False
         Height          =   600
         Left            =   120
         Picture         =   "frmInstituciones.frx":0CC6
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Guardar"
         Top             =   1920
         Width           =   600
      End
      Begin VB.CommandButton cmdCancelar 
         Enabled         =   0   'False
         Height          =   600
         Left            =   120
         Picture         =   "frmInstituciones.frx":1108
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Cancelar"
         Top             =   2760
         Width           =   600
      End
   End
   Begin VB.Frame Frame2 
      Height          =   2655
      Left            =   120
      TabIndex        =   1
      Top             =   4440
      Width           =   7815
      Begin MSDataGridLib.DataGrid dtgInstituciones 
         Bindings        =   "frmInstituciones.frx":154A
         Height          =   2175
         Left            =   240
         TabIndex        =   26
         Top             =   360
         Width           =   7455
         _ExtentX        =   13150
         _ExtentY        =   3836
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
         ColumnCount     =   9
         BeginProperty Column00 
            DataField       =   "Institucion"
            Caption         =   "Institucion"
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
            DataField       =   "Direccion"
            Caption         =   "Direccion"
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
         BeginProperty Column03 
            DataField       =   "Provincia"
            Caption         =   "Provincia"
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
            DataField       =   "TelefonoEnRed"
            Caption         =   "TelefonoEnRed"
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
            DataField       =   "Telefono Local"
            Caption         =   "Telefono Local"
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
            DataField       =   "EMail"
            Caption         =   "EMail"
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
            DataField       =   "Http"
            Caption         =   "Http"
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
            DataField       =   "Director"
            Caption         =   "Director"
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
            EndProperty
            BeginProperty Column01 
            EndProperty
            BeginProperty Column02 
            EndProperty
            BeginProperty Column03 
            EndProperty
            BeginProperty Column04 
            EndProperty
            BeginProperty Column05 
            EndProperty
            BeginProperty Column06 
            EndProperty
            BeginProperty Column07 
            EndProperty
            BeginProperty Column08 
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Datos"
      Height          =   3975
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   6855
      Begin VB.TextBox txtHttp 
         Enabled         =   0   'False
         Height          =   285
         Left            =   3360
         TabIndex        =   25
         Top             =   3360
         Width           =   3255
      End
      Begin VB.TextBox txtEMail 
         Enabled         =   0   'False
         Height          =   285
         Left            =   240
         TabIndex        =   23
         Top             =   3360
         Width           =   3015
      End
      Begin VB.TextBox txtTelefonoLocal 
         Enabled         =   0   'False
         Height          =   285
         Left            =   3240
         TabIndex        =   21
         Top             =   2640
         Width           =   3375
      End
      Begin VB.TextBox txtTelefonoEnRed 
         Enabled         =   0   'False
         Height          =   285
         Left            =   240
         TabIndex        =   19
         Top             =   2640
         Width           =   2895
      End
      Begin VB.TextBox txtDirector 
         Enabled         =   0   'False
         Height          =   285
         Left            =   3120
         TabIndex        =   17
         Top             =   1920
         Width           =   3495
      End
      Begin VB.TextBox txtProvincia 
         Enabled         =   0   'False
         Height          =   285
         Left            =   240
         TabIndex        =   15
         Top             =   1920
         Width           =   2775
      End
      Begin VB.TextBox txtLocalidad 
         Enabled         =   0   'False
         Height          =   285
         Left            =   3720
         TabIndex        =   13
         Top             =   1200
         Width           =   2895
      End
      Begin VB.TextBox txtDireccion 
         Enabled         =   0   'False
         Height          =   285
         Left            =   240
         TabIndex        =   11
         Top             =   1200
         Width           =   3375
      End
      Begin VB.TextBox txtInstitucion 
         Enabled         =   0   'False
         Height          =   285
         Left            =   240
         TabIndex        =   9
         Top             =   600
         Width           =   6375
      End
      Begin VB.Label Label8 
         Caption         =   "Página en Internet"
         Height          =   255
         Left            =   3360
         TabIndex        =   24
         Top             =   3120
         Width           =   1575
      End
      Begin VB.Label Label7 
         Caption         =   "E-Mail"
         Height          =   255
         Left            =   240
         TabIndex        =   22
         Top             =   3120
         Width           =   615
      End
      Begin VB.Label Label6 
         Caption         =   "Telefono local"
         Height          =   255
         Left            =   3240
         TabIndex        =   20
         Top             =   2400
         Width           =   1215
      End
      Begin VB.Label Label5 
         Caption         =   "Telefono en Red"
         Height          =   255
         Left            =   240
         TabIndex        =   18
         Top             =   2400
         Width           =   1575
      End
      Begin VB.Label Label4 
         Caption         =   "Director"
         Height          =   255
         Left            =   3120
         TabIndex        =   16
         Top             =   1680
         Width           =   735
      End
      Begin VB.Label Provincia 
         Caption         =   "Provincia"
         Height          =   255
         Left            =   240
         TabIndex        =   14
         Top             =   1680
         Width           =   975
      End
      Begin VB.Label Label3 
         Caption         =   "Localidad"
         Height          =   255
         Left            =   3720
         TabIndex        =   12
         Top             =   960
         Width           =   855
      End
      Begin VB.Label Label2 
         Caption         =   "Dirección"
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Top             =   960
         Width           =   1695
      End
      Begin VB.Label Label1 
         Caption         =   "Nombre de la Institución"
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   360
         Width           =   1815
      End
   End
   Begin VB.Label Label9 
      Caption         =   "Buscar"
      Height          =   255
      Left            =   120
      TabIndex        =   28
      Top             =   4200
      Width           =   735
   End
End
Attribute VB_Name = "frmInstituciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Conexion As New Connection
Dim Resultado As New Recordset
Dim Estado As String
Private Sub cmdAgregar_Click()
    Estado = "Agregando"
    cmdAgregar.Enabled = False
    cmdModificar.Enabled = False
    cmdGuardar.Enabled = True
    cmdCancelar.Enabled = True
    cmdSalir.Enabled = False
    LimpiarDatos
    HabilitarDatos
    txtInstitucion.SetFocus
End Sub

Private Sub cmdCancelar_Click()
    cmdAgregar.Enabled = True
    cmdModificar.Enabled = True
    cmdGuardar.Enabled = False
    cmdCancelar.Enabled = False
    cmdSalir.Enabled = True
    DeshabilitarDatos
    MostrarDatos
End Sub

Private Sub cmdGuardar_Click()
    Conexion.Open
    If Estado = "Agregando" Then
        InstitucionNueva = txtInstitucion
        Set Resultado = Conexion.Execute("SELECT MAx(Codigo) as Ultimo FROM Instituciones")
        Ultimo = Resultado!Ultimo + 1
        Conexion.Execute ("INSERT INTO Instituciones ( Codigo, Institucion, Direccion, Localidad, Provincia, TelefonoEnRed, [Telefono Local], EMail, Http, Director ) VALUES  (" & Ultimo & ",'" & txtInstitucion & "','" & txtDireccion & "','" & txtLocalidad & "','" & txtProvincia & "','" & txtTelefonoEnRed & "','" & txtTelefonoLocal & "','" & txtEmail & "','" & txtHttp & "','" & txtDireccion & "')")
    Else
        RegistroActual = adoInstituciones.Recordset!Codigo
        Conexion.Execute ("UPDATE Instituciones SET Instituciones.Institucion = '" & txtInstitucion & "', Instituciones.Direccion = '" & txtDireccion & "', Instituciones.Localidad = '" & txtLocalidad & "', Instituciones.Provincia = '" & txtProvincia & "', Instituciones.TelefonoEnRed = '" & txtTelefonoEnRed & "', Instituciones.[Telefono Local] = '" & txtTelefonoLocal & "', Instituciones.EMail = '" & txtEmail & "', Instituciones.Http = '" & txtHttp & "', Instituciones.Director = '" & txtDirector & "' WHERE Instituciones.Codigo=" & adoInstituciones.Recordset!Codigo)
    End If
    Conexion.Close
    adoInstituciones.Refresh
    If Estado = "Modificando" Then adoInstituciones.Recordset.Find ("Codigo=" & RegistroActual)
    If Estado = Agregando Then adoInstituciones.Recordset.Find ("Institucion=" & InstitucionNueva)
    cmdAgregar.Enabled = True
    cmdModificar.Enabled = True
    cmdGuardar.Enabled = False
    cmdCancelar.Enabled = False
    cmdSalir.Enabled = True
    DeshabilitarDatos
End Sub

Private Sub cmdModificar_Click()
    Estado = "Modificando"
    cmdAgregar.Enabled = False
    cmdModificar.Enabled = False
    cmdGuardar.Enabled = True
    cmdCancelar.Enabled = True
    cmdSalir.Enabled = False
    HabilitarDatos
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Function HabilitarDatos()
    txtDireccion.Enabled = True
    txtDirector.Enabled = True
    txtEmail.Enabled = True
    txtHttp.Enabled = True
    txtLocalidad.Enabled = True
    txtProvincia.Enabled = True
    txtInstitucion.Enabled = True
    txtTelefonoEnRed.Enabled = True
    txtTelefonoLocal.Enabled = True
End Function

Private Function DeshabilitarDatos()
    txtDireccion.Enabled = False
    txtDirector.Enabled = False
    txtEmail.Enabled = False
    txtHttp.Enabled = False
    txtLocalidad.Enabled = False
    txtProvincia.Enabled = False
    txtInstitucion.Enabled = False
    txtTelefonoEnRed.Enabled = False
    txtTelefonoLocal.Enabled = False
End Function

Private Function LimpiarDatos()
    txtDireccion = ""
    txtDirector = ""
    txtEmail = ""
    txtHttp = ""
    txtLocalidad = ""
    txtProvincia = ""
    txtInstitucion = ""
    txtTelefonoEnRed = ""
    txtTelefonoLocal = ""
End Function

Private Sub dtgInstituciones_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    If adoInstituciones.Recordset.EOF = False Then MostrarDatos
End Sub

Private Sub Form_Load()
    Conexion.ConnectionString = ("DSN=Instituto")
    MostrarDatos
    If adoInstituciones.Recordset.RecordCount > 0 Then cmdModificar.Enabled = True
End Sub

Private Function MostrarDatos()
If adoInstituciones.Recordset.RecordCount > 0 Then
    With adoInstituciones.Recordset
    If !Direccion <> vacio Then
        txtDireccion = !Direccion
    Else
        txtDireccion = ""
    End If
    If !Director <> vacio Then
        txtDirector = !Director
    Else
        txtDirector = ""
    End If
    If !email <> vacio Then
        txtEmail = !email
    Else
        txtEmail = ""
    End If
    If !Http <> vacio Then
        txtHttp = !Http
    Else
        txtHttp = ""
    End If
    If !Localidad <> vacio Then
        txtLocalidad = !Localidad
    Else
        txtLocalidad = ""
    End If
    If !Provincia <> vacio Then
        txtProvincia = !Provincia
    Else
        txtProvincia = ""
    End If
    If !Institucion <> vacio Then
        txtInstitucion = !Institucion
    Else
        txtInstitucion = ""
    End If
    If !TelefonoEnRed <> vacio Then
        txtTelefonoEnRed = !TelefonoEnRed
    Else
        txtTelefonoEnRed = ""
    End If
    If ![Telefono Local] <> vacio Then
        txtTelefonoLocal = ![Telefono Local]
    Else
        txtTelefonoLocal = ""
    End If
    End With
End If
End Function

Private Sub txtBuscador_Change()
    adoInstituciones.Recordset.MoveFirst
    adoInstituciones.Recordset.Find ("Institucion >='" & txtBuscador & "'")
    If adoInstituciones.Recordset.EOF = True Then adoInstituciones.Recordset.MoveLast
    MostrarDatos
End Sub

Private Sub txtDireccion_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtLocalidad.SetFocus
End Sub

Private Sub txtDirector_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtTelefonoEnRed.SetFocus
End Sub

Private Sub txtEMail_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtHttp.SetFocus
End Sub

Private Sub txtHttp_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cmdGuardar.SetFocus
End Sub

Private Sub txtInstitucion_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtDireccion.SetFocus
End Sub

Private Sub txtLocalidad_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtProvincia.SetFocus
End Sub

Private Sub txtProvincia_KeyPress(KeyAscii As Integer)
     If KeyAscii = 13 Then txtDirector.SetFocus
End Sub

Private Sub txtTelefonoEnRed_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtTelefonoLocal.SetFocus
End Sub

Private Sub txtTelefonoLocal_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtEmail.SetFocus
End Sub
