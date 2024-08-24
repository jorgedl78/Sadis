VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmPersonal 
   BackColor       =   &H00404040&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Personal"
   ClientHeight    =   8250
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11310
   Icon            =   "Personal.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   8250
   ScaleWidth      =   11310
   Begin VB.CommandButton cmdPass 
      Height          =   320
      Left            =   6840
      Picture         =   "Personal.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   54
      Top             =   3360
      Width           =   440
   End
   Begin VB.TextBox txtContrasena 
      Enabled         =   0   'False
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   4920
      MaxLength       =   25
      MultiLine       =   -1  'True
      PasswordChar    =   "*"
      TabIndex        =   44
      Top             =   3360
      Width           =   1695
   End
   Begin VB.TextBox txtUsuario 
      Enabled         =   0   'False
      Height          =   285
      Left            =   1560
      MaxLength       =   20
      MultiLine       =   -1  'True
      TabIndex        =   42
      Top             =   3360
      Width           =   1695
   End
   Begin VB.CheckBox chkEliminado 
      BackColor       =   &H00404040&
      Enabled         =   0   'False
      Height          =   255
      Left            =   9240
      TabIndex        =   20
      Top             =   1200
      Width           =   255
   End
   Begin MSAdodcLib.Adodc adoPersonal 
      Height          =   330
      Left            =   9840
      Top             =   7200
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   582
      ConnectMode     =   3
      CursorLocation  =   2
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
      Caption         =   "Adodc1"
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
   Begin VB.TextBox txtCodigo 
      Alignment       =   2  'Center
      BackColor       =   &H00B0FFB0&
      Enabled         =   0   'False
      Height          =   285
      Left            =   8880
      TabIndex        =   3
      Top             =   120
      Width           =   735
   End
   Begin VB.TextBox txtLugar 
      Enabled         =   0   'False
      Height          =   285
      Left            =   6000
      TabIndex        =   9
      Top             =   840
      Width           =   1455
   End
   Begin VB.CheckBox chkActivo 
      BackColor       =   &H00404040&
      Enabled         =   0   'False
      Height          =   255
      Left            =   9240
      TabIndex        =   11
      Top             =   840
      Width           =   255
   End
   Begin VB.TextBox txtRegistro 
      Enabled         =   0   'False
      Height          =   285
      Left            =   1560
      MaxLength       =   75
      MultiLine       =   -1  'True
      TabIndex        =   34
      Top             =   2640
      Width           =   1695
   End
   Begin VB.TextBox txtCodigoPostal 
      Enabled         =   0   'False
      Height          =   285
      Left            =   8640
      MaxLength       =   8
      TabIndex        =   26
      Top             =   1560
      Width           =   855
   End
   Begin VB.TextBox txtLocalidad 
      Enabled         =   0   'False
      Height          =   285
      Left            =   5880
      TabIndex        =   24
      Top             =   1560
      Width           =   2175
   End
   Begin VB.TextBox txtSexo 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      Height          =   285
      Left            =   1560
      MaxLength       =   1
      TabIndex        =   5
      Top             =   840
      Width           =   375
   End
   Begin VB.ComboBox txtCalificacion2 
      Enabled         =   0   'False
      Height          =   315
      ItemData        =   "Personal.frx":388B
      Left            =   7920
      List            =   "Personal.frx":38AE
      TabIndex        =   38
      Top             =   2640
      Width           =   1575
   End
   Begin VB.ComboBox txtCalificacion1 
      Enabled         =   0   'False
      Height          =   315
      ItemData        =   "Personal.frx":391A
      Left            =   4920
      List            =   "Personal.frx":393D
      TabIndex        =   36
      Top             =   2640
      Width           =   1575
   End
   Begin VB.TextBox txtTelefono 
      Enabled         =   0   'False
      Height          =   285
      Left            =   1560
      TabIndex        =   28
      Top             =   1920
      Width           =   2535
   End
   Begin VB.TextBox txtEmail 
      Enabled         =   0   'False
      Height          =   285
      Left            =   5040
      TabIndex        =   30
      Top             =   1920
      Width           =   4455
   End
   Begin VB.TextBox txtFechaNac 
      Enabled         =   0   'False
      Height          =   285
      Left            =   3720
      TabIndex        =   7
      Top             =   840
      Width           =   1455
   End
   Begin VB.TextBox txtDomicilio 
      Enabled         =   0   'False
      Height          =   285
      Left            =   1560
      TabIndex        =   22
      Top             =   1560
      Width           =   3135
   End
   Begin VB.TextBox txtObservaciones 
      Enabled         =   0   'False
      Height          =   285
      Left            =   1560
      MultiLine       =   -1  'True
      TabIndex        =   40
      Top             =   3000
      Width           =   7935
   End
   Begin VB.ComboBox txtNombre 
      DataField       =   "nombre"
      DataSource      =   "Personal"
      Height          =   315
      ItemData        =   "Personal.frx":39A9
      Left            =   1680
      List            =   "Personal.frx":39AB
      TabIndex        =   1
      Top             =   90
      Width           =   6255
   End
   Begin VB.CommandButton cmdGuardar 
      BackColor       =   &H00FFE0C0&
      Caption         =   "Guardar"
      Enabled         =   0   'False
      Height          =   375
      Left            =   9960
      Style           =   1  'Graphical
      TabIndex        =   48
      Top             =   2460
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancelar 
      BackColor       =   &H00FFE0C0&
      Caption         =   "Cancelar"
      Enabled         =   0   'False
      Height          =   375
      Left            =   9960
      Style           =   1  'Graphical
      TabIndex        =   49
      Top             =   3060
      Width           =   1215
   End
   Begin VB.Frame frmBuscar 
      BackColor       =   &H00404040&
      Caption         =   " Buscar "
      ForeColor       =   &H00FFFFFF&
      Height          =   660
      Left            =   120
      TabIndex        =   50
      Top             =   3840
      Width           =   11055
      Begin VB.TextBox txtBuscarNombre 
         Height          =   285
         Left            =   2400
         MaxLength       =   40
         TabIndex        =   52
         Top             =   240
         Width           =   8535
      End
      Begin VB.Label Label15 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Apellido y Nombres:"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   840
         TabIndex        =   51
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.CommandButton cmdEliminar 
      BackColor       =   &H00FFE0C0&
      Caption         =   "Eliminar"
      Height          =   375
      Left            =   9960
      Style           =   1  'Graphical
      TabIndex        =   47
      Top             =   1500
      Width           =   1215
   End
   Begin VB.CommandButton cmdModificar 
      BackColor       =   &H00FFE0C0&
      Caption         =   "Modificar"
      Height          =   375
      Left            =   9960
      Style           =   1  'Graphical
      TabIndex        =   46
      Top             =   900
      Width           =   1215
   End
   Begin VB.CommandButton cmdNuevo 
      BackColor       =   &H00FFE0C0&
      Caption         =   "Nuevo"
      Height          =   375
      Left            =   9960
      Style           =   1  'Graphical
      TabIndex        =   45
      Top             =   120
      Width           =   1215
   End
   Begin VB.TextBox txtTitulo 
      Enabled         =   0   'False
      Height          =   285
      Left            =   1560
      MaxLength       =   75
      MultiLine       =   -1  'True
      TabIndex        =   32
      Top             =   2280
      Width           =   7935
   End
   Begin VB.TextBox txtFoja 
      Enabled         =   0   'False
      Height          =   285
      Left            =   6000
      MaxLength       =   8
      TabIndex        =   18
      Top             =   1200
      Width           =   1455
   End
   Begin VB.ComboBox cmbDocumentoTipo 
      Enabled         =   0   'False
      Height          =   315
      ItemData        =   "Personal.frx":39AD
      Left            =   2280
      List            =   "Personal.frx":39BD
      TabIndex        =   14
      Text            =   "CUIL"
      Top             =   1170
      Width           =   855
   End
   Begin VB.TextBox txtDocumentoNumero 
      Enabled         =   0   'False
      Height          =   285
      Left            =   3720
      MaxLength       =   13
      TabIndex        =   16
      Top             =   1200
      Width           =   1455
   End
   Begin MSDataGridLib.DataGrid dbgPersonal 
      Bindings        =   "Personal.frx":39D4
      Height          =   3495
      Left            =   120
      TabIndex        =   53
      Top             =   4680
      Width           =   11055
      _ExtentX        =   19500
      _ExtentY        =   6165
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
      Caption         =   "Personal"
      ColumnCount     =   20
      BeginProperty Column00 
         DataField       =   "Codigo"
         Caption         =   "Código"
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
      BeginProperty Column01 
         DataField       =   "Nombre"
         Caption         =   "Apellido y Nombres"
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
         DataField       =   "Sexo"
         Caption         =   "Sexo"
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
         DataField       =   "Foja"
         Caption         =   "Nº Foja"
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
      BeginProperty Column06 
         DataField       =   "FechaNacimiento"
         Caption         =   "Fecha Nac."
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
         DataField       =   "Lugar"
         Caption         =   "Lugar Nac."
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
      BeginProperty Column09 
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
      BeginProperty Column10 
         DataField       =   "Postal"
         Caption         =   "C. Postal"
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
         DataField       =   "Telefono"
         Caption         =   "Teléfono"
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
      BeginProperty Column12 
         DataField       =   "Email"
         Caption         =   "E-mail"
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
      BeginProperty Column13 
         DataField       =   "Titulos"
         Caption         =   "Títulos"
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
      BeginProperty Column14 
         DataField       =   "Registro"
         Caption         =   "Nº Registro"
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
      BeginProperty Column15 
         DataField       =   "Calificacion1"
         Caption         =   "Calificación 1"
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
      BeginProperty Column16 
         DataField       =   "Calificacion2"
         Caption         =   "Calificación 2"
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
      BeginProperty Column17 
         DataField       =   "Observaciones"
         Caption         =   "Observaciones"
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
      BeginProperty Column18 
         DataField       =   "TrabajaActualmente"
         Caption         =   "Activo"
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
      BeginProperty Column19 
         DataField       =   "Eliminado"
         Caption         =   "Eliminado"
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
            Locked          =   -1  'True
            ColumnWidth     =   585,071
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1904,882
         EndProperty
         BeginProperty Column02 
            Alignment       =   2
            ColumnWidth     =   450,142
         EndProperty
         BeginProperty Column03 
            Alignment       =   2
            ColumnWidth     =   494,929
         EndProperty
         BeginProperty Column04 
            Alignment       =   2
            ColumnWidth     =   1170,142
         EndProperty
         BeginProperty Column05 
            Alignment       =   2
            ColumnWidth     =   720
         EndProperty
         BeginProperty Column06 
            Alignment       =   2
            ColumnWidth     =   900,284
         EndProperty
         BeginProperty Column07 
            ColumnWidth     =   1260,284
         EndProperty
         BeginProperty Column08 
            ColumnWidth     =   1230,236
         EndProperty
         BeginProperty Column09 
            ColumnWidth     =   1019,906
         EndProperty
         BeginProperty Column10 
            Alignment       =   2
            ColumnWidth     =   870,236
         EndProperty
         BeginProperty Column11 
            ColumnWidth     =   1154,835
         EndProperty
         BeginProperty Column12 
            ColumnWidth     =   1739,906
         EndProperty
         BeginProperty Column13 
            ColumnWidth     =   1739,906
         EndProperty
         BeginProperty Column14 
            Alignment       =   2
            ColumnWidth     =   915,024
         EndProperty
         BeginProperty Column15 
            Alignment       =   2
            ColumnWidth     =   1019,906
         EndProperty
         BeginProperty Column16 
            Alignment       =   2
            ColumnWidth     =   1049,953
         EndProperty
         BeginProperty Column17 
            ColumnWidth     =   1739,906
         EndProperty
         BeginProperty Column18 
            Alignment       =   2
            ColumnWidth     =   540,284
         EndProperty
         BeginProperty Column19 
            Alignment       =   2
            ColumnWidth     =   764,787
         EndProperty
      EndProperty
   End
   Begin VB.Image imgOcultar 
      Enabled         =   0   'False
      Height          =   255
      Left            =   9120
      Picture         =   "Personal.frx":39EE
      Top             =   3360
      Visible         =   0   'False
      Width           =   330
   End
   Begin VB.Image imgMostrar 
      Enabled         =   0   'False
      Height          =   255
      Left            =   8640
      Picture         =   "Personal.frx":6D1A
      Top             =   3360
      Visible         =   0   'False
      Width           =   330
   End
   Begin VB.Label Label22 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Contraseña:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   3480
      TabIndex        =   43
      Top             =   3360
      Width           =   1335
   End
   Begin VB.Label Label21 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Usuario:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   41
      Top             =   3360
      Width           =   1335
   End
   Begin VB.Label lblEliminado 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Eliminado:"
      ForeColor       =   &H00FFFFFF&
      Height          =   252
      Left            =   8280
      TabIndex        =   19
      Top             =   1200
      Width           =   852
   End
   Begin VB.Label Label16 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Código:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   8160
      TabIndex        =   2
      Top             =   120
      Width           =   615
   End
   Begin VB.Label Label10 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Lugar:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   5400
      TabIndex        =   8
      Top             =   840
      Width           =   495
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Activo:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   8520
      TabIndex        =   10
      Top             =   840
      Width           =   615
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Nº de Registro:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   33
      Top             =   2640
      Width           =   1335
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "CP:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   8280
      TabIndex        =   25
      Top             =   1560
      Width           =   255
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Localidad:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   4920
      TabIndex        =   23
      Top             =   1560
      Width           =   855
   End
   Begin VB.Line Line6 
      BorderColor     =   &H00FFFFFF&
      X1              =   0
      X2              =   11280
      Y1              =   600
      Y2              =   600
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Sexo (M/F):"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   840
      Width           =   1335
   End
   Begin VB.Line Line5 
      BorderColor     =   &H00FFFFFF&
      X1              =   11280
      X2              =   9720
      Y1              =   2160
      Y2              =   2160
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Calificación 2:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   6720
      TabIndex        =   37
      Top             =   2640
      Width           =   1095
   End
   Begin VB.Label Label17 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Calificación 1:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   3720
      TabIndex        =   35
      Top             =   2640
      Width           =   1095
   End
   Begin VB.Label Label20 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Teléfono/s:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   27
      Top             =   1920
      Width           =   1335
   End
   Begin VB.Label Label19 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "e-Mail:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   4080
      TabIndex        =   29
      Top             =   1920
      Width           =   855
   End
   Begin VB.Label lblFechaNac 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Fecha Nacimiento:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2160
      TabIndex        =   6
      Top             =   840
      Width           =   1455
   End
   Begin VB.Label Label12 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Domicilio:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   21
      Top             =   1560
      Width           =   1335
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00FFFFFF&
      X1              =   9720
      X2              =   9720
      Y1              =   3720
      Y2              =   0
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      X1              =   0
      X2              =   11280
      Y1              =   3720
      Y2              =   3720
   End
   Begin VB.Label Label18 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Observaciones:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   39
      Top             =   3000
      Width           =   1335
   End
   Begin VB.Label Label14 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Tipo:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1800
      TabIndex        =   13
      Top             =   1200
      Width           =   375
   End
   Begin VB.Label Label13 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Nº:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   3360
      TabIndex        =   15
      Top             =   1200
      Width           =   255
   End
   Begin VB.Label Label11 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Título/s:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   31
      Top             =   2280
      Width           =   1335
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Foja:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   5520
      TabIndex        =   17
      Top             =   1200
      Width           =   375
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Documento:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   480
      TabIndex        =   12
      Top             =   1200
      Width           =   975
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Apellido y Nombres:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1455
   End
End
Attribute VB_Name = "frmPersonal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rstPersonal As Recordset
Dim rstMovimientos As Recordset
Dim rstConceptos As Recordset
Dim Conexion As New Connection

Private Sub cmdPass_Click()
    If txtContrasena.PasswordChar = "*" Then
        txtContrasena.PasswordChar = ""
        cmdPass.Picture = imgOcultar.Picture
    Else
        txtContrasena.PasswordChar = "*"
        cmdPass.Picture = imgMostrar.Picture
    End If
End Sub

Private Sub dbgPersonal_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    If rstPersonal.EditMode <> adEditNone Then Exit Sub
    txtNombre.Text = adoPersonal.Recordset!Nombre: AutoLoad
End Sub

Private Sub Form_Load()
    Conexion.ConnectionString = ("DSN=Instituto")
    Conexion.Mode = adModeReadWrite
    Conexion.Open
    Set rstPersonal = New ADODB.Recordset
    rstPersonal.Open "SELECT Codigo, Nombre FROM Personal ORDER BY Nombre", Conexion, adOpenKeyset, adLockOptimistic, adCmdText
    
    'Set rstConceptos = New ADODB.Recordset
    'rstConceptos.Open "SELECT * FROM Conceptos ORDER BY nombre", Conexion, adOpenKeyset, adLockOptimistic, adCmdText
    '
    'While rstPersonal.EOF = False
    '    If InStr(1, rstPersonal!Nombre, ",") = 0 Then enPersonal = rstPersonal!Nombre Else enPersonal = Left(rstPersonal!Nombre, InStr(1, rstPersonal!Nombre, ",") - 1)
    '    While rstConceptos.EOF = False
    '        If Left(rstConceptos!Nombre, Len(enPersonal)) = enPersonal Then
    '            If rstConceptos!doctipo <> "" Then rstPersonal!Tipo = rstConceptos!doctipo
    '            If rstConceptos!docnumero <> "" Then rstPersonal!Documento = rstConceptos!docnumero
    '            If rstConceptos!foja <> "" Then rstPersonal!foja = rstConceptos!foja
    '            If rstConceptos!fechanac <> "" Then rstPersonal!FechaNacimiento = rstConceptos!fechanac
    '            If rstConceptos!domicilio <> "" Then rstPersonal!domicilio = rstConceptos!domicilio
    '            If rstConceptos!telefono <> "" Then rstPersonal!telefono = rstConceptos!telefono
    '            If rstConceptos!email <> "" Then rstPersonal!email = rstConceptos!email
    '            If rstConceptos!titreg <> "" Then rstPersonal!Titulos = rstConceptos!titreg
    '            If rstConceptos!concepto1 <> "" Then rstPersonal!Calificacion1 = rstConceptos!concepto1
    '            If rstConceptos!concepto2 <> "" Then rstPersonal!Calificacion2 = rstConceptos!concepto2
    '            If rstConceptos!Observaciones <> "" Then rstPersonal!Observaciones = rstConceptos!Observaciones
    '            If rstConceptos!Usuario <> "" Then rstPersonal!Usuario = rstConceptos!Usuario
    '            If rstConceptos!Contrasena <> "" Then rstPersonal!Contrasena = rstConceptos!Contrasena
    '        End If
    '        rstConceptos.MoveNext
    '    Wend
    '    rstConceptos.MoveFirst
    '    rstPersonal.MoveNext
    'Wend
    cmbDocumentoTipo.Clear
    cmbDocumentoTipo.AddItem "CUIL", 0
    cmbDocumentoTipo.AddItem "DNI", 1
    cmbDocumentoTipo.AddItem "LC", 2
    cmbDocumentoTipo.AddItem "LE", 3
End Sub

Private Sub Form_Activate()
    txtNombre.Clear
    rstPersonal.Close
    rstPersonal.Open "SELECT Nombre FROM Personal ORDER BY Nombre", Conexion, adOpenKeyset, adLockOptimistic, adCmdText
    rstPersonal.Requery
    While rstPersonal.EOF = False
        If opersonal <> rstPersonal!Nombre Then txtNombre.AddItem rstPersonal!Nombre
        opersonal = rstPersonal!Nombre
        rstPersonal.MoveNext
    Wend
    rstPersonal.MoveFirst
    txtNombre.ListIndex = 0
End Sub

Private Sub cmdNuevo_Click()
    rstPersonal.Close
    rstPersonal.Open "SELECT * FROM Personal ORDER BY Codigo", Conexion, adOpenKeyset, adLockOptimistic, adCmdText
    NuevoCodigo = 0
    If rstPersonal.RecordCount > 0 Then
        For i = 0 To rstPersonal.RecordCount - 1
            If rstPersonal!Codigo > i Then
                NuevoCodigo = i
                Exit For
            End If
            rstPersonal.MoveNext
        Next
        If rstPersonal.EOF Then NuevoCodigo = i
    End If
    rstPersonal.AddNew
    txtCodigo.Text = NuevoCodigo
    cmdNuevo.Enabled = False
    cmdModificar.Enabled = False
    cmdEliminar.Enabled = False
    cmdGuardar.Enabled = True
    cmdCancelar.Enabled = True
    frmBuscar.Enabled = False
    cmbDocumentoTipo.Enabled = True
    txtDocumentoNumero.Enabled = True
    txtFoja.Enabled = True
    txtSexo.Enabled = True
    txtFechaNac.Enabled = True
    txtLugar.Enabled = True
    txtDomicilio.Enabled = True
    txtLocalidad.Enabled = True
    txtCodigoPostal.Enabled = True
    txtTelefono.Enabled = True
    txtEmail.Enabled = True
    txtTitulo.Enabled = True
    txtRegistro.Enabled = True
    txtCalificacion1.Enabled = True
    txtCalificacion2.Enabled = True
    txtObservaciones.Enabled = True
    txtUsuario.Enabled = True
    txtContrasena.Enabled = True
    chkActivo.Enabled = True
    chkEliminado.Enabled = True
End Sub

Private Sub cmdModificar_Click()
    cmdNuevo.Enabled = False
    cmdModificar.Enabled = False
    cmdEliminar.Enabled = False
    cmdGuardar.Enabled = True
    cmdCancelar.Enabled = True
    frmBuscar.Enabled = False
    cmbDocumentoTipo.Enabled = True
    txtDocumentoNumero.Enabled = True
    txtFoja.Enabled = True
    txtSexo.Enabled = True
    txtFechaNac.Enabled = True
    txtLugar.Enabled = True
    txtDomicilio.Enabled = True
    txtLocalidad.Enabled = True
    txtCodigoPostal.Enabled = True
    txtTelefono.Enabled = True
    txtEmail.Enabled = True
    txtTitulo.Enabled = True
    txtRegistro.Enabled = True
    txtCalificacion1.Enabled = True
    txtCalificacion2.Enabled = True
    txtObservaciones.Enabled = True
    txtUsuario.Enabled = True
    txtContrasena.Enabled = True
    chkActivo.Enabled = True
    chkEliminado.Enabled = True
End Sub

Private Sub cmdGuardar_Click()
    On Error Resume Next
    If Len(txtNombre.Text) = 0 Then MsgBox "Debe especificar el Nombre.", , "Personal": Exit Sub
    'If Len(cmbDocumentoTipo) = 0 Then cmbDocumentoTipo.Text = "CUIL"
    'If Val(txtDocumentoNumero) = 0 Then txtDocumentoNumero.Text = "0"
    'If Val(txtFoja.Text) = 0 Then txtFoja.Text = "0"
    If rstPersonal.RecordCount = 0 Then
        rstPersonal.AddNew
        rstPersonal!Codigo = txtCodigo.Text
        rstPersonal!Nombre = txtNombre.Text
        rstPersonal!Sexo = txtSexo.Text
        rstPersonal!Tipo = cmbDocumentoTipo.Text
        rstPersonal!Documento = txtDocumentoNumero.Text
        rstPersonal!foja = txtFoja.Text
        rstPersonal!FechaNacimiento = txtFechaNac.Text
        rstPersonal!domicilio = txtDomicilio.Text
        rstPersonal!Lugar = txtLugar.Text
        rstPersonal!Localidad = txtLocalidad.Text
        rstPersonal!Postal = txtCodigoPostal.Text
        rstPersonal!telefono = txtTelefono.Text
        rstPersonal!email = txtEmail.Text
        rstPersonal!Titulos = txtTitulo.Text
        rstPersonal!Registro = txtRegistro.Text
        rstPersonal!Calificacion1 = txtCalificacion1.Text
        rstPersonal!Calificacion2 = txtCalificacion2.Text
        rstPersonal!Observaciones = txtObservaciones.Text
        rstPersonal!Usuario = txtUsuario.Text
        rstPersonal!Contrasena = txtContrasena.Text
        rstPersonal!TrabajaActualmente = chkActivo.Value
        rstPersonal!Eliminado = chkEliminado.Value
    Else
        rstPersonal!Codigo = txtCodigo.Text
        rstPersonal!Nombre = txtNombre.Text
        rstPersonal!Sexo = txtSexo.Text
        rstPersonal!Tipo = cmbDocumentoTipo.Text
        rstPersonal!Documento = txtDocumentoNumero.Text
        rstPersonal!foja = txtFoja.Text
        rstPersonal!FechaNacimiento = txtFechaNac.Text
        rstPersonal!domicilio = txtDomicilio.Text
        rstPersonal!Lugar = txtLugar.Text
        rstPersonal!Localidad = txtLocalidad.Text
        rstPersonal!Postal = txtCodigoPostal.Text
        rstPersonal!telefono = txtTelefono.Text
        rstPersonal!email = txtEmail.Text
        rstPersonal!Titulos = txtTitulo.Text
        rstPersonal!Registro = txtRegistro.Text
        rstPersonal!Calificacion1 = txtCalificacion1.Text
        rstPersonal!Calificacion2 = txtCalificacion2.Text
        rstPersonal!Observaciones = txtObservaciones.Text
        rstPersonal!Usuario = txtUsuario.Text
        rstPersonal!Contrasena = txtContrasena.Text
        rstPersonal!TrabajaActualmente = chkActivo.Value
        rstPersonal!Eliminado = chkEliminado.Value
    End If
    rstPersonal.Update
    rstPersonal.Requery
    
    adoPersonal.Recordset.Requery
    
    cmdNuevo.Enabled = True
    cmdModificar.Enabled = True
    cmdEliminar.Enabled = True
    cmdGuardar.Enabled = False
    cmdCancelar.Enabled = False
    frmBuscar.Enabled = True
    cmbDocumentoTipo.Enabled = False
    txtDocumentoNumero.Enabled = False
    txtFoja.Enabled = False
    txtSexo.Enabled = False
    txtFechaNac.Enabled = False
    txtLugar.Enabled = False
    txtDomicilio.Enabled = False
    txtLocalidad.Enabled = False
    txtCodigoPostal.Enabled = False
    txtTelefono.Enabled = False
    txtEmail.Enabled = False
    txtTitulo.Enabled = False
    txtRegistro.Enabled = False
    txtCalificacion1.Enabled = False
    txtCalificacion2.Enabled = False
    txtObservaciones.Enabled = False
    txtUsuario.Enabled = False
    txtContrasena.Enabled = False
    chkActivo.Enabled = False
    chkEliminado.Enabled = True
End Sub

Private Sub cmdCancelar_Click()
    rstPersonal.CancelUpdate
    cmdNuevo.Enabled = True
    cmdModificar.Enabled = True
    cmdEliminar.Enabled = True
    cmdGuardar.Enabled = False
    cmdCancelar.Enabled = False
    frmBuscar.Enabled = True
    cmbDocumentoTipo.Enabled = False
    txtDocumentoNumero.Enabled = False
    txtFoja.Enabled = False
    txtSexo.Enabled = False
    txtFechaNac.Enabled = False
    txtLugar.Enabled = False
    txtDomicilio.Enabled = False
    txtLocalidad.Enabled = False
    txtCodigoPostal.Enabled = False
    txtTelefono.Enabled = False
    txtEmail.Enabled = False
    txtTitulo.Enabled = False
    txtRegistro.Enabled = False
    txtCalificacion1.Enabled = False
    txtCalificacion2.Enabled = False
    txtObservaciones.Enabled = False
    txtUsuario.Enabled = False
    txtContrasena.Enabled = False
    chkActivo.Enabled = False
    chkEliminado.Enabled = False
End Sub

Private Sub cmdEliminar_Click()
    If rstPersonal.RecordCount > 0 Then
        rstPersonal!Eliminado = True
        rstPersonal.Update
        rstPersonal.Requery
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Conexion.Close
End Sub

Private Sub txtBuscarNombre_Change()
    If txtBuscarNombre.Text = "" Then Exit Sub
    buscar = "Nombre LIKE '" & txtBuscarNombre.Text & "*'"
    rstPersonal.MoveFirst
    rstPersonal.Find buscar
    If rstPersonal.EOF = False Then
        txtNombre.Text = rstPersonal!Nombre
        AutoLoad
    End If
End Sub

Private Sub txtNombre_Click()
    AutoLoad
End Sub

Private Sub AutoLoad()
    rstPersonal.Close
    rstPersonal.Open "SELECT * FROM Personal ORDER BY Nombre", Conexion, adOpenKeyset, adLockOptimistic, adCmdText
    txtSexo.Text = ""
    txtFechaNac.Text = ""
    txtFoja.Text = ""
    txtLugar.Text = ""
    cmbDocumentoTipo.Text = ""
    txtDocumentoNumero.Text = ""
    txtDomicilio.Text = ""
    txtLocalidad.Text = ""
    txtCodigoPostal.Text = ""
    txtTelefono.Text = ""
    txtEmail.Text = ""
    txtTitulo.Text = ""
    txtRegistro.Text = ""
    txtCalificacion1.Text = ""
    txtCalificacion2.Text = ""
    txtObservaciones.Text = ""
    txtUsuario.Text = ""
    txtContrasena.Text = ""
    rstPersonal.MoveLast
    rstPersonal.MoveFirst
    rstPersonal.Find "Nombre = '" & txtNombre.Text & "'"
    If rstPersonal.EOF = False Then
        adoPersonal.Recordset.Find "Nombre = '" & txtNombre.Text & "'", , , 1
        If rstPersonal!Codigo <> "" Then txtCodigo.Text = rstPersonal!Codigo
        If rstPersonal!Sexo <> "" Then txtSexo.Text = rstPersonal!Sexo
        If rstPersonal!Tipo <> "" Then cmbDocumentoTipo.Text = rstPersonal!Tipo
        If rstPersonal!Documento <> "" Then txtDocumentoNumero.Text = rstPersonal!Documento
        If rstPersonal!foja <> "" Then txtFoja.Text = rstPersonal!foja
        If rstPersonal!FechaNacimiento <> "" Then txtFechaNac.Text = rstPersonal!FechaNacimiento
        If rstPersonal!domicilio <> "" Then txtDomicilio.Text = rstPersonal!domicilio
        If rstPersonal!Lugar <> "" Then txtLugar.Text = rstPersonal!Lugar
        If rstPersonal!Localidad <> "" Then txtLocalidad.Text = rstPersonal!Localidad
        If rstPersonal!Postal <> "" Then txtCodigoPostal.Text = rstPersonal!Postal
        If rstPersonal!telefono <> "" Then txtTelefono.Text = rstPersonal!telefono
        If rstPersonal!email <> "" Then txtEmail.Text = rstPersonal!email
        If rstPersonal!Titulos <> "" Then txtTitulo.Text = rstPersonal!Titulos
        If rstPersonal!Registro <> "" Then txtRegistro.Text = rstPersonal!Registro
        If rstPersonal!Calificacion1 <> "" Then txtCalificacion1.Text = rstPersonal!Calificacion1
        If rstPersonal!Calificacion2 <> "" Then txtCalificacion2.Text = rstPersonal!Calificacion2
        If rstPersonal!Observaciones <> "" Then txtObservaciones.Text = rstPersonal!Observaciones
        If rstPersonal!Usuario <> "" Then txtUsuario.Text = rstPersonal!Usuario
        If rstPersonal!Contrasena <> "" Then txtContrasena.Text = rstPersonal!Contrasena
        If rstPersonal!TrabajaActualmente = True Then chkActivo.Value = 1 Else chkActivo.Value = 0
        If rstPersonal!Eliminado = True Then chkEliminado.Value = 1 Else chkEliminado.Value = 0
    End If
End Sub
