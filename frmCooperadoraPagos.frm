VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "msadodc.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Begin VB.Form frmCooperadoraPagos 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Pagos de Cooperadora"
   ClientHeight    =   6855
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4230
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   6855
   ScaleWidth      =   4230
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame frInformes 
      Caption         =   "Informes"
      Height          =   1455
      Left            =   120
      TabIndex        =   12
      Top             =   5280
      Width           =   3975
      Begin Crystal.CrystalReport rptDetalleDePagos 
         Left            =   1080
         Top             =   480
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   262150
         ReportFileName  =   "listpago.rpt"
         WindowTitle     =   "Detalle de Aportes de Cooperadora"
         WindowControlBox=   -1  'True
         WindowMaxButton =   -1  'True
         WindowMinButton =   -1  'True
      End
      Begin VB.CommandButton cmdDetalleDePagos 
         Caption         =   "Detalle de Pagos por Carrera"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   120
         TabIndex        =   13
         Top             =   240
         Width           =   1455
      End
      Begin Crystal.CrystalReport RptDetalleDePagosPorApellido 
         Left            =   3000
         Top             =   360
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   262150
         ReportFileName  =   "listape.rpt"
         WindowTitle     =   "Listado de Alumnos para Cooperadora"
         WindowControlBox=   -1  'True
         WindowMaxButton =   -1  'True
         WindowMinButton =   -1  'True
      End
      Begin VB.CommandButton cmdDetalleDePagoPorApellido 
         Caption         =   "Detalle de Pagos por Apellido"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   2280
         TabIndex        =   14
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.Frame frDetalleDePagos 
      Caption         =   "Detalle de Pagos"
      Height          =   4095
      Left            =   120
      TabIndex        =   6
      Top             =   1200
      Width           =   3975
      Begin Crystal.CrystalReport CrystalReport1 
         Left            =   2400
         Top             =   360
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   262150
         ReportFileName  =   "C:\Proyecto\Instituto\listape.rpt"
      End
      Begin MSAdodcLib.Adodc adoConceptos 
         Height          =   375
         Left            =   600
         Top             =   3000
         Visible         =   0   'False
         Width           =   2055
         _ExtentX        =   3625
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
         RecordSource    =   $"frmCooperadoraPagos.frx":0000
         Caption         =   "Conceptos"
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
      Begin MSDataGridLib.DataGrid dtgConceptos 
         Bindings        =   "frmCooperadoraPagos.frx":0129
         Height          =   2415
         Left            =   360
         TabIndex        =   11
         Top             =   1560
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   4260
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
            DataField       =   "Concepto"
            Caption         =   "Concepto"
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
            DataField       =   "Importe"
            Caption         =   "Importe"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "$ 0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   11274
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column02 
            DataField       =   "Cancelado"
            Caption         =   "Cancelado"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   5
               Format          =   ""
               HaveTrueFalseNull=   1
               TrueValue       =   "Si"
               FalseValue      =   ""
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
            EndProperty
            BeginProperty Column01 
               Alignment       =   1
            EndProperty
            BeginProperty Column02 
               Alignment       =   2
            EndProperty
         EndProperty
      End
      Begin VB.TextBox txtPermiso 
         Height          =   315
         Left            =   120
         MaxLength       =   5
         TabIndex        =   1
         Top             =   480
         Width           =   615
      End
      Begin VB.CommandButton cmdBuscarAlumno 
         Caption         =   "Buscar..."
         Height          =   315
         Left            =   960
         TabIndex        =   8
         Top             =   480
         Width           =   855
      End
      Begin VB.Label lblDocumento 
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
         Left            =   1920
         TabIndex        =   10
         Top             =   480
         Width           =   1935
      End
      Begin VB.Label lblNombre 
         Caption         =   "Nombre:"
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
         Left            =   120
         TabIndex        =   9
         Top             =   960
         Width           =   3735
      End
      Begin VB.Line Line1 
         X1              =   120
         X2              =   3840
         Y1              =   1440
         Y2              =   1440
      End
      Begin VB.Label Label1 
         Caption         =   "Permiso:"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.Frame frAño 
      Height          =   1095
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4095
      Begin VB.CommandButton cmdSalir 
         Height          =   600
         Left            =   3240
         Picture         =   "frmCooperadoraPagos.frx":0144
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Salir"
         Top             =   240
         Width           =   600
      End
      Begin VB.TextBox txtAño 
         Height          =   315
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   480
         Width           =   735
      End
      Begin MSComCtl2.UpDown UpDown1 
         Height          =   375
         Left            =   840
         TabIndex        =   3
         Top             =   450
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   661
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin VB.Label Label4 
         Caption         =   "Año:"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   495
      End
   End
End
Attribute VB_Name = "frmCooperadoraPagos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Conexion As New Connection
Dim Resultado As New Recordset

Private Sub cmdDetalleDePagoPorApellido_Click()
    If frmIdentificacion.Permisos!AgregarPagosCooperadora = False Then Exit Sub
    Me.MousePointer = 11
    Conexion.Open
    Conexion.Execute ("DELETE * FROM [Listado Cooperadora]")
    
    'inserto los alumnos matriculados en este año
    Conexion.Execute ("INSERT INTO [Listado Cooperadora] ( Permiso, Alumno, Tipo, Documento, Ano ) SELECT DISTINCT [Cooperadora Pagos].Alumno, Alumnos.Nombre, Alumnos.Tipo, Alumnos.Documento, [Cooperadora Pagos].Año FROM [Cooperadora Pagos] INNER JOIN Alumnos ON [Cooperadora Pagos].Alumno = Alumnos.Permiso WHERE [Cooperadora Pagos].Año=" & txtAño)
    
    'inserto los nombres de carreras que esta cursando cada alumno
    Conexion.Execute ("UPDATE (([Listado Cooperadora] INNER JOIN Finales ON (Finales.Ano = [Listado Cooperadora].Ano) AND ([Listado Cooperadora].Permiso = Finales.Alumno)) INNER JOIN Materias ON Finales.Materia = Materias.Codigo) INNER JOIN Carreras ON Materias.Carrera = Carreras.Codigo SET [Listado Cooperadora].Carreras = [Carreras].[Nombre]")
    
    'inserto los importes cancelados en los respectivos conceptos
    Conexion.Execute ("UPDATE [Listado Cooperadora] INNER JOIN [Cooperadora Pagos] ON [Listado Cooperadora].Permiso = [Cooperadora Pagos].Alumno SET [Listado Cooperadora].Importe0 = [Cooperadora Pagos ].[Importe]WHERE [Cooperadora Pagos].Cancelado=True AND [Cooperadora Pagos].Concepto=0 AND [Cooperadora Pagos].Año = " & txtAño)
    Conexion.Execute ("UPDATE [Listado Cooperadora] INNER JOIN [Cooperadora Pagos] ON [Listado Cooperadora].Permiso = [Cooperadora Pagos].Alumno SET [Listado Cooperadora].Importe1 = [Cooperadora Pagos ].[Importe]WHERE [Cooperadora Pagos].Cancelado=True AND [Cooperadora Pagos].Concepto=1 AND [Cooperadora Pagos].Año = " & txtAño)
    Conexion.Execute ("UPDATE [Listado Cooperadora] INNER JOIN [Cooperadora Pagos] ON [Listado Cooperadora].Permiso = [Cooperadora Pagos].Alumno SET [Listado Cooperadora].Importe2 = [Cooperadora Pagos ].[Importe]WHERE [Cooperadora Pagos].Cancelado=True AND [Cooperadora Pagos].Concepto=2 AND [Cooperadora Pagos].Año = " & txtAño)
    Conexion.Execute ("UPDATE [Listado Cooperadora] INNER JOIN [Cooperadora Pagos] ON [Listado Cooperadora].Permiso = [Cooperadora Pagos].Alumno SET [Listado Cooperadora].Importe3 = [Cooperadora Pagos ].[Importe]WHERE [Cooperadora Pagos].Cancelado=True AND [Cooperadora Pagos].Concepto=3 AND [Cooperadora Pagos].Año = " & txtAño)
    Conexion.Execute ("UPDATE [Listado Cooperadora] INNER JOIN [Cooperadora Pagos] ON [Listado Cooperadora].Permiso = [Cooperadora Pagos].Alumno SET [Listado Cooperadora].Importe4 = [Cooperadora Pagos ].[Importe]WHERE [Cooperadora Pagos].Cancelado=True AND [Cooperadora Pagos].Concepto=4 AND [Cooperadora Pagos].Año = " & txtAño)
    Conexion.Execute ("UPDATE [Listado Cooperadora] INNER JOIN [Cooperadora Pagos] ON [Listado Cooperadora].Permiso = [Cooperadora Pagos].Alumno SET [Listado Cooperadora].Importe5 = [Cooperadora Pagos ].[Importe]WHERE [Cooperadora Pagos].Cancelado=True AND [Cooperadora Pagos].Concepto=5 AND [Cooperadora Pagos].Año = " & txtAño)
    Conexion.Execute ("UPDATE [Listado Cooperadora] INNER JOIN [Cooperadora Pagos] ON [Listado Cooperadora].Permiso = [Cooperadora Pagos].Alumno SET [Listado Cooperadora].Importe6 = [Cooperadora Pagos ].[Importe]WHERE [Cooperadora Pagos].Cancelado=True AND [Cooperadora Pagos].Concepto=6 AND [Cooperadora Pagos].Año = " & txtAño)
    Conexion.Execute ("UPDATE [Listado Cooperadora] INNER JOIN [Cooperadora Pagos] ON [Listado Cooperadora].Permiso = [Cooperadora Pagos].Alumno SET [Listado Cooperadora].Importe7 = [Cooperadora Pagos ].[Importe]WHERE [Cooperadora Pagos].Cancelado=True AND [Cooperadora Pagos].Concepto=7 AND [Cooperadora Pagos].Año = " & txtAño)
    Conexion.Execute ("UPDATE [Listado Cooperadora] INNER JOIN [Cooperadora Pagos] ON [Listado Cooperadora].Permiso = [Cooperadora Pagos].Alumno SET [Listado Cooperadora].Importe8 = [Cooperadora Pagos ].[Importe]WHERE [Cooperadora Pagos].Cancelado=True AND [Cooperadora Pagos].Concepto=8 AND [Cooperadora Pagos].Año = " & txtAño)
    Conexion.Execute ("UPDATE [Listado Cooperadora] INNER JOIN [Cooperadora Pagos] ON [Listado Cooperadora].Permiso = [Cooperadora Pagos].Alumno SET [Listado Cooperadora].Importe9 = [Cooperadora Pagos ].[Importe]WHERE [Cooperadora Pagos].Cancelado=True AND [Cooperadora Pagos].Concepto=9 AND [Cooperadora Pagos].Año = " & txtAño)
    Conexion.Execute ("UPDATE [Listado Cooperadora] INNER JOIN [Cooperadora Pagos] ON [Listado Cooperadora].Permiso = [Cooperadora Pagos].Alumno SET [Listado Cooperadora].Importe10 = [Cooperadora Pagos ].[Importe]WHERE [Cooperadora Pagos].Cancelado=True AND [Cooperadora Pagos].Concepto=10 AND [Cooperadora Pagos].Año = " & txtAño)
    Conexion.Execute ("UPDATE [Listado Cooperadora] INNER JOIN [Cooperadora Pagos] ON [Listado Cooperadora].Permiso = [Cooperadora Pagos].Alumno SET [Listado Cooperadora].Importe11 = [Cooperadora Pagos ].[Importe]WHERE [Cooperadora Pagos].Cancelado=True AND [Cooperadora Pagos].Concepto=11 AND [Cooperadora Pagos].Año = " & txtAño)
    Conexion.Execute ("UPDATE [Listado Cooperadora] INNER JOIN [Cooperadora Pagos] ON [Listado Cooperadora].Permiso = [Cooperadora Pagos].Alumno SET [Listado Cooperadora].Importe12 = [Cooperadora Pagos ].[Importe]WHERE [Cooperadora Pagos].Cancelado=True AND [Cooperadora Pagos].Concepto=12 AND [Cooperadora Pagos].Año = " & txtAño)
    Conexion.Close
    RptDetalleDePagosPorApellido.PrintReport
    'para que no queden datos en la tabla por las dudas que abran el archivo rpt
    Conexion.Open
    Conexion.Execute ("DELETE * FROM [Listado Cooperadora]")
    Conexion.Close
    Me.MousePointer = 0
End Sub

Private Sub cmdDetalleDePagos_Click()
    If frmIdentificacion.Permisos!AgregarPagosCooperadora = False Then Exit Sub
    Me.MousePointer = 11
    Conexion.Open
    Conexion.Execute ("DELETE * FROM [Listado Cooperadora]")
    
    'inserto los alumnos matriculados en este año
    Conexion.Execute ("INSERT INTO [Listado Cooperadora] ( Permiso, Alumno, Tipo, Documento, Ano ) SELECT DISTINCT [Cooperadora Pagos].Alumno, Alumnos.Nombre, Alumnos.Tipo, Alumnos.Documento, [Cooperadora Pagos].Año FROM [Cooperadora Pagos] INNER JOIN Alumnos ON [Cooperadora Pagos].Alumno = Alumnos.Permiso WHERE [Cooperadora Pagos].Año=" & txtAño)
    
    'inserto los nombres de carreras que esta cursando cada alumno
    Conexion.Execute ("UPDATE (([Listado Cooperadora] INNER JOIN Finales ON (Finales.Ano = [Listado Cooperadora].Ano) AND ([Listado Cooperadora].Permiso = Finales.Alumno)) INNER JOIN Materias ON Finales.Materia = Materias.Codigo) INNER JOIN Carreras ON Materias.Carrera = Carreras.Codigo SET [Listado Cooperadora].Carreras = [Carreras].[Nombre]")
    
    'inserto los importes cancelados en los respectivos conceptos
    Conexion.Execute ("UPDATE [Listado Cooperadora] INNER JOIN [Cooperadora Pagos] ON [Listado Cooperadora].Permiso = [Cooperadora Pagos].Alumno SET [Listado Cooperadora].Importe0 = [Cooperadora Pagos ].[Importe]WHERE [Cooperadora Pagos].Cancelado=True AND [Cooperadora Pagos].Concepto=0 AND [Cooperadora Pagos].Año = " & txtAño)
    Conexion.Execute ("UPDATE [Listado Cooperadora] INNER JOIN [Cooperadora Pagos] ON [Listado Cooperadora].Permiso = [Cooperadora Pagos].Alumno SET [Listado Cooperadora].Importe1 = [Cooperadora Pagos ].[Importe]WHERE [Cooperadora Pagos].Cancelado=True AND [Cooperadora Pagos].Concepto=1 AND [Cooperadora Pagos].Año = " & txtAño)
    Conexion.Execute ("UPDATE [Listado Cooperadora] INNER JOIN [Cooperadora Pagos] ON [Listado Cooperadora].Permiso = [Cooperadora Pagos].Alumno SET [Listado Cooperadora].Importe2 = [Cooperadora Pagos ].[Importe]WHERE [Cooperadora Pagos].Cancelado=True AND [Cooperadora Pagos].Concepto=2 AND [Cooperadora Pagos].Año = " & txtAño)
    Conexion.Execute ("UPDATE [Listado Cooperadora] INNER JOIN [Cooperadora Pagos] ON [Listado Cooperadora].Permiso = [Cooperadora Pagos].Alumno SET [Listado Cooperadora].Importe3 = [Cooperadora Pagos ].[Importe]WHERE [Cooperadora Pagos].Cancelado=True AND [Cooperadora Pagos].Concepto=3 AND [Cooperadora Pagos].Año = " & txtAño)
    Conexion.Execute ("UPDATE [Listado Cooperadora] INNER JOIN [Cooperadora Pagos] ON [Listado Cooperadora].Permiso = [Cooperadora Pagos].Alumno SET [Listado Cooperadora].Importe4 = [Cooperadora Pagos ].[Importe]WHERE [Cooperadora Pagos].Cancelado=True AND [Cooperadora Pagos].Concepto=4 AND [Cooperadora Pagos].Año = " & txtAño)
    Conexion.Execute ("UPDATE [Listado Cooperadora] INNER JOIN [Cooperadora Pagos] ON [Listado Cooperadora].Permiso = [Cooperadora Pagos].Alumno SET [Listado Cooperadora].Importe5 = [Cooperadora Pagos ].[Importe]WHERE [Cooperadora Pagos].Cancelado=True AND [Cooperadora Pagos].Concepto=5 AND [Cooperadora Pagos].Año = " & txtAño)
    Conexion.Execute ("UPDATE [Listado Cooperadora] INNER JOIN [Cooperadora Pagos] ON [Listado Cooperadora].Permiso = [Cooperadora Pagos].Alumno SET [Listado Cooperadora].Importe6 = [Cooperadora Pagos ].[Importe]WHERE [Cooperadora Pagos].Cancelado=True AND [Cooperadora Pagos].Concepto=6 AND [Cooperadora Pagos].Año = " & txtAño)
    Conexion.Execute ("UPDATE [Listado Cooperadora] INNER JOIN [Cooperadora Pagos] ON [Listado Cooperadora].Permiso = [Cooperadora Pagos].Alumno SET [Listado Cooperadora].Importe7 = [Cooperadora Pagos ].[Importe]WHERE [Cooperadora Pagos].Cancelado=True AND [Cooperadora Pagos].Concepto=7 AND [Cooperadora Pagos].Año = " & txtAño)
    Conexion.Execute ("UPDATE [Listado Cooperadora] INNER JOIN [Cooperadora Pagos] ON [Listado Cooperadora].Permiso = [Cooperadora Pagos].Alumno SET [Listado Cooperadora].Importe8 = [Cooperadora Pagos ].[Importe]WHERE [Cooperadora Pagos].Cancelado=True AND [Cooperadora Pagos].Concepto=8 AND [Cooperadora Pagos].Año = " & txtAño)
    Conexion.Execute ("UPDATE [Listado Cooperadora] INNER JOIN [Cooperadora Pagos] ON [Listado Cooperadora].Permiso = [Cooperadora Pagos].Alumno SET [Listado Cooperadora].Importe9 = [Cooperadora Pagos ].[Importe]WHERE [Cooperadora Pagos].Cancelado=True AND [Cooperadora Pagos].Concepto=9 AND [Cooperadora Pagos].Año = " & txtAño)
    Conexion.Execute ("UPDATE [Listado Cooperadora] INNER JOIN [Cooperadora Pagos] ON [Listado Cooperadora].Permiso = [Cooperadora Pagos].Alumno SET [Listado Cooperadora].Importe10 = [Cooperadora Pagos ].[Importe]WHERE [Cooperadora Pagos].Cancelado=True AND [Cooperadora Pagos].Concepto=10 AND [Cooperadora Pagos].Año = " & txtAño)
    Conexion.Execute ("UPDATE [Listado Cooperadora] INNER JOIN [Cooperadora Pagos] ON [Listado Cooperadora].Permiso = [Cooperadora Pagos].Alumno SET [Listado Cooperadora].Importe11 = [Cooperadora Pagos ].[Importe]WHERE [Cooperadora Pagos].Cancelado=True AND [Cooperadora Pagos].Concepto=11 AND [Cooperadora Pagos].Año = " & txtAño)
    Conexion.Execute ("UPDATE [Listado Cooperadora] INNER JOIN [Cooperadora Pagos] ON [Listado Cooperadora].Permiso = [Cooperadora Pagos].Alumno SET [Listado Cooperadora].Importe12 = [Cooperadora Pagos ].[Importe]WHERE [Cooperadora Pagos].Cancelado=True AND [Cooperadora Pagos].Concepto=12 AND [Cooperadora Pagos].Año = " & txtAño)
    Conexion.Close
    rptDetalleDePagos.PrintReport
    'para que no queden datos en la tabla por las dudas que abran el archivo rpt
    Conexion.Open
    Conexion.Execute ("DELETE * FROM [Listado Cooperadora]")
    Conexion.Close
    Me.MousePointer = 0
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub dtgConceptos_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then txtPermiso.SetFocus
    If KeyAscii = 13 Then
        If frmIdentificacion.Permisos!AgregarPagosCooperadora = True Then
            AgregarPago
        Else
            MsgBox ("Usted no tiene permiso para ingresar pagos de Cooperadora")
        End If
    End If
    
End Sub

Private Sub Form_Load()
    Conexion.ConnectionString = ("DSN=Instituto")
    txtAño = Format(Date, "yyyy")
End Sub

Private Sub txtPermiso_GotFocus()
    txtPermiso = ""
    dtgConceptos.Enabled = False
End Sub

Private Sub txtPermiso_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Conexion.Open
        Set Resultado = Conexion.Execute("SELECT Alumnos.Nombre, Alumnos.Tipo, Alumnos.Documento From Alumnos WHERE Alumnos.Permiso= " & txtPermiso & " AND Alumnos.Eliminado=False")
            If Resultado.EOF = False Then 'el alumno existe
                lblNombre = Resultado!Nombre
                lblDocumento = Resultado!Tipo & " " & Resultado!documento
                adoConceptos.RecordSource = "SELECT [Cooperadora Conceptos].Concepto, [Cooperadora Pagos].Importe, [Cooperadora Pagos].Cancelado, [Cooperadora Conceptos].Codigo FROM [Cooperadora Pagos] INNER JOIN [Cooperadora Conceptos] ON [Cooperadora Pagos].Concepto = [Cooperadora Conceptos].Codigo Where [Cooperadora Pagos].Alumno = " & txtPermiso & " AND [Cooperadora Pagos].Año =" & txtAño & " ORDER BY [Cooperadora Conceptos].Codigo"
                adoConceptos.Refresh
                If adoConceptos.Recordset.EOF = False Then 'el alumno se matriculo ese año
                    dtgConceptos.Enabled = True
                    dtgConceptos.SetFocus
                Else
                    MsgBox ("El alumno no se matriculó en este año o no tiene el plan generado"): Conexion.Close: txtPermiso.SetFocus: Exit Sub
                End If
            Else
                MsgBox ("El alumno no existe"): Conexion.Close: txtPermiso.SetFocus: Exit Sub
            End If
        Conexion.Close
    End If
End Sub

Private Sub UpDown1_DownClick()
    txtAño = Val(txtAño) - 1
End Sub

Private Sub UpDown1_UpClick()
    txtAño = Val(txtAño) + 1
End Sub

Private Function AgregarPago()
    'Exit Function
    With frmCooperadoraAgregarPago
        .txtImporte = Format(adoConceptos.Recordset!Importe, "0.00")
        .Show 1
    End With
End Function
