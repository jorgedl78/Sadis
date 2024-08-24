VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "msadodc.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Begin VB.Form frmCooperadoraPlan 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Plan de Cooperadora"
   ClientHeight    =   7410
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3720
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   7410
   ScaleWidth      =   3720
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame frListados 
      Caption         =   "Listados"
      Height          =   1215
      Left            =   0
      TabIndex        =   11
      Top             =   6120
      Width           =   3615
      Begin Crystal.CrystalReport rptDetalleDePagos 
         Left            =   960
         Top             =   720
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   262150
         ReportFileName  =   "listpago.rpt"
         WindowTitle     =   "Detalle de Aportes de Cooperadora"
         WindowControlBox=   -1  'True
         WindowMaxButton =   -1  'True
         WindowMinButton =   -1  'True
      End
      Begin VB.CommandButton cmdListadoGeneral 
         Caption         =   "Listado General"
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
         TabIndex        =   12
         Top             =   360
         Width           =   1455
      End
   End
   Begin VB.Frame frPlanAnual 
      Caption         =   "Plan Anual"
      Height          =   4935
      Left            =   0
      TabIndex        =   6
      Top             =   1200
      Width           =   3615
      Begin VB.CommandButton cmdGenerarPlan 
         Caption         =   "Generar Plan"
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
         Left            =   2520
         TabIndex        =   13
         Top             =   4200
         Width           =   975
      End
      Begin MSAdodcLib.Adodc adoPlanAnual 
         Height          =   375
         Left            =   960
         Top             =   3480
         Visible         =   0   'False
         Width           =   2295
         _ExtentX        =   4048
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
         RecordSource    =   $"frmCooperadoraPlan.frx":0000
         Caption         =   "PlanAnual"
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
      Begin VB.CommandButton cmdEliminar 
         Enabled         =   0   'False
         Height          =   600
         Left            =   1680
         Picture         =   "frmCooperadoraPlan.frx":0126
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Borrar"
         Top             =   4200
         Width           =   600
      End
      Begin VB.CommandButton cmdModificar 
         Enabled         =   0   'False
         Height          =   600
         Left            =   960
         Picture         =   "frmCooperadoraPlan.frx":0568
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Modificar"
         Top             =   4200
         Width           =   600
      End
      Begin VB.CommandButton cmdAgregar 
         Enabled         =   0   'False
         Height          =   600
         Left            =   120
         Picture         =   "frmCooperadoraPlan.frx":09AA
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Agregar"
         Top             =   4200
         Width           =   600
      End
      Begin MSDataGridLib.DataGrid dtgPlanAnual 
         Bindings        =   "frmCooperadoraPlan.frx":0DEC
         Height          =   3855
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   3375
         _ExtentX        =   5953
         _ExtentY        =   6800
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
         ColumnCount     =   2
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
         SplitCount      =   1
         BeginProperty Split0 
            MarqueeStyle    =   4
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
               Alignment       =   1
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame frAño 
      Height          =   1095
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3615
      Begin VB.CommandButton cmdSalir 
         Height          =   600
         Left            =   2640
         Picture         =   "frmCooperadoraPlan.frx":0E07
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Salir"
         Top             =   360
         Width           =   600
      End
      Begin VB.CommandButton cmdMostrar 
         Caption         =   "Mostrar"
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
         Left            =   1200
         TabIndex        =   4
         Top             =   360
         Width           =   1095
      End
      Begin VB.TextBox txtAño 
         Height          =   315
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   1
         Top             =   480
         Width           =   735
      End
      Begin MSComCtl2.UpDown UpDown1 
         Height          =   375
         Left            =   840
         TabIndex        =   2
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
         TabIndex        =   3
         Top             =   240
         Width           =   495
      End
   End
End
Attribute VB_Name = "frmCooperadoraPlan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Conexion As New Connection
Dim Resultado As New Recordset
Dim Alumnos As New Recordset
Dim PlanGenerado As New Recordset

Private Sub cmdGenerarPlan_Click()
    'Exit Sub
    Respuesta = MsgBox("¿Está seguro de generar el plan de cooperadora para los alumnos matriculados en el año " & txtAño & "?", vbYesNo, "Confirmar")
    If Respuesta = vbNo Then Exit Sub
    Me.MousePointer = 11
    Generados = 0
    Conexion.Open
    'levanto todos los alumnos que cursan alguna materia en este año
    Set Alumnos = Conexion.Execute("SELECT Distinct Finales.Alumno From Finales Where Finales.Ano = " & Val(txtAño) & " ORDER BY Finales.Alumno")
    'levanto los registros de la tabla para ver si ya se generó algún alumno matriculado ese año
    Set Resultado = Conexion.Execute("SELECT [Cooperadora Pagos].Año From [Cooperadora Pagos] WHERE [Cooperadora Pagos].Año=" & Val(txtAño))
    If Resultado.EOF = False Then 'uno por uno
        While Alumnos.EOF = False
            Set Resultado = Conexion.Execute("SELECT Alumno From [Cooperadora Pagos] WHERE Alumno = " & Alumnos!Alumno)
            If Resultado.EOF = True Then 'no esta generado el plan para este alumno
                Conexion.Execute ("INSERT INTO [Cooperadora Pagos] ( Alumno, Año, Concepto, Importe ) SELECT Alumnos.Permiso, [Cooperadora Plan].Año, [Cooperadora Plan].Concepto, [Cooperadora Plan].Importe From Alumnos, [Cooperadora Plan] WHERE Alumnos.Permiso=" & Alumnos!Alumno & " AND [Cooperadora Plan].Año = " & txtAño)
                Generados = Generados + 1
            End If
            Alumnos.MoveNext
        Wend
        If Generados = 0 Then
            MsgBox ("No se agregó ningún alumno al plan")
        Else
            MsgBox ("Se agregaron " & Generados & " alumnos matriculados al plan")
        End If
    Else 'genero todos los matriculados
        While Alumnos.EOF = False
            Conexion.Execute ("INSERT INTO [Cooperadora Pagos] ( Alumno, Año, Concepto, Importe ) SELECT Alumnos.Permiso, [Cooperadora Plan].Año, [Cooperadora Plan].Concepto, [Cooperadora Plan].Importe From Alumnos, [Cooperadora Plan] WHERE Alumnos.Permiso=" & Alumnos!Alumno & " AND [Cooperadora Plan].Año = " & Val(txtAño))
            Alumnos.MoveNext
            Generados = Generados + 1
        Wend
    MsgBox ("Se generó el plan para " & Generados & " alumnos matriculados")
    End If
    Conexion.Close
    Me.MousePointer = 0
End Sub

Private Sub cmdListadoGeneral_Click()
    Me.MousePointer = 11
    Conexion.Open
    Conexion.Execute ("DELETE * FROM [Listado Cooperadora]")
    
    'inserto los alumnos matriculados en este año
    Conexion.Execute ("INSERT INTO [Listado Cooperadora] ( Permiso, Alumno, Tipo, Documento, Ano ) SELECT DISTINCT [Cooperadora Pagos].Alumno, Alumnos.Nombre, Alumnos.Tipo, Alumnos.Documento, [Cooperadora Pagos].Año FROM [Cooperadora Pagos] INNER JOIN Alumnos ON [Cooperadora Pagos].Alumno = Alumnos.Permiso WHERE [Cooperadora Pagos].Año=" & txtAño)
    
    'inserto los nombres de carreras que esta cursando cada alumno
    Conexion.Execute ("UPDATE (([Listado Cooperadora] INNER JOIN Finales ON (Finales.Ano = [Listado Cooperadora].Ano) AND ([Listado Cooperadora].Permiso = Finales.Alumno)) INNER JOIN Materias ON Finales.Materia = Materias.Codigo) INNER JOIN Carreras ON Materias.Carrera = Carreras.Codigo SET [Listado Cooperadora].Carreras = [Carreras].[Nombre]")

    Conexion.Close
    rptDetalleDePagos.PrintReport
    Me.MousePointer = 0
End Sub

Private Sub cmdMostrar_Click()
    adoPlanAnual.RecordSource = "SELECT [Cooperadora Conceptos].Concepto, [Cooperadora Plan].Importe, [Cooperadora Plan].Generada, [Cooperadora Conceptos].Codigo FROM [Cooperadora Conceptos] INNER JOIN [Cooperadora Plan] ON [Cooperadora Conceptos].Codigo = [Cooperadora Plan].Concepto Where [Cooperadora Plan].Año = " & Val(txtAño) & " ORDER BY [Cooperadora Conceptos].Codigo"
    adoPlanAnual.Refresh
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Conexion.ConnectionString = ("DSN=Instituto")
    txtAño = Format(Date, "yyyy")
End Sub

Private Sub UpDown1_DownClick()
    txtAño = Val(txtAño) - 1
End Sub

Private Sub UpDown1_UpClick()
    txtAño = Val(txtAño) + 1
End Sub

