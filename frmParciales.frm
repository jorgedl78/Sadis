VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Begin VB.Form frmParciales 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Parciales y Asistencias por Divisiones"
   ClientHeight    =   10620
   ClientLeft      =   45
   ClientTop       =   240
   ClientWidth     =   14910
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10620
   ScaleWidth      =   14910
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin Crystal.CrystalReport rptAsistenciaCompletada 
      Left            =   7320
      Top             =   8640
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "asistcom.rpt"
   End
   Begin MSAdodcLib.Adodc AdoFinal 
      Height          =   375
      Left            =   10560
      Top             =   6240
      Visible         =   0   'False
      Width           =   1815
      _ExtentX        =   3201
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
      RecordSource    =   "SELECT Aprobada FROM Finales WHERE Alumno = 0 AND Materia=0 AND Cursada=1"
      Caption         =   "adoFinal"
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
   Begin MSAdodcLib.Adodc adoCorrelativas 
      Height          =   375
      Left            =   8760
      Top             =   6240
      Visible         =   0   'False
      Width           =   1575
      _ExtentX        =   2778
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
      RecordSource    =   $"frmParciales.frx":0000
      Caption         =   "adoCorrelativas"
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
   Begin VB.Frame Frame2 
      Caption         =   "Asistencia"
      Height          =   1935
      Left            =   11880
      TabIndex        =   78
      Top             =   6600
      Width           =   2895
      Begin VB.CommandButton cmdAsistenciaCompleta 
         Caption         =   "Completa"
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
         Left            =   1680
         TabIndex        =   88
         Top             =   780
         Width           =   1095
      End
      Begin VB.CommandButton cmdAsistencia 
         Caption         =   "Planilla"
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
         Left            =   1680
         TabIndex        =   87
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton cmdIngresarAsistencia 
         Caption         =   "Ingresar Asistencia"
         Enabled         =   0   'False
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
         Left            =   240
         TabIndex        =   81
         Top             =   240
         Width           =   1335
      End
      Begin VB.CommandButton cmdVerAsistencia 
         Caption         =   "Ver Asistencia"
         Enabled         =   0   'False
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
         Left            =   240
         TabIndex        =   80
         Top             =   780
         Width           =   1335
      End
      Begin VB.CommandButton cmdCalcularAsistencia 
         Caption         =   "Calcular Asistencia"
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
         Left            =   240
         TabIndex        =   79
         ToolTipText     =   "Calcula y Actualiza el Porcentaje de Asistencia de estos alumnos matriculados"
         Top             =   1320
         Width           =   1335
      End
   End
   Begin VB.Frame frDetalles 
      Enabled         =   0   'False
      Height          =   1815
      Left            =   7200
      TabIndex        =   21
      Top             =   3240
      Width           =   7575
      Begin VB.TextBox txtPorcentajeAsistencia 
         Enabled         =   0   'False
         Height          =   285
         Left            =   3000
         TabIndex        =   73
         Top             =   1260
         Width           =   615
      End
      Begin VB.CheckBox chkAsistencia 
         Alignment       =   1  'Right Justify
         Caption         =   "Aprobó Asistencia"
         Enabled         =   0   'False
         Height          =   195
         Left            =   120
         TabIndex        =   71
         Top             =   1320
         Width           =   1695
      End
      Begin VB.TextBox txtTotalizador 
         Height          =   285
         Left            =   5760
         TabIndex        =   37
         Top             =   840
         Width           =   495
      End
      Begin VB.TextBox txtParcial2 
         Height          =   285
         Left            =   1320
         TabIndex        =   36
         Top             =   840
         Width           =   495
      End
      Begin VB.TextBox txtRecuperatorio2 
         Height          =   285
         Left            =   1920
         TabIndex        =   35
         Top             =   840
         Width           =   495
      End
      Begin VB.TextBox txtParcial1 
         Height          =   285
         Left            =   120
         TabIndex        =   34
         Top             =   840
         Width           =   495
      End
      Begin VB.TextBox txtRecuperatorio1 
         Height          =   285
         Left            =   720
         TabIndex        =   33
         Top             =   840
         Width           =   495
      End
      Begin VB.TextBox txtPractico4 
         Height          =   285
         Left            =   4440
         TabIndex        =   32
         Top             =   840
         Width           =   495
      End
      Begin VB.TextBox txtPractico3 
         Height          =   285
         Left            =   3840
         TabIndex        =   31
         Top             =   840
         Width           =   495
      End
      Begin VB.TextBox txtPractico2 
         Height          =   285
         Left            =   3240
         TabIndex        =   30
         Top             =   840
         Width           =   495
      End
      Begin VB.TextBox txtPractico1 
         Height          =   285
         Left            =   2640
         TabIndex        =   29
         Top             =   840
         Width           =   495
      End
      Begin VB.TextBox txtPractico5 
         Height          =   285
         Left            =   5040
         TabIndex        =   28
         Top             =   840
         Width           =   495
      End
      Begin VB.CheckBox chkPromocion 
         Height          =   255
         Left            =   6600
         TabIndex        =   38
         Top             =   840
         Width           =   255
      End
      Begin VB.CheckBox chkCursada 
         Height          =   255
         Left            =   7080
         TabIndex        =   39
         Top             =   840
         Width           =   255
      End
      Begin VB.Label lblPerdioCursada 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Cursada Vencida"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   240
         Left            =   5835
         TabIndex        =   76
         Top             =   1275
         Visible         =   0   'False
         Width           =   1635
      End
      Begin VB.Line Line16 
         BorderWidth     =   2
         X1              =   7440
         X2              =   7440
         Y1              =   240
         Y2              =   1200
      End
      Begin VB.Line Line5 
         BorderWidth     =   2
         X1              =   6960
         X2              =   6960
         Y1              =   240
         Y2              =   1200
      End
      Begin VB.Line Line15 
         BorderWidth     =   2
         X1              =   120
         X2              =   7440
         Y1              =   1200
         Y2              =   1200
      End
      Begin VB.Line Line3 
         BorderWidth     =   2
         X1              =   5640
         X2              =   5640
         Y1              =   240
         Y2              =   1200
      End
      Begin VB.Line Line2 
         BorderWidth     =   2
         X1              =   2520
         X2              =   2520
         Y1              =   240
         Y2              =   1200
      End
      Begin VB.Line Line6 
         BorderWidth     =   2
         X1              =   120
         X2              =   5640
         Y1              =   480
         Y2              =   480
      End
      Begin VB.Line Line17 
         BorderWidth     =   2
         X1              =   90
         X2              =   90
         Y1              =   240
         Y2              =   1200
      End
      Begin VB.Line Line14 
         BorderWidth     =   2
         X1              =   120
         X2              =   7440
         Y1              =   240
         Y2              =   240
      End
      Begin VB.Line Line13 
         X1              =   4980
         X2              =   4980
         Y1              =   480
         Y2              =   1200
      End
      Begin VB.Line Line12 
         X1              =   4380
         X2              =   4380
         Y1              =   480
         Y2              =   1200
      End
      Begin VB.Line Line11 
         X1              =   3780
         X2              =   3780
         Y1              =   480
         Y2              =   1200
      End
      Begin VB.Line Line10 
         X1              =   3180
         X2              =   3180
         Y1              =   480
         Y2              =   1200
      End
      Begin VB.Line Line9 
         X1              =   1860
         X2              =   1860
         Y1              =   480
         Y2              =   1200
      End
      Begin VB.Line Line8 
         X1              =   660
         X2              =   660
         Y1              =   480
         Y2              =   1200
      End
      Begin VB.Line Line7 
         BorderWidth     =   2
         X1              =   120
         X2              =   7440
         Y1              =   720
         Y2              =   720
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Prom."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   6405
         TabIndex        =   26
         ToolTipText     =   "Promocionó la materia"
         Top             =   360
         Width           =   510
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Totaliz."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   5715
         TabIndex        =   25
         ToolTipText     =   "Examen totalizador en mesas de final"
         Top             =   360
         Width           =   600
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Trabajos Prácticos"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   3315
         TabIndex        =   24
         Top             =   240
         Width           =   1545
      End
      Begin VB.Line Line4 
         BorderWidth     =   2
         X1              =   6360
         X2              =   6360
         Y1              =   240
         Y2              =   1200
      End
      Begin VB.Line Line1 
         BorderWidth     =   2
         X1              =   1260
         X2              =   1260
         Y1              =   240
         Y2              =   1200
      End
      Begin VB.Label Label12 
         Alignment       =   2  'Center
         Caption         =   "Nota"
         Height          =   255
         Left            =   1320
         TabIndex        =   41
         Top             =   480
         Width           =   495
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "2º Parcial"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   1560
         TabIndex        =   23
         Top             =   240
         Width           =   750
      End
      Begin VB.Label Label13 
         Alignment       =   2  'Center
         Caption         =   "Recup."
         Height          =   255
         Left            =   720
         TabIndex        =   42
         Top             =   480
         Width           =   495
      End
      Begin VB.Label Label11 
         Alignment       =   2  'Center
         Caption         =   "Nota"
         Height          =   255
         Left            =   120
         TabIndex        =   40
         Top             =   480
         Width           =   495
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "1º Parcial"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   360
         TabIndex        =   22
         Top             =   240
         Width           =   750
      End
      Begin VB.Label Label14 
         Alignment       =   2  'Center
         Caption         =   "Recup."
         Height          =   255
         Left            =   1920
         TabIndex        =   43
         Top             =   480
         Width           =   495
      End
      Begin VB.Label Label19 
         Alignment       =   2  'Center
         Caption         =   "1º"
         Height          =   255
         Left            =   2640
         TabIndex        =   48
         Top             =   480
         Width           =   495
      End
      Begin VB.Label Label18 
         Alignment       =   2  'Center
         Caption         =   "2º"
         Height          =   255
         Left            =   3240
         TabIndex        =   47
         Top             =   480
         Width           =   495
      End
      Begin VB.Label Label17 
         Alignment       =   2  'Center
         Caption         =   "3º"
         Height          =   255
         Left            =   3840
         TabIndex        =   46
         Top             =   480
         Width           =   495
      End
      Begin VB.Label Label16 
         Alignment       =   2  'Center
         Caption         =   "4º"
         Height          =   255
         Left            =   4440
         TabIndex        =   45
         Top             =   480
         Width           =   495
      End
      Begin VB.Label Label15 
         Alignment       =   2  'Center
         Caption         =   "5º"
         Height          =   255
         Left            =   5040
         TabIndex        =   44
         Top             =   480
         Width           =   495
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         Caption         =   "Curs."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   6960
         TabIndex        =   27
         ToolTipText     =   "Aprobó la cursada"
         Top             =   360
         Width           =   465
      End
      Begin VB.Label Label25 
         BackStyle       =   0  'Transparent
         Caption         =   "Porcentaje"
         Height          =   255
         Left            =   2160
         TabIndex        =   72
         Top             =   1320
         Width           =   855
      End
   End
   Begin VB.Frame frIngresoDeNotas 
      Caption         =   "Ingreso de Notas"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   7200
      TabIndex        =   60
      Top             =   6600
      Visible         =   0   'False
      Width           =   4455
      Begin VB.CommandButton cmdIngresarNota 
         Height          =   495
         Left            =   1920
         Picture         =   "frmParciales.frx":0074
         Style           =   1  'Graphical
         TabIndex        =   65
         ToolTipText     =   "Aceptar"
         Top             =   1320
         Width           =   495
      End
      Begin VB.TextBox txtNota 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   840
         MaxLength       =   4
         TabIndex        =   64
         Top             =   1440
         Width           =   735
      End
      Begin VB.Label lblAlumnoNota 
         Caption         =   "lblAlumnoNota"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   66
         Top             =   720
         Width           =   5055
      End
      Begin VB.Label Label23 
         Caption         =   "Nota:"
         Height          =   255
         Left            =   240
         TabIndex        =   63
         Top             =   1440
         Width           =   495
      End
      Begin VB.Label lblDescripcionDeNota 
         Caption         =   "DescripcionDeNota"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1920
         TabIndex        =   62
         Top             =   310
         Width           =   3135
      End
      Begin VB.Label Label22 
         Caption         =   "Ingresando Notas de:"
         Height          =   255
         Left            =   120
         TabIndex        =   61
         Top             =   360
         Width           =   1575
      End
   End
   Begin VB.Frame Frame1 
      Height          =   855
      Left            =   7200
      TabIndex        =   53
      Top             =   2040
      Width           =   7575
      Begin VB.Label lblSituacion 
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
         Left            =   5160
         TabIndex        =   56
         Top             =   480
         Width           =   2295
      End
      Begin VB.Label Label21 
         Caption         =   "Situación:"
         Height          =   255
         Left            =   5280
         TabIndex        =   55
         Top             =   240
         Width           =   855
      End
      Begin VB.Line Line18 
         X1              =   5040
         X2              =   5040
         Y1              =   120
         Y2              =   840
      End
      Begin VB.Label lblAlumno 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   54
         Top             =   240
         Width           =   4815
      End
   End
   Begin VB.Frame frPlanillas 
      Caption         =   "Planillas"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   7200
      TabIndex        =   49
      Top             =   9000
      Width           =   7455
      Begin VB.CommandButton cmdPlanillaParcualesCompleta 
         Caption         =   "Parciales Completa"
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
         Left            =   2280
         TabIndex        =   92
         Top             =   480
         Width           =   975
      End
      Begin VB.CommandButton cmdParcialeB 
         Caption         =   "Parciales ""B"""
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
         TabIndex        =   91
         Top             =   480
         Width           =   975
      End
      Begin VB.CommandButton cmdNroCursada 
         Caption         =   "Nº Cursada"
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
         Left            =   6120
         TabIndex        =   86
         Top             =   480
         Width           =   1215
      End
      Begin Crystal.CrystalReport rptPlanillaAnalitica 
         Left            =   5040
         Top             =   120
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   262150
         ReportFileName  =   "plaanali.rpt"
         WindowTitle     =   "Planilla Analítica"
         WindowControlBox=   -1  'True
         WindowMaxButton =   -1  'True
         WindowMinButton =   -1  'True
      End
      Begin VB.CommandButton cmdPara_desdoblamiento 
         Caption         =   "Planilla Analítica"
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
         Left            =   3360
         TabIndex        =   84
         Top             =   480
         Width           =   1095
      End
      Begin Crystal.CrystalReport rptRecuperatorioAsistencia 
         Left            =   3360
         Top             =   120
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   262150
         ReportFileName  =   "recupera.rpt"
         WindowControlBox=   -1  'True
         WindowMaxButton =   -1  'True
         WindowMinButton =   -1  'True
      End
      Begin VB.CommandButton cmdRecuperatorioAsistencia 
         Caption         =   "Recuperatorio Asistencia"
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
         Left            =   4680
         TabIndex        =   82
         Top             =   480
         Width           =   1335
      End
      Begin Crystal.CrystalReport rptParciales 
         Left            =   1200
         Top             =   0
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   262150
         ReportFileName  =   "parcial.rpt"
         WindowTitle     =   "Planilla de Parciales"
         WindowControlBox=   -1  'True
         WindowMaxButton =   -1  'True
         WindowMinButton =   -1  'True
      End
      Begin VB.CommandButton cmdParciales 
         Caption         =   "Parciales"
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
         TabIndex        =   50
         Top             =   480
         Width           =   975
      End
      Begin Crystal.CrystalReport rptAsistencia 
         Left            =   720
         Top             =   480
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   262150
         ReportFileName  =   "asistenc.rpt"
         WindowTitle     =   "Planilla de Asistencia"
         WindowControlBox=   -1  'True
         WindowMaxButton =   -1  'True
         WindowMinButton =   -1  'True
      End
   End
   Begin VB.Frame frComandos 
      Enabled         =   0   'False
      Height          =   975
      Left            =   7200
      TabIndex        =   17
      Top             =   5160
      Width           =   7575
      Begin VB.CommandButton cmdPasarALibre 
         Caption         =   "Pasar a Libre"
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
         Left            =   3600
         TabIndex        =   89
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton cmdIngresaRecuAsistencia 
         Caption         =   "Ingresar Rec. Asistencia"
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
         Left            =   4680
         TabIndex        =   83
         Top             =   240
         Width           =   1335
      End
      Begin VB.CommandButton cmdAprobarAsistenciaATodos 
         Caption         =   "Aprobar Asistencia a Todos"
         Enabled         =   0   'False
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
         Left            =   4800
         TabIndex        =   74
         Top             =   120
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.CommandButton cmdImprimeActaPromocion 
         Caption         =   "Imprime Acta Promoción"
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
         Left            =   6120
         TabIndex        =   70
         Top             =   240
         Width           =   1335
      End
      Begin VB.CommandButton cmdIngresarNotas 
         Caption         =   "Ingresar Notas"
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
         Left            =   2400
         TabIndex        =   69
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton cmdCancelar 
         Enabled         =   0   'False
         Height          =   600
         Left            =   1680
         Picture         =   "frmParciales.frx":04B6
         Style           =   1  'Graphical
         TabIndex        =   20
         ToolTipText     =   "Cancelar"
         Top             =   240
         Width           =   600
      End
      Begin VB.CommandButton cmdGuardar 
         Enabled         =   0   'False
         Height          =   600
         Left            =   960
         Picture         =   "frmParciales.frx":08F8
         Style           =   1  'Graphical
         TabIndex        =   19
         ToolTipText     =   "Guardar"
         Top             =   240
         Width           =   600
      End
      Begin VB.CommandButton cmdModificar 
         Enabled         =   0   'False
         Height          =   600
         Left            =   240
         Picture         =   "frmParciales.frx":0D3A
         Style           =   1  'Graphical
         TabIndex        =   18
         ToolTipText     =   "Modificar"
         Top             =   240
         Width           =   600
      End
   End
   Begin VB.Frame frMaterias 
      Enabled         =   0   'False
      Height          =   855
      Left            =   0
      TabIndex        =   12
      Top             =   960
      Width           =   14775
      Begin VB.ComboBox cbDivision 
         Height          =   315
         Left            =   9000
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   360
         Width           =   735
      End
      Begin MSDataListLib.DataCombo dtcMaterias 
         Bindings        =   "frmParciales.frx":117C
         Height          =   315
         Left            =   120
         TabIndex        =   14
         Top             =   360
         Width           =   8775
         _ExtentX        =   15478
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "Nombre"
         BoundColumn     =   "Codigo"
         Text            =   ""
      End
      Begin MSAdodcLib.Adodc adoMaterias 
         Height          =   330
         Left            =   5160
         Top             =   600
         Visible         =   0   'False
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   582
         ConnectMode     =   0
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
         RecordSource    =   $"frmParciales.frx":1196
         Caption         =   "Materias"
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
      Begin VB.Label lblCodigoProfesor 
         Height          =   135
         Left            =   10680
         TabIndex        =   59
         Top             =   120
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label lblProfesor 
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
         Left            =   9960
         TabIndex        =   58
         Top             =   360
         Width           =   4575
      End
      Begin VB.Label Label20 
         Caption         =   "Profesor:"
         Height          =   255
         Left            =   9960
         TabIndex        =   57
         Top             =   120
         Width           =   735
      End
      Begin VB.Label Label4 
         Caption         =   "División:"
         Height          =   255
         Left            =   9000
         TabIndex        =   16
         Top             =   120
         Width           =   615
      End
      Begin VB.Label Label3 
         Caption         =   "Materias"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   120
         Width           =   615
      End
   End
   Begin VB.Frame frMatriculados 
      Caption         =   "Matriculados"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   8655
      Left            =   0
      TabIndex        =   10
      Top             =   1920
      Width           =   6975
      Begin VB.CommandButton cmdNotificacionPersonal 
         Caption         =   "Notificación"
         Height          =   840
         Left            =   5520
         Picture         =   "frmParciales.frx":1297
         Style           =   1  'Graphical
         TabIndex        =   90
         ToolTipText     =   "Modificar"
         Top             =   7560
         Width           =   960
      End
      Begin VB.CommandButton cmdLibres 
         Caption         =   "Libres"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   3360
         TabIndex        =   85
         Top             =   7560
         Width           =   975
      End
      Begin VB.CommandButton cmdPerdioCursada 
         Caption         =   "Perdió Cursada"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   2280
         TabIndex        =   75
         Top             =   7560
         Width           =   975
      End
      Begin VB.CommandButton cmdQuitar 
         Caption         =   "Quitar"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   1200
         TabIndex        =   52
         Top             =   7560
         Width           =   975
      End
      Begin VB.CommandButton cmdMatricular 
         Caption         =   "Matricular"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   120
         TabIndex        =   51
         Top             =   7560
         Width           =   975
      End
      Begin MSAdodcLib.Adodc adoMatriculados 
         Height          =   375
         Left            =   1200
         Top             =   2520
         Visible         =   0   'False
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   661
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
         RecordSource    =   $"frmParciales.frx":16D9
         Caption         =   "Matriculados"
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
      Begin MSDataGridLib.DataGrid dtgMatriculados 
         Bindings        =   "frmParciales.frx":1913
         Height          =   6975
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Width           =   6375
         _ExtentX        =   11245
         _ExtentY        =   12303
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
         SplitCount      =   1
         BeginProperty Split0 
            MarqueeStyle    =   4
            BeginProperty Column00 
               ColumnWidth     =   720
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   4949,858
            EndProperty
         EndProperty
      End
      Begin VB.CommandButton cmdCursadasVencidas 
         Caption         =   "Cursadas vencidas"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Index           =   1
         Left            =   4440
         TabIndex        =   77
         Top             =   7560
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label lblTotalMatriculados 
         Caption         =   "0"
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
         Left            =   1800
         TabIndex        =   68
         Top             =   7320
         Width           =   855
      End
      Begin VB.Label Label24 
         Caption         =   "Total Matriculados:"
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
         TabIndex        =   67
         Top             =   7320
         Width           =   2415
      End
   End
   Begin VB.Frame frCarrera 
      Height          =   975
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   14895
      Begin VB.ComboBox cbCurso 
         Height          =   315
         Left            =   11760
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   480
         Width           =   615
      End
      Begin VB.TextBox txtAño 
         Height          =   285
         Left            =   12480
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   480
         Width           =   495
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
         Left            =   13200
         TabIndex        =   2
         Top             =   360
         Width           =   855
      End
      Begin VB.CommandButton cmdSalir 
         Height          =   550
         Left            =   14160
         Picture         =   "frmParciales.frx":1931
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Salir"
         Top             =   330
         Width           =   550
      End
      Begin MSComCtl2.UpDown UpDown1 
         Height          =   375
         Left            =   12960
         TabIndex        =   4
         Top             =   420
         Width           =   240
         _ExtentX        =   450
         _ExtentY        =   661
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSAdodcLib.Adodc adoCarreras 
         Height          =   330
         Left            =   1920
         Top             =   120
         Visible         =   0   'False
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   582
         ConnectMode     =   0
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
         RecordSource    =   $"frmParciales.frx":1D73
         Caption         =   "Carreras"
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
      Begin MSDataListLib.DataCombo dtcCarreras 
         Bindings        =   "frmParciales.frx":1E77
         Height          =   315
         Left            =   120
         TabIndex        =   6
         Top             =   480
         Width           =   11535
         _ExtentX        =   20346
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "Nombre"
         BoundColumn     =   "Codigo"
         Text            =   ""
      End
      Begin VB.Label Label1 
         Caption         =   "Carrera:"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label2 
         Caption         =   "Año:"
         Height          =   255
         Left            =   12600
         TabIndex        =   8
         Top             =   240
         Width           =   375
      End
      Begin VB.Label lblCurso 
         Caption         =   "Curso:"
         Height          =   255
         Left            =   11760
         TabIndex        =   7
         Top             =   240
         Width           =   495
      End
   End
End
Attribute VB_Name = "frmParciales"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Conexion As New Connection
Dim Conectar As New Connection
Dim Resultado As New Recordset
Dim Promocionados As New Recordset
Dim Auxiliar As New Recordset
Dim VolverAVer As String
Dim Profesor(10) As String
Dim CodigoProfesor(10) As Integer
Dim TotalHoras As Integer
Dim HastaFecha As Date
Dim Presentes As New Recordset
Dim Ausentes As New Recordset
Dim NumeroCursada As New Recordset
Dim PorcentajeAsistencia As New Recordset
Dim CursadasVencidas As New Recordset
Dim CorrelativasVencidas As New Recordset
Dim CorrelativaPorAlumnoVencida As New Recordset
Dim RecuperanAsistencia As New Recordset
Dim Correlativas As New Recordset
Dim CorrelativasAprobadas As New Recordset
Dim CursadaNumero As Single
Dim DebeCorrelativa As String
Dim NombreCorrelativa(30) As String
Dim TotalCorrelativas As Integer
Dim TotalQueDebe As Integer
Dim PlanillaAnalitica As New Recordset
Dim TotalCursadas As New Recordset
Dim FechasAsistenciaTotal As New Recordset
Dim FechasAsistencia As New Recordset

Private Sub cbCurso_Click()
    NuevasMaterias
End Sub

Private Sub cbDivision_Click()
    If cbDivision = "" Then
        lblProfesor = ""
    Else
        lblProfesor = Profesor(cbDivision - 1)
        lblCodigoProfesor = CodigoProfesor(cbDivision - 1)
    End If
    If VolverAVer = "Si" Then VerMatriculados
End Sub

Private Sub chkCursada_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
   If chkCursada.Value = 1 Then
      If adoMatriculados.Recordset!PerdioCursada = True Then
         chkCursada.Value = 0
         Respuesta = MsgBox("Esta cursada ya se ha vencido", vbCritical, "Imposible aprobar la cursada")
         Exit Sub
       End If
   End If
End Sub

Private Sub cmdAprobarAsistenciaATodos_Click()
    Respuesta = MsgBox("¿Está seguro de aprobar la asistencia para todos los alumnos matriculados?", vbYesNo, "Confirmar")
    If Respuesta = vbNo Then Exit Sub
    Conexion.Open
    Conexion.Execute ("UPDATE Finales SET Finales.Asistencia = True, AsistenciaPorcentaje = 0 WHERE (((Finales.Materia)=" & dtcMaterias.BoundText & ") AND ((Finales.Ano)=" & txtAño & ") AND ((Finales.Division)=" & cbDivision & "))")
    Conexion.Close
    adoMatriculados.Refresh
End Sub

Private Sub cmdAsistencia_Click()
    Conexion.Open
    Conexion.Execute ("DELETE * FROM [Planilla Asistencia]")
    Conexion.Execute ("INSERT INTO [Planilla Asistencia] ( Carrera, Materia, Curso, Permiso, Alumno, Año, Profesor, Division, Establecimiento ) SELECT Carreras.Nombre AS Carrera, Materias.Nombre AS Materia, Materias.Curso AS Curso, Alumnos.Permiso AS Permiso, Alumnos.Nombre AS Alumno, Finales.Ano AS Año, Personal.Nombre AS Profesor, Finales.Division, Parametros.NombreInstitucion FROM Parametros, (Divisiones INNER JOIN Personal ON Divisiones.Profesor = Personal.Codigo) INNER JOIN (((Alumnos INNER JOIN Finales ON Alumnos.Permiso = Finales.Alumno) INNER JOIN Materias ON Finales.Materia = Materias.Codigo) INNER JOIN Carreras ON Materias.Carrera = Carreras.Codigo) ON (Divisiones.Ano = Finales.Ano) AND (Divisiones.Materia = Finales.Materia) Where (((Finales.Ano) = " & txtAño & ") And ((Finales.Division) = " & cbDivision & ") And ((Materias.Codigo) = " & dtcMaterias.BoundText & ") And ((Divisiones.Division) = " & cbDivision & ")) AND Finales.PerdioCursada = 0 AND Finales.Libre = 0 ORDER BY Alumnos.Nombre")
    Conexion.Close
    rptAsistencia.PrintReport
End Sub

Private Sub cmdAsistenciaCompleta_Click()
    Conexion.Open
    Conexion.Execute ("DELETE * FROM [PlanillaAsistenciaCompletada]")
    Conexion.Execute ("INSERT INTO [PlanillaAsistenciaCompletada] ( Carrera, Materia, Curso, Permiso, Alumno, Año, Profesor, Division, Establecimiento ) SELECT Carreras.Nombre AS Carrera, Materias.Nombre AS Materia, Materias.Curso AS Curso, Alumnos.Permiso AS Permiso, Alumnos.Nombre AS Alumno, Finales.Ano AS Año, Personal.Nombre AS Profesor, Finales.Division, Parametros.NombreInstitucion FROM Parametros, (Divisiones INNER JOIN Personal ON Divisiones.Profesor = Personal.Codigo) INNER JOIN (((Alumnos INNER JOIN Finales ON Alumnos.Permiso = Finales.Alumno) INNER JOIN Materias ON Finales.Materia = Materias.Codigo) INNER JOIN Carreras ON Materias.Carrera = Carreras.Codigo) ON (Divisiones.Ano = Finales.Ano) AND (Divisiones.Materia = Finales.Materia) Where (((Finales.Ano) = " & txtAño & ") And ((Finales.Division) = " & cbDivision & ") And ((Materias.Codigo) = " & dtcMaterias.BoundText & ") And ((Divisiones.Division) = " & cbDivision & ")) AND Finales.PerdioCursada = 0 AND Finales.Libre = 0 ORDER BY Alumnos.Nombre")
    Set NumeroCursada = Conexion.Execute("SELECT Divisiones.Numero From Divisiones WHERE Divisiones.Materia=" & frmParciales.dtcMaterias.BoundText & " AND Divisiones.Ano=" & frmParciales.txtAño & " AND Divisiones.Division=" & frmParciales.cbDivision)
    CursadaNumero = NumeroCursada!Numero
    'Cuento cuantas horas de asistencia se dieron y declaro array
    Set FechasAsistenciaTotal = Conexion.Execute("SELECT count(Asistencias.Fecha) as Total From Asistencias Where (((Asistencias.Numero) = " & CursadaNumero & ") And ((Asistencias.Agente) = " & adoMatriculados.Recordset!Permiso & "))")
    Dim totalfechas As Integer
    totalfechas = FechasAsistenciaTotal!Total
    Dim ListaFechas()
    ReDim ListaFechas(totalfechas)
    
    'traigo la fecha de cada hora y la actualizo en el reporte para usar de encabezado
    'y guardo cada fecha en la lista para ir a buscar cada presente
    i = 1
    Set FechasAsistencia = Conexion.Execute("Select Fecha From Asistencias Where (((Asistencias.Numero) = " & CursadaNumero & ") And ((Asistencias.Agente) = " & adoMatriculados.Recordset!Permiso & "))")
    While FechasAsistencia.EOF = False
        ListaFechas(i - 1) = FechasAsistencia!Fecha
        Conexion.Execute ("update [PlanillaAsistenciaCompletada] set D" & i & "= '" & Mid(FechasAsistencia!Fecha, 1, 2) & " " & Mid(FechasAsistencia!Fecha, 4, 2) & "'")
        FechasAsistencia.MoveNext
        i = i + 1
    Wend
    
    'recorro los alumnos que forman el reporte
    Dim Presentes As New Recordset
    Dim alumnosAsistencia As New Recordset
    Set alumnosAsistencia = Conexion.Execute("SELECT distinct PlanillaAsistenciaCompletada.Permiso FROM PlanillaAsistenciaCompletada")
    While alumnosAsistencia.EOF = False
        i = 1
        Set Presentes = Conexion.Execute("SELECT Fecha, Presente From Asistencias Where Numero = " & CursadaNumero & " AND Agente=" & alumnosAsistencia!Permiso & " ORDER BY Asistencias.Agente, Asistencias.Fecha")
        While Presentes.EOF = False
            p = "P"
            If Presentes!Presente = False Then
                p = "A"
            End If
            Conexion.Execute ("update [PlanillaAsistenciaCompletada] set " & i & "= '" & p & "' WHERE Permiso=" & alumnosAsistencia!Permiso)
            Presentes.MoveNext
            i = i + 1
        Wend
        alumnosAsistencia.MoveNext
    Wend
    
    
    
    Conexion.Close
    rptAsistenciaCompletada.PrintReport
End Sub

Private Sub cmdCancelar_Click()
    cmdModificar.Enabled = True
    cmdGuardar.Enabled = False
    cmdCancelar.Enabled = False
    frCarrera.Enabled = True
    frMaterias.Enabled = True
    frMatriculados.Enabled = True
    frPlanillas.Enabled = True
    cmdIngresarNota.Enabled = True
    frDetalles.Enabled = False
End Sub

Private Sub cmdCursadasVencidas_Click(Index As Integer)
   TotalVencidas = 0
   Respuesta = InputBox("Ingreso desde el año que se vencieron", "Año de vencimiento")
   Conexion.Open
   'levanto las cursadas que estan para marcar como vencidas
   Set CursadasVencidas = Conexion.Execute("SELECT Alumnos.Permiso,Finales.Materia, Alumnos.Nombre, Carreras.Resolucion, Carreras.Nombre, Materias.Nombre, Materias.Curso, Finales.Ano, Finales.Cursada, finales.Aprobada FROM ((Finales INNER JOIN Alumnos ON Finales.Alumno = Alumnos.Permiso) INNER JOIN Materias ON Finales.Materia = Materias.Codigo) INNER JOIN Carreras ON Materias.Carrera = Carreras.Codigo Where (((Finales.Ano) <= " & Respuesta & ") And ((Finales.Cursada) = True) And ((Finales.Aprobada) = False)  AND ((Carreras.Codigo)=63) ) ORDER BY Carreras.Nombre, Materias.Curso, Materias.Nombre, Alumnos.Nombre")
   CursadasVencidas.MoveFirst
   While CursadasVencidas.EOF = False
      Conexion.Execute ("UPDATE Finales SET Finales.PerdioCursada = True, Finales.Cursada = False WHERE (((Finales.Alumno)=" & CursadasVencidas!Permiso & ") AND ((Finales.Materia)=" & CursadasVencidas!Materia & ") AND ((Finales.Ano)=" & Respuesta & "))")
      Conexion.Execute ("INSERT INTO Cursadas_Vencidas ( Permiso, Materia, Ano_Cursada, Fecha_Proceso, Detalle ) values(" & CursadasVencidas!Permiso & "," & CursadasVencidas!Materia & "," & Respuesta & ",#" & DateValue(Date) & "#,1 )")
      Conexion.Execute ("insert into Mensajes_alumnos (Permiso, Fecha,Asunto,Detalle) values (" & CursadasVencidas!Permiso & ",#" & DateValue(Date) & "#,'Cursada Vencida','Se vencio la cursada de la materia " & CursadasVencidas!Nombre & "')")
      'busco las materias correlativas en cadena
      Set CorrelativasVencidas = Conexion.Execute("SELECT Correlativas.Principal From Correlativas WHERE Correlativas.Correlativa=" & CursadasVencidas!Materia)
      While CorrelativasVencidas.EOF = False
         'me fijo cada correlativa si el alumno la tiene aprobada
         Set CorrelativaPorAlumnoVencida = Conexion.Execute("SELECT Finales.Ano, Materias.Nombre FROM Finales INNER JOIN Materias ON Finales.Materia = Materias.Codigo WHERE Finales.Alumno=" & CursadasVencidas!Permiso & " AND Finales.Materia=" & CorrelativasVencidas!Principal & " AND Finales.Cursada=True AND Finales.Aprobada=False")
         While CorrelativaPorAlumnoVencida.EOF = False
            Conexion.Execute ("UPDATE Finales SET Finales.PerdioCursada = True, Finales.Cursada = False WHERE (((Finales.Alumno)=" & CursadasVencidas!Permiso & ") AND ((Finales.Materia)=" & CorrelativasVencidas!Principal & ") AND ((Finales.Ano)=" & CorrelativaPorAlumnoVencida!Ano & "))")
            Conexion.Execute ("INSERT INTO Cursadas_Vencidas ( Permiso, Materia, Ano_Cursada, Fecha_Proceso, Detalle ) values(" & CursadasVencidas!Permiso & "," & CorrelativasVencidas!Principal & "," & CorrelativaPorAlumnoVencida!Ano & ",#" & DateValue(Date) & "#,2 )")
            Conexion.Execute ("insert into Mensajes_alumnos (Permiso, Fecha,Asunto,Detalle) values (" & CursadasVencidas!Permiso & ",#" & DateValue(Date) & "#,'Cursada Vencida','Se vencio la cursada de la materia " & CorrelativaPorAlumnoVencida!Nombre & "')")
            CorrelativaPorAlumnoVencida.MoveNext
         Wend
         CorrelativasVencidas.MoveNext
      Wend
      CursadasVencidas.MoveNext
      TotalVencidas = TotalVencidas + 1
   Wend
   Conexion.Close
   MsgBox ("Se vencieron y actualizaron " & TotalVencidas & " cursadas")
End Sub

Private Sub cmdGuardar_Click()
    If chkPromocion.Value = 1 And chkCursada.Value = 0 Then
       MsgBox ("No puede promocinar sin aprobar la cursada")
       Exit Sub
    End If
    
    cmdModificar.Enabled = True
    cmdGuardar.Enabled = False
    cmdCancelar.Enabled = False
    frCarrera.Enabled = True
    frMaterias.Enabled = True
    frMatriculados.Enabled = True
    frPlanillas.Enabled = True
    cmdIngresarNota.Enabled = True
    frDetalles.Enabled = False
    GuardarDatos
    dtgMatriculados.SetFocus
End Sub

Private Sub cmdImprimeActaPromocion_Click()
    Respuesta = MsgBox("Está seguro de generar e imprimir las actas de los alumnos que promocionaron", vbYesNo, "Imprimir Actas")
    If Respuesta = vbNo Then Exit Sub
    Conexion.Open
    Set Resultado = Conexion.Execute("SELECT Numero From Mesas WHERE (((Mesas.Materia)=" & dtcMaterias.BoundText & ") AND ((Mesas.Division)=" & cbDivision & ") AND ((Mesas.Turno)=13) AND ((Mesas.Ano)=" & txtAño & "))")
    If Resultado.EOF = False Then Respuesta = MsgBox("El acta para esta materia ya fue impresa", vbOKOnly, "Imposible Imprimir"): Conexion.Close: Exit Sub
    Set Resultado = Conexion.Execute("SELECT NombreInstitucion FROM Parametros")
    NombreInstitucion = Resultado!NombreInstitucion
    Set Promocionados = Conexion.Execute("SELECT Alumnos.Permiso, Alumnos.Nombre, Finales.Cursada, Finales.Asistencia, Finales.Promocion FROM (((Finales INNER JOIN Alumnos ON Finales.Alumno = Alumnos.Permiso) INNER JOIN Materias ON Finales.Materia = Materias.Codigo) INNER JOIN CarrerasHechas ON (Materias.Carrera = CarrerasHechas.Carrera) AND (Alumnos.Permiso = CarrerasHechas.Permiso)) INNER JOIN Condicion ON CarrerasHechas.Condición = Condicion.Codigo Where (((Finales.Promocion) = True) And ((Finales.Materia) = " & dtcMaterias.BoundText & ") And ((Finales.Ano) = " & txtAño & ") And ((Finales.Division) = " & cbDivision & ")) ORDER BY Alumnos.Nombre")
    If Promocionados.EOF = True Then MsgBox ("Ningún alumno ha promocionado esta materia"): Conexion.Close: Exit Sub
    Conexion.Execute ("INSERT INTO Mesas ( Materia, Division, Turno, Ano, Fecha, Hora, Lugar, Titular, TipoDeMesa ) VALUES (" & dtcMaterias.BoundText & "," & cbDivision & ",13," & txtAño & ",'" & DateValue(Date) & "','" & TimeValue(Time) & "','Acta de Promoción sin Examen'," & lblCodigoProfesor & ", 2 )")
    Set Resultado = Conexion.Execute("SELECT Numero From Mesas WHERE (((Mesas.Materia)=" & dtcMaterias.BoundText & ") AND ((Mesas.Division)=" & cbDivision & ") AND ((Mesas.Turno)=13) AND ((Mesas.Ano)=" & txtAño & "))")
    NumeroDeMesa = Resultado!Numero
    'simulo las inscripciones de los alumnos
    While Promocionados.EOF = False
        Conexion.Execute ("INSERT INTO Inscripciones ( Mesa, Alumno, Cursada ) VALUES (" & NumeroDeMesa & "," & Promocionados!Permiso & "," & txtAño & ")")
        Promocionados.MoveNext
    Wend
    'genero los numeros de acta
    Promocionados.MoveFirst
    OrdenActa = 1
    TotalActas = 1
    While Promocionados.EOF = False
        Conexion.Execute ("UPDATE Inscripciones SET Inscripciones.Acta = " & TotalActas & " WHERE Inscripciones.Mesa=" & NumeroDeMesa & " AND Inscripciones.Alumno=" & Promocionados!Permiso)
        OrdenActa = OrdenActa + 1
        Promocionados.MoveNext
        If OrdenActa = 26 And Promocionados.EOF = False Then
            TotalActas = TotalActas + 1
            OrdenActa = 1
        End If
    Wend
    'imprimo las actas
    For i = 0 To TotalActas - 1
        Set Promocionados = Conexion.Execute("SELECT Alumnos.Permiso, Alumnos.Nombre, Inscripciones.Cursada, ([Alumnos].[Tipo] & ' ' & [Alumnos].[Documento]) AS Documento FROM Inscripciones INNER JOIN Alumnos ON Inscripciones.Alumno = Alumnos.Permiso Where (((Inscripciones.FechaBorrado) Is Null) And ((Inscripciones.Mesa) = " & NumeroDeMesa & ") And ((Inscripciones.Acta) = " & i + 1 & "))ORDER BY Alumnos.Nombre")
        Set Auxiliar = Conexion.Execute("SELECT Count(Alumno) AS TotalAlumnos FROM Inscripciones Where Mesa = " & NumeroDeMesa & " And Acta = " & i + 1)
        MsgBox ("Imprimir acta Nº " & i + 1)
        With frmImprimeActas
        .lblEstablecimiento = NombreInstitucion
        .lblTituloDeActa = "Acta de Promoción sin Examen"
        .lblCarrera = dtcCarreras.Text
        .lblMateria = dtcMaterias.BoundText & " " & dtcMaterias.Text
        .lblMesa = NumeroDeMesa
        .lblActa = i + 1
        .lblFecha = Format(Date, "dd/mm/yyyy")
        .lblHora = ""
        .lblLugar = "Sede"
        .lblCurso = cbCurso.Text
        .lblCursada = txtAño
        .lblTitular = lblProfesor
        .lblIntegrante1 = "Directivo"
        .lblIntegrante2 = "Secretario"
        .lblTotalAlumnos = Auxiliar!TotalAlumnos
        For j = 1 To Auxiliar!TotalAlumnos
            .lblPermiso(j) = Promocionados!Permiso
            .lblAlumno(j) = Promocionados!Nombre
            .lblDocumento(j) = Format(Promocionados!documento, "##,###,###")
            .lblOrden(j).Visible = True
            .lblPermiso(j).Visible = True
            .lblEscritoNota(j).Visible = True
            .lnlEscritoLetras(j).Visible = True
            .lblOralNota(j).Visible = True
            .lnlOralLetras(j).Visible = True
            .lblFinalNota(j).Visible = True
            .lnlFinalLetras(j).Visible = True
            .lblAlumno(j).Visible = True
            .lblDocumento(j).Visible = True
            Promocionados.MoveNext
        Next j
        End With
        frmImprimeActas.PrintForm
        Unload frmImprimeActas
         'Agrego en la tabla Actas el acta con los datos correspondientes
        Conexion.Execute ("INSERT INTO Actas ( Mesa, Acta, Total, Ano, Division ) VALUES (" & NumeroDeMesa & ", " & i + 1 & ", " & Auxiliar!TotalAlumnos & ", " & txtAño & ", " & cbDivision & ")")
    Next i
    Conexion.Execute ("UPDATE Mesas SET Mesas.Impresas = True, Mesas.Actas = " & i & " WHERE Mesas.Numero = " & NumeroDeMesa)
    Conexion.Close
End Sub

Private Sub cmdIngresarAsistencia_Click()
    frmIngresarAsistencia.Show 1
End Sub

Private Sub cmdIngresaRecuAsistencia_Click()
If chkAsistencia = 1 Then MsgBox ("El alumno ya a aprobado la asistencia"): Exit Sub
'controlo que solo se cargue una vez el recuperatorio de asistencia
 Conectar.Open
 Set NumeroCursada = Conectar.Execute("SELECT Divisiones.Numero From Divisiones WHERE Divisiones.Materia=" & frmParciales.dtcMaterias.BoundText & " AND Divisiones.Ano=" & frmParciales.txtAño & " AND Divisiones.Division=" & frmParciales.cbDivision)
 CursadaNumero = NumeroCursada!Numero
 Set Resultado = Conectar.Execute("SELECT Numero FROM Recuperatorios_de_Asistencia WHERE Permiso = " & adoMatriculados.Recordset!Permiso & " AND Numero = " & CursadaNumero)
 If Resultado.EOF = True Then
    frmIngresarRecAsistencia.Show 1
 Else
    MsgBox ("Ya se ha ingresado el Recuperatorio de Asistencia para este Alumno")
 End If
 Conectar.Close
End Sub

Private Sub cmdIngresarNota_Click()
    If Val(txtNota) > 10 Then MsgBox ("Nota Incorrecta"): txtNota = "": Exit Sub
    Conexion.Open
    Select Case lblDescripcionDeNota
    Case "1º Parcial": Conexion.Execute ("UPDATE Finales SET Finales.Parcial1 = " & Val(txtNota) & " WHERE Finales.Alumno=" & adoMatriculados.Recordset!Permiso & " AND Finales.Materia=" & dtcMaterias.BoundText & " AND Finales.Ano=" & txtAño & " AND Finales.Division=" & cbDivision)
    Case "1º Recuperatorio": Conexion.Execute ("UPDATE Finales SET Finales.Recuperatorio1 = " & Val(txtNota) & " WHERE Finales.Alumno=" & adoMatriculados.Recordset!Permiso & " AND Finales.Materia=" & dtcMaterias.BoundText & " AND Finales.Ano=" & txtAño & " AND Finales.Division=" & cbDivision)
    Case "2º Parcial": Conexion.Execute ("UPDATE Finales SET Finales.Parcial2 = " & Val(txtNota) & " WHERE Finales.Alumno=" & adoMatriculados.Recordset!Permiso & " AND Finales.Materia=" & dtcMaterias.BoundText & " AND Finales.Ano=" & txtAño & " AND Finales.Division=" & cbDivision)
    Case "2º Recuperatorio": Conexion.Execute ("UPDATE Finales SET Finales.Recuperatorio2 = " & Val(txtNota) & " WHERE Finales.Alumno=" & adoMatriculados.Recordset!Permiso & " AND Finales.Materia=" & dtcMaterias.BoundText & " AND Finales.Ano=" & txtAño & " AND Finales.Division=" & cbDivision)
    Case "Trabajo Práctico Nº 1": Conexion.Execute ("UPDATE Finales SET Finales.Practico1 = " & Val(txtNota) & " WHERE Finales.Alumno=" & adoMatriculados.Recordset!Permiso & " AND Finales.Materia=" & dtcMaterias.BoundText & " AND Finales.Ano=" & txtAño & " AND Finales.Division=" & cbDivision)
    Case "Trabajo Práctico Nº 2": Conexion.Execute ("UPDATE Finales SET Finales.Practico2 = " & Val(txtNota) & " WHERE Finales.Alumno=" & adoMatriculados.Recordset!Permiso & " AND Finales.Materia=" & dtcMaterias.BoundText & " AND Finales.Ano=" & txtAño & " AND Finales.Division=" & cbDivision)
    Case "Trabajo Práctico Nº 3": Conexion.Execute ("UPDATE Finales SET Finales.Practico3 = " & Val(txtNota) & " WHERE Finales.Alumno=" & adoMatriculados.Recordset!Permiso & " AND Finales.Materia=" & dtcMaterias.BoundText & " AND Finales.Ano=" & txtAño & " AND Finales.Division=" & cbDivision)
    Case "Trabajo Práctico Nº 4": Conexion.Execute ("UPDATE Finales SET Finales.Practico4 = " & Val(txtNota) & " WHERE Finales.Alumno=" & adoMatriculados.Recordset!Permiso & " AND Finales.Materia=" & dtcMaterias.BoundText & " AND Finales.Ano=" & txtAño & " AND Finales.Division=" & cbDivision)
    Case "Trabajo Práctico Nº 5": Conexion.Execute ("UPDATE Finales SET Finales.Practico5 = " & Val(txtNota) & " WHERE Finales.Alumno=" & adoMatriculados.Recordset!Permiso & " AND Finales.Materia=" & dtcMaterias.BoundText & " AND Finales.Ano=" & txtAño & " AND Finales.Division=" & cbDivision)
    End Select
    Conexion.Close
    adoMatriculados.Recordset.MoveNext
    If adoMatriculados.Recordset.EOF = True Then
        frIngresoDeNotas.Visible = False
        dtcMaterias_Change
    Else 'ingreso la nota
        lblAlumnoNota = adoMatriculados.Recordset!Nombre
        txtNota = ""
    End If
End Sub

Private Sub cmdIngresarNotas_Click()
    frmIngresoDeParciales.Show 1
    If frIngresoDeNotas.Visible = True Then '  se decidio realizar el ingreso de notas
        lblAlumnoNota = adoMatriculados.Recordset!Nombre
        txtNota = ""
        txtNota.SetFocus
    End If
End Sub

Private Sub cmdLibres_Click()
    frmAlumnosLibres.Show 1
End Sub

Private Sub cmdMatricular_Click()
    frmMatricular.Show 1
End Sub

Private Sub cmdModificar_Click()
    cmdModificar.Enabled = False
    cmdGuardar.Enabled = True
    cmdCancelar.Enabled = True
    frCarrera.Enabled = False
    frMaterias.Enabled = False
    frMatriculados.Enabled = False
    frPlanillas.Enabled = False
    cmdIngresarNota.Enabled = False
    frDetalles.Enabled = True
End Sub

Private Sub cmdMostrar_Click()
    cmdMostrar.Enabled = False
    adoMaterias.RecordSource = "SELECT Materias.Codigo, Materias.Nombre FROM Divisiones INNER JOIN Materias ON Divisiones.Materia = Materias.Codigo Where Materias.Curso = " & cbCurso.Text & " And Divisiones.Ano = " & txtAño & " And Materias.Carrera = " & dtcCarreras.BoundText & " And Divisiones.Division = 1 ORDER BY Materias.Curso"
    adoMaterias.Refresh
    If adoMaterias.Recordset.RecordCount > 0 Then
        dtcMaterias.BoundText = adoMaterias.Recordset!Codigo
        frMaterias.Enabled = True
        frMatriculados.Enabled = True
    Else
        frMaterias.Enabled = False
        frMatriculados.Enabled = False
        MsgBox ("No se creó ninguna división")
    End If
End Sub

Private Sub cmdNotificacionPersonal_Click()
    frmAgregarNotificacion.lblTipoDeNotificacion = "Por Materia"
    frmAgregarNotificacion.lblMateria = dtcMaterias.BoundText
    frmAgregarNotificacion.Show 1
End Sub

Private Sub cmdNroCursada_Click()
   Conectar.Open
   Set NumeroCursada = Conectar.Execute("SELECT Divisiones.Numero From Divisiones WHERE Divisiones.Materia=" & frmParciales.dtcMaterias.BoundText & " AND Divisiones.Ano=" & frmParciales.txtAño & " AND Divisiones.Division=" & frmParciales.cbDivision)
   CursadaNumero = NumeroCursada!Numero
   MsgBox (NumeroCursada!Numero)
   Conectar.Close
End Sub

Private Sub cmdPara_desdoblamiento_Click()
    Recursantes_libres = 0
    Me.MousePointer = 11
    Conexion.Open
    Conexion.Execute ("DELETE * From rpt_Planilla_analitica")
    Conexion.Execute ("INSERT INTO rpt_Planilla_analitica ( Alumno, Nombre, Documento, Ingreso, Libre, Division ) SELECT Finales.Alumno, Alumnos.Nombre, Alumnos.Documento, CarrerasHechas.Ingreso, Finales.Libre, Finales.Division FROM (((Finales INNER JOIN Materias ON Finales.Materia = Materias.Codigo) INNER JOIN Carreras ON Materias.Carrera = Carreras.Codigo) INNER JOIN CarrerasHechas ON (Materias.Carrera = CarrerasHechas.Carrera) AND (Finales.Alumno = CarrerasHechas.Permiso)) INNER JOIN Alumnos ON Finales.Alumno = Alumnos.Permiso WHERE Finales.Materia=" & dtcMaterias.BoundText & " AND Finales.Ano=" & txtAño & " AND Finales.Division = " & cbDivision)
    Conexion.Close
    Conectar.Open
    PlanillaAnalitica.Open "SELECT * from rpt_Planilla_analitica", Conectar, adOpenDynamic, adLockPessimistic
    While PlanillaAnalitica.EOF = False
    Conexion.Open
     TotalCursadas.Open "SELECT Count(Alumno) As total FROM Finales WHERE Alumno=" & PlanillaAnalitica!Alumno & " AND Materia=" & dtcMaterias.BoundText, Conexion, adOpenDynamic, adLockPessimistic
     If TotalCursadas!Total > 1 Then
        PlanillaAnalitica!Recursante = 1
     Else
        If txtAño - PlanillaAnalitica!Ingreso + 1 > cbCurso Then
           PlanillaAnalitica!Atraso_academico = 1
        Else
           PlanillaAnalitica!Cohorte = 1
        End If
     End If
     Conexion.Close
     If PlanillaAnalitica!Libre = True Then
        PlanillaAnalitica!Cursada_libre = 1
     Else
        PlanillaAnalitica!Presencial = 1
     End If
     If PlanillaAnalitica!Libre = True And PlanillaAnalitica!Recursante = 0 Then
        Recursantes_libres = Recursantes_libres + 1
     End If
     PlanillaAnalitica.Update
     PlanillaAnalitica.MoveNext
    Wend
    Conectar.Execute ("UPDATE rpt_Planilla_analitica, Parametros SET rpt_Planilla_analitica.Carrera = '" & dtcCarreras & "', rpt_Planilla_analitica.Curso = " & cbCurso & ", rpt_Planilla_analitica.Codigo = " & dtcMaterias.BoundText & ", rpt_Planilla_analitica.Materia = '" & dtcMaterias & "', rpt_Planilla_analitica.Profesor = '" & lblProfesor & "', rpt_Planilla_analitica.Institutcion = parametros.nombreinstitucion, rpt_Planilla_analitica.Recursantes_libres=" & Recursantes_libres)
    Conectar.Close
    Me.MousePointer = 0
    rptPlanillaAnalitica.PrintReport
 End Sub

Private Sub cmdParcialeB_Click()
    cn.Open
    codigo_materia = dtcMaterias.BoundText
    division = cbDivision
    Anio = txtAño

    rptPlanillaParcialesB.WindowState = 2
    rptPlanillaParcialesB.Show 1
    Unload rptPlanillaParcialesB
    cn.Close
End Sub

Private Sub cmdParciales_Click()
    Conexion.Open
    Conexion.Execute ("DELETE * FROM [Planilla Asistencia]")
    Conexion.Execute ("INSERT INTO [Planilla Asistencia] ( Carrera, Materia, Curso, Permiso, Alumno, Año, Profesor, Division, Establecimiento ) SELECT Carreras.Nombre AS Carrera, Materias.Nombre AS Materia, Materias.Curso AS Curso, Alumnos.Permiso AS Permiso, Alumnos.Nombre AS Alumno, Finales.Ano AS Año, Personal.Nombre AS Profesor, Finales.Division, Parametros.NombreInstitucion FROM Parametros, (Divisiones INNER JOIN Personal ON Divisiones.Profesor = Personal.Codigo) INNER JOIN (((Alumnos INNER JOIN Finales ON Alumnos.Permiso = Finales.Alumno) INNER JOIN Materias ON Finales.Materia = Materias.Codigo) INNER JOIN Carreras ON Materias.Carrera = Carreras.Codigo) ON (Divisiones.Ano = Finales.Ano) AND (Divisiones.Materia = Finales.Materia) Where (((Finales.Ano) = " & txtAño & ") And ((Finales.Division) = " & cbDivision & ") And ((Materias.Codigo) = " & dtcMaterias.BoundText & ") And ((Divisiones.Division) = " & cbDivision & ")) AND Finales.PerdioCursada = 0 AND Finales.Libre=0 ORDER BY Alumnos.Nombre")
    Conexion.Close
    rptParciales.PrintReport
End Sub

Private Sub cmdPasarALibre_Click()
    Respuesta = MsgBox("Va a pasar al alumno " & adoMatriculados.Recordset!Nombre & " a condiciòn de libre. ¿Continúa?", vbYesNo, "Está seguro")
    If Respuesta = vbYes Then
        Conexion.Open
        Conexion.Execute ("UPDATE Finales SET Libre = True WHERE Finales.Alumno=" & adoMatriculados.Recordset!Permiso & " AND Finales.Materia=" & dtcMaterias.BoundText & " AND Finales.Ano=" & txtAño)
        Conexion.Close
        VerMatriculados
    End If
End Sub

Private Sub cmdPerdioCursada_Click()
    Respuesta = MsgBox("El alumno " & adoMatriculados.Recordset!Nombre & " perdió la cursada?", vbYesNo, "Está seguro")
    If Respuesta = vbYes Then
       Conexion.Open
       Conexion.Execute ("UPDATE Finales SET PerdioCursada = True AND Cursada = False WHERE Alumno = " & adoMatriculados.Recordset!Permiso & " AND Materia = " & dtcMaterias.BoundText & " AND Ano = " & txtAño)
       Conexion.Close
       VerMatriculados
    End If
End Sub

Private Sub cmdPlanillaParcualesCompleta_Click()
    cn.Open
    codigo_materia = dtcMaterias.BoundText
    division = cbDivision
    Anio = txtAño

    rptPlanillaCursadaCompleta.WindowState = 2
    rptPlanillaCursadaCompleta.Show 1
    Unload rptPlanillaCursadaCompleta
    cn.Close
End Sub

Private Sub cmdQuitar_Click()
    Respuesta = MsgBox("Esta seguro de borrar al alumno de la cursada de: " & Chr(13) & dtcMaterias.Text, vbYesNo, "Borrar cursada")
    If Respuesta = vbYes Then
        Me.MousePointer = 11
        Conexion.Open
        Conexion.Execute ("DELETE * FROM Finales WHERE Alumno = " & adoMatriculados.Recordset!Permiso & " AND Materia = " & dtcMaterias.BoundText & " AND Ano = " & txtAño)
        Respuesta = MsgBox("¿Desea eliminar todas las asistencias ingresadas para este alumno en esta materia?", vbYesNo, "Borrar Asistencias")
        If Respuesta = vbYes Then
            Conexion.Execute ("DELETE * From Asistencias WHERE (((Numero)=" & CursadaNumero & ") AND ((Asistencias.Agente)=" & adoMatriculados.Recordset!Permiso & "))")
        End If
        Conexion.Close
        dtcMaterias_Change
        Me.MousePointer = 0
    End If
End Sub

Private Sub cmdRecuperatorioAsistencia_Click()
Conexion.Open
Conexion.Execute ("DELETE * FROM rptRecuperatorioDeAsistencia")
Conexion.Execute ("INSERT INTO rptRecuperatorioDeAsistencia ( Carrera, Curso, Materia, Division, Permiso, Documento, Alumno, Ano, Cursada, AsistenciaPorcentaje ) SELECT Carreras.Nombre, Materias.Curso, Materias.Nombre, Finales.Division, Alumnos.Permiso, Alumnos.Documento, Alumnos.Nombre, Finales.Ano, Finales.Cursada, Finales.AsistenciaPorcentaje FROM ((Finales INNER JOIN Alumnos ON Finales.Alumno = Alumnos.Permiso) INNER JOIN Materias ON Finales.Materia = Materias.Codigo) INNER JOIN Carreras ON Materias.Carrera = Carreras.Codigo Where (((Finales.Division) = " & cbDivision & ") And ((Finales.Ano) = " & txtAño & ") And ((Materias.Codigo) = " & adoMaterias.Recordset!Codigo & ") And ((Finales.Asistencia) = False)) AND Finales.Libre=0 ORDER BY Alumnos.Nombre")
Conexion.Execute ("Update rptRecuperatorioDeAsistencia set AproboCursada='Si' where Cursada = 1")
Conexion.Execute ("Update rptRecuperatorioDeAsistencia set AproboCursada='No' where Cursada = 0")
Conexion.Execute ("Update rptRecuperatorioDeAsistencia set Calificacion='Recursa', Aprobada='Recursa' where AsistenciaPorcentaje < 50")
Conexion.Execute ("Update rptRecuperatorioDeAsistencia set Profesor='" & lblProfesor & "'")
Conexion.Close
rptRecuperatorioAsistencia.PrintReport
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub cmdTodasLasVencidas_Click(Index As Integer)
   
End Sub

Private Sub cmdVerAsistencia_Click()
    frmVerAsistencia.Show 1
End Sub

Private Sub dtcCarreras_Change()
    NuevasMaterias
    adoCarreras.Recordset.MoveFirst
    adoCarreras.Recordset.Find ("Codigo=" & dtcCarreras.BoundText)
    lblCurso = adoCarreras.Recordset!Medida
    cbCurso.Clear
    For i = 0 To adoCarreras.Recordset!Años - 1
        cbCurso.List(i) = i + 1
    Next i
    cbCurso.Text = cbCurso.List(0)
End Sub

Private Sub dtcMaterias_Change()
    Conexion.Open
    Set Resultado = Conexion.Execute("SELECT Divisiones.Numero, Divisiones.Division, Personal.Nombre, Personal.Codigo FROM Divisiones INNER JOIN Personal ON Divisiones.Profesor = Personal.Codigo WHERE Divisiones.Materia=" & dtcMaterias.BoundText & " AND Divisiones.Ano=" & txtAño & " Order By Divisiones.Division")
    CursadaNumero = Resultado!Numero
    i = 0
    cbDivision.Clear
    While Resultado.EOF = False
        cbDivision.List(i) = i + 1 'Resultado!Division
        Profesor(i) = Resultado!Nombre
        CodigoProfesor(i) = Resultado!Codigo
        i = i + 1
        Resultado.MoveNext
    Wend
    Conexion.Close
    VolverAVer = "No" 'para no llamar a la funcion VerMatriculados Cuando cambie el cbDivision
    cbDivision.Text = cbCurso.List(0)
    VolverAVer = "Si"
    adoMaterias.Recordset.MoveFirst
    adoMaterias.Recordset.Find ("Codigo='" & dtcMaterias.BoundText & "'")
    VerMatriculados
    MostrarDatos
End Sub

Private Sub dtgMatriculados_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    MostrarDatos
End Sub

Private Sub Form_Load()
    Conexion.ConnectionString = ("DSN=Instituto")
    Conectar.ConnectionString = ("DSN=Instituto")
    dtcCarreras.BoundText = adoCarreras.Recordset!Codigo
    txtAño = Format(Date, "yyyy")
End Sub

Private Sub txtAño_Change()
    NuevasMaterias
End Sub

Private Sub txtNota_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cmdIngresarNota_Click
    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 And KeyAscii <> 46 Then KeyAscii = 0
End Sub

Private Sub txtParcial1_KeyPress(KeyAscii As Integer)
    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 And KeyAscii <> 46 Then KeyAscii = 0
End Sub

Private Sub txtParcial2_KeyPress(KeyAscii As Integer)
    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 And KeyAscii <> 46 Then KeyAscii = 0
End Sub

Private Sub txtPractico1_KeyPress(KeyAscii As Integer)
    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 And KeyAscii <> 46 Then KeyAscii = 0
End Sub

Private Sub txtPractico2_KeyPress(KeyAscii As Integer)
    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 And KeyAscii <> 46 Then KeyAscii = 0
End Sub

Private Sub txtPractico3_KeyPress(KeyAscii As Integer)
    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 And KeyAscii <> 46 Then KeyAscii = 0
End Sub

Private Sub txtPractico4_KeyPress(KeyAscii As Integer)
    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 And KeyAscii <> 46 Then KeyAscii = 0
End Sub

Private Sub txtPractico5_KeyPress(KeyAscii As Integer)
    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 And KeyAscii <> 46 Then KeyAscii = 0
End Sub

Private Sub txtRecuperatorio1_KeyPress(KeyAscii As Integer)
    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 And KeyAscii <> 46 Then KeyAscii = 0
End Sub

Private Sub txtRecuperatorio2_KeyPress(KeyAscii As Integer)
    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 And KeyAscii <> 46 Then KeyAscii = 0
End Sub

Private Sub txtTotalizador_KeyPress(KeyAscii As Integer)
    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 And KeyAscii <> 46 Then KeyAscii = 0
End Sub

Private Sub UpDown1_DownClick()
    txtAño = Val(txtAño) - 1
End Sub

Private Sub UpDown1_UpClick()
    txtAño = Val(txtAño) + 1
End Sub

Private Function MostrarDatos()
    With adoMatriculados.Recordset
    If .RecordCount > 0 And .EOF = False Then
        lblAlumno = adoMatriculados.Recordset!Nombre
        lblSituacion = adoMatriculados.Recordset!Condicion
        txtParcial1 = Format(!Parcial1, "0.00")
        txtParcial2 = Format(!Parcial2, "0.00")
        txtRecuperatorio1 = Format(!Recuperatorio1, "0.00")
        txtRecuperatorio2 = Format(!Recuperatorio2, "0.00")
        txtPractico1 = Format(!Practico1, "0.00")
        txtPractico2 = Format(!Practico2, "0.00")
        txtPractico3 = Format(!Practico3, "0.00")
        txtPractico4 = Format(!Practico4, "0.00")
        txtPractico5 = Format(!Practico5, "0.00")
        If !Promocion = True Then
            chkPromocion.Value = 1
        Else
            chkPromocion.Value = 0
        End If
        txtTotalizador = !Totalizador
        If !Cursada = True Then
            chkCursada.Value = 1
        Else
            chkCursada.Value = 0
        End If
        If !Asistencia = True Then
            chkAsistencia.Value = 1
        Else
            chkAsistencia.Value = 0
        End If
        txtPorcentajeAsistencia = Format(!AsistenciaPorcentaje, "0.00")
        If !PerdioCursada = True Then
           lblPerdioCursada.Visible = True
        Else
           lblPerdioCursada.Visible = False
        End If
    End If
    End With
End Function

Public Function VerMatriculados()
    adoMatriculados.RecordSource = "SELECT  Alumnos.Permiso, Alumnos.Nombre, Finales.Parcial1, Finales.Parcial2, Finales.Totalizador, Finales.Recuperatorio1, Finales.Recuperatorio2, Finales.Practico1, Finales.Practico2, Finales.Practico3, Finales.Practico4, Finales.Practico5, Finales.Cursada, Finales.Asistencia, Finales.[AsistenciaPorcentaje], Finales.Promocion, Finales.Equivalencia, Finales.Establecimiento, Condicion.Condicion, Finales.PerdioCursada FROM (((Finales INNER JOIN Alumnos ON Finales.Alumno = Alumnos.Permiso) INNER JOIN Materias ON Finales.Materia = Materias.Codigo) INNER JOIN CarrerasHechas ON (Alumnos.Permiso = CarrerasHechas.Permiso) AND (Materias.Carrera = CarrerasHechas.Carrera)) INNER JOIN Condicion ON CarrerasHechas.Condición = Condicion.Codigo Where Finales.Materia = " & dtcMaterias.BoundText & " And Finales.Ano = " & txtAño & " And Finales.Division = " & cbDivision & " AND Finales.Libre=0 ORDER BY Alumnos.Nombre"
    adoMatriculados.Refresh
    If adoMatriculados.Recordset.RecordCount > 0 Then
        cmdVerAsistencia.Enabled = True
        frComandos.Enabled = True
        frPlanillas.Enabled = True
        frMatriculados.Enabled = True
        frCarrera.Enabled = True
        frMaterias.Enabled = True
        cmdLibres.Enabled = True
        If frmIdentificacion.Permisos!ModificarMatriculados = True Then
            cmdMatricular.Enabled = True
            cmdQuitar.Enabled = True
            cmdPerdioCursada.Enabled = True
        Else
            cmdMatricular.Enabled = False
            cmdQuitar.Enabled = False
        End If
        If frmIdentificacion.Permisos!ModificarParciales = True Then
            cmdAprobarAsistenciaATodos.Enabled = True
            cmdModificar.Enabled = True
            cmdIngresarNotas.Enabled = True
            cmdImprimeActaPromocion.Enabled = True
        Else
            cmdAprobarAsistenciaATodos.Enabled = False
            cmdModificar.Enabled = False
            cmdIngresarNotas.Enabled = False
            cmdImprimeActaPromocion.Enabled = False
        End If
        If frmIdentificacion.Permisos!IngresarAsistencia = True Then
            cmdIngresarAsistencia.Enabled = True
        Else
            cmdIngresarAsistencia.Enabled = False
        End If
        lblTotalMatriculados = adoMatriculados.Recordset.RecordCount
    Else
        cmdVerAsistencia.Enabled = False
        frComandos.Enabled = False
        frPlanillas.Enabled = False
        cmdModificar.Enabled = False
        'cmdLibres.Enabled = False
        If frmIdentificacion.Permisos!ModificarMatriculados = True Then
            cmdMatricular.Enabled = True
        End If
        cmdQuitar.Enabled = False
        cmdPerdioCursada.Enabled = False
        cmdIngresarAsistencia.Enabled = False
    End If
End Function

Private Function NuevasMaterias()
    frMaterias.Enabled = False
    frMatriculados.Enabled = False
    frComandos.Enabled = False
    frPlanillas.Enabled = False
    cmdMostrar.Enabled = True
    cmdIngresarNotas.Enabled = False
    cmdIngresarAsistencia.Enabled = False
End Function

Private Function GuardarDatos()
    SeGuarda = "Si"
    AlumnoActual = adoMatriculados.Recordset.Bookmark
    If chkPromocion.Value = 1 And chkAsistencia.Value = 0 Then
       MsgBox ("El alumno no puede promocionar ya que no aprobó la asistencia.")
       Me.MousePointer = 0
       SeGuarda = "No"
    End If
    If chkPromocion.Value = 1 Then 'el alumno promociono la materia, chequear si aprobo asistencia y no debe correlativas
        adoCorrelativas.RecordSource = "SELECT Correlativas.Principal, Correlativas.Correlativa, Materias.Nombre, Materias.Curso FROM Correlativas INNER JOIN Materias ON Correlativas.Correlativa = Materias.Codigo Where (((Correlativas.Principal) = " & adoMaterias.Recordset!Codigo & "))ORDER BY Correlativas.Correlativa    "
        adoCorrelativas.Refresh
        If adoCorrelativas.Recordset.RecordCount > 0 Then 'la materia elegida tiene correlativas
           'levanto las materias que son correlativas pero tiene aprobadas
            AdoFinal.RecordSource = "SELECT Correlativas.Correlativa FROM Correlativas INNER JOIN Finales ON Correlativas.Correlativa = Finales.Materia Where (((Correlativas.Principal) = " & adoMaterias.Recordset!Codigo & ") And ((Finales.Alumno) = " & adoMatriculados.Recordset!Permiso & ") And ((Finales.Aprobada) = True))"
            AdoFinal.Refresh
            TotalQueDebe = adoCorrelativas.Recordset.RecordCount - AdoFinal.Recordset.RecordCount
            If TotalQueDebe > 0 Then 'Debe alguna correlativa
                TotalCorrelativas = 0
                If AdoFinal.Recordset.RecordCount = 0 Then 'debe todas las correlativas
                    adoCorrelativas.Recordset.MoveFirst
                    For i = 1 To TotalQueDebe
                        NombreCorrelativa(i) = adoCorrelativas.Recordset!Curso & "°- " & adoCorrelativas.Recordset!Nombre
                        adoCorrelativas.Recordset.MoveNext
                    Next i
                Else ' debe solo alguna/s correlativa/s
                    adoCorrelativas.Recordset.MoveFirst
                    For i = 1 To adoCorrelativas.Recordset.RecordCount
                        AdoFinal.Recordset.MoveFirst
                        AdoFinal.Recordset.Find ("Correlativa=" & adoCorrelativas.Recordset!Correlativa)
                        If AdoFinal.Recordset.BOF = True Or AdoFinal.Recordset.EOF = True Then 'no encontro la materia en los finales aprobados (entonces la debe)
                            TotalCorrelativas = TotalCorrelativas + 1
                            NombreCorrelativa(TotalCorrelativas) = adoCorrelativas.Recordset!Curso & "°- " & adoCorrelativas.Recordset!Nombre
                            AdoFinal.Recordset.MoveFirst
                        End If
                        adoCorrelativas.Recordset.MoveNext
                    Next i
                End If
                StrinCorrelativa = ""
                For i = 1 To TotalQueDebe
                    StrinCorrelativas = StrinCorrelativas & NombreCorrelativa(i) & Chr(13) & Chr(13)
                Next i
                MsgBox ("El alumno no puede promocionar: debe las siguientes correlativas:" & Chr(13) & Chr(13) & StrinCorrelativas)
                Me.MousePointer = 0
                SeGuarda = "No"
            End If
        End If
    End If
    
    If SeGuarda = "Si" Then
        Conexion.Open
        Conexion.Execute ("UPDATE Finales SET Finales.Parcial1 = " & Replace(Val(txtParcial1), ",", ".") & ", Finales.Parcial2 = " & Replace(Val(txtParcial2), ",", ".") & ", Finales.Totalizador = " & Replace(Val(txtTotalizador), ",", ".") & ", Finales.Recuperatorio1 = " & Replace(Val(txtRecuperatorio1), ",", ".") & ", Finales.Recuperatorio2 = " & Replace(Val(txtRecuperatorio2), ",", ".") & ", Finales.Practico1 = " & Replace(Val(txtPractico1), ",", ".") & ", Finales.Practico2 = " & Replace(Val(txtPractico2), ",", ".") & ", Finales.Practico3 = " & Replace(Val(txtPractico3), ",", ".") & ", Finales.Practico4 = " & Replace(Val(txtPractico4), ",", ".") & ", Finales.Practico5 = " & Replace(Val(txtPractico5), ",", ".") & ", Finales.AsistenciaPorcentaje=" & Replace(Val(txtPorcentajeAsistencia), ",", ".") & " WHERE Finales.Alumno=" & adoMatriculados.Recordset!Permiso & " AND Finales.Materia=" & dtcMaterias.BoundText & " AND Finales.Ano=" & txtAño)
        If chkAsistencia = 1 Then
            Conexion.Execute ("UPDATE Finales SET Asistencia = True WHERE Finales.Alumno=" & adoMatriculados.Recordset!Permiso & " AND Finales.Materia=" & dtcMaterias.BoundText & " AND Finales.Ano=" & txtAño)
        Else
            Conexion.Execute ("UPDATE Finales SET Asistencia = False WHERE Finales.Alumno=" & adoMatriculados.Recordset!Permiso & " AND Finales.Materia=" & dtcMaterias.BoundText & " AND Finales.Ano=" & txtAño)
        End If
        If chkPromocion.Value = 1 Then
            If chkCursada.Value = 1 Then
                Conexion.Execute ("UPDATE Finales SET Finales.Cursada = True, Finales.Promocion = True WHERE Finales.Alumno=" & adoMatriculados.Recordset!Permiso & " AND Finales.Materia=" & dtcMaterias.BoundText & " AND Finales.Ano=" & txtAño)
            Else
                Conexion.Execute ("UPDATE Finales SET Finales.Cursada = False, Finales.Promocion = True WHERE Finales.Alumno=" & adoMatriculados.Recordset!Permiso & " AND Finales.Materia=" & dtcMaterias.BoundText & " AND Finales.Ano=" & txtAño)
            End If
        Else
            If chkCursada.Value = 1 Then
                Conexion.Execute ("UPDATE Finales SET Finales.Cursada = True, Finales.Promocion = False WHERE Finales.Alumno=" & adoMatriculados.Recordset!Permiso & " AND Finales.Materia=" & dtcMaterias.BoundText & " AND Finales.Ano=" & txtAño)
            Else
                Conexion.Execute ("UPDATE Finales SET Finales.Cursada = False, Finales.Promocion = False WHERE Finales.Alumno=" & adoMatriculados.Recordset!Permiso & " AND Finales.Materia=" & dtcMaterias.BoundText & " AND Finales.Ano=" & txtAño)
            End If
        End If
        Conexion.Close
    End If
    adoMatriculados.Refresh
    adoMatriculados.Recordset.Bookmark = AlumnoActual
End Function

Private Sub cmdCalcularAsistencia_Click()
    Respuesta = MsgBox("A continuación se perderán los resultados de la asistencia y se recalcularán teniendo en cuenta el registro diario de asistencia." & Chr(13) & "¿Desea continuar?", vbYesNo, "Atención")
    If Respuesta = vbNo Then Exit Sub
    Me.MousePointer = 11
    Conexion.Open
    Conectar.Open
    Set NumeroCursada = Conexion.Execute("SELECT Divisiones.Numero From Divisiones WHERE Divisiones.Materia=" & frmParciales.dtcMaterias.BoundText & " AND Divisiones.Ano=" & frmParciales.txtAño & " AND Divisiones.Division=" & frmParciales.cbDivision)
    CursadaNumero = NumeroCursada!Numero
    Set PorcentajeAsistencia = Conexion.Execute("SELECT Parametros.PorcentajeAsistencia FROM Parametros")
    PorcentajeAsistencias = PorcentajeAsistencia!PorcentajeAsistencia
    Set Auxiliar = Conectar.Execute("SELECT max(Fecha) as UltimaFecha from Asistencias WHERE Numero=" & CursadaNumero)
    HastaFecha = Auxiliar!UltimaFecha
    Conectar.Execute ("INSERT INTO AsistenciaTemporal ( Numero, Agente, Presente ) SELECT Asistencias.Numero, Asistencias.Agente, Asistencias.Presente From Asistencias WHERE Asistencias.Numero=" & CursadaNumero)
    Set Horas = Conectar.Execute("SELECT Count(Asistencias.Presente) AS cantidad, Asistencias.Agente FROM Asistencias WHERE (((Asistencias.Numero)=" & CursadaNumero & ")) GROUP BY Asistencias.Agente ORDER BY Count(Asistencias.Presente) desc")
    'Set Horas = Conectar.Execute("SELECT count( AsistenciaTemporal.Numero) as total From AsistenciaTemporal WHERE (((AsistenciaTemporal.Agente)=" & adoMatriculados.Recordset!Permiso & "))")
    TotalHoras = Horas!Cantidad
    adoMatriculados.Recordset.MoveFirst
    For i = 1 To adoMatriculados.Recordset.RecordCount
        'Set Presentes = Conexion.Execute("SELECT Count([Presente]) AS Presentes From Asistencias WHERE (((Asistencias.Presente)=True) AND ((Asistencias.Agente)=" & adoMatriculados.Recordset!Permiso & " ) AND ((Asistencias.Numero)=" & CursadaNumero & "))")
        'Set Ausentes = Conexion.Execute("SELECT Count([Presente]) AS Ausentes From Asistencias WHERE (((Asistencias.Presente)=False) AND ((Asistencias.Agente)=" & adoMatriculados.Recordset!Permiso & " ) AND ((Asistencias.Numero)=" & CursadaNumero & "))")
        Set Presentes = Conectar.Execute("SELECT sum(AsistenciaTemporal.Presente)*-1 AS Presentes FROM AsistenciaTemporal WHERE ((AsistenciaTemporal.Numero)=" & CursadaNumero & ") AND ((AsistenciaTemporal.Agente)=" & adoMatriculados.Recordset!Permiso & ")")
        
        'TotalPresentes = Presentes!Presentes
        If Presentes!Presentes <> Nulo Then
            TotalPresentes = Presentes!Presentes
        Else
            TotalPresentes = 0
        End If
        'If Presentes.EOF = False Then
        '    TotalPresentes = Presentes!Presentes
        'Else
        '    TotalPresentes = 0
        'End If
        'If Ausentes.EOF = False Then
        '    TotalAusentes = Ausentes!Ausentes
        'Else
        '    TotalAusentes = 0
        'End If
        'porcentaje = Int((TotalPresentes / (TotalPresentes + TotalAusentes)) * 100)
        porcentaje = Int(((TotalPresentes * 100) / TotalHoras))
        If porcentaje >= PorcentajeAsistencias Then
            Set Auxiliar = Conexion.Execute("UPDATE Finales SET Finales.[AsistenciaPorcentaje] = " & porcentaje & ", Finales.Asistencia = True, Finales.AsistenciaHasta = '" & DateValue(HastaFecha) & "' WHERE (((Finales.Alumno)=" & adoMatriculados.Recordset!Permiso & ") AND ((Finales.Materia)=" & dtcMaterias.BoundText & ") AND ((Finales.Ano)=" & txtAño & ") AND ((Finales.Division)=" & cbDivision & "))")
        Else
            Set Auxiliar = Conexion.Execute("UPDATE Finales SET Finales.[AsistenciaPorcentaje] = " & porcentaje & ", Finales.Asistencia = False,Finales.AsistenciaHasta = '" & DateValue(HastaFecha) & "' WHERE (((Finales.Alumno)=" & adoMatriculados.Recordset!Permiso & ") AND ((Finales.Materia)=" & dtcMaterias.BoundText & ") AND ((Finales.Ano)=" & txtAño & ") AND ((Finales.Division)=" & cbDivision & "))")
        End If
        adoMatriculados.Recordset.MoveNext
    Next i
    Conectar.Execute ("delete * from AsistenciaTemporal WHERE Numero=" & CursadaNumero)
    Conectar.Close
    Conexion.Close
    VerMatriculados
    Me.MousePointer = 0
End Sub

