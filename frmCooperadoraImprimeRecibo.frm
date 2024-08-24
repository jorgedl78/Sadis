VERSION 5.00
Begin VB.Form frmCooperadoraImprimeRecibo 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   7995
   ClientLeft      =   2025
   ClientTop       =   60
   ClientWidth     =   11700
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   9996.249
   ScaleMode       =   0  'User
   ScaleWidth      =   11700
   ShowInTaskbar   =   0   'False
   Begin VB.Label lblOrden 
      BackColor       =   &H00FFFFFF&
      Caption         =   "9876"
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
      Left            =   240
      TabIndex        =   121
      Top             =   0
      Width           =   855
   End
   Begin VB.Line Line2 
      BorderStyle     =   3  'Dot
      X1              =   11640
      X2              =   11640
      Y1              =   1500.375
      Y2              =   3150.788
   End
   Begin VB.Label lblPermisoOctubre 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Label1"
      ForeColor       =   &H80000008&
      Height          =   165
      Left            =   4440
      TabIndex        =   120
      Top             =   2640
      Width           =   3135
   End
   Begin VB.Label lblPermisoSetiembre 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Label1"
      ForeColor       =   &H80000008&
      Height          =   165
      Left            =   8640
      TabIndex        =   119
      Top             =   2640
      Width           =   3015
   End
   Begin VB.Label lblPermisoAgosto 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Label1"
      ForeColor       =   &H80000008&
      Height          =   165
      Left            =   240
      TabIndex        =   118
      Top             =   4440
      Width           =   3255
   End
   Begin VB.Label lblPermisoJulio 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Label1"
      ForeColor       =   &H80000008&
      Height          =   165
      Left            =   4560
      TabIndex        =   117
      Top             =   4440
      Width           =   3015
   End
   Begin VB.Label lblPermisoJunio 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Label1"
      ForeColor       =   &H80000008&
      Height          =   165
      Left            =   8640
      TabIndex        =   116
      Top             =   4440
      Width           =   2895
   End
   Begin VB.Label lblPermisoMayo 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Label1"
      ForeColor       =   &H80000008&
      Height          =   165
      Left            =   240
      TabIndex        =   115
      Top             =   6240
      Width           =   3135
   End
   Begin VB.Label lblPermisoAbril 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Label1"
      ForeColor       =   &H80000008&
      Height          =   165
      Left            =   4560
      TabIndex        =   114
      Top             =   6240
      Width           =   3135
   End
   Begin VB.Label lblPermisoAporteVoluntario 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Label1"
      ForeColor       =   &H80000008&
      Height          =   165
      Left            =   8640
      TabIndex        =   113
      Top             =   6240
      Width           =   2775
   End
   Begin VB.Label lblPermisoNoviembre 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Label1"
      ForeColor       =   &H80000008&
      Height          =   165
      Left            =   240
      TabIndex        =   112
      Top             =   2640
      Width           =   3255
   End
   Begin VB.Label lblReciboNoviembre 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Label5"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   2280
      TabIndex        =   111
      Top             =   2880
      Width           =   975
   End
   Begin VB.Label lblReciboOctubre 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Label5"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   6360
      TabIndex        =   110
      Top             =   2880
      Width           =   975
   End
   Begin VB.Label lblReciboSetiembre 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Label5"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   10680
      TabIndex        =   109
      Top             =   2880
      Width           =   975
   End
   Begin VB.Label lblReciboAgosto 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Label5"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   2040
      TabIndex        =   108
      Top             =   4680
      Width           =   975
   End
   Begin VB.Label lblReciboJulio 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Label5"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   6240
      TabIndex        =   107
      Top             =   4680
      Width           =   975
   End
   Begin VB.Label lblReciboJunio 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Label5"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   10320
      TabIndex        =   106
      Top             =   4680
      Width           =   975
   End
   Begin VB.Label lblReciboMayo 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Label5"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   2280
      TabIndex        =   105
      Top             =   6480
      Width           =   975
   End
   Begin VB.Label lblReciboAbril 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Label5"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   6240
      TabIndex        =   104
      Top             =   6480
      Width           =   975
   End
   Begin VB.Label lblReciboAportevoluntario 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Label5"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   10560
      TabIndex        =   103
      Top             =   6480
      Width           =   975
   End
   Begin VB.Label lblDomicilioEnJunin 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "lblDomicilio"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   1080
      TabIndex        =   102
      Top             =   840
      Width           =   10455
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Alumno:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   240
      TabIndex        =   101
      Top             =   360
      Width           =   975
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "En Junín:"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   240
      TabIndex        =   100
      Top             =   840
      Width           =   855
   End
   Begin VB.Label Label24 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Fecha"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   99
      Top             =   6720
      Width           =   735
   End
   Begin VB.Label lblLocalidad 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "lblLocalidad"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   9480
      TabIndex        =   98
      Top             =   600
      Width           =   2175
   End
   Begin VB.Label Label19 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Localidad:"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   8400
      TabIndex        =   97
      Top             =   600
      Width           =   975
   End
   Begin VB.Label lblCobrador 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "lblCobrador"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   8160
      TabIndex        =   96
      Top             =   0
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label lblDomicilio 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "lblDomicilio"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   1200
      TabIndex        =   95
      Top             =   600
      Width           =   7095
   End
   Begin VB.Label lblTipoYNumeroDocumento 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "lblTipoYNumeroDocumento"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   8160
      TabIndex        =   94
      Top             =   240
      Width           =   2295
   End
   Begin VB.Label lblAno 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "2009"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   10680
      TabIndex        =   93
      Top             =   0
      Width           =   735
   End
   Begin VB.Label lblAlumno 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "lblAlumno"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   1200
      TabIndex        =   92
      Top             =   360
      Width           =   6495
   End
   Begin VB.Label Label15 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Planilla de control de cobranza - Cooperadora Instituto Nº 20"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   1320
      TabIndex        =   91
      Top             =   0
      Width           =   6495
   End
   Begin VB.Line Line13 
      BorderStyle     =   3  'Dot
      BorderWidth     =   2
      X1              =   11880
      X2              =   11880
      Y1              =   1500.375
      Y2              =   3150.788
   End
   Begin VB.Line Line10 
      BorderStyle     =   3  'Dot
      X1              =   240
      X2              =   240
      Y1              =   1500.375
      Y2              =   3150.788
   End
   Begin VB.Label Label28 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Firma"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   10
      Left            =   10680
      TabIndex        =   90
      Top             =   3840
      Width           =   615
   End
   Begin VB.Label Label29 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "...................................."
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   10
      Left            =   10200
      TabIndex        =   89
      Top             =   3600
      Width           =   1695
   End
   Begin VB.Label Label28 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Firma"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   9
      Left            =   6360
      TabIndex        =   88
      Top             =   3840
      Width           =   615
   End
   Begin VB.Label Label29 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "...................................."
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   9
      Left            =   5880
      TabIndex        =   87
      Top             =   3600
      Width           =   1695
   End
   Begin VB.Label Label28 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Firma"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   8
      Left            =   2280
      TabIndex        =   86
      Top             =   3840
      Width           =   615
   End
   Begin VB.Label Label29 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "...................................."
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   8
      Left            =   1800
      TabIndex        =   85
      Top             =   3600
      Width           =   1695
   End
   Begin VB.Label Label28 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Firma"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   7
      Left            =   10560
      TabIndex        =   84
      Top             =   5640
      Width           =   615
   End
   Begin VB.Label Label29 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "...................................."
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   7
      Left            =   10080
      TabIndex        =   83
      Top             =   5400
      Width           =   1695
   End
   Begin VB.Label Label28 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Firma"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   6
      Left            =   6360
      TabIndex        =   82
      Top             =   5640
      Width           =   615
   End
   Begin VB.Label Label29 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "...................................."
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   6
      Left            =   5880
      TabIndex        =   81
      Top             =   5400
      Width           =   1695
   End
   Begin VB.Label Label28 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Firma"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   5
      Left            =   2280
      TabIndex        =   80
      Top             =   5640
      Width           =   615
   End
   Begin VB.Label Label29 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "...................................."
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   5
      Left            =   1800
      TabIndex        =   79
      Top             =   5400
      Width           =   1695
   End
   Begin VB.Label Label28 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Firma"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   4
      Left            =   10440
      TabIndex        =   78
      Top             =   7440
      Width           =   615
   End
   Begin VB.Label Label29 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "...................................."
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   4
      Left            =   9960
      TabIndex        =   77
      Top             =   7200
      Width           =   1695
   End
   Begin VB.Label Label28 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Firma"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   3
      Left            =   6360
      TabIndex        =   76
      Top             =   7440
      Width           =   615
   End
   Begin VB.Label Label29 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "...................................."
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   3
      Left            =   5880
      TabIndex        =   75
      Top             =   7200
      Width           =   1695
   End
   Begin VB.Label Label28 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Firma"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   2
      Left            =   2160
      TabIndex        =   74
      Top             =   7440
      Width           =   615
   End
   Begin VB.Label Label29 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "...................................."
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   2
      Left            =   1680
      TabIndex        =   73
      Top             =   7200
      Width           =   1695
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Fecha "
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   1
      Left            =   8520
      TabIndex        =   72
      Top             =   6720
      Width           =   735
   End
   Begin VB.Label lblConceptoAporteVoluntario 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Aporte Vol.\10"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   1
      Left            =   8640
      TabIndex        =   71
      Top             =   6480
      Width           =   1335
   End
   Begin VB.Label lblImporteAporteVoluntario 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "$30"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   1
      Left            =   8760
      TabIndex        =   70
      Top             =   7080
      Width           =   615
   End
   Begin VB.Label lblFechaPagoAporteVoluntario 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   ".........../............/......................"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   1
      Left            =   9480
      TabIndex        =   69
      Top             =   6720
      Width           =   2175
   End
   Begin VB.Label lblFechaPagoAbril 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   ".........../............/......................"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   1
      Left            =   5400
      TabIndex        =   68
      Top             =   6720
      Width           =   2175
   End
   Begin VB.Label lblImporteAbril 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "$10"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   1
      Left            =   4560
      TabIndex        =   67
      Top             =   7080
      Width           =   735
   End
   Begin VB.Label lblConceptoAbril 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Abril\10"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   1
      Left            =   4560
      TabIndex        =   66
      Top             =   6480
      Width           =   855
   End
   Begin VB.Label Label24 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Fecha "
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   1
      Left            =   4440
      TabIndex        =   65
      Top             =   6720
      Width           =   735
   End
   Begin VB.Label lblFechaPagoMayo 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   ".........../............/......................"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   1
      Left            =   1200
      TabIndex        =   64
      Top             =   6720
      Width           =   2175
   End
   Begin VB.Label lblImporteMayo 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "$10"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   1
      Left            =   360
      TabIndex        =   63
      Top             =   7080
      Width           =   735
   End
   Begin VB.Label lblConceptoMayo 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Mayo\10"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   1
      Left            =   240
      TabIndex        =   62
      Top             =   6480
      Width           =   975
   End
   Begin VB.Label Label72 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Fecha "
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   1
      Left            =   8520
      TabIndex        =   61
      Top             =   4920
      Width           =   735
   End
   Begin VB.Label lblConceptoJunio 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Junio\10"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   1
      Left            =   8640
      TabIndex        =   60
      Top             =   4680
      Width           =   855
   End
   Begin VB.Label lblImporteJunio 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "$10"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   1
      Left            =   8760
      TabIndex        =   59
      Top             =   5280
      Width           =   735
   End
   Begin VB.Label lblFechaPagoJunio 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   ".........../............/......................"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   1
      Left            =   9600
      TabIndex        =   58
      Top             =   4920
      Width           =   2175
   End
   Begin VB.Label lblFechaPagoJulio 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   ".........../............/......................"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   1
      Left            =   5400
      TabIndex        =   57
      Top             =   4920
      Width           =   2175
   End
   Begin VB.Label lblImporteJulio 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "$10"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   1
      Left            =   4920
      TabIndex        =   56
      Top             =   5280
      Width           =   735
   End
   Begin VB.Label lblConceptoJulio 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Julio\10"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   1
      Left            =   4560
      TabIndex        =   55
      Top             =   4680
      Width           =   735
   End
   Begin VB.Label Label53 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Fecha "
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   1
      Left            =   4440
      TabIndex        =   54
      Top             =   4920
      Width           =   735
   End
   Begin VB.Label lblFechaPagoAgosto 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   ".........../............/......................"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   1
      Left            =   1320
      TabIndex        =   53
      Top             =   4920
      Width           =   2175
   End
   Begin VB.Label lblImporteAgosto 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "$10"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   1
      Left            =   600
      TabIndex        =   52
      Top             =   5280
      Width           =   735
   End
   Begin VB.Label lblConceptoAgosto 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Agosto\10"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   1
      Left            =   240
      TabIndex        =   51
      Top             =   4680
      Width           =   1095
   End
   Begin VB.Label Label41 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Fecha"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   1
      Left            =   0
      TabIndex        =   50
      Top             =   4920
      Width           =   855
   End
   Begin VB.Label Label108 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Fecha"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   1
      Left            =   8520
      TabIndex        =   49
      Top             =   3120
      Width           =   735
   End
   Begin VB.Label lblConceptoSetiembre 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Setiembre\10"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   1
      Left            =   8640
      TabIndex        =   48
      Top             =   2880
      Width           =   1455
   End
   Begin VB.Label lblImporteSetiembre 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "$10"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   1
      Left            =   8760
      TabIndex        =   47
      Top             =   3480
      Width           =   735
   End
   Begin VB.Label lblFechaPagoSetiembre 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   ".........../............/......................"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   1
      Left            =   9720
      TabIndex        =   46
      Top             =   3120
      Width           =   2175
   End
   Begin VB.Label lblFechaPagoOctubre 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   ".........../............/......................"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   1
      Left            =   5400
      TabIndex        =   45
      Top             =   3120
      Width           =   2175
   End
   Begin VB.Label lblImporteOctubre 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "$10"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   1
      Left            =   4800
      TabIndex        =   44
      Top             =   3480
      Width           =   615
   End
   Begin VB.Label lblConceptoOctubre 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Octubre\10"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   1
      Left            =   4440
      TabIndex        =   43
      Top             =   2880
      Width           =   1095
   End
   Begin VB.Label Label89 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Fecha"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   1
      Left            =   4320
      TabIndex        =   42
      Top             =   3120
      Width           =   735
   End
   Begin VB.Label lblFechaPagoNoviembre 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   ".........../............/......................"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   1
      Left            =   1200
      TabIndex        =   41
      Top             =   3120
      Width           =   2175
   End
   Begin VB.Label lblImporteNoviembre 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "$10"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   1
      Left            =   480
      TabIndex        =   40
      Top             =   3480
      Width           =   615
   End
   Begin VB.Label lblConceptoNoviembre 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Noviembre\10"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   1
      Left            =   240
      TabIndex        =   39
      Top             =   2880
      Width           =   1335
   End
   Begin VB.Label Label77 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Fecha"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   38
      Top             =   3120
      Width           =   735
   End
   Begin VB.Label lblConceptoSetiembre 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Setiembre"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   0
      Left            =   6360
      TabIndex        =   37
      Top             =   1440
      Width           =   975
   End
   Begin VB.Label lblImporteSetiembre 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "$8"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   0
      Left            =   7560
      TabIndex        =   36
      Top             =   1440
      Width           =   375
   End
   Begin VB.Label lblFechaPagoSetiembre 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   ".........../............/........."
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   0
      Left            =   8040
      TabIndex        =   35
      Top             =   1440
      Width           =   1695
   End
   Begin VB.Label lblReciboNumeroSetiembre 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "lblReciboNumeroSetiembre"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   0
      Left            =   9840
      TabIndex        =   34
      Top             =   1440
      Width           =   975
   End
   Begin VB.Label lblReciboNumeroOctubre 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "lblReciboNumeroOctubre"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   0
      Left            =   9840
      TabIndex        =   33
      Top             =   1680
      Width           =   1095
   End
   Begin VB.Label lblFechaPagoOctubre 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   ".........../............/........."
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   0
      Left            =   8040
      TabIndex        =   32
      Top             =   1680
      Width           =   1815
   End
   Begin VB.Label lblImporteOctubre 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "$8"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   0
      Left            =   7560
      TabIndex        =   31
      Top             =   1680
      Width           =   375
   End
   Begin VB.Label lblConceptoOctubre 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Octubre"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   0
      Left            =   6360
      TabIndex        =   30
      Top             =   1680
      Width           =   855
   End
   Begin VB.Label lblReciboNumeroNoviembre 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "lblReciboNumeroNoviembre"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   0
      Left            =   9840
      TabIndex        =   29
      Top             =   1920
      Width           =   1095
   End
   Begin VB.Label lblFechaPagoNoviembre 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   ".........../............/........."
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   0
      Left            =   8040
      TabIndex        =   28
      Top             =   1920
      Width           =   1695
   End
   Begin VB.Label lblImporteNoviembre 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "$8"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   0
      Left            =   7560
      TabIndex        =   27
      Top             =   1920
      Width           =   375
   End
   Begin VB.Label lblConceptoNoviembre 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Noviembre"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   0
      Left            =   6360
      TabIndex        =   26
      Top             =   1920
      Width           =   1215
   End
   Begin VB.Line Line3 
      BorderStyle     =   3  'Dot
      X1              =   240
      X2              =   12000
      Y1              =   3000.75
      Y2              =   3000.75
   End
   Begin VB.Label lblConceptoJunio 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Junio"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   0
      Left            =   360
      TabIndex        =   25
      Top             =   1920
      Width           =   975
   End
   Begin VB.Label lblImporteJunio 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "$8"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   0
      Left            =   1680
      TabIndex        =   24
      Top             =   1920
      Width           =   615
   End
   Begin VB.Label lblFechaPagoJunio 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   ".........../............/..........."
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   0
      Left            =   2280
      TabIndex        =   23
      Top             =   1920
      Width           =   1935
   End
   Begin VB.Label lblReciboNumeroJunio 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "lblReciboNumeroJunio"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   0
      Left            =   4200
      TabIndex        =   22
      Top             =   1920
      Width           =   1455
   End
   Begin VB.Label lblReciboNumeroJulio 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "lblReciboNumeroJulio"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   0
      Left            =   4200
      TabIndex        =   21
      Top             =   2160
      Width           =   1455
   End
   Begin VB.Label lblFechaPagoJulio 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   ".........../............/..........."
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   0
      Left            =   2280
      TabIndex        =   20
      Top             =   2160
      Width           =   1935
   End
   Begin VB.Label lblImporteJulio 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "$8"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   0
      Left            =   1680
      TabIndex        =   19
      Top             =   2160
      Width           =   615
   End
   Begin VB.Label lblConceptoJulio 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Julio"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   0
      Left            =   360
      TabIndex        =   18
      Top             =   2160
      Width           =   1335
   End
   Begin VB.Label lblReciboNumeroAgosto 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "lblReciboNumeroAgosto"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   0
      Left            =   9840
      TabIndex        =   17
      Top             =   1200
      Width           =   1095
   End
   Begin VB.Label lblFechaPagoAgosto 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   ".........../............/........."
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   0
      Left            =   8040
      TabIndex        =   16
      Top             =   1200
      Width           =   1695
   End
   Begin VB.Label lblImporteAgosto 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "$8"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   0
      Left            =   7560
      TabIndex        =   15
      Top             =   1200
      Width           =   375
   End
   Begin VB.Label lblConceptoAgosto 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Agosto"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   0
      Left            =   6360
      TabIndex        =   14
      Top             =   1200
      Width           =   975
   End
   Begin VB.Label lblConceptoMayo 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Mayo"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   0
      Left            =   360
      TabIndex        =   13
      Top             =   1680
      Width           =   975
   End
   Begin VB.Label lblImporteMayo 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "$8"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   0
      Left            =   1680
      TabIndex        =   12
      Top             =   1680
      Width           =   615
   End
   Begin VB.Label lblFechaPagoMayo 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   ".........../............/..........."
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   0
      Left            =   2280
      TabIndex        =   11
      Top             =   1680
      Width           =   1815
   End
   Begin VB.Label lblReciboNumeroMayo 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "lblReciboNumeroMayo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   0
      Left            =   4200
      TabIndex        =   10
      Top             =   1680
      Width           =   1455
   End
   Begin VB.Label lblConceptoAbril 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Abril"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   0
      Left            =   360
      TabIndex        =   9
      Top             =   1440
      Width           =   1335
   End
   Begin VB.Label lblImporteAbril 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "$8"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   0
      Left            =   1680
      TabIndex        =   8
      Top             =   1440
      Width           =   615
   End
   Begin VB.Label lblFechaPagoAbril 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   ".........../............/..........."
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   0
      Left            =   2280
      TabIndex        =   7
      Top             =   1440
      Width           =   1815
   End
   Begin VB.Label lblReciboNumeroAbril 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "lblReciboNumeroAbril"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   0
      Left            =   4200
      TabIndex        =   6
      Top             =   1440
      Width           =   1455
   End
   Begin VB.Label lblReciboNumeroaporteVoluntario 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "lblReciboNumeroaporteVoluntario"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   0
      Left            =   4200
      TabIndex        =   5
      Top             =   1200
      Width           =   1455
   End
   Begin VB.Label lblFechaPagoAporteVoluntario 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   ".........../............/..........."
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   0
      Left            =   2280
      TabIndex        =   4
      Top             =   1200
      Width           =   1815
   End
   Begin VB.Label lblImporteAporteVoluntario 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "$30"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   0
      Left            =   1680
      TabIndex        =   3
      Top             =   1200
      Width           =   495
   End
   Begin VB.Label lblConceptoAporteVoluntario 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Ap. Voluntario"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   0
      Left            =   360
      TabIndex        =   2
      Top             =   1200
      Width           =   1455
   End
   Begin VB.Line Line1 
      BorderStyle     =   3  'Dot
      X1              =   240
      X2              =   12240
      Y1              =   1400.35
      Y2              =   1400.35
   End
   Begin VB.Label Label6 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Domicilio:"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   600
      Width           =   975
   End
   Begin VB.Label Label4 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Año:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   10080
      TabIndex        =   0
      Top             =   0
      Width           =   615
   End
End
Attribute VB_Name = "frmCooperadoraImprimeRecibo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Activate()
'Unload Me
'Me.PrintForm
End Sub

