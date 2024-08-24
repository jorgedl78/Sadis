VERSION 5.00
Begin VB.Form frmImprimeActas 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   ClientHeight    =   16995
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11865
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   297
   ScaleMode       =   0  'User
   ScaleWidth      =   212.962
   ShowInTaskbar   =   0   'False
   Begin VB.Label lblDivision 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "division"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   11040
      TabIndex        =   315
      Top             =   2160
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label Label4 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Div:"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   10560
      TabIndex        =   314
      Top             =   2160
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label lblLocalidad 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Presidente"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   -360
      TabIndex        =   313
      Top             =   13080
      Width           =   1935
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "(L) = Cursada Libre"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   9840
      TabIndex        =   312
      Top             =   14520
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Label lblLibre 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   0
      Left            =   2400
      TabIndex        =   311
      Top             =   480
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.Label CuadroIngresado 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
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
      Height          =   375
      Index           =   1
      Left            =   3960
      TabIndex        =   310
      Top             =   13560
      Width           =   4095
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Confeccionada por el Profesor"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   1440
      TabIndex        =   309
      Top             =   13680
      Width           =   2295
   End
   Begin VB.Label CuadroIngresado 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
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
      Height          =   375
      Index           =   0
      Left            =   3960
      TabIndex        =   308
      Top             =   14040
      Width           =   4095
   End
   Begin VB.Label Label29 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Ingresada al sistema por:"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   1440
      TabIndex        =   307
      Top             =   14040
      Width           =   2295
   End
   Begin VB.Label lblLugar 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "lblLugar"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   9960
      TabIndex        =   306
      Top             =   1800
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Lugar:"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   9360
      TabIndex        =   305
      Top             =   1800
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label lblHora 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "lblHora"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   9960
      TabIndex        =   304
      Top             =   1560
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label28 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Hora:"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   9360
      TabIndex        =   303
      Top             =   1560
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label Label27 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Vocal"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   4320
      TabIndex        =   302
      Top             =   15960
      Width           =   3615
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label26 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Vocal"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   8280
      TabIndex        =   301
      Top             =   15960
      Width           =   3135
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label25 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Presidente"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   960
      TabIndex        =   300
      Top             =   15960
      Width           =   3135
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label24 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   " ....................................................."
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   4320
      TabIndex        =   299
      Top             =   15240
      Width           =   3615
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "....................................................."
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   8280
      TabIndex        =   298
      Top             =   15240
      Width           =   3135
   End
   Begin VB.Label lblTitular 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "lblTitular"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   960
      TabIndex        =   297
      Top             =   15600
      Width           =   3135
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblIntegrante1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "lblIntegrante1"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   4320
      TabIndex        =   296
      Top             =   15600
      Width           =   3615
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblIntegrante2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "lblIntegrante2"
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   8280
      TabIndex        =   295
      Top             =   15600
      Width           =   3135
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblPermiso 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   26
      Left            =   1200
      TabIndex        =   294
      Top             =   12240
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label lblPermiso 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   25
      Left            =   1200
      TabIndex        =   293
      Top             =   11880
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label lblPermiso 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   24
      Left            =   1200
      TabIndex        =   292
      Top             =   11520
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label lblPermiso 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   23
      Left            =   1200
      TabIndex        =   291
      Top             =   11160
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label lblPermiso 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   22
      Left            =   1200
      TabIndex        =   290
      Top             =   10800
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label lblPermiso 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   21
      Left            =   1200
      TabIndex        =   289
      Top             =   10440
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label lblPermiso 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   20
      Left            =   1200
      TabIndex        =   288
      Top             =   10080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label lblPermiso 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   19
      Left            =   1200
      TabIndex        =   287
      Top             =   9720
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label lblPermiso 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   18
      Left            =   1200
      TabIndex        =   286
      Top             =   9360
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label lblPermiso 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   17
      Left            =   1200
      TabIndex        =   285
      Top             =   9000
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label lblPermiso 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   16
      Left            =   1200
      TabIndex        =   284
      Top             =   8640
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label lblPermiso 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   15
      Left            =   1200
      TabIndex        =   283
      Top             =   8280
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label lblPermiso 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   14
      Left            =   1200
      TabIndex        =   282
      Top             =   7920
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label lblPermiso 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   13
      Left            =   1200
      TabIndex        =   281
      Top             =   7560
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label lblPermiso 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   12
      Left            =   1200
      TabIndex        =   280
      Top             =   7200
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label lblPermiso 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   11
      Left            =   1200
      TabIndex        =   279
      Top             =   6840
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label lblPermiso 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   10
      Left            =   1200
      TabIndex        =   278
      Top             =   6480
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label lblPermiso 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   9
      Left            =   1200
      TabIndex        =   277
      Top             =   6120
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label lblPermiso 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   8
      Left            =   1200
      TabIndex        =   276
      Top             =   5760
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label lblPermiso 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   7
      Left            =   1200
      TabIndex        =   275
      Top             =   5400
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label lblPermiso 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   6
      Left            =   1200
      TabIndex        =   274
      Top             =   5040
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label lblPermiso 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   5
      Left            =   1200
      TabIndex        =   273
      Top             =   4680
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label lblPermiso 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   4
      Left            =   1200
      TabIndex        =   272
      Top             =   4320
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label lblPermiso 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   3
      Left            =   1200
      TabIndex        =   271
      Top             =   3960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label lblPermiso 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   2
      Left            =   1200
      TabIndex        =   270
      Top             =   3600
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label lblPermiso 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   1
      Left            =   1200
      TabIndex        =   269
      Top             =   3240
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label lblPermiso 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Permiso"
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
      Height          =   615
      Index           =   0
      Left            =   1200
      TabIndex        =   268
      Top             =   2640
      Width           =   735
   End
   Begin VB.Label Label22 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Folio:"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   9240
      TabIndex        =   267
      Top             =   840
      Width           =   495
   End
   Begin VB.Label Label21 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   345
      Left            =   10680
      TabIndex        =   266
      Top             =   12720
      Width           =   735
   End
   Begin VB.Label Label20 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   345
      Left            =   10680
      TabIndex        =   265
      Top             =   13080
      Width           =   735
   End
   Begin VB.Label Label19 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   345
      Left            =   10680
      TabIndex        =   264
      Top             =   13440
      Width           =   735
   End
   Begin VB.Label lnlFinalLetras 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
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
      Height          =   375
      Index           =   26
      Left            =   10560
      TabIndex        =   263
      Top             =   12240
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label lnlFinalLetras 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
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
      Height          =   375
      Index           =   25
      Left            =   10560
      TabIndex        =   262
      Top             =   11880
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label lnlFinalLetras 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
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
      Height          =   375
      Index           =   24
      Left            =   10560
      TabIndex        =   261
      Top             =   11520
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label lnlFinalLetras 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
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
      Height          =   375
      Index           =   23
      Left            =   10560
      TabIndex        =   260
      Top             =   11160
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label lnlFinalLetras 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
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
      Height          =   375
      Index           =   22
      Left            =   10560
      TabIndex        =   259
      Top             =   10800
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label lnlFinalLetras 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
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
      Height          =   375
      Index           =   21
      Left            =   10560
      TabIndex        =   258
      Top             =   10440
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label lnlFinalLetras 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
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
      Height          =   375
      Index           =   20
      Left            =   10560
      TabIndex        =   257
      Top             =   10080
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label lnlFinalLetras 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
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
      Height          =   375
      Index           =   19
      Left            =   10560
      TabIndex        =   256
      Top             =   9720
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label lnlFinalLetras 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
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
      Height          =   375
      Index           =   18
      Left            =   10560
      TabIndex        =   255
      Top             =   9360
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label lnlFinalLetras 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
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
      Height          =   375
      Index           =   17
      Left            =   10560
      TabIndex        =   254
      Top             =   9000
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label lnlFinalLetras 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
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
      Height          =   375
      Index           =   16
      Left            =   10560
      TabIndex        =   253
      Top             =   8640
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label lnlFinalLetras 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
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
      Height          =   375
      Index           =   15
      Left            =   10560
      TabIndex        =   252
      Top             =   8280
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label lnlFinalLetras 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
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
      Height          =   375
      Index           =   14
      Left            =   10560
      TabIndex        =   251
      Top             =   7920
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label lnlFinalLetras 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
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
      Height          =   375
      Index           =   13
      Left            =   10560
      TabIndex        =   250
      Top             =   7560
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label lnlFinalLetras 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
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
      Height          =   375
      Index           =   12
      Left            =   10560
      TabIndex        =   249
      Top             =   7200
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label lnlFinalLetras 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
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
      Height          =   375
      Index           =   11
      Left            =   10560
      TabIndex        =   248
      Top             =   6840
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label lnlFinalLetras 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
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
      Height          =   375
      Index           =   10
      Left            =   10560
      TabIndex        =   247
      Top             =   6480
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label lnlFinalLetras 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
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
      Height          =   375
      Index           =   9
      Left            =   10560
      TabIndex        =   246
      Top             =   6120
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label lnlFinalLetras 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
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
      Height          =   375
      Index           =   8
      Left            =   10560
      TabIndex        =   245
      Top             =   5760
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label lnlFinalLetras 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
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
      Height          =   375
      Index           =   7
      Left            =   10560
      TabIndex        =   244
      Top             =   5400
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label lnlFinalLetras 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
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
      Height          =   375
      Index           =   6
      Left            =   10560
      TabIndex        =   243
      Top             =   5040
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label lnlFinalLetras 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
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
      Height          =   375
      Index           =   5
      Left            =   10560
      TabIndex        =   242
      Top             =   4680
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label lnlFinalLetras 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
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
      Height          =   375
      Index           =   4
      Left            =   10560
      TabIndex        =   241
      Top             =   4320
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label lnlFinalLetras 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
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
      Height          =   375
      Index           =   3
      Left            =   10560
      TabIndex        =   240
      Top             =   3960
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label lnlFinalLetras 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
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
      Height          =   375
      Index           =   2
      Left            =   10560
      TabIndex        =   239
      Top             =   3600
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label lnlFinalLetras 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
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
      Height          =   375
      Index           =   1
      Left            =   10560
      TabIndex        =   238
      Top             =   3240
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label lnlOralLetras 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
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
      Height          =   375
      Index           =   26
      Left            =   9000
      TabIndex        =   237
      Top             =   12240
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label lnlOralLetras 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
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
      Height          =   375
      Index           =   25
      Left            =   9000
      TabIndex        =   236
      Top             =   11880
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label lnlOralLetras 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
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
      Height          =   375
      Index           =   24
      Left            =   9000
      TabIndex        =   235
      Top             =   11520
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label lnlOralLetras 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
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
      Height          =   375
      Index           =   23
      Left            =   9000
      TabIndex        =   234
      Top             =   11160
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label lnlOralLetras 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
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
      Height          =   375
      Index           =   22
      Left            =   9000
      TabIndex        =   233
      Top             =   10800
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label lnlOralLetras 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
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
      Height          =   375
      Index           =   21
      Left            =   9000
      TabIndex        =   232
      Top             =   10440
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label lnlOralLetras 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
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
      Height          =   375
      Index           =   20
      Left            =   9000
      TabIndex        =   231
      Top             =   10080
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label lnlOralLetras 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
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
      Height          =   375
      Index           =   19
      Left            =   9000
      TabIndex        =   230
      Top             =   9720
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label lnlOralLetras 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
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
      Height          =   375
      Index           =   18
      Left            =   9000
      TabIndex        =   229
      Top             =   9360
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label lnlOralLetras 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
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
      Height          =   375
      Index           =   17
      Left            =   9000
      TabIndex        =   228
      Top             =   9000
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label lnlOralLetras 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
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
      Height          =   375
      Index           =   16
      Left            =   9000
      TabIndex        =   227
      Top             =   8640
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label lnlOralLetras 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
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
      Height          =   375
      Index           =   15
      Left            =   9000
      TabIndex        =   226
      Top             =   8280
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label lnlOralLetras 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
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
      Height          =   375
      Index           =   14
      Left            =   9000
      TabIndex        =   225
      Top             =   7920
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label lnlOralLetras 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
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
      Height          =   375
      Index           =   13
      Left            =   9000
      TabIndex        =   224
      Top             =   7560
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label lnlOralLetras 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
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
      Height          =   375
      Index           =   12
      Left            =   9000
      TabIndex        =   223
      Top             =   7200
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label lnlOralLetras 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
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
      Height          =   375
      Index           =   11
      Left            =   9000
      TabIndex        =   222
      Top             =   6840
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label lnlOralLetras 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
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
      Height          =   375
      Index           =   10
      Left            =   9000
      TabIndex        =   221
      Top             =   6480
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label lnlOralLetras 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
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
      Height          =   375
      Index           =   9
      Left            =   9000
      TabIndex        =   220
      Top             =   6120
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label lnlOralLetras 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
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
      Height          =   375
      Index           =   8
      Left            =   9000
      TabIndex        =   219
      Top             =   5760
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label lnlOralLetras 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
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
      Height          =   375
      Index           =   7
      Left            =   9000
      TabIndex        =   218
      Top             =   5400
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label lnlOralLetras 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
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
      Height          =   375
      Index           =   6
      Left            =   9000
      TabIndex        =   217
      Top             =   5040
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label lnlOralLetras 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
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
      Height          =   375
      Index           =   5
      Left            =   9000
      TabIndex        =   216
      Top             =   4680
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label lnlOralLetras 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
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
      Height          =   375
      Index           =   4
      Left            =   9000
      TabIndex        =   215
      Top             =   4320
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label lnlOralLetras 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
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
      Height          =   375
      Index           =   3
      Left            =   9000
      TabIndex        =   214
      Top             =   3960
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label lnlOralLetras 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
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
      Height          =   375
      Index           =   2
      Left            =   9000
      TabIndex        =   213
      Top             =   3600
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label lnlOralLetras 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
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
      Height          =   375
      Index           =   1
      Left            =   9000
      TabIndex        =   212
      Top             =   3240
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label lnlFinalLetras 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
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
      Height          =   375
      Index           =   0
      Left            =   9840
      TabIndex        =   211
      Top             =   720
      Width           =   1215
   End
   Begin VB.Label lnlEscritoLetras 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   26
      Left            =   7440
      TabIndex        =   210
      Top             =   12240
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label lnlEscritoLetras 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   25
      Left            =   7440
      TabIndex        =   209
      Top             =   11880
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label lnlEscritoLetras 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   24
      Left            =   7440
      TabIndex        =   208
      Top             =   11520
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label lnlEscritoLetras 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   23
      Left            =   7440
      TabIndex        =   207
      Top             =   11160
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label lnlEscritoLetras 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   22
      Left            =   7440
      TabIndex        =   206
      Top             =   10800
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label lnlEscritoLetras 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   21
      Left            =   7440
      TabIndex        =   205
      Top             =   10440
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label lnlEscritoLetras 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   20
      Left            =   7440
      TabIndex        =   204
      Top             =   10080
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label lnlEscritoLetras 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   19
      Left            =   7440
      TabIndex        =   203
      Top             =   9720
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label lnlEscritoLetras 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   18
      Left            =   7440
      TabIndex        =   202
      Top             =   9360
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label lnlEscritoLetras 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   17
      Left            =   7440
      TabIndex        =   201
      Top             =   9000
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label lnlEscritoLetras 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   16
      Left            =   7440
      TabIndex        =   200
      Top             =   8640
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label lnlEscritoLetras 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   15
      Left            =   7440
      TabIndex        =   199
      Top             =   8280
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label lnlEscritoLetras 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   14
      Left            =   7440
      TabIndex        =   198
      Top             =   7920
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label lnlEscritoLetras 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   13
      Left            =   7440
      TabIndex        =   197
      Top             =   7560
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label lnlEscritoLetras 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   12
      Left            =   7440
      TabIndex        =   196
      Top             =   7200
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label lnlEscritoLetras 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   11
      Left            =   7440
      TabIndex        =   195
      Top             =   6840
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label lnlEscritoLetras 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   10
      Left            =   7440
      TabIndex        =   194
      Top             =   6480
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label lnlEscritoLetras 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   9
      Left            =   7440
      TabIndex        =   193
      Top             =   6120
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label lnlEscritoLetras 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   8
      Left            =   7440
      TabIndex        =   192
      Top             =   5760
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label lnlEscritoLetras 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   7
      Left            =   7440
      TabIndex        =   191
      Top             =   5400
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label lnlEscritoLetras 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   6
      Left            =   7440
      TabIndex        =   190
      Top             =   5040
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label lnlEscritoLetras 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   5
      Left            =   7440
      TabIndex        =   189
      Top             =   4680
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label lnlEscritoLetras 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   4
      Left            =   7440
      TabIndex        =   188
      Top             =   4320
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label lnlEscritoLetras 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   3
      Left            =   7440
      TabIndex        =   187
      Top             =   3960
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label lnlEscritoLetras 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   2
      Left            =   7440
      TabIndex        =   186
      Top             =   3600
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label lnlEscritoLetras 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   1
      Left            =   7440
      TabIndex        =   185
      Top             =   3240
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label lblFinalNota 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
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
      Height          =   375
      Index           =   26
      Left            =   9960
      TabIndex        =   184
      Top             =   12240
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label lblFinalNota 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
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
      Height          =   375
      Index           =   25
      Left            =   9960
      TabIndex        =   183
      Top             =   11880
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label lblFinalNota 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
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
      Height          =   375
      Index           =   24
      Left            =   9960
      TabIndex        =   182
      Top             =   11520
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label lblFinalNota 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
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
      Height          =   375
      Index           =   23
      Left            =   9960
      TabIndex        =   181
      Top             =   11160
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label lblFinalNota 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
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
      Height          =   375
      Index           =   22
      Left            =   9960
      TabIndex        =   180
      Top             =   10800
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label lblFinalNota 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
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
      Height          =   375
      Index           =   21
      Left            =   9960
      TabIndex        =   179
      Top             =   10440
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label lblFinalNota 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
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
      Height          =   375
      Index           =   20
      Left            =   9960
      TabIndex        =   178
      Top             =   10080
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label lblFinalNota 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
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
      Height          =   375
      Index           =   19
      Left            =   9960
      TabIndex        =   177
      Top             =   9720
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label lblFinalNota 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
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
      Height          =   375
      Index           =   18
      Left            =   9960
      TabIndex        =   176
      Top             =   9360
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label lblFinalNota 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
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
      Height          =   375
      Index           =   17
      Left            =   9960
      TabIndex        =   175
      Top             =   9000
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label lblFinalNota 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
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
      Height          =   375
      Index           =   16
      Left            =   9960
      TabIndex        =   174
      Top             =   8640
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label lblFinalNota 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
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
      Height          =   375
      Index           =   15
      Left            =   9960
      TabIndex        =   173
      Top             =   8280
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label lblFinalNota 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
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
      Height          =   375
      Index           =   14
      Left            =   9960
      TabIndex        =   172
      Top             =   7920
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label lblFinalNota 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
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
      Height          =   375
      Index           =   13
      Left            =   9960
      TabIndex        =   171
      Top             =   7560
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label lblFinalNota 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
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
      Height          =   375
      Index           =   12
      Left            =   9960
      TabIndex        =   170
      Top             =   7200
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label lblFinalNota 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
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
      Height          =   375
      Index           =   11
      Left            =   9960
      TabIndex        =   169
      Top             =   6840
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label lblFinalNota 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
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
      Height          =   375
      Index           =   10
      Left            =   9960
      TabIndex        =   168
      Top             =   6480
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label lblFinalNota 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
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
      Height          =   375
      Index           =   9
      Left            =   9960
      TabIndex        =   167
      Top             =   6120
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label lblFinalNota 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
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
      Height          =   375
      Index           =   8
      Left            =   9960
      TabIndex        =   166
      Top             =   5760
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label lblFinalNota 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
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
      Height          =   375
      Index           =   7
      Left            =   9960
      TabIndex        =   165
      Top             =   5400
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label lblFinalNota 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
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
      Height          =   375
      Index           =   6
      Left            =   9960
      TabIndex        =   164
      Top             =   5040
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label lblFinalNota 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
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
      Height          =   375
      Index           =   5
      Left            =   9960
      TabIndex        =   163
      Top             =   4680
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label lblFinalNota 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
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
      Height          =   375
      Index           =   4
      Left            =   9960
      TabIndex        =   162
      Top             =   4320
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label lblFinalNota 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
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
      Height          =   375
      Index           =   3
      Left            =   9960
      TabIndex        =   161
      Top             =   3960
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label lblFinalNota 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
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
      Height          =   375
      Index           =   2
      Left            =   9960
      TabIndex        =   160
      Top             =   3600
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label lblFinalNota 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
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
      Height          =   375
      Index           =   1
      Left            =   9960
      TabIndex        =   159
      Top             =   3240
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label lblOralNota 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
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
      Height          =   375
      Index           =   26
      Left            =   8400
      TabIndex        =   158
      Top             =   12240
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label lblOralNota 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
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
      Height          =   375
      Index           =   25
      Left            =   8400
      TabIndex        =   157
      Top             =   11880
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label lblOralNota 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
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
      Height          =   375
      Index           =   24
      Left            =   8400
      TabIndex        =   156
      Top             =   11520
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label lblOralNota 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
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
      Height          =   375
      Index           =   23
      Left            =   8400
      TabIndex        =   155
      Top             =   11160
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label lblOralNota 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
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
      Height          =   375
      Index           =   22
      Left            =   8400
      TabIndex        =   154
      Top             =   10800
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label lblOralNota 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
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
      Height          =   375
      Index           =   21
      Left            =   8400
      TabIndex        =   153
      Top             =   10440
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label lblOralNota 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
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
      Height          =   375
      Index           =   20
      Left            =   8400
      TabIndex        =   152
      Top             =   10080
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label lblOralNota 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
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
      Height          =   375
      Index           =   19
      Left            =   8400
      TabIndex        =   151
      Top             =   9720
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label lblOralNota 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
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
      Height          =   375
      Index           =   18
      Left            =   8400
      TabIndex        =   150
      Top             =   9360
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label lblOralNota 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
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
      Height          =   375
      Index           =   17
      Left            =   8400
      TabIndex        =   149
      Top             =   9000
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label lblOralNota 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
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
      Height          =   375
      Index           =   16
      Left            =   8400
      TabIndex        =   148
      Top             =   8640
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label lblOralNota 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
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
      Height          =   375
      Index           =   15
      Left            =   8400
      TabIndex        =   147
      Top             =   8280
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label lblOralNota 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
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
      Height          =   375
      Index           =   14
      Left            =   8400
      TabIndex        =   146
      Top             =   7920
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label lblOralNota 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
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
      Height          =   375
      Index           =   13
      Left            =   8400
      TabIndex        =   145
      Top             =   7560
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label lblOralNota 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
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
      Height          =   375
      Index           =   12
      Left            =   8400
      TabIndex        =   144
      Top             =   7200
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label lblOralNota 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
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
      Height          =   375
      Index           =   11
      Left            =   8400
      TabIndex        =   143
      Top             =   6840
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label lblOralNota 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
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
      Height          =   375
      Index           =   10
      Left            =   8400
      TabIndex        =   142
      Top             =   6480
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label lblOralNota 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
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
      Height          =   375
      Index           =   9
      Left            =   8400
      TabIndex        =   141
      Top             =   6120
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label lblOralNota 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
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
      Height          =   375
      Index           =   8
      Left            =   8400
      TabIndex        =   140
      Top             =   5760
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label lblOralNota 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
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
      Height          =   375
      Index           =   7
      Left            =   8400
      TabIndex        =   139
      Top             =   5400
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label lblOralNota 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
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
      Height          =   375
      Index           =   6
      Left            =   8400
      TabIndex        =   138
      Top             =   5040
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label lblOralNota 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
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
      Height          =   375
      Index           =   5
      Left            =   8400
      TabIndex        =   137
      Top             =   4680
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label lblOralNota 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
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
      Height          =   375
      Index           =   4
      Left            =   8400
      TabIndex        =   136
      Top             =   4320
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label lblOralNota 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
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
      Height          =   375
      Index           =   3
      Left            =   8400
      TabIndex        =   135
      Top             =   3960
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label lblOralNota 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
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
      Height          =   375
      Index           =   2
      Left            =   8400
      TabIndex        =   134
      Top             =   3600
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label lblOralNota 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
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
      Height          =   375
      Index           =   1
      Left            =   8400
      TabIndex        =   133
      Top             =   3240
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label lblEscritoNota 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   26
      Left            =   6840
      TabIndex        =   132
      Top             =   12240
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label lblEscritoNota 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   25
      Left            =   6840
      TabIndex        =   131
      Top             =   11880
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label lblEscritoNota 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   24
      Left            =   6840
      TabIndex        =   130
      Top             =   11520
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label lblEscritoNota 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   23
      Left            =   6840
      TabIndex        =   129
      Top             =   11160
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label lblEscritoNota 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   22
      Left            =   6840
      TabIndex        =   128
      Top             =   10800
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label lblEscritoNota 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   21
      Left            =   6840
      TabIndex        =   127
      Top             =   10440
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label lblEscritoNota 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   20
      Left            =   6840
      TabIndex        =   126
      Top             =   10080
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label lblEscritoNota 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   19
      Left            =   6840
      TabIndex        =   125
      Top             =   9720
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label lblEscritoNota 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   18
      Left            =   6840
      TabIndex        =   124
      Top             =   9360
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label lblEscritoNota 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   17
      Left            =   6840
      TabIndex        =   123
      Top             =   9000
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label lblEscritoNota 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   16
      Left            =   6840
      TabIndex        =   122
      Top             =   8640
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label lblEscritoNota 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   15
      Left            =   6840
      TabIndex        =   121
      Top             =   8280
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label lblEscritoNota 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   14
      Left            =   6840
      TabIndex        =   120
      Top             =   7920
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label lblEscritoNota 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   13
      Left            =   6840
      TabIndex        =   119
      Top             =   7560
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label lblEscritoNota 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   12
      Left            =   6840
      TabIndex        =   118
      Top             =   7200
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label lblEscritoNota 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   11
      Left            =   6840
      TabIndex        =   117
      Top             =   6840
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label lblEscritoNota 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   10
      Left            =   6840
      TabIndex        =   116
      Top             =   6480
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label lblEscritoNota 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   9
      Left            =   6840
      TabIndex        =   115
      Top             =   6120
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label lblEscritoNota 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   8
      Left            =   6840
      TabIndex        =   114
      Top             =   5760
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label lblEscritoNota 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   7
      Left            =   6840
      TabIndex        =   113
      Top             =   5400
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label lblEscritoNota 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   6
      Left            =   6840
      TabIndex        =   112
      Top             =   5040
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label lblEscritoNota 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   5
      Left            =   6840
      TabIndex        =   111
      Top             =   4680
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label lblEscritoNota 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   4
      Left            =   6840
      TabIndex        =   110
      Top             =   4320
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label lblEscritoNota 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   3
      Left            =   6840
      TabIndex        =   109
      Top             =   3960
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label lblEscritoNota 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   2
      Left            =   6840
      TabIndex        =   108
      Top             =   3600
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label lblEscritoNota 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   1
      Left            =   6840
      TabIndex        =   107
      Top             =   3240
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label lblOral 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Oral"
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
      Height          =   615
      Index           =   0
      Left            =   8400
      TabIndex        =   106
      Top             =   2640
      Width           =   1575
   End
   Begin VB.Label lblCalificaciónDefinitiva 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Calificación Definitiva"
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
      Height          =   615
      Index           =   1
      Left            =   9960
      TabIndex        =   105
      Top             =   2640
      Width           =   1455
   End
   Begin VB.Label lblEscrito 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Escrito"
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
      Height          =   615
      Index           =   0
      Left            =   6840
      TabIndex        =   104
      Top             =   2640
      Width           =   1575
   End
   Begin VB.Label lblAlumno 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   26
      Left            =   3480
      TabIndex        =   103
      Top             =   12240
      Visible         =   0   'False
      Width           =   3399
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblAlumno 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   25
      Left            =   3480
      TabIndex        =   102
      Top             =   11880
      Visible         =   0   'False
      Width           =   3399
   End
   Begin VB.Label lblAlumno 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   24
      Left            =   3480
      TabIndex        =   101
      Top             =   11520
      Visible         =   0   'False
      Width           =   3399
   End
   Begin VB.Label lblAlumno 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   23
      Left            =   3480
      TabIndex        =   100
      Top             =   11160
      Visible         =   0   'False
      Width           =   3399
   End
   Begin VB.Label lblAlumno 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   22
      Left            =   3480
      TabIndex        =   99
      Top             =   10800
      Visible         =   0   'False
      Width           =   3399
   End
   Begin VB.Label lblAlumno 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   21
      Left            =   3480
      TabIndex        =   98
      Top             =   10440
      Visible         =   0   'False
      Width           =   3399
   End
   Begin VB.Label lblAlumno 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   20
      Left            =   3480
      TabIndex        =   97
      Top             =   10080
      Visible         =   0   'False
      Width           =   3399
   End
   Begin VB.Label lblAlumno 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   19
      Left            =   3480
      TabIndex        =   96
      Top             =   9720
      Visible         =   0   'False
      Width           =   3399
   End
   Begin VB.Label lblAlumno 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   18
      Left            =   3480
      TabIndex        =   95
      Top             =   9360
      Visible         =   0   'False
      Width           =   3399
   End
   Begin VB.Label lblAlumno 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   17
      Left            =   3480
      TabIndex        =   94
      Top             =   9000
      Visible         =   0   'False
      Width           =   3399
   End
   Begin VB.Label lblAlumno 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   16
      Left            =   3480
      TabIndex        =   93
      Top             =   8640
      Visible         =   0   'False
      Width           =   3399
   End
   Begin VB.Label lblAlumno 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   15
      Left            =   3480
      TabIndex        =   92
      Top             =   8280
      Visible         =   0   'False
      Width           =   3399
   End
   Begin VB.Label lblAlumno 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   14
      Left            =   3480
      TabIndex        =   91
      Top             =   7920
      Visible         =   0   'False
      Width           =   3399
   End
   Begin VB.Label lblAlumno 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   13
      Left            =   3480
      TabIndex        =   90
      Top             =   7560
      Visible         =   0   'False
      Width           =   3399
   End
   Begin VB.Label lblAlumno 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   12
      Left            =   3480
      TabIndex        =   89
      Top             =   7200
      Visible         =   0   'False
      Width           =   3399
   End
   Begin VB.Label lblAlumno 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   11
      Left            =   3480
      TabIndex        =   88
      Top             =   6840
      Visible         =   0   'False
      Width           =   3399
   End
   Begin VB.Label lblAlumno 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   10
      Left            =   3480
      TabIndex        =   87
      Top             =   6480
      Visible         =   0   'False
      Width           =   3399
   End
   Begin VB.Label lblAlumno 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   9
      Left            =   3480
      TabIndex        =   86
      Top             =   6120
      Visible         =   0   'False
      Width           =   3399
   End
   Begin VB.Label lblAlumno 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   8
      Left            =   3480
      TabIndex        =   85
      Top             =   5760
      Visible         =   0   'False
      Width           =   3399
   End
   Begin VB.Label lblAlumno 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   7
      Left            =   3480
      TabIndex        =   84
      Top             =   5400
      Visible         =   0   'False
      Width           =   3399
   End
   Begin VB.Label lblAlumno 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   6
      Left            =   3480
      TabIndex        =   83
      Top             =   5040
      Visible         =   0   'False
      Width           =   3399
   End
   Begin VB.Label lblAlumno 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   5
      Left            =   3480
      TabIndex        =   82
      Top             =   4680
      Visible         =   0   'False
      Width           =   3399
   End
   Begin VB.Label lblAlumno 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   4
      Left            =   3480
      TabIndex        =   81
      Top             =   4320
      Visible         =   0   'False
      Width           =   3399
   End
   Begin VB.Label lblAlumno 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   3
      Left            =   3480
      TabIndex        =   80
      Top             =   3960
      Visible         =   0   'False
      Width           =   3399
   End
   Begin VB.Label lblAlumno 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   2
      Left            =   3480
      TabIndex        =   79
      Top             =   3600
      Visible         =   0   'False
      Width           =   3399
   End
   Begin VB.Label lblAlumno 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   1
      Left            =   3480
      TabIndex        =   78
      Top             =   3240
      Visible         =   0   'False
      Width           =   3399
   End
   Begin VB.Label lblAlumno 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Alumno"
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
      Height          =   615
      Index           =   0
      Left            =   3480
      TabIndex        =   77
      Top             =   2640
      Width           =   3375
   End
   Begin VB.Label lblDocumento 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   26
      Left            =   1920
      TabIndex        =   76
      Top             =   12240
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Label lblDocumento 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   25
      Left            =   1920
      TabIndex        =   75
      Top             =   11880
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Label lblDocumento 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   24
      Left            =   1920
      TabIndex        =   74
      Top             =   11520
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Label lblDocumento 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   23
      Left            =   1920
      TabIndex        =   73
      Top             =   11160
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Label lblDocumento 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   22
      Left            =   1920
      TabIndex        =   72
      Top             =   10800
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Label lblDocumento 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   21
      Left            =   1920
      TabIndex        =   71
      Top             =   10440
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Label lblDocumento 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   20
      Left            =   1920
      TabIndex        =   70
      Top             =   10080
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Label lblDocumento 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   19
      Left            =   1920
      TabIndex        =   69
      Top             =   9720
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Label lblDocumento 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   18
      Left            =   1920
      TabIndex        =   68
      Top             =   9360
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Label lblDocumento 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   17
      Left            =   1920
      TabIndex        =   67
      Top             =   9000
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Label lblDocumento 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   16
      Left            =   1920
      TabIndex        =   66
      Top             =   8640
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Label lblDocumento 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   15
      Left            =   1920
      TabIndex        =   65
      Top             =   8280
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Label lblDocumento 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   14
      Left            =   1920
      TabIndex        =   64
      Top             =   7920
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Label lblDocumento 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   13
      Left            =   1920
      TabIndex        =   63
      Top             =   7560
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Label lblDocumento 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   12
      Left            =   1920
      TabIndex        =   62
      Top             =   7200
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Label lblDocumento 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   11
      Left            =   1920
      TabIndex        =   61
      Top             =   6840
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Label lblDocumento 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   10
      Left            =   1920
      TabIndex        =   60
      Top             =   6480
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Label lblDocumento 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   9
      Left            =   1920
      TabIndex        =   59
      Top             =   6120
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Label lblDocumento 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   8
      Left            =   1920
      TabIndex        =   58
      Top             =   5760
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Label lblDocumento 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   7
      Left            =   1920
      TabIndex        =   57
      Top             =   5400
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Label lblDocumento 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   6
      Left            =   1920
      TabIndex        =   56
      Top             =   5040
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Label lblDocumento 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   5
      Left            =   1920
      TabIndex        =   55
      Top             =   4680
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Label lblDocumento 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   4
      Left            =   1920
      TabIndex        =   54
      Top             =   4320
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Label lblDocumento 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   3
      Left            =   1920
      TabIndex        =   53
      Top             =   3960
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Label lblDocumento 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   2
      Left            =   1920
      TabIndex        =   52
      Top             =   3600
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Label lblDocumento 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   1
      Left            =   1920
      TabIndex        =   51
      Top             =   3240
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Label lblDocumento 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Documento"
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
      Height          =   615
      Index           =   0
      Left            =   1920
      TabIndex        =   50
      Top             =   2640
      Width           =   1575
   End
   Begin VB.Label lblOrden 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "26"
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   26
      Left            =   960
      TabIndex        =   49
      Top             =   12240
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label lblOrden 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "25"
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   25
      Left            =   960
      TabIndex        =   48
      Top             =   11880
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label lblOrden 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "24"
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   24
      Left            =   960
      TabIndex        =   47
      Top             =   11520
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label lblOrden 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "23"
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   23
      Left            =   960
      TabIndex        =   46
      Top             =   11160
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label lblOrden 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "22"
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   22
      Left            =   960
      TabIndex        =   45
      Top             =   10800
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label lblOrden 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "21"
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   21
      Left            =   960
      TabIndex        =   44
      Top             =   10440
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label lblOrden 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "20"
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   20
      Left            =   960
      TabIndex        =   43
      Top             =   10080
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label lblOrden 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "19"
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   19
      Left            =   960
      TabIndex        =   42
      Top             =   9720
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label lblOrden 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "18"
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   18
      Left            =   960
      TabIndex        =   41
      Top             =   9360
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label lblOrden 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "17"
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   17
      Left            =   960
      TabIndex        =   40
      Top             =   9000
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label lblOrden 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "16"
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   16
      Left            =   960
      TabIndex        =   39
      Top             =   8640
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label lblOrden 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "15"
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   15
      Left            =   960
      TabIndex        =   38
      Top             =   8280
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label lblOrden 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "14"
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   14
      Left            =   960
      TabIndex        =   37
      Top             =   7920
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label lblOrden 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "13"
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   13
      Left            =   960
      TabIndex        =   36
      Top             =   7560
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label lblOrden 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "12"
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   12
      Left            =   960
      TabIndex        =   35
      Top             =   7200
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label lblOrden 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "11"
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   11
      Left            =   960
      TabIndex        =   34
      Top             =   6840
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label lblOrden 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "10"
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   10
      Left            =   960
      TabIndex        =   33
      Top             =   6480
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label lblOrden 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "9"
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   9
      Left            =   960
      TabIndex        =   32
      Top             =   6120
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label lblOrden 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "8"
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   8
      Left            =   960
      TabIndex        =   31
      Top             =   5760
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label lblOrden 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "7"
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   7
      Left            =   960
      TabIndex        =   30
      Top             =   5400
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label lblOrden 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "6"
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   6
      Left            =   960
      TabIndex        =   29
      Top             =   5040
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label lblOrden 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "5"
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   5
      Left            =   960
      TabIndex        =   28
      Top             =   4680
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label lblOrden 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "4"
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   4
      Left            =   960
      TabIndex        =   27
      Top             =   4320
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label lblOrden 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "3"
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   3
      Left            =   960
      TabIndex        =   26
      Top             =   3960
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label lblOrden 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "2"
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   2
      Left            =   960
      TabIndex        =   25
      Top             =   3600
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label lblOrden 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "1"
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   1
      Left            =   960
      TabIndex        =   24
      Top             =   3240
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label lblOrden 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Nº"
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
      Height          =   615
      Index           =   0
      Left            =   960
      TabIndex        =   23
      Top             =   2640
      Width           =   255
   End
   Begin VB.Label lblActa 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Acta"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   11040
      TabIndex        =   22
      Top             =   120
      Width           =   615
   End
   Begin VB.Label lblMesa 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Mesa"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   9720
      TabIndex        =   21
      Top             =   120
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label Label18 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Acta:"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   10440
      TabIndex        =   20
      Top             =   120
      Width           =   495
   End
   Begin VB.Label Label13 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Mesa:"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   9000
      TabIndex        =   19
      Top             =   120
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label lblCursada 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Cursada"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   3720
      TabIndex        =   18
      Top             =   0
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label lblCurso 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Curso"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   9960
      TabIndex        =   17
      Top             =   2160
      Width           =   375
   End
   Begin VB.Label lblFecha 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Fecha"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   10080
      TabIndex        =   16
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Label lblMateria 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Carrera:"
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   1920
      TabIndex        =   15
      Top             =   2040
      Width           =   7215
   End
   Begin VB.Label lblCarrera 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Carrera:"
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   1920
      TabIndex        =   14
      Top             =   1440
      Width           =   7335
   End
   Begin VB.Label lblTotalAlumnos 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Total"
      ForeColor       =   &H80000008&
      Height          =   345
      Left            =   10680
      TabIndex        =   13
      Top             =   14040
      Width           =   735
   End
   Begin VB.Line Line6 
      X1              =   189.539
      X2              =   206.77
      Y1              =   243.262
      Y2              =   243.262
   End
   Begin VB.Label Label17 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Total de Alumnos:"
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   8640
      TabIndex        =   12
      Top             =   14040
      Width           =   1935
   End
   Begin VB.Label Label16 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Aplazados:"
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   9120
      TabIndex        =   11
      Top             =   13200
      Width           =   1455
   End
   Begin VB.Label Label15 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Aprobados:"
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   9120
      TabIndex        =   10
      Top             =   12840
      Width           =   1455
   End
   Begin VB.Label Label14 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Ausentes:"
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   9360
      TabIndex        =   9
      Top             =   13560
      Width           =   1215
   End
   Begin VB.Label Label12 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   ", ........................................ de ........................................... de 20........"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   1680
      TabIndex        =   8
      Top             =   13080
      Width           =   4905
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "....................................................."
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   960
      TabIndex        =   7
      Top             =   15240
      Width           =   3135
   End
   Begin VB.Label Label9 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Cursada:"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   2880
      TabIndex        =   6
      Top             =   0
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Label8 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Curso:"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   9360
      TabIndex        =   5
      Top             =   2160
      Width           =   495
   End
   Begin VB.Label Label7 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Fecha:"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   9360
      TabIndex        =   4
      Top             =   1320
      Width           =   615
   End
   Begin VB.Label Label6 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Materia:"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   960
      TabIndex        =   3
      Top             =   2040
      Width           =   855
   End
   Begin VB.Label Label5 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Carrera:"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   960
      TabIndex        =   2
      Top             =   1440
      Width           =   735
   End
   Begin VB.Label lblEstablecimiento 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Instituto Superior de Formación Docente y Técnica Nº 20"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   960
      TabIndex        =   1
      Top             =   1080
      Width           =   7935
   End
   Begin VB.Label lblTituloDeActa 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Acta Volante de Examenes "
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
      Left            =   3840
      TabIndex        =   0
      Top             =   360
      Width           =   4695
   End
End
Attribute VB_Name = "frmImprimeActas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub lblAnioCursada_Click(Index As Integer)

End Sub
