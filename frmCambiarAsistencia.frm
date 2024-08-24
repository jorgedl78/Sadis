VERSION 5.00
Begin VB.Form frmCambiarAsistencia 
   Caption         =   "Modificar Asistencia"
   ClientHeight    =   3510
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14145
   LinkTopic       =   "Form1"
   ScaleHeight     =   3510
   ScaleWidth      =   14145
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Height          =   1815
      Left            =   120
      TabIndex        =   5
      Top             =   1440
      Width           =   13935
      Begin VB.CommandButton cmdGuardar 
         Height          =   600
         Left            =   4920
         Picture         =   "frmCambiarAsistencia.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Guardar"
         Top             =   600
         Width           =   600
      End
      Begin VB.CommandButton cmdCancelar 
         Height          =   600
         Left            =   7440
         Picture         =   "frmCambiarAsistencia.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Cancelar"
         Top             =   600
         Width           =   600
      End
      Begin VB.CheckBox chkPresente 
         Caption         =   "Presente"
         Height          =   255
         Left            =   480
         TabIndex        =   8
         Top             =   720
         Width           =   1575
      End
      Begin VB.Label lblNumero 
         Caption         =   "lblNumero"
         Height          =   375
         Left            =   360
         TabIndex        =   12
         Top             =   1440
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Label lblPermiso 
         Caption         =   "lblPermiso"
         Height          =   375
         Left            =   360
         TabIndex        =   11
         Top             =   1200
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Label lblFecha 
         Caption         =   "lblFecha"
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
         Left            =   1200
         TabIndex        =   7
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label3 
         Caption         =   "Fecha:"
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
         Left            =   480
         TabIndex        =   6
         Top             =   360
         Width           =   735
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1335
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   13935
      Begin VB.Label lblMateria 
         Caption         =   "lblMateria"
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
         Left            =   1200
         TabIndex        =   4
         Top             =   315
         Width           =   12615
      End
      Begin VB.Label lblAlumno 
         Caption         =   "lblAlumno"
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
         Left            =   1200
         TabIndex        =   3
         Top             =   840
         Width           =   12615
      End
      Begin VB.Shape Shape1 
         Height          =   375
         Left            =   240
         Shape           =   4  'Rounded Rectangle
         Top             =   720
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Alumno:"
         Height          =   255
         Left            =   360
         TabIndex        =   2
         Top             =   795
         Width           =   615
      End
      Begin VB.Shape Shape2 
         Height          =   375
         Left            =   240
         Shape           =   4  'Rounded Rectangle
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label2 
         Caption         =   "Materia:"
         Height          =   255
         Left            =   360
         TabIndex        =   1
         Top             =   315
         Width           =   615
      End
   End
End
Attribute VB_Name = "frmCambiarAsistencia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Conexion As New Connection

Private Sub cmdCancelar_Click()
    Unload Me
End Sub

Private Sub cmdGuardar_Click()
    Me.MousePointer = 11
    Conexion.Open
    Conexion.Execute ("UPDATE Asistencias SET Presente = " & chkPresente.Value & " WHERE Agente=" & lblPermiso & " and Numero = " & lblNumero & " and Fecha=#" & Format(lblFecha, "mm/dd/yyyy") & "#")
    Conexion.Close
    Me.MousePointer = 1
    Unload Me
End Sub

Private Sub Form_Load()
    Conexion.ConnectionString = ("DSN=Instituto")
End Sub
