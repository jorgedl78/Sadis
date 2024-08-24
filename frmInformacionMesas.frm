VERSION 5.00
Begin VB.Form frmInformacionMesas 
   Caption         =   "Información Sobre Mesas de Exámenes"
   ClientHeight    =   8595
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   8595
   ScaleWidth      =   11880
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdActasSinIngresar 
      Caption         =   "Actas sin Ingresar"
      Height          =   615
      Left            =   360
      TabIndex        =   5
      Top             =   6480
      Width           =   4575
   End
   Begin VB.CommandButton cmdAlumnosAprobadosPorTurnoYCarrera 
      Caption         =   "Alumnos Aprobados por Turno"
      Height          =   615
      Left            =   360
      TabIndex        =   4
      Top             =   5160
      Width           =   4575
   End
   Begin VB.CommandButton cmdMesasPorCarrera 
      Caption         =   "Mesas por Carrera"
      Height          =   615
      Left            =   360
      TabIndex        =   3
      Top             =   3720
      Width           =   4575
   End
   Begin VB.CommandButton MesasPorProfesores 
      Caption         =   "Mesas para entregar a profesores"
      Height          =   615
      Left            =   360
      TabIndex        =   2
      Top             =   2040
      Width           =   4575
   End
   Begin VB.CommandButton cmdSalir 
      Height          =   840
      Left            =   10320
      Picture         =   "frmInformacionMesas.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Salir"
      Top             =   7200
      Width           =   960
   End
   Begin VB.CommandButton cmdMesasPorDia 
      Caption         =   "Mesas Por Dia Para Firmar"
      Height          =   615
      Left            =   360
      TabIndex        =   0
      Top             =   600
      Width           =   4575
   End
End
Attribute VB_Name = "frmInformacionMesas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdActasSinIngresar_Click()
    frmInformacionMesasActasSinIngresar.Show 1
End Sub

Private Sub cmdAlumnosAprobadosPorTurnoYCarrera_Click()
    frmAlumnosAprobadosPOrTurnoyCarrera.Show 1
End Sub

Private Sub cmdMesasPorCarrera_Click()
    frmInformacionMesasPorCarrera.Show 1
End Sub

Private Sub cmdMesasPorDia_Click()
    frmInformacionMesasPorDia.Show 1
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub MesasPorProfesores_Click()
    frmInformacionMesasParaProfesores.Show 1
End Sub
