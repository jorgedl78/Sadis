VERSION 5.00
Begin VB.Form frmNewsAlumnos 
   BackColor       =   &H00FFFFFF&
   ClientHeight    =   6900
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   8100
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   Moveable        =   0   'False
   ScaleHeight     =   6900
   ScaleWidth      =   8100
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      BackColor       =   &H0080C0FF&
      Caption         =   "Continuar..."
      DisabledPicture =   "frmNewsAlumnos.frx":0000
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5400
      MaskColor       =   &H000000FF&
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   6120
      Width           =   2175
   End
   Begin VB.Image Image2 
      Height          =   1560
      Left            =   6120
      Picture         =   "frmNewsAlumnos.frx":0442
      Top             =   2520
      Width           =   915
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "1 GB de capacidad!!!"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008080&
      Height          =   375
      Left            =   1800
      TabIndex        =   6
      Top             =   2280
      Width           =   2415
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Sorteo: 29/06/07 por lotería nacional (17:30 hs)"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008080&
      Height          =   975
      Left            =   240
      TabIndex        =   5
      Top             =   5640
      Width           =   3015
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Buscá tu número en el istado al lado de esta compu...!!"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008080&
      Height          =   375
      Left            =   960
      TabIndex        =   4
      Top             =   5160
      Width           =   6255
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Con solo tener tus aportes al día participas automáticamente "
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008080&
      Height          =   1455
      Left            =   480
      TabIndex        =   3
      Top             =   3360
      Width           =   4575
   End
   Begin VB.Image Image3 
      Height          =   480
      Left            =   240
      Picture         =   "frmNewsAlumnos.frx":0CF6
      Top             =   5040
      Width           =   480
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H00008080&
      BorderWidth     =   5
      Height          =   6855
      Left            =   0
      Top             =   0
      Width           =   8055
   End
   Begin VB.Image Image1 
      Height          =   1050
      Left            =   6480
      Picture         =   "frmNewsAlumnos.frx":1138
      Stretch         =   -1  'True
      Top             =   120
      Width           =   1080
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00C0C0C0&
      BorderColor     =   &H00808080&
      BorderWidth     =   3
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   975
      Left            =   3600
      Shape           =   4  'Rounded Rectangle
      Top             =   5880
      Width           =   4455
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Atención!!!!!!!"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   480
      TabIndex        =   0
      Top             =   240
      Width           =   1815
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H0080C0FF&
      BorderColor     =   &H000080FF&
      BorderStyle     =   6  'Inside Solid
      BorderWidth     =   5
      FillColor       =   &H0080C0FF&
      FillStyle       =   0  'Solid
      Height          =   855
      Left            =   0
      Shape           =   4  'Rounded Rectangle
      Top             =   0
      Width           =   4695
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Ganate este MP3!!"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   975
      Left            =   360
      TabIndex        =   1
      Top             =   1440
      Width           =   6255
   End
End
Attribute VB_Name = "frmNewsAlumnos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Label4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Label4.MousePointer = a
End Sub

Private Sub Command1_Click()
    Unload Me
End Sub
