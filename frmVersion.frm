VERSION 5.00
Begin VB.Form frmVersion 
   Caption         =   "Version"
   ClientHeight    =   2550
   ClientLeft      =   5400
   ClientTop       =   4410
   ClientWidth     =   5595
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   2550
   ScaleWidth      =   5595
   Begin VB.Frame Frame1 
      Height          =   1215
      Left            =   240
      TabIndex        =   1
      Top             =   240
      Width           =   4935
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "Versión 1.1.22"
         Height          =   375
         Left            =   1920
         TabIndex        =   3
         Top             =   720
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "SADIS"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1920
         TabIndex        =   2
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "Aceptar"
      Height          =   495
      Left            =   2040
      TabIndex        =   0
      Top             =   1800
      Width           =   1215
   End
End
Attribute VB_Name = "frmVersion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAceptar_Click()
    Unload Me
End Sub

