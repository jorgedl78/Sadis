VERSION 5.00
Begin VB.Form frmModificarFolioActa 
   Caption         =   "Cambiar Folio"
   ClientHeight    =   3660
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4695
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   3660
   ScaleWidth      =   4695
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtFolioActual 
      Alignment       =   1  'Right Justify
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
      Height          =   375
      Left            =   1440
      TabIndex        =   9
      Top             =   1200
      Width           =   975
   End
   Begin VB.Frame Frame1 
      Height          =   2055
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   4455
      Begin VB.TextBox txtLibroNuevo 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2400
         TabIndex        =   11
         Top             =   600
         Width           =   975
      End
      Begin VB.TextBox txtFolioNuevo 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2400
         TabIndex        =   10
         Top             =   1080
         Width           =   975
      End
      Begin VB.TextBox txtLibroActual 
         Alignment       =   1  'Right Justify
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
         Height          =   375
         Left            =   1320
         TabIndex        =   8
         Top             =   600
         Width           =   975
      End
      Begin VB.Label lblMesa 
         Caption         =   "lblMesa"
         Height          =   375
         Left            =   3600
         TabIndex        =   12
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "Nuevo"
         Height          =   255
         Left            =   2520
         TabIndex        =   7
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Actual"
         Height          =   255
         Left            =   1440
         TabIndex        =   6
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Libro"
         Height          =   255
         Left            =   720
         TabIndex        =   5
         Top             =   720
         Width           =   495
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Folio"
         Height          =   255
         Left            =   840
         TabIndex        =   4
         Top             =   1200
         Width           =   375
      End
   End
   Begin VB.Frame Frame2 
      Height          =   1215
      Left            =   120
      TabIndex        =   0
      Top             =   2280
      Width           =   4455
      Begin VB.CommandButton cmdCancelar 
         Height          =   650
         Left            =   2640
         Picture         =   "frmModificarFolioActa.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   2
         TabStop         =   0   'False
         ToolTipText     =   "Cancelar"
         Top             =   360
         Width           =   650
      End
      Begin VB.CommandButton cmdAceptar 
         Height          =   650
         Left            =   1200
         Picture         =   "frmModificarFolioActa.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   1
         TabStop         =   0   'False
         ToolTipText     =   "Aceptar"
         Top             =   360
         Width           =   650
      End
   End
End
Attribute VB_Name = "frmModificarFolioActa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Conexion As New Connection
Private Sub cmdAceptar_Click()
    If txtLibroNuevo.Text = "" Or txtFolioNuevo.Text = "" Then MsgBox ("Debe completar los valores nuevos"): Exit Sub
    If Not IsNumeric(txtLibroNuevo) Then MsgBox ("El valor del libro no es numérico"): txtLibroNuevo.SetFocus: Exit Sub
    If Not IsNumeric(txtFolioNuevo) Then MsgBox ("El valor del folio no es numérico"): txtLibroNuevo.SetFocus: Exit Sub

    Respuesta = MsgBox("Confirma la actualización?", vbYesNo, "Atención")
    If Respuesta = vbYes Then
        Me.MousePointer = 11
        Conexion.Open
        'MsgBox ("UPDATE Actas set Libro=" & txtLibroNuevo & ", Folio=" & txtLibroNuevo & " WHERE Mesa=" & lblMesa)
        'MsgBox ("UPDATE Finales set Libro=" & txtLibroNuevo & ", Folio=" & txtLibroNuevo & " WHERE Mesa=" & lblMesa)
        Conexion.Execute ("UPDATE Actas set Libro=" & txtLibroNuevo & ", Folio=" & txtFolioNuevo & " WHERE Mesa=" & lblMesa)
        Conexion.Execute ("UPDATE Finales set Libro=" & txtLibroNuevo & ", Folio=" & txtFolioNuevo & " WHERE Mesa=" & lblMesa)
        Conexion.Close
        Me.MousePointer = 0
        Unload Me
    End If
End Sub

Private Sub cmdCancelar_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
    txtLibroNuevo.SetFocus
End Sub

Private Sub Form_Load()
    Conexion.ConnectionString = ("DSN=Instituto")

End Sub
