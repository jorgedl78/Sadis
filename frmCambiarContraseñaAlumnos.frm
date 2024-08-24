VERSION 5.00
Begin VB.Form frmCambiarContraseñaAlumnos 
   BackColor       =   &H0080FFFF&
   Caption         =   "Cambio de Contraseña"
   ClientHeight    =   3690
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   2910
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   Moveable        =   0   'False
   ScaleHeight     =   3690
   ScaleWidth      =   2910
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancelar 
      Height          =   650
      Left            =   1680
      Picture         =   "frmCambiarContraseñaAlumnos.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   7
      TabStop         =   0   'False
      ToolTipText     =   "Cancelar"
      Top             =   2880
      Width           =   650
   End
   Begin VB.CommandButton cmdAceptar 
      Height          =   650
      Left            =   600
      Picture         =   "frmCambiarContraseñaAlumnos.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   6
      TabStop         =   0   'False
      ToolTipText     =   "Aceptar"
      Top             =   2880
      Width           =   650
   End
   Begin VB.TextBox txtConfirmar 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   480
      MaxLength       =   6
      PasswordChar    =   "*"
      TabIndex        =   5
      Top             =   2280
      Width           =   1935
   End
   Begin VB.TextBox txtNueva 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   480
      MaxLength       =   6
      PasswordChar    =   "*"
      TabIndex        =   4
      Top             =   1320
      Width           =   1935
   End
   Begin VB.TextBox txtActual 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   480
      MaxLength       =   6
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   360
      Width           =   1935
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Contraseña nueva:"
      Height          =   255
      Left            =   480
      TabIndex        =   2
      Top             =   1080
      Width           =   1455
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Confirmar nueva contraseña:"
      Height          =   255
      Left            =   480
      TabIndex        =   1
      Top             =   2040
      Width           =   2175
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Contraseña Actual:"
      Height          =   255
      Left            =   480
      TabIndex        =   0
      Top             =   120
      Width           =   1455
   End
End
Attribute VB_Name = "frmCambiarContraseñaAlumnos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAceptar_Click()
    If txtActual <> frmLoginAlumno.ContraseñaReal Then
        Respuesta = MsgBox("La contraseña actual es incorrecta", 0, "Error")
        txtActual = ""
        txtActual.SetFocus
        Exit Sub
    End If
    If txtNueva <> txtConfirmar Then
        Respuesta = MsgBox("No coincide la nueva contraseña", 0, "Error")
        txtNueva = ""
        txtConfirmar = ""
        txtNueva.SetFocus
        Exit Sub
    End If
    frmLoginAlumno.ContraseñaReal = txtNueva
    frmLoginAlumno.GuardarContraseña
    Unload Me
End Sub

Private Sub cmdCancelar_Click()
    Unload Me
End Sub

Private Sub txtActual_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If txtActual <> frmLoginAlumno.ContraseñaReal Then
            Respuesta = MsgBox("La contraseña actual es incorrecta", 0, "Error")
            txtActual = ""
            txtNueva = ""
            txtConfirmar = ""
            txtActual.SetFocus
            Exit Sub
        End If
        txtNueva.SetFocus
    End If
    If ((KeyAscii >= 48 And KeyAscii <= 57) Or (KeyAscii >= 65 And KeyAscii <= 90) Or (KeyAscii >= 97 And KeyAscii <= 122)) And Len(ContraseñaParcial) < 6 Or KeyAscii = 8 Then
    Else
        KeyAscii = 0
    End If
End Sub

Private Sub txtConfirmar_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cmdAceptar_Click
    If ((KeyAscii >= 48 And KeyAscii <= 57) Or (KeyAscii >= 65 And KeyAscii <= 90) Or (KeyAscii >= 97 And KeyAscii <= 122)) And Len(ContraseñaParcial) < 6 Or KeyAscii = 8 Then
    Else
        KeyAscii = 0
    End If
End Sub

Private Sub txtNueva_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtConfirmar.SetFocus
    If ((KeyAscii >= 48 And KeyAscii <= 57) Or (KeyAscii >= 65 And KeyAscii <= 90) Or (KeyAscii >= 97 And KeyAscii <= 122)) And Len(ContraseñaParcial) < 6 Or KeyAscii = 8 Then
    Else
        KeyAscii = 0
    End If
End Sub
