VERSION 5.00
Begin VB.Form frmCambiarContraseñaUsuarios 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cambiar Contraseña"
   ClientHeight    =   3690
   ClientLeft      =   4620
   ClientTop       =   2790
   ClientWidth     =   2790
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   3690
   ScaleWidth      =   2790
   ShowInTaskbar   =   0   'False
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
      Left            =   360
      MaxLength       =   6
      PasswordChar    =   "*"
      TabIndex        =   4
      Top             =   360
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
      Left            =   360
      MaxLength       =   6
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   1320
      Width           =   1935
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
      Left            =   360
      MaxLength       =   6
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   2280
      Width           =   1935
   End
   Begin VB.CommandButton cmdAceptar 
      Height          =   650
      Left            =   480
      Picture         =   "frmCambiarContraseñaUsuarios.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   1
      TabStop         =   0   'False
      ToolTipText     =   "Aceptar"
      Top             =   2880
      Width           =   650
   End
   Begin VB.CommandButton cmdCancelar 
      Height          =   650
      Left            =   1560
      Picture         =   "frmCambiarContraseñaUsuarios.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   0
      TabStop         =   0   'False
      ToolTipText     =   "Cancelar"
      Top             =   2880
      Width           =   650
   End
   Begin VB.Label Label1 
      Caption         =   "Contraseña Actual:"
      Height          =   255
      Left            =   360
      TabIndex        =   7
      Top             =   120
      Width           =   1455
   End
   Begin VB.Label Label2 
      Caption         =   "Confirmar nueva contraseña:"
      Height          =   255
      Left            =   360
      TabIndex        =   6
      Top             =   2040
      Width           =   2175
   End
   Begin VB.Label Label3 
      Caption         =   "Contraseña nueva:"
      Height          =   255
      Left            =   360
      TabIndex        =   5
      Top             =   1080
      Width           =   1455
   End
End
Attribute VB_Name = "frmCambiarContraseñaUsuarios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAceptar_Click()
    frmIdentificacion.ContraseñaAComprobar = txtActual
    frmIdentificacion.ComprobarContraseña
    If frmIdentificacion.adoUsuarios.Recordset!Contraseña <> frmIdentificacion.ContraseñaVerdadera Then
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
    frmIdentificacion.NuevaContraseña = txtNueva
    frmIdentificacion.GuardarNuevaContraseña
    Unload Me
    frmIdentificacion.dtcUsuarios_Change
End Sub

Private Sub cmdCancelar_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
    txtActual.SetFocus
End Sub

Private Sub txtActual_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        frmIdentificacion.ContraseñaAComprobar = txtActual
        frmIdentificacion.ComprobarContraseña
        If frmIdentificacion.adoUsuarios.Recordset!Contraseña <> frmIdentificacion.ContraseñaVerdadera Then
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

