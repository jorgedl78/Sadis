VERSION 5.00
Begin VB.Form frmLoginAlumno 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Identificación"
   ClientHeight    =   3195
   ClientLeft      =   5070
   ClientTop       =   3045
   ClientWidth     =   2490
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   2490
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdAceptar 
      Height          =   650
      Left            =   840
      Picture         =   "frmLoginAlumno.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   2160
      Width           =   650
   End
   Begin VB.TextBox txtDocumento 
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
      Left            =   480
      MaxLength       =   8
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   480
      Width           =   1455
   End
   Begin VB.TextBox txtContraseña 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      IMEMode         =   3  'DISABLE
      Left            =   480
      MaxLength       =   6
      PasswordChar    =   "*"
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   1560
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label lblContreseña 
      Caption         =   "Contraseña:"
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
      TabIndex        =   3
      Top             =   1200
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "Documento:"
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
      TabIndex        =   2
      Top             =   120
      Width           =   1455
   End
End
Attribute VB_Name = "frmLoginAlumno"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Contraseña As String
Public ContraseñaReal As String
Dim Correcto(64) As String
Dim Modo1(64) As String
Dim Modo2(64) As String
Dim Modo3(64) As String
Dim Modo4(64) As String
Dim Caracter As String
Dim NuevoUsuario As String

Private Sub cmdAceptar_Click()
If txtDocumento = "" Then txtDocumento.SetFocus: Exit Sub
If txtDocumento = "1" Then
    frmConexionAlumnos.SalirPorCompleto = "Si"
    Unload Me
    Exit Sub
End If
'******* Est lo uso por si algun alumno en particular lo bloqueo para ingresar***
'If txtDocumento = 30928487 Or txtDocumento = 30573464 Or txtDocumento = 32772598 Or txtDocumento = 31526741 Or txtDocumento = 31813759 Or txtDocumento = 29652084 Or txtDocumento = 30573238 Or txtDocumento = 31062070 Or txtDocumento = 32988393 Or txtDocumento = 32195521 Or txtDocumento = 32456531 Or txtDocumento = 31730022 Or txtDocumento = 30928291 Or txtDocumento = 32527486 Or txtDocumento = 31062300 Or txtDocumento = 32209746 Or txtDocumento = 32209707 Or txtDocumento = 31114903 Or txtDocumento = 32773029 Or txtDocumento = 31941334 Or txtDocumento = 32066360 Or txtDocumento = 32772612 Then
'    Respuesta = MsgBox("Por favor consulte al personal administrativo", vbOKOnly, "Imposible acceder")
'    txtDocumento = ""
'    txtDocumento.SetFocus
'    Exit Sub
'End If

    With frmConexionAlumnos
    .adoParametros.Refresh
    If txtContraseña.Visible = False Then
        .adoAlumnos.RecordSource = "SELECT * FROM Alumnos WHERE Documento = " & Val(txtDocumento) & " AND Eliminado = False"
        .adoAlumnos.Refresh
        If .adoAlumnos.Recordset.RecordCount >= 1 Then
            If .adoAlumnos.Recordset!Contrasena <> "" Then
            Else
                NuevoUsuario = "Si"
                MsgBox ("A continuación deberá definir una contraseña " & Chr(13) & "utilizando las letras mayúsculas o mimnúsculas " & Chr(13) & "y los numeros del 0 al 9")
            End If
            lblContreseña.Visible = True
            txtContraseña.Visible = True
            txtContraseña.SetFocus
        Else
            MsgBox ("El documento no existe")
            txtDocumento = ""
            Exit Sub
        End If
    Else
        If NuevoUsuario = "Si" Then
            ContraseñaReal = txtContraseña
            GuardarContraseña
            NuevoUsuario = "No"
            cmdAceptar_Click
        Else
            VerContraseña
            If Contraseña <> .adoAlumnos.Recordset!Contrasena Then
                MsgBox ("Contraseña Incorrecta")
                Universal = txtContraseña
                txtContraseña = ""
                txtContraseña.SetFocus
                If Universal <> "Univer" Then 'para poder entrar sin saber la contraseña
                    Exit Sub
                End If
            End If
            ContraseñaReal = txtContraseña
            Unload Me
            .frDatosAlumno.Visible = True
            .frMenu.Visible = True
        End If
        'frmNewsAlumnos.Show 1
        'If adoAlumnos.Recordset!Contraseña Then
        '    MsgBox ("Consulte en administración por documentación faltante")
         '   Exit Sub
        'End If
        If .adoAlumnos.Recordset!BloquearAutogestion = -1 Then
            .lblNotificacion = "Documentación Faltante. Consultar en administración"
            .cmdInscripcionFinales.Enabled = False
            .cmdMatriculacion.Enabled = False
        End If
        frmCambiarDatosPersonales.Show 1
        'frmConexionAlumnos.ControlaCooperadora
    End If
    End With
End Sub



Private Sub Form_Load()
    NuevoUsuario = "No"
    Caracter = ""
    Correcto(1) = "0"
    Correcto(2) = "1"
    Correcto(3) = "2"
    Correcto(4) = "3"
    Correcto(5) = "4"
    Correcto(6) = "5"
    Correcto(7) = "6"
    Correcto(8) = "7"
    Correcto(9) = "8"
    Correcto(10) = "9"
    Correcto(11) = "a"
    Correcto(12) = "b"
    Correcto(13) = "c"
    Correcto(14) = "d"
    Correcto(15) = "e"
    Correcto(16) = "f"
    Correcto(17) = "g"
    Correcto(18) = "h"
    Correcto(19) = "i"
    Correcto(20) = "j"
    Correcto(21) = "k"
    Correcto(22) = "l"
    Correcto(23) = "m"
    Correcto(24) = "n"
    Correcto(25) = "ñ"
    Correcto(26) = "o"
    Correcto(27) = "p"
    Correcto(28) = "q"
    Correcto(29) = "r"
    Correcto(30) = "s"
    Correcto(31) = "t"
    Correcto(32) = "u"
    Correcto(33) = "v"
    Correcto(34) = "w"
    Correcto(35) = "x"
    Correcto(36) = "y"
    Correcto(37) = "z"
    Correcto(38) = "A"
    Correcto(39) = "B"
    Correcto(40) = "C"
    Correcto(41) = "D"
    Correcto(42) = "E"
    Correcto(43) = "F"
    Correcto(44) = "G"
    Correcto(45) = "H"
    Correcto(46) = "I"
    Correcto(47) = "J"
    Correcto(48) = "K"
    Correcto(49) = "L"
    Correcto(50) = "M"
    Correcto(51) = "N"
    Correcto(52) = "Ñ"
    Correcto(53) = "O"
    Correcto(54) = "P"
    Correcto(55) = "Q"
    Correcto(56) = "R"
    Correcto(57) = "S"
    Correcto(58) = "T"
    Correcto(59) = "U"
    Correcto(60) = "V"
    Correcto(61) = "W"
    Correcto(62) = "X"
    Correcto(63) = "Y"
    Correcto(64) = "Z"
'MODO1
    Modo1(1) = "H"
    Modo1(2) = "u"
    Modo1(3) = "2"
    Modo1(4) = "f"
    Modo1(5) = "q"
    Modo1(6) = "c"
    Modo1(7) = "w"
    Modo1(8) = "K"
    Modo1(9) = "b"
    Modo1(10) = "k"
    Modo1(11) = "t"
    Modo1(12) = "i"
    Modo1(13) = "5"
    Modo1(14) = "A"
    Modo1(15) = "7"
    Modo1(16) = "X"
    Modo1(17) = "l"
    Modo1(18) = "n"
    Modo1(19) = "y"
    Modo1(20) = "6"
    Modo1(21) = "v"
    Modo1(22) = "R"
    Modo1(23) = "E"
    Modo1(24) = "9"
    Modo1(25) = "L"
    Modo1(26) = "0"
    Modo1(27) = "p"
    Modo1(28) = "Z"
    Modo1(29) = "U"
    Modo1(30) = "e"
    Modo1(31) = "F"
    Modo1(32) = "Q"
    Modo1(33) = "x"
    Modo1(34) = "G"
    Modo1(35) = "m"
    Modo1(36) = "d"
    Modo1(37) = "W"
    Modo1(38) = "C"
    Modo1(39) = "D"
    Modo1(40) = "z"
    Modo1(41) = "Y"
    Modo1(42) = "8"
    Modo1(43) = "s"
    Modo1(44) = "I"
    Modo1(45) = "1"
    Modo1(46) = "V"
    Modo1(47) = "B"
    Modo1(48) = "h"
    Modo1(49) = "3"
    Modo1(50) = "g"
    Modo1(51) = "ñ"
    Modo1(52) = "J"
    Modo1(53) = "r"
    Modo1(54) = "S"
    Modo1(55) = "4"
    Modo1(56) = "j"
    Modo1(57) = "M"
    Modo1(58) = "N"
    Modo1(59) = "a"
    Modo1(60) = "O"
    Modo1(61) = "Ñ"
    Modo1(62) = "T"
    Modo1(63) = "o"
    Modo1(64) = "P"
'MODO2
    Modo2(1) = "z"
    Modo2(2) = "Ñ"
    Modo2(3) = "e"
    Modo2(4) = "y"
    Modo2(5) = "c"
    Modo2(6) = "9"
    Modo2(7) = "t"
    Modo2(8) = "Q"
    Modo2(9) = "s"
    Modo2(10) = "G"
    Modo2(11) = "O"
    Modo2(12) = "T"
    Modo2(13) = "5"
    Modo2(14) = "v"
    Modo2(15) = "d"
    Modo2(16) = "X"
    Modo2(17) = "I"
    Modo2(18) = "J"
    Modo2(19) = "h"
    Modo2(20) = "o"
    Modo2(21) = "M"
    Modo2(22) = "Y"
    Modo2(23) = "l"
    Modo2(24) = "K"
    Modo2(25) = "C"
    Modo2(26) = "D"
    Modo2(27) = "U"
    Modo2(28) = "F"
    Modo2(29) = "B"
    Modo2(30) = "8"
    Modo2(31) = "W"
    Modo2(32) = "V"
    Modo2(33) = "r"
    Modo2(34) = "ñ"
    Modo2(35) = "j"
    Modo2(36) = "7"
    Modo2(37) = "R"
    Modo2(38) = "6"
    Modo2(39) = "b"
    Modo2(40) = "3"
    Modo2(41) = "4"
    Modo2(42) = "1"
    Modo2(43) = "n"
    Modo2(44) = "0"
    Modo2(45) = "x"
    Modo2(46) = "L"
    Modo2(47) = "P"
    Modo2(48) = "a"
    Modo2(49) = "f"
    Modo2(50) = "A"
    Modo2(51) = "u"
    Modo2(52) = "2"
    Modo2(53) = "Z"
    Modo2(54) = "m"
    Modo2(55) = "p"
    Modo2(56) = "E"
    Modo2(57) = "w"
    Modo2(58) = "i"
    Modo2(59) = "S"
    Modo2(60) = "k"
    Modo2(61) = "q"
    Modo2(62) = "g"
    Modo2(63) = "N"
    Modo2(64) = "H"
'MODO3
    Modo3(1) = "F"
    Modo3(2) = "G"
    Modo3(3) = "B"
    Modo3(4) = "H"
    Modo3(5) = "g"
    Modo3(6) = "t"
    Modo3(7) = "l"
    Modo3(8) = "J"
    Modo3(9) = "M"
    Modo3(10) = "X"
    Modo3(11) = "D"
    Modo3(12) = "C"
    Modo3(13) = "m"
    Modo3(14) = "h"
    Modo3(15) = "x"
    Modo3(16) = "ñ"
    Modo3(17) = "s"
    Modo3(18) = "Ñ"
    Modo3(19) = "y"
    Modo3(20) = "8"
    Modo3(21) = "N"
    Modo3(22) = "c"
    Modo3(23) = "V"
    Modo3(24) = "2"
    Modo3(25) = "q"
    Modo3(26) = "0"
    Modo3(27) = "r"
    Modo3(28) = "6"
    Modo3(29) = "R"
    Modo3(30) = "j"
    Modo3(31) = "O"
    Modo3(32) = "4"
    Modo3(33) = "o"
    Modo3(34) = "E"
    Modo3(35) = "p"
    Modo3(36) = "A"
    Modo3(37) = "f"
    Modo3(38) = "1"
    Modo3(39) = "Z"
    Modo3(40) = "K"
    Modo3(41) = "u"
    Modo3(42) = "Q"
    Modo3(43) = "d"
    Modo3(44) = "9"
    Modo3(45) = "W"
    Modo3(46) = "3"
    Modo3(47) = "k"
    Modo3(48) = "i"
    Modo3(49) = "v"
    Modo3(50) = "5"
    Modo3(51) = "w"
    Modo3(52) = "U"
    Modo3(53) = "n"
    Modo3(54) = "b"
    Modo3(55) = "T"
    Modo3(56) = "L"
    Modo3(57) = "I"
    Modo3(58) = "7"
    Modo3(59) = "S"
    Modo3(60) = "e"
    Modo3(61) = "Y"
    Modo3(62) = "P"
    Modo3(63) = "z"
    Modo3(64) = "a"
'MODO4
    Modo4(1) = "Ñ"
    Modo4(2) = "I"
    Modo4(3) = "O"
    Modo4(4) = "E"
    Modo4(5) = "L"
    Modo4(6) = "C"
    Modo4(7) = "N"
    Modo4(8) = "6"
    Modo4(9) = "Q"
    Modo4(10) = "F"
    Modo4(11) = "J"
    Modo4(12) = "0"
    Modo4(13) = "y"
    Modo4(14) = "A"
    Modo4(15) = "U"
    Modo4(16) = "P"
    Modo4(17) = "u"
    Modo4(18) = "f"
    Modo4(19) = "K"
    Modo4(20) = "4"
    Modo4(21) = "z"
    Modo4(22) = "i"
    Modo4(23) = "Z"
    Modo4(24) = "x"
    Modo4(25) = "Y"
    Modo4(26) = "T"
    Modo4(27) = "w"
    Modo4(28) = "D"
    Modo4(29) = "W"
    Modo4(30) = "M"
    Modo4(31) = "H"
    Modo4(32) = "d"
    Modo4(33) = "8"
    Modo4(34) = "7"
    Modo4(35) = "v"
    Modo4(36) = "t"
    Modo4(37) = "p"
    Modo4(38) = "a"
    Modo4(39) = "S"
    Modo4(40) = "R"
    Modo4(41) = "l"
    Modo4(42) = "k"
    Modo4(43) = "j"
    Modo4(44) = "2"
    Modo4(45) = "ñ"
    Modo4(46) = "G"
    Modo4(47) = "n"
    Modo4(48) = "m"
    Modo4(49) = "r"
    Modo4(50) = "5"
    Modo4(51) = "X"
    Modo4(52) = "9"
    Modo4(53) = "g"
    Modo4(54) = "B"
    Modo4(55) = "o"
    Modo4(56) = "c"
    Modo4(57) = "e"
    Modo4(58) = "b"
    Modo4(59) = "V"
    Modo4(60) = "s"
    Modo4(61) = "q"
    Modo4(62) = "3"
    Modo4(63) = "h"
    Modo4(64) = "1"
    Exit Sub
End Sub

Private Sub lblContreseñaTexto_Click()
    txtContraseña.SetFocus
End Sub

Private Sub txtDocumento_Click()
    txtDocumento.Enabled = True
    txtDocumento.SetFocus
    txtContraseña = ""
    lblContreseña.Visible = False
    txtContraseña.Visible = False
End Sub

Private Sub txtContraseña_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cmdAceptar_Click
        Exit Sub
    End If
    If ((KeyAscii >= 48 And KeyAscii <= 57) Or (KeyAscii >= 65 And KeyAscii <= 90) Or (KeyAscii >= 97 And KeyAscii <= 122)) And Len(ContraseñaParcial) < 6 Or KeyAscii = 8 Then
   
    Else
        KeyAscii = 0
    End If
End Sub

Private Sub txtDocumento_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If txtDocumento = "" Then
            KeyAscii = 0: Exit Sub
        Else
            cmdAceptar_Click
        End If
    End If
    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 Then KeyAscii = 0
End Sub

Private Function VerContraseña()
    Contraseña = ""
    With frmConexionAlumnos.adoAlumnos.Recordset
    For i = 1 To Len(txtContraseña)
        Caracter = Mid(txtContraseña, i, 1)
        For j = 1 To 64
            If Correcto(j) = Caracter Then Ubicacion = j
        Next j
        If !Modo = 1 Then
            Contraseña = Contraseña & Modo1(Ubicacion)
        ElseIf !Modo = 2 Then
            Contraseña = Contraseña & Modo2(Ubicacion)
        ElseIf !Modo = 3 Then
            Contraseña = Contraseña & Modo3(Ubicacion)
        Else
            Contraseña = Contraseña & Modo4(Ubicacion)
        End If
    Next i
    End With
End Function

Public Function GuardarContraseña()
    Contraseña = ""
    Randomize   ' Inicializa el generador de números aleatorios.
    Modo = Int((4 * Rnd) + 1)   ' Genera valores aleatorios entre 1 y 4.
    For i = 1 To Len(ContraseñaReal)
        Caracter = Mid(ContraseñaReal, i, 1)
        For j = 1 To 64
            If Correcto(j) = Caracter Then Ubicacion = j
        Next j
        If Modo = 1 Then
            Contraseña = Contraseña & Modo1(Ubicacion)
        ElseIf Modo = 2 Then
            Contraseña = Contraseña & Modo2(Ubicacion)
        ElseIf Modo = 3 Then
            Contraseña = Contraseña & Modo3(Ubicacion)
        Else
            Contraseña = Contraseña & Modo4(Ubicacion)
        End If
    Next i
    With frmConexionAlumnos.adoAlumnos.Recordset
    !Contrasena = Contraseña
    !Modo = Modo
    .Update
    End With
End Function
