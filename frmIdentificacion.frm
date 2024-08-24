VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Begin VB.Form frmIdentificacion 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Identificación"
   ClientHeight    =   3030
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3360
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3030
   ScaleWidth      =   3360
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   2895
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   3135
      Begin Crystal.CrystalReport CrystalReport1 
         Left            =   120
         Top             =   2160
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   262150
         ReportFileName  =   "D:\Proyecto\Instituto\recibo.rpt"
      End
      Begin VB.CommandButton cmdIngresar 
         Height          =   615
         Left            =   1080
         Picture         =   "frmIdentificacion.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Aceptar"
         Top             =   2160
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.CommandButton cmdFin 
         Height          =   615
         Left            =   2400
         Picture         =   "frmIdentificacion.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Cancelar"
         Top             =   2160
         Width           =   615
      End
      Begin VB.CommandButton cmdCambiar 
         Caption         =   "Cambiar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1920
         TabIndex        =   5
         Top             =   1560
         Visible         =   0   'False
         Width           =   1095
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
         Height          =   375
         IMEMode         =   3  'DISABLE
         Left            =   120
         MaxLength       =   6
         PasswordChar    =   "*"
         TabIndex        =   4
         Top             =   1560
         Visible         =   0   'False
         Width           =   1695
      End
      Begin MSDataListLib.DataCombo dtcUsuarios 
         Bindings        =   "frmIdentificacion.frx":0884
         Height          =   360
         Left            =   120
         TabIndex        =   1
         Top             =   600
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   741
         _Version        =   393216
         Style           =   2
         ListField       =   "Identificacion"
         BoundColumn     =   "Usuario"
         Text            =   "Elija el Usuario"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSAdodcLib.Adodc adoUsuarios 
         Height          =   330
         Left            =   1080
         Top             =   120
         Visible         =   0   'False
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   582
         ConnectMode     =   1
         CursorLocation  =   3
         IsolationLevel  =   -1
         ConnectionTimeout=   15
         CommandTimeout  =   30
         CursorType      =   2
         LockType        =   1
         CommandType     =   1
         CursorOptions   =   0
         CacheSize       =   50
         MaxRecords      =   0
         BOFAction       =   0
         EOFAction       =   0
         ConnectStringType=   3
         Appearance      =   1
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Orientation     =   0
         Enabled         =   -1
         Connect         =   "DSN=Instituto"
         OLEDBString     =   ""
         OLEDBFile       =   ""
         DataSourceName  =   "Instituto"
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   "SELECT * FROM Usuarios WHERE Eliminado = 0 ORDER BY Identificacion"
         Caption         =   "Usuarios"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _Version        =   393216
      End
      Begin VB.Label Label1 
         Caption         =   "Usuario:"
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
         Left            =   120
         TabIndex        =   2
         Top             =   360
         Width           =   1695
      End
      Begin VB.Label lblContraseña 
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
         Left            =   120
         TabIndex        =   3
         Top             =   1200
         Visible         =   0   'False
         Width           =   1695
      End
   End
End
Attribute VB_Name = "frmIdentificacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Const WS_EX_APPWINDOW As Long = &H40000
Private Const GWL_EXSTYLE As Long = (-20)
Private Const SW_HIDE As Long = 0
Private Const SW_SHOW As Long = 5
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
Private mbActivated As Boolean

Dim Contraseña As String
Public NuevaContraseña As String
Public ContraseñaVerdadera As String
Public ContraseñaAComprobar As String
Dim Caracter As String
Dim Modo1(64) As String
Dim Modo2(64) As String
Dim Modo3(64) As String
Dim Modo4(64) As String
Dim Correcto(64) As String
Dim ConectarPermisos As New Connection
Dim Ejecutar As New Connection
Public Permisos As New Recordset



Private Sub cmdCambiar_Click()
    frmCambiarContraseñaUsuarios.Show 1
End Sub

Private Sub cmdFin_Click()
    End
End Sub

Private Sub cmdIngresar_Click()
    ContraseñaAComprobar = txtContraseña
    ComprobarContraseña
    If Contraseña <> adoUsuarios.Recordset!Contraseña Then
        MsgBox ("Contraseña Incorrecta")
        txtContraseña = ""
        txtContraseña.SetFocus
        Exit Sub
    Else
        Set Permisos = ConectarPermisos.Execute("SELECT * FROM Permisos WHERE Usuario =" & dtcUsuarios.BoundText)
        If dtcUsuarios.BoundText = 14 Then
            With frmMenuPrincipal
                .mnAlumnos.Visible = False
                .mnPersonal.Visible = False
                .mnCarreras.Visible = False
                .mnHorarios.Visible = False
                .mnExamenes.Visible = False
                .mnCertificados.Visible = False
                .mnCooperadora.Visible = False
                .mnConfiguracion.Visible = False
                .mnInformesYControles.Visible = False
            End With
        End If
        frmMenuPrincipal.Show 1
    End If
End Sub





Public Sub dtcUsuarios_Change()
    adoUsuarios.Recordset.MoveFirst
    adoUsuarios.Recordset.Find ("Usuario=" & dtcUsuarios.BoundText)
    lblContraseña.Visible = True
    txtContraseña.Visible = True
    cmdCambiar.Visible = True
    cmdIngresar.Visible = True
    txtContraseña = ""
    txtContraseña.SetFocus
End Sub

Public Function ComprobarContraseña()
    Contraseña = ""
    Caracter = ""
    With adoUsuarios.Recordset
    For i = 1 To Len(ContraseñaAComprobar)
        Caracter = Mid(ContraseñaAComprobar, i, 1)
        'ubico a que posicion del orden correcto pertenece este caracter
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
    ContraseñaVerdadera = Contraseña
End Function

Private Sub Form_Activate()
If Not m_bActivated Then
    m_bActivated = True
    Call SetWindowLong(hWnd, GWL_EXSTYLE, GetWindowLong(hWnd, GWL_EXSTYLE) Or WS_EX_APPWINDOW)
    Call ShowWindow(hWnd, SW_HIDE)
    Call ShowWindow(hWnd, SW_SHOW)
End If
End Sub

Private Sub Form_Load()
    ConectarPermisos.ConnectionString = ("DSN=Instituto")
    Ejecutar.ConnectionString = ("DSN=Instituto")
    ConectarPermisos.Open
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
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ConectarPermisos.Close
End Sub

Public Function GuardarNuevaContraseña()
    Contraseña = ""
    Randomize   ' Inicializa el generador de números aleatorios.
    Modo = Int((4 * Rnd) + 1)   ' Genera valores aleatorios entre 1 y 4.
    For i = 1 To Len(NuevaContraseña)
        Caracter = Mid(NuevaContraseña, i, 1)
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
    Ejecutar.Open
    Ejecutar.Execute ("UPDATE Usuarios SET Contraseña = '" & Contraseña & "', Modo =" & Modo & "  WHERE Usuario=" & dtcUsuarios.BoundText)
    Ejecutar.Close
    UsuarioActual = dtcUsuarios.BoundText
    adoUsuarios.Refresh
    dtcUsuarios.BoundText = UsuarioActual
End Function


Private Sub txtContraseña_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cmdIngresar_Click
        Exit Sub
    End If
    If ((KeyAscii >= 48 And KeyAscii <= 57) Or (KeyAscii >= 65 And KeyAscii <= 90) Or (KeyAscii >= 97 And KeyAscii <= 122)) And Len(ContraseñaParcial) < 6 Or KeyAscii = 8 Then
   
    Else
        KeyAscii = 0
    End If
End Sub
