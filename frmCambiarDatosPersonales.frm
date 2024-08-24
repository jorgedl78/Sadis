VERSION 5.00
Begin VB.Form frmCambiarDatosPersonales 
   BackColor       =   &H00C0FFFF&
   ClientHeight    =   6570
   ClientLeft      =   2385
   ClientTop       =   1770
   ClientWidth     =   7425
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   6570
   ScaleWidth      =   7425
   Begin VB.CommandButton cmdConfirmar 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Confirmar"
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
      Left            =   2760
      MouseIcon       =   "frmCambiarDatosPersonales.frx":0000
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   5880
      Width           =   1935
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H0080FFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3975
      Left            =   240
      TabIndex        =   0
      Top             =   1440
      Width           =   6855
      Begin VB.TextBox txtDomicilioEnJunin 
         Height          =   375
         Left            =   1800
         MaxLength       =   30
         TabIndex        =   12
         Top             =   2040
         Width           =   3855
      End
      Begin VB.TextBox txtCodigoPostal 
         Height          =   375
         Left            =   1800
         MaxLength       =   10
         TabIndex        =   11
         Top             =   1560
         Width           =   1455
      End
      Begin VB.TextBox txtEmail 
         Height          =   375
         Left            =   1800
         MaxLength       =   50
         TabIndex        =   10
         Top             =   3000
         Width           =   4815
      End
      Begin VB.TextBox txtTelefono 
         Height          =   375
         Left            =   1800
         MaxLength       =   20
         TabIndex        =   9
         Top             =   2520
         Width           =   3015
      End
      Begin VB.TextBox txtLocalidad 
         Height          =   375
         Left            =   1800
         MaxLength       =   20
         TabIndex        =   8
         Top             =   1080
         Width           =   2295
      End
      Begin VB.TextBox txtDomicilio 
         Height          =   375
         Left            =   1800
         MaxLength       =   30
         TabIndex        =   7
         Top             =   600
         Width           =   3855
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "E-Mail:"
         Height          =   255
         Left            =   1080
         TabIndex        =   6
         Top             =   3120
         Width           =   615
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Celular:"
         Height          =   255
         Left            =   1080
         TabIndex        =   5
         Top             =   2640
         Width           =   615
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Teléfonos:"
         Height          =   255
         Left            =   360
         TabIndex        =   4
         Top             =   2160
         Width           =   1335
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Cod. Postal:"
         Height          =   255
         Left            =   840
         TabIndex        =   3
         Top             =   1680
         Width           =   855
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Localidad:"
         Height          =   255
         Left            =   960
         TabIndex        =   2
         Top             =   1200
         Width           =   735
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Domicilio:"
         Height          =   255
         Left            =   960
         TabIndex        =   1
         Top             =   720
         Width           =   735
      End
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "A continuación le solicitamos que actualice sus datos personales a efectos de mantener un óptimo contacto."
      Height          =   495
      Left            =   240
      TabIndex        =   15
      Top             =   840
      Width           =   7215
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Confirmación de sus datos personales"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   375
      Left            =   840
      TabIndex        =   14
      Top             =   120
      Width           =   5655
   End
End
Attribute VB_Name = "frmCambiarDatosPersonales"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Conexion As New Connection
Dim DatosPersonales As New Recordset
Public miyu As String

Private Sub cmdConfirmar_Click()
   Conexion.Open
   Conexion.Execute ("UPDATE Alumnos SET Alumnos.Domicilio = '" & txtDomicilio & "', Alumnos.Localidad = '" & txtLocalidad & "', Alumnos.Postal = '" & txtCodigoPostal & "', Alumnos.DomicilioEnJunin = '" & txtDomicilioEnJunin & "', Alumnos.Telefono = '" & txtTelefono & "', Alumnos.Correo = '" & txtEmail & "' WHERE (((Alumnos.Permiso)=" & frmConexionAlumnos.adoAlumnos.Recordset!Permiso & "))")
   Conexion.Close
   Unload Me
End Sub

Private Sub Form_Load()
    Conexion.ConnectionString = ("DSN=Instituto")
    Conexion.Open
    Set DatosPersonales = Conexion.Execute("SELECT Alumnos.Domicilio, Alumnos.Localidad, Alumnos.Postal, Alumnos.DomicilioEnJunin, Alumnos.Telefono, Alumnos.Correo From Alumnos WHERE (((Alumnos.Permiso)=" & frmConexionAlumnos.adoAlumnos.Recordset!Permiso & "))")
    If DatosPersonales!domicilio <> vacio Then
       txtDomicilio = DatosPersonales!domicilio
    Else
       txtDomicilio = ""
    End If
    If DatosPersonales!Localidad <> vacio Then
       txtLocalidad = DatosPersonales!Localidad
    Else
       txtLocalidad = ""
    End If
    If DatosPersonales!Postal <> vacio Then
       txtCodigoPostal = DatosPersonales!Postal
    Else
       txtCodigoPostal = ""
    End If
    If DatosPersonales!DomicilioEnJunin <> vacio Then
       txtDomicilioEnJunin = DatosPersonales!DomicilioEnJunin
    Else
       txtDomicilioEnJunin = ""
    End If
    If DatosPersonales!telefono <> vacio Then
       txtTelefono = DatosPersonales!telefono
    Else
       txtTelefono = ""
    End If
    If DatosPersonales!Correo <> vacio Then
       txtEmail = DatosPersonales!Correo
    Else
       txtEmail = ""
    End If
    Conexion.Close
End Sub


