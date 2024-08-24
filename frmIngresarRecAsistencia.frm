VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmIngresarRecAsistencia 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Recuperatorio de Asistencia"
   ClientHeight    =   4755
   ClientLeft      =   6810
   ClientTop       =   3810
   ClientWidth     =   6120
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4755
   ScaleWidth      =   6120
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Height          =   2295
      Left            =   240
      TabIndex        =   4
      Top             =   1200
      Width           =   5535
      Begin VB.TextBox txtNota 
         Height          =   375
         Left            =   2160
         TabIndex        =   8
         Top             =   1080
         Width           =   975
      End
      Begin VB.CheckBox chkAprobo 
         Caption         =   "Check1"
         Height          =   255
         Left            =   2280
         TabIndex        =   7
         Top             =   1680
         Width           =   255
      End
      Begin MSComCtl2.DTPicker DTPFechaRecuperatorio 
         Height          =   375
         Left            =   2160
         TabIndex        =   5
         Top             =   480
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         Format          =   51838977
         CurrentDate     =   40127
      End
      Begin VB.Label Label2 
         Caption         =   "Nota:"
         Height          =   255
         Left            =   1560
         TabIndex        =   10
         Top             =   1200
         Width           =   495
      End
      Begin VB.Label Label3 
         Caption         =   "Aprobó:"
         Height          =   255
         Left            =   1440
         TabIndex        =   9
         Top             =   1680
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Fecha del Recuperatorio:"
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   480
         Width           =   1935
      End
   End
   Begin VB.CommandButton cmdConfirmar 
      Height          =   615
      Left            =   1800
      Picture         =   "frmIngresarRecAsistencia.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Aceptar"
      Top             =   3840
      Width           =   855
   End
   Begin VB.CommandButton cmdCancelar 
      Height          =   615
      Left            =   3480
      Picture         =   "frmIngresarRecAsistencia.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Cancelar"
      Top             =   3840
      Width           =   855
   End
   Begin VB.Label lblAlumno 
      Caption         =   "Label1"
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
      Left            =   360
      TabIndex        =   3
      Top             =   240
      Width           =   5295
   End
   Begin VB.Label lblMateria 
      Caption         =   "Label1"
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
      Left            =   360
      TabIndex        =   2
      Top             =   720
      Width           =   5295
   End
End
Attribute VB_Name = "frmIngresarRecAsistencia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim CursadaNumero As Integer
Dim NotaAprobacion As Double
Dim Conectar As New Connection
Dim Resultado As New Recordset

Private Sub cmdIngresarNota_Click()

End Sub

Private Sub cmdCancelar_Click()
    Unload Me
End Sub

Private Sub cmdConfirmar_Click()
    If Val(txtNota) < NotaAprobacion And chkAprobo = 1 Then
       MsgBox ("La nota no corresponde para aprobar el exámen")
       Exit Sub
    End If
    If Val(txtNota) >= NotaAprobacion And chkAprobo = 0 Then
       MsgBox ("La nota determina que el exámen ha sido aprobado")
       Exit Sub
    End If
    Conectar.Open
    Set NumeroCursada = Conectar.Execute("SELECT Divisiones.Numero From Divisiones WHERE Divisiones.Materia=" & frmParciales.dtcMaterias.BoundText & " AND Divisiones.Ano=" & frmParciales.txtAño & " AND Divisiones.Division=" & frmParciales.cbDivision)
    CursadaNumero = NumeroCursada!Numero
    Conectar.Execute ("INSERT INTO Recuperatorios_de_Asistencia ( Numero, Permiso, Nota, Aprobo, Fecha_recuperatorio, Fecha_registracion, Usuario ) values (" & NumeroCursada!Numero & "," & frmParciales.adoMatriculados.Recordset!Permiso & "," & Val(txtNota) & "," & chkAprobo & ",'" & DateValue(DTPFechaRecuperatorio.Value) & "','" & Date & "'," & frmIdentificacion.adoUsuarios.Recordset!Usuario & ")")
    If chkAprobo = 1 Then
        Conectar.Execute ("UPDATE FINALES SET Asistencia=1 WHERE Alumno=" & frmParciales.adoMatriculados.Recordset!Permiso & " AND Materia = " & frmParciales.dtcMaterias.BoundText & " and ano=" & frmParciales.txtAño)
        AlumnoActual = frmParciales.adoMatriculados.Recordset.Bookmark
        frmParciales.adoMatriculados.Refresh
        frmParciales.adoMatriculados.Recordset.Bookmark = AlumnoActual
    End If
    Conectar.Close
    Unload Me
End Sub

Private Sub Form_Activate()
    lblAlumno = frmParciales.adoMatriculados.Recordset!Nombre
    lblMateria = frmParciales.dtcMaterias
    Conectar.Open
    Set Resultado = Conectar.Execute("SELECT NotaAprobacionFinal FROM Parametros")
    NotaAprobacion = Resultado!NotaAprobacionFinal
    Conectar.Close
End Sub

Private Sub Form_Load()
    Conectar.ConnectionString = ("DSN=Instituto")
    DTPFechaRecuperatorio = Date
End Sub

Private Sub txtNota_KeyPress(KeyAscii As Integer)
    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 And KeyAscii <> 46 Then KeyAscii = 0
End Sub
