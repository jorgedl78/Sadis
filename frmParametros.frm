VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmParametros 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Parámetros del Sistema"
   ClientHeight    =   8925
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7935
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   8925
   ScaleWidth      =   7935
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Caption         =   "Control de Autogestión"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   120
      TabIndex        =   24
      Top             =   6120
      Width           =   7575
      Begin VB.CommandButton cmdDesbloquearAtogestion 
         Caption         =   "Desbloquear Autogestión"
         Height          =   615
         Left            =   3840
         TabIndex        =   26
         Top             =   480
         Width           =   1575
      End
      Begin VB.CommandButton cmdBloquearAutogestion 
         Caption         =   "Bloquear Autogestión"
         Height          =   615
         Left            =   1320
         TabIndex        =   25
         Top             =   480
         Width           =   1575
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1335
      Left            =   120
      TabIndex        =   13
      Top             =   7440
      Width           =   7575
      Begin VB.CommandButton cmdSalir 
         Height          =   705
         Left            =   4800
         Picture         =   "frmParametros.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   15
         ToolTipText     =   "Salir"
         Top             =   360
         Width           =   720
      End
      Begin VB.CommandButton cmdModificar 
         Caption         =   "Modificar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   960
         TabIndex        =   14
         Top             =   480
         Width           =   1335
      End
   End
   Begin VB.Frame frParametros 
      Caption         =   "Parametros"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6015
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7575
      Begin VB.CheckBox chkHabilitarMatriculacion 
         Caption         =   "Check1"
         Height          =   255
         Left            =   4560
         TabIndex        =   23
         Top             =   5400
         Width           =   255
      End
      Begin MSComCtl2.DTPicker dtpLimiteAsistencia 
         Height          =   495
         Left            =   4560
         TabIndex        =   20
         Top             =   4680
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   873
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   46530561
         CurrentDate     =   40121
      End
      Begin VB.CheckBox chkControlaTurno 
         Height          =   195
         Left            =   4560
         TabIndex        =   19
         Top             =   3840
         Width           =   255
      End
      Begin VB.TextBox txtAñoTurnoControl 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   6600
         TabIndex        =   18
         Text            =   "2002"
         Top             =   4080
         Width           =   855
      End
      Begin VB.TextBox txtPorcentajeAsistencia 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   4560
         TabIndex        =   12
         Text            =   "80.00"
         Top             =   3240
         Width           =   855
      End
      Begin VB.TextBox txtNotaPromocion 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   4560
         TabIndex        =   11
         Text            =   "6.00"
         Top             =   2640
         Width           =   855
      End
      Begin VB.TextBox txtNotaAprobacion 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   4560
         TabIndex        =   10
         Text            =   "4.00"
         Top             =   2040
         Width           =   855
      End
      Begin VB.TextBox txtAñoMAtriculacion 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   4560
         TabIndex        =   6
         Text            =   "2002"
         Top             =   1440
         Width           =   855
      End
      Begin VB.TextBox txtAñoLlamado 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   4560
         TabIndex        =   4
         Text            =   "2002"
         Top             =   840
         Width           =   855
      End
      Begin MSDataListLib.DataCombo dtcMeses 
         Bindings        =   "frmParametros.frx":0442
         Height          =   420
         Left            =   4560
         TabIndex        =   2
         Top             =   360
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   741
         _Version        =   393216
         Style           =   2
         ListField       =   "Nombre"
         BoundColumn     =   "Numero"
         Text            =   ""
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
      Begin MSAdodcLib.Adodc adoMeses 
         Height          =   330
         Left            =   4680
         Top             =   120
         Visible         =   0   'False
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   582
         ConnectMode     =   0
         CursorLocation  =   3
         IsolationLevel  =   -1
         ConnectionTimeout=   15
         CommandTimeout  =   30
         CursorType      =   3
         LockType        =   3
         CommandType     =   2
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
         RecordSource    =   "Meses"
         Caption         =   "Meses"
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
      Begin MSDataListLib.DataCombo dtcMesesTurnoControl 
         Bindings        =   "frmParametros.frx":0459
         Height          =   420
         Left            =   4560
         TabIndex        =   17
         Top             =   4080
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   741
         _Version        =   393216
         Style           =   2
         ListField       =   "Nombre"
         BoundColumn     =   "Numero"
         Text            =   ""
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
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         Caption         =   "Habilitar matriculación:"
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
         Left            =   1320
         TabIndex        =   22
         Top             =   5400
         Width           =   3015
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         Caption         =   "Límite de Asistencia:"
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
         TabIndex        =   21
         Top             =   4800
         Width           =   4215
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Caption         =   "No permitir Inscribir en Final si se anoto en el turno"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   120
         TabIndex        =   16
         Top             =   3840
         Width           =   4215
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "Porcentaje para Aprobar la Asistencia: %"
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
         Left            =   240
         TabIndex        =   9
         Top             =   3360
         Width           =   4335
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Nota de Promoción sin Exámen:"
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
         Left            =   120
         TabIndex        =   8
         Top             =   2760
         Width           =   4215
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Nota de Aprobación:"
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
         Left            =   120
         TabIndex        =   7
         Top             =   2160
         Width           =   4215
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Año de Matriculación:"
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
         Left            =   120
         TabIndex        =   5
         Top             =   1440
         Width           =   4215
      End
      Begin VB.Label lblAñoDeLlamado 
         Alignment       =   1  'Right Justify
         Caption         =   "Año de Llamado a Exámenes Finales:"
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
         Left            =   120
         TabIndex        =   3
         Top             =   840
         Width           =   4215
      End
      Begin VB.Label lblTurnoLLamado 
         Alignment       =   1  'Right Justify
         Caption         =   "Turno de Llamado a Exámenes Finales:"
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
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   4215
      End
   End
End
Attribute VB_Name = "frmParametros"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Conexion As New Connection
Dim Parametros As New Recordset

Private Sub cmdBloquearAutogestion_Click()
    Respuesta = MsgBox("A continuación va a bloquear la autogestión para todos los alumnos", vbYesNo, "Atención!!")
    If Respuesta = vbNo Then
        Exit Sub
    Else
        Me.MousePointer = 11
        Conexion.Open
        Conexion.Execute ("UPDATE Alumnos SET BloquearAutogestion=True")
        Conexion.Close
        Me.MousePointer = 1
    End If
End Sub

Private Sub cmdDesbloquearAtogestion_Click()
    Respuesta = MsgBox("A continuación va a desbloquear la autogestión para todos los alumnos", vbYesNo, "Atención!!")
    If Respuesta = vbNo Then
        Exit Sub
    Else
        Me.MousePointer = 11
        Conexion.Open
        Conexion.Execute ("UPDATE Alumnos SET BloquearAutogestion=False")
        Conexion.Close
        Me.MousePointer = 1
    End If
End Sub

Private Sub cmdModificar_Click()
    If cmdModificar.Caption = "Modificar" Then
        cmdModificar.Caption = "Guardar"
        frParametros.Enabled = True
        cmdSalir.Enabled = False
    Else 'guardar
        Conexion.Open
        If chkControlaTurno.Value = True Then
            controlaturno = "True"
        Else
            controlaturno = "False"
        End If
        Conexion.Execute ("UPDATE Parametros SET Parametros.TurnoLlamado = " & dtcMeses.BoundText & ", Parametros.AñoLlamado = " & Val(txtAñoLlamado) & ", Parametros.AñoMatriculacion = " & Val(txtAñoMAtriculacion) & ", Parametros.NotaAprobacionFinal = " & Replace(Val(txtNotaAprobacion), ",", ".") & ", Parametros.NotaPromocionFinal = " & Replace(Val(txtNotaPromocion), ",", ".") & ", Parametros.PorcentajeAsistencia = " & Replace(Val(txtPorcentajeAsistencia), ",", ".") & ",ControlaInscripcionAnterior=" & chkControlaTurno.Value & ", TurnoControl = " & dtcMesesTurnoControl.BoundText & ", AñoControl = " & txtAñoTurnoControl & ", Limite_Asistencia = '" & DateValue(dtpLimiteAsistencia.Value) & "', Habilitar_matriculacion=" & chkHabilitarMatriculacion.Value)
        Conexion.Close
        frParametros.Enabled = False
        cmdSalir.Enabled = True
        cmdModificar.Caption = "Modificar"
        LevantarDatos
    End If
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Conexion.ConnectionString = ("DSN=Instituto")
    LevantarDatos
End Sub

Private Function LevantarDatos()
    Conexion.Open
    Set Parametros = Conexion.Execute("SELECT * FROM Parametros")
    With Parametros
    dtcMeses.BoundText = !TurnoLlamado
    txtAñoLlamado = !AñoLlamado
    txtAñoMAtriculacion = !AñoMatriculacion
    txtNotaAprobacion = Format(!NotaAprobacionFinal, "0.00")
    txtNotaPromocion = Format(!NotaPromocionFinal, "0.00")
    txtPorcentajeAsistencia = Format(!PorcentajeAsistencia, "0.00")
    If !ControlaInscripcionAnterior = True Then
        chkControlaTurno.Value = 1
    Else
        chkControlaTurno.Value = 0
    End If
    If !Habilitar_matriculacion = True Then
       chkHabilitarMatriculacion = 1
    Else
        chkHabilitarMatriculacion = 0
    End If
    dtcMesesTurnoControl.BoundText = !TurnoControl
    txtAñoTurnoControl = !AñoControl
    dtpLimiteAsistencia.Value = !Limite_Asistencia
    End With
    Conexion.Close
End Function
