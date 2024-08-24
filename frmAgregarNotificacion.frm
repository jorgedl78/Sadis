VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmAgregarNotificacion 
   Caption         =   "Nueva Notificación"
   ClientHeight    =   5355
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   8730
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   Moveable        =   0   'False
   ScaleHeight     =   5355
   ScaleWidth      =   8730
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      Height          =   1455
      Left            =   240
      TabIndex        =   5
      Top             =   120
      Width           =   8175
      Begin MSComCtl2.DTPicker dtpFechaNotificacion 
         Height          =   375
         Left            =   3480
         TabIndex        =   7
         Top             =   600
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         Format          =   142868481
         CurrentDate     =   37672
      End
      Begin MSComCtl2.DTPicker dtpFechaCaducidad 
         Height          =   375
         Left            =   6480
         TabIndex        =   8
         Top             =   600
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         Format          =   142868481
         CurrentDate     =   37672
      End
      Begin VB.Label lblMateria 
         Caption         =   "lblMateria"
         Height          =   255
         Left            =   4680
         TabIndex        =   13
         Top             =   1080
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Label lblPermiso 
         Caption         =   "lblPermiso"
         Height          =   255
         Left            =   3240
         TabIndex        =   12
         Top             =   1080
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.Label lblTipoDeNotificacion 
         Alignment       =   2  'Center
         Caption         =   "lblTipoDeNotificacion"
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
         Left            =   480
         TabIndex        =   11
         Top             =   720
         Width           =   2175
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Caption         =   "Tipo de Notificación"
         Height          =   255
         Left            =   480
         TabIndex        =   10
         Top             =   480
         Width           =   2175
      End
      Begin VB.Label Label2 
         Caption         =   "Hasta el:"
         Height          =   255
         Left            =   6480
         TabIndex        =   9
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Desde el:"
         Height          =   255
         Left            =   3480
         TabIndex        =   6
         Top             =   360
         Width           =   1095
      End
   End
   Begin VB.Frame Frame2 
      Height          =   1215
      Left            =   240
      TabIndex        =   1
      Top             =   3960
      Width           =   8175
      Begin VB.CommandButton cmdAceptar 
         Height          =   650
         Left            =   1800
         Picture         =   "frmAgregarNotificacion.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   3
         TabStop         =   0   'False
         ToolTipText     =   "Aceptar"
         Top             =   360
         Width           =   650
      End
      Begin VB.CommandButton cmdCancelar 
         Height          =   650
         Left            =   4800
         Picture         =   "frmAgregarNotificacion.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   2
         TabStop         =   0   'False
         ToolTipText     =   "Cancelar"
         Top             =   360
         Width           =   650
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Descripción (Máximo 100 caracteres)"
      Height          =   2055
      Left            =   240
      TabIndex        =   0
      Top             =   1680
      Width           =   8175
      Begin VB.TextBox txtNotificacion 
         Height          =   1455
         Left            =   240
         MaxLength       =   100
         MultiLine       =   -1  'True
         TabIndex        =   4
         Top             =   360
         Width           =   7575
      End
   End
End
Attribute VB_Name = "frmAgregarNotificacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Conexion As New Connection

Private Sub cmdAceptar_Click()
    If txtNotificacion = "" Then MsgBox ("Debe describir la notificación"): Exit Sub
    Respuesta = MsgBox("Confirma la Notificación?", vbYesNo, "Atención")
    If Respuesta = vbYes Then
        Me.MousePointer = 11
        Conexion.Open
        If lblTipoDeNotificacion = "Personal" Then
            Conexion.Execute ("INSERT INTO Notificaciones (Fecha, Notificaciones, idPermiso, Caduca) VALUES ('" & DateValue(dtpFechaNotificacion.Value) & "','" & txtNotificacion.Text & "'," & lblPermiso & ", '" & DateValue(dtpFechaCaducidad.Value) & "')")
        End If
        If lblTipoDeNotificacion = "Por Materia" Then
            Conexion.Execute ("INSERT INTO Notificaciones (Fecha, Notificaciones, idMateria, Caduca) VALUES ('" & DateValue(dtpFechaNotificacion.Value) & "','" & txtNotificacion.Text & "', " & lblMateria & ", '" & DateValue(dtpFechaCaducidad.Value) & "')")
        End If
        If lblTipoDeNotificacion = "General" Then
            Conexion.Execute ("INSERT INTO Notificaciones (Fecha, Notificaciones, Caduca) VALUES ('" & DateValue(dtpFechaNotificacion.Value) & "','" & txtNotificacion.Text & "','" & DateValue(dtpFechaCaducidad.Value) & "')")
        End If
        Conexion.Close
        Me.MousePointer = 0
        Unload Me
    End If

End Sub

Private Sub cmdCancelar_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Conexion.ConnectionString = ("DSN=Instituto")
    dtpFechaNotificacion.Value = Date
    dtpFechaCaducidad.Value = Date + 7
End Sub
