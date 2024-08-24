VERSION 5.00
Begin VB.Form frmCooperadoraAgregarPago 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Agregar Pago"
   ClientHeight    =   2280
   ClientLeft      =   7605
   ClientTop       =   2130
   ClientWidth     =   2475
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   2280
   ScaleWidth      =   2475
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Height          =   2175
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2415
      Begin VB.CommandButton cmdAgregarPago 
         Height          =   615
         Left            =   240
         Picture         =   "frmCooperadoraAgregarPago.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Aceptar"
         Top             =   1440
         Width           =   615
      End
      Begin VB.CommandButton cmdCancelarCorrelativa 
         Height          =   615
         Left            =   1560
         Picture         =   "frmCooperadoraAgregarPago.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Cancelar"
         Top             =   1440
         Width           =   615
      End
      Begin VB.CheckBox chkCancelado 
         Alignment       =   1  'Right Justify
         Caption         =   "Cancelado:"
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
         Top             =   840
         Value           =   1  'Checked
         Width           =   1455
      End
      Begin VB.TextBox txtImporte 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1080
         TabIndex        =   2
         Top             =   360
         Width           =   735
      End
      Begin VB.Line Line1 
         X1              =   0
         X2              =   2400
         Y1              =   1320
         Y2              =   1320
      End
      Begin VB.Label Label1 
         Caption         =   "Importe: $"
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
         TabIndex        =   1
         Top             =   360
         Width           =   975
      End
   End
End
Attribute VB_Name = "frmCooperadoraAgregarPago"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Conexion As New Connection
Private Sub cmdAgregarPago_Click()
    With frmCooperadoraPagos
    Conexion.Open
    Conexion.Execute ("UPDATE [Cooperadora Pagos] SET [Cooperadora Pagos].Importe = " & Val(txtImporte) & ", [Cooperadora Pagos].Cancelado = " & chkCancelado.Value & " WHERE ((([Cooperadora Pagos].Alumno)=" & .txtPermiso & ") AND (([Cooperadora Pagos].Año)=" & .txtAño & ") AND (([Cooperadora Pagos].Concepto)=" & .adoConceptos.Recordset!Codigo & "))")
    Conexion.Close
    .adoConceptos.Refresh
    End With
    Unload Me
End Sub

Private Sub cmdCancelarCorrelativa_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Conexion.ConnectionString = ("DSN=Instituto")
End Sub

Private Sub txtImporte_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then Unload Me
    If KeyAscii = 13 Then cmdAgregarPago_Click
End Sub
