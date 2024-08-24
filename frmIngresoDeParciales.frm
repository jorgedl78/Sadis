VERSION 5.00
Begin VB.Form frmIngresoDeParciales 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   2805
   ClientLeft      =   8970
   ClientTop       =   4485
   ClientWidth     =   2055
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   2805
   ScaleWidth      =   2055
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdAgregarCorrelativa 
      Height          =   615
      Left            =   240
      Picture         =   "frmIngresoDeParciales.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Aceptar"
      Top             =   2040
      Width           =   615
   End
   Begin VB.CommandButton cmdCancelarCorrelativa 
      Height          =   615
      Left            =   1200
      Picture         =   "frmIngresoDeParciales.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Cancelar"
      Top             =   2040
      Width           =   615
   End
   Begin VB.ListBox lbNotaAIngresar 
      Height          =   1815
      ItemData        =   "frmIngresoDeParciales.frx":0884
      Left            =   120
      List            =   "frmIngresoDeParciales.frx":08A3
      TabIndex        =   0
      Top             =   120
      Width           =   1815
   End
End
Attribute VB_Name = "frmIngresoDeParciales"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAgregarCorrelativa_Click()
    If lbNotaAIngresar.Text = "" Then MsgBox ("Debe elegir la nota a ingresar"): Exit Sub
    With frmParciales
    .frIngresoDeNotas.Visible = True
    .lblDescripcionDeNota = lbNotaAIngresar.Text
    .frMaterias.Enabled = False
    .frCarrera.Enabled = False
    .frMatriculados.Enabled = False
    .frComandos.Enabled = False
    .frPlanillas.Enabled = False
    .adoMatriculados.Recordset.MoveFirst
    End With
    Unload Me
End Sub

Private Sub cmdCancelarCorrelativa_Click()
    Unload Me
End Sub

