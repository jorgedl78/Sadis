VERSION 5.00
Begin VB.Form frmAgregarMateria 
   Caption         =   "Agregar Materia"
   ClientHeight    =   1980
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   2115
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   Moveable        =   0   'False
   ScaleHeight     =   1980
   ScaleWidth      =   2115
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCancelar 
      Height          =   495
      Left            =   1080
      Picture         =   "frmAgregarMateria.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1320
      Width           =   735
   End
   Begin VB.CommandButton cmdAceptar 
      Height          =   495
      Left            =   120
      Picture         =   "frmAgregarMateria.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1320
      Width           =   735
   End
   Begin VB.TextBox txtOrden 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   1200
      TabIndex        =   2
      Top             =   720
      Width           =   615
   End
   Begin VB.ComboBox cbAño 
      Height          =   315
      ItemData        =   "frmAgregarMateria.frx":0884
      Left            =   1200
      List            =   "frmAgregarMateria.frx":0886
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   240
      Width           =   615
   End
   Begin VB.Label Label2 
      Caption         =   "Orden Interno:"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   720
      Width           =   1095
   End
   Begin VB.Label lblMedida 
      Alignment       =   1  'Right Justify
      Caption         =   "Año:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   975
   End
End
Attribute VB_Name = "frmAgregarMateria"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Modalidad As String
Dim Conexion As New Connection

Private Sub cbAño_Click()
    txtOrden.SetFocus
End Sub

Private Sub cbAño_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Aceptar
End Sub

Private Sub cmdAceptar_Click()
    If cbAño.Text = "" Then
        MsgBox ("Deebe especificar el curso"): Exit Sub
    End If
    If txtOrden.Text = "" Then
        MsgBox ("Debe especificar el orden interno"): Exit Sub
    End If
    
    If IsNumeric(txtOrden.Text) = False Then
        MsgBox ("El valor debe ser numérico"): Exit Sub
    End If
    
    Codigo = frmPlanes.txtCodigoCarrera & Format(cbAño, "00") & Format(txtOrden, "00")
    Conexion.Open
    Dim rs As ADODB.Recordset
    ' Crear y abrir un Recordset
    Set rs = Conexion.Execute("SELECT Materias.Codigo FROM Materias WHERE Materias.Codigo=" & Codigo)
    If rs.EOF = False Then
       MsgBox ("Este código ya ha sido utilizado"): rs.Close: Set rs = Nothing: Conexion.Close: Exit Sub
    End If
    rs.Close
    Set rs = Nothing
    Conexion.Close
    Aceptar
End Sub


Private Sub cmdCancelar_Click()
    frmPlanes.cmdCancelarMateria_Click
    Unload Me
End Sub

Private Sub Form_Load()
    Modalidad = frmPlanes.dtcModalidadCarrera.BoundText
    frmPlanes.adoModalidadCarreras.Recordset.MoveFirst
    frmPlanes.adoModalidadCarreras.Recordset.Find ("Codigo = " & Modalidad)
    lblMedida = frmPlanes.adoModalidadCarreras.Recordset!Medida & ":"
    Conexion.ConnectionString = ("DSN=Instituto")
End Sub

Private Sub txtOrden_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Aceptar
End Sub
Private Function Aceptar()
    frmPlanes.txtCursoMateria = cbAño.Text
    frmPlanes.txtCursoAnalitico = cbAño.Text
    frmPlanes.txtCodigoBloqueadoMateria = frmPlanes.adoCarreras.Recordset!Codigo & Format(cbAño.Text, "00")
    frmPlanes.txtCodigoDesbloqueadoMateria = Format(txtOrden.Text, "00")
    Unload Me
End Function
