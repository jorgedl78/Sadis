VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmNotificaciones 
   Caption         =   "Notificaciones"
   ClientHeight    =   8880
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   12705
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   8880
   ScaleWidth      =   12705
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Height          =   7575
      Left            =   240
      TabIndex        =   3
      Top             =   120
      Width           =   12255
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
         Height          =   7095
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   12015
         _ExtentX        =   21193
         _ExtentY        =   12515
         _Version        =   393216
         Rows            =   1
         Cols            =   5
         FixedRows       =   0
         FixedCols       =   0
         BackColorFixed  =   8421504
         BackColorSel    =   -2147483642
         FocusRect       =   0
         SelectionMode   =   1
         _NumberOfBands  =   1
         _Band(0).Cols   =   5
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1095
      Left            =   240
      TabIndex        =   0
      Top             =   7680
      Width           =   12255
      Begin VB.CommandButton cmdCancelar 
         Height          =   765
         Left            =   6480
         Picture         =   "frmNotificaciones.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   2
         TabStop         =   0   'False
         ToolTipText     =   "Cancelar"
         Top             =   240
         Width           =   1845
      End
      Begin VB.CommandButton cmdNotificacionPersonal 
         Caption         =   "Notificación General"
         Height          =   720
         Left            =   2400
         Picture         =   "frmNotificaciones.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Modificar"
         Top             =   240
         Width           =   1800
      End
   End
End
Attribute VB_Name = "frmNotificaciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Conexion As New Connection
Private Sub cmdCancelar_Click()
    Unload Me
End Sub

Private Sub cmdNotificacionPersonal_Click()
    frmAgregarNotificacion.lblTipoDeNotificacion = "General"
    frmAgregarNotificacion.Show 1
End Sub

Private Sub Form_Load()
    Conexion.ConnectionString = ("DSN=Instituto")
    Conexion.Open
    Dim rs As ADODB.Recordset
    Set rs = Conexion.Execute("SELECT Notificaciones.Fecha, Notificaciones.Notificaciones, Notificaciones.Caduca, Alumnos.Nombre as Alumno, Materias.Nombre as Materia FROM (Notificaciones LEFT JOIN Alumnos ON Notificaciones.idPermiso = Alumnos.Permiso) LEFT JOIN Materias ON Notificaciones.idMateria = Materias.Codigo ORDER BY Notificaciones.Fecha DESC")
    a = 1
    MSHFlexGrid1.TextArray(0) = "Fecha"
    MSHFlexGrid1.TextArray(1) = "Notificación"
    MSHFlexGrid1.TextArray(2) = "Caduca"
    MSHFlexGrid1.TextArray(3) = "Alumno"
    MSHFlexGrid1.TextArray(4) = "Materia"
    MSHFlexGrid1.ColWidth(0) = 1200
    MSHFlexGrid1.ColWidth(1) = 6000
    MSHFlexGrid1.ColWidth(2) = 1200
    MSHFlexGrid1.ColWidth(3) = 5000
    MSHFlexGrid1.ColWidth(4) = 5000
    For i = 0 To 4
        MSHFlexGrid1.Col = i
        MSHFlexGrid1.CellBackColor = &H808080
    Next i
    Do While rs.EOF = False
        Me.MSHFlexGrid1.Rows = Me.MSHFlexGrid1.Rows + 1
        MSHFlexGrid1.TextMatrix(a, 0) = rs!Fecha
        MSHFlexGrid1.TextMatrix(a, 1) = rs!Notificaciones
        MSHFlexGrid1.TextMatrix(a, 2) = rs!Caduca
        If rs!Alumno <> vacio Then MSHFlexGrid1.TextMatrix(a, 3) = rs!Alumno
        If rs!Materia <> vacio Then MSHFlexGrid1.TextMatrix(a, 4) = rs!Materia

        rs.MoveNext
        a = a + 1
    Loop
    rs.Close
    Set rs = Nothing
    Conexion.Close
    'ColorearTitulos
    MSHFlexGrid1.Row = 1
End Sub
