VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Begin VB.Form frmOrdenPlan 
   Caption         =   "Ordenamiento del Plan de Estudios"
   ClientHeight    =   9600
   ClientLeft      =   1500
   ClientTop       =   1620
   ClientWidth     =   11880
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   9600
   ScaleWidth      =   11880
   Begin VB.Frame Frame2 
      Caption         =   "Plan de Estudios"
      Height          =   8295
      Left            =   240
      TabIndex        =   5
      Top             =   120
      Width           =   11535
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
         Height          =   7935
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   11295
         _ExtentX        =   19923
         _ExtentY        =   13996
         _Version        =   393216
         Rows            =   1
         Cols            =   5
         FixedRows       =   0
         BackColorFixed  =   8421504
         BackColorSel    =   -2147483642
         FocusRect       =   0
         SelectionMode   =   1
         _NumberOfBands  =   1
         _Band(0).Cols   =   5
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H8000000C&
      BorderStyle     =   0  'None
      Height          =   1035
      Left            =   240
      TabIndex        =   0
      Top             =   8520
      Width           =   11535
      Begin VB.CommandButton cmdGuardar 
         Height          =   840
         Left            =   6240
         Picture         =   "frmOrdenPlan.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Guardar"
         Top             =   120
         Width           =   960
      End
      Begin VB.CommandButton cmdSalir 
         Height          =   795
         Left            =   9240
         Picture         =   "frmOrdenPlan.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Salir"
         Top             =   120
         Width           =   915
      End
      Begin VB.CommandButton cmdSubir 
         Height          =   855
         Left            =   720
         Picture         =   "frmOrdenPlan.frx":0884
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   120
         Width           =   975
      End
      Begin VB.CommandButton cmdBajar 
         Height          =   855
         Left            =   2040
         Picture         =   "frmOrdenPlan.frx":114E
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   120
         Width           =   975
      End
   End
End
Attribute VB_Name = "frmOrdenPlan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Conexion As New Connection
Dim Materias As New Recordset

Private Sub cmdBajar_Click()
    If MSHFlexGrid1.Row = -1 Then Exit Sub
    If MSHFlexGrid1.Row >= MSHFlexGrid1.Rows - 1 Then Exit Sub
    filaactual = MSHFlexGrid1.Row
    WCurso = MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Row, 1)
    WCodigo = MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Row, 2)
    WNombre = MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Row, 3)
    WDetalle = MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Row, 4)
    
    MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Row, 1) = MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Row + 1, 1)
    MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Row, 2) = MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Row + 1, 2)
    MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Row, 3) = MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Row + 1, 3)
    MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Row, 4) = MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Row + 1, 4)
    
    MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Row + 1, 1) = WCurso
    MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Row + 1, 2) = WCodigo
    MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Row + 1, 3) = WNombre
    MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Row + 1, 4) = WDetalle
    
    ColorearTitulos
    MSHFlexGrid1.Row = filaactual + 1
    MSHFlexGrid1.ColSel = 1
    MSHFlexGrid1.SetFocus
End Sub

Private Sub cmdGuardar_Click()
    Respuesta = MsgBox("¿Está seguro de guardar los cambios? ", vbYesNo, "Guardar")
    If Respuesta = vbYes Then
        With MSHFlexGrid1
        Conexion.Open
        For i = 1 To .Rows - 1
            Conexion.Execute ("update materias set ordenplan=" & .TextMatrix(i, 0) & " where codigo=" & .TextMatrix(i, 2))
        Next i
        Conexion.Close
        End With
        Unload Me
    End If
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub



Private Sub cmdSubir_Click()
    If MSHFlexGrid1.Row = -1 Then Exit Sub
    If MSHFlexGrid1.Row <= 1 Then Exit Sub
    filaactual = MSHFlexGrid1.Row
    WCurso = MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Row, 1)
    WCodigo = MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Row, 2)
    WNombre = MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Row, 3)
    WDetalle = MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Row, 4)
    
    MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Row, 1) = MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Row - 1, 1)
    MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Row, 2) = MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Row - 1, 2)
    MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Row, 3) = MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Row - 1, 3)
    MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Row, 4) = MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Row - 1, 4)
    
    MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Row - 1, 1) = WCurso
    MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Row - 1, 2) = WCodigo
    MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Row - 1, 3) = WNombre
    MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Row - 1, 4) = WDetalle
    
    ColorearTitulos
    MSHFlexGrid1.Row = filaactual - 1
    MSHFlexGrid1.ColSel = 1
    MSHFlexGrid1.SetFocus
 End Sub

Private Sub Form_Load()
    Conexion.ConnectionString = ("DSN=Instituto")
    Conexion.Open
    Dim rs As ADODB.Recordset
    Set rs = Conexion.Execute("SELECT * FROM Materias WHERE Carrera = " & frmPlanes.dtcCarreras.BoundText & " AND Eliminada = 0 ORDER BY OrdenPlan,Curso, Codigo")
    a = 1
    MSHFlexGrid1.TextArray(0) = "Orden"
    MSHFlexGrid1.TextArray(1) = "Curso"
    MSHFlexGrid1.TextArray(2) = "Código"
    MSHFlexGrid1.TextArray(3) = "Nombre"
    MSHFlexGrid1.TextArray(4) = "Detalle"
    MSHFlexGrid1.ColWidth(0) = 600
    MSHFlexGrid1.ColWidth(1) = 600
    MSHFlexGrid1.ColWidth(2) = 800
    MSHFlexGrid1.ColWidth(3) = 10000
    MSHFlexGrid1.ColWidth(4) = 0
    For i = 1 To 4
        MSHFlexGrid1.Col = i
        MSHFlexGrid1.CellBackColor = &H808080
    Next i
    Do While rs.EOF = False
        Me.MSHFlexGrid1.Rows = Me.MSHFlexGrid1.Rows + 1
        MSHFlexGrid1.TextMatrix(a, 0) = a
        MSHFlexGrid1.TextMatrix(a, 1) = rs!Curso
        MSHFlexGrid1.TextMatrix(a, 2) = rs!Codigo
        MSHFlexGrid1.TextMatrix(a, 3) = rs!Nombre
        MSHFlexGrid1.TextMatrix(a, 4) = rs!Detalle
        rs.MoveNext
        a = a + 1
    Loop
    rs.Close
    Set rs = Nothing
    Conexion.Close
    ColorearTitulos
    MSHFlexGrid1.Row = 1
 End Sub

Private Sub ColorearTitulos()
    For j = 1 To MSHFlexGrid1.Rows - 1
            For h = 1 To 4
                MSHFlexGrid1.Row = j
                MSHFlexGrid1.Col = h
                If MSHFlexGrid1.TextMatrix(j, 4) = "4" Then
                    MSHFlexGrid1.CellBackColor = &H8000000F
                Else
                    MSHFlexGrid1.CellBackColor = &HFFFFFF
                End If
            Next h
    Next j

End Sub
