VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "msadodc.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmSuplentes 
   Caption         =   "Designación de Suplentes para Exámenes Finales"
   ClientHeight    =   6345
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10170
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   6345
   ScaleWidth      =   10170
   StartUpPosition =   2  'CenterScreen
   Begin MSAdodcLib.Adodc adoControlSuperposicion 
      Height          =   330
      Left            =   5400
      Top             =   0
      Visible         =   0   'False
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   582
      ConnectMode     =   0
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
      RecordSource    =   $"frmSuplentes.frx":0000
      Caption         =   "ControlSuperposicion"
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
   Begin VB.Frame Frame1 
      Caption         =   "Agregar y Borrar Suplentes"
      Height          =   5055
      Left            =   5040
      TabIndex        =   3
      Top             =   840
      Width           =   4815
      Begin MSAdodcLib.Adodc adoCondicion 
         Height          =   330
         Left            =   2040
         Top             =   1440
         Visible         =   0   'False
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   582
         ConnectMode     =   0
         CursorLocation  =   3
         IsolationLevel  =   -1
         ConnectionTimeout=   15
         CommandTimeout  =   30
         CursorType      =   3
         LockType        =   3
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
         RecordSource    =   $"frmSuplentes.frx":00D6
         Caption         =   "Condicion"
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
      Begin VB.Frame frBotones 
         BackColor       =   &H8000000C&
         BorderStyle     =   0  'None
         Height          =   855
         Left            =   240
         TabIndex        =   14
         Top             =   3840
         Width           =   4455
         Begin VB.CommandButton cmdCancelar 
            Enabled         =   0   'False
            Height          =   600
            Left            =   2760
            Picture         =   "frmSuplentes.frx":014E
            Style           =   1  'Graphical
            TabIndex        =   21
            ToolTipText     =   "Cancelar"
            Top             =   120
            Width           =   600
         End
         Begin VB.CommandButton cmdGuardar 
            Enabled         =   0   'False
            Height          =   600
            Left            =   1920
            Picture         =   "frmSuplentes.frx":0590
            Style           =   1  'Graphical
            TabIndex        =   20
            ToolTipText     =   "Guardar"
            Top             =   120
            Width           =   600
         End
         Begin VB.CommandButton cmdSalir 
            Height          =   600
            Left            =   3600
            Picture         =   "frmSuplentes.frx":09D2
            Style           =   1  'Graphical
            TabIndex        =   17
            ToolTipText     =   "Salir"
            Top             =   120
            Width           =   600
         End
         Begin VB.CommandButton cmdAgregar 
            Height          =   600
            Left            =   240
            Picture         =   "frmSuplentes.frx":0E14
            Style           =   1  'Graphical
            TabIndex        =   16
            ToolTipText     =   "Agregar"
            Top             =   120
            Width           =   600
         End
         Begin VB.CommandButton cmdEliminar 
            Height          =   600
            Left            =   1080
            Picture         =   "frmSuplentes.frx":1256
            Style           =   1  'Graphical
            TabIndex        =   15
            ToolTipText     =   "Borrar"
            Top             =   120
            Width           =   600
         End
      End
      Begin VB.TextBox txtHora 
         Enabled         =   0   'False
         Height          =   315
         Left            =   240
         Locked          =   -1  'True
         TabIndex        =   6
         Text            =   "00"
         Top             =   1680
         Width           =   375
      End
      Begin VB.TextBox txtMinutos 
         Enabled         =   0   'False
         Height          =   315
         Left            =   960
         Locked          =   -1  'True
         TabIndex        =   4
         Text            =   "00"
         Top             =   1680
         Width           =   375
      End
      Begin MSComCtl2.UpDown upHora 
         Height          =   375
         Left            =   600
         TabIndex        =   5
         Top             =   1650
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   661
         _Version        =   393216
         Max             =   24
         Min             =   -1
         Enabled         =   0   'False
      End
      Begin MSComCtl2.DTPicker dtpFecha 
         Height          =   315
         Left            =   240
         TabIndex        =   7
         Top             =   840
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         Format          =   52822017
         CurrentDate     =   37553
      End
      Begin MSComCtl2.UpDown upMinutos 
         Height          =   375
         Left            =   1320
         TabIndex        =   8
         Top             =   1650
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   661
         _Version        =   393216
         Increment       =   10
         Max             =   60
         Min             =   -10
         Enabled         =   0   'False
      End
      Begin MSAdodcLib.Adodc adoPersonal 
         Height          =   330
         Left            =   2760
         Top             =   2400
         Visible         =   0   'False
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   582
         ConnectMode     =   0
         CursorLocation  =   3
         IsolationLevel  =   -1
         ConnectionTimeout=   15
         CommandTimeout  =   30
         CursorType      =   3
         LockType        =   3
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
         RecordSource    =   "SELECT Codigo, Nombre FROM Personal WHERE Eliminado = 0 AND TrabajaActualmente = 1 ORDER BY Nombre"
         Caption         =   "Personal"
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
      Begin MSDataListLib.DataCombo dtcTitular 
         Bindings        =   "frmSuplentes.frx":1698
         Height          =   315
         Left            =   240
         TabIndex        =   12
         Top             =   2760
         Width           =   4455
         _ExtentX        =   7858
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         Style           =   2
         ListField       =   "Nombre"
         BoundColumn     =   "Codigo"
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo dtcCondición 
         Bindings        =   "frmSuplentes.frx":16B2
         Height          =   315
         Left            =   1800
         TabIndex        =   18
         Top             =   840
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         Style           =   2
         ListField       =   "Nombre"
         BoundColumn     =   "Codigo"
         Text            =   ""
      End
      Begin VB.Label Label1 
         Caption         =   "Condición:"
         Height          =   255
         Left            =   1800
         TabIndex        =   19
         Top             =   600
         Width           =   855
      End
      Begin VB.Label Label7 
         Caption         =   "Titular:"
         Height          =   255
         Left            =   240
         TabIndex        =   13
         Top             =   2520
         Width           =   615
      End
      Begin VB.Label Label9 
         Caption         =   "Hora:"
         Height          =   255
         Left            =   240
         TabIndex        =   11
         Top             =   1440
         Width           =   615
      End
      Begin VB.Label Label10 
         Caption         =   "Fecha:"
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Top             =   600
         Width           =   615
      End
      Begin VB.Label Label13 
         Caption         =   "Minutos:"
         Height          =   255
         Left            =   960
         TabIndex        =   9
         Top             =   1440
         Width           =   615
      End
   End
   Begin MSAdodcLib.Adodc adoSuplentes 
      Height          =   330
      Left            =   2760
      Top             =   5520
      Visible         =   0   'False
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
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
      RecordSource    =   $"frmSuplentes.frx":16CD
      Caption         =   "Suplentes"
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
   Begin MSDataGridLib.DataGrid dtgSuplentes 
      Bindings        =   "frmSuplentes.frx":1842
      Height          =   4935
      Left            =   240
      TabIndex        =   0
      Top             =   960
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   8705
      _Version        =   393216
      AllowUpdate     =   0   'False
      HeadLines       =   1
      RowHeight       =   15
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   4
      BeginProperty Column00 
         DataField       =   "Fecha"
         Caption         =   "Fecha"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "d-mmm"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   11274
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "Hora"
         Caption         =   "Hora"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "HH:mm"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   11274
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column02 
         DataField       =   "Nombre"
         Caption         =   "Profesor"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   11274
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column03 
         DataField       =   "Condicion"
         Caption         =   "Condicion"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   11274
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         MarqueeStyle    =   4
         RecordSelectors =   0   'False
         BeginProperty Column00 
            ColumnWidth     =   945.071
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   599.811
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1950.236
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   975.118
         EndProperty
      EndProperty
   End
   Begin VB.Label lblAño 
      Caption         =   "Año"
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
      Left            =   3480
      TabIndex        =   2
      Top             =   240
      Width           =   1215
   End
   Begin VB.Label lblTurno 
      Caption         =   "Turno"
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
      TabIndex        =   1
      Top             =   240
      Width           =   2535
   End
End
Attribute VB_Name = "frmSuplentes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Conexion As New Connection
Dim Resultado As Recordset

Private Sub cmdAgregar_Click()
    cmdAgregar.Enabled = False
    cmdGuardar.Enabled = True
    cmdCancelar.Enabled = True
    cmdSalir.Enabled = False
    dtpFecha.Enabled = True
    dtcCondición.Enabled = True
    upHora.Value = 0
    upMinutos.Value = 0
    upHora.Enabled = True
    upMinutos.Enabled = True
    dtcTitular.Text = ""
    dtcTitular.Enabled = True
End Sub

Private Sub cmdCancelar_Click()
    cmdAgregar.Enabled = True
    cmdGuardar.Enabled = False
    cmdCancelar.Enabled = False
    cmdSalir.Enabled = True
    dtpFecha.Enabled = False
    dtcCondición.Enabled = False
    upHora.Enabled = False
    upMinutos.Enabled = False
    dtcTitular.Enabled = False
    MostrarSuplente
End Sub

Private Sub cmdEliminar_Click()
    If adoSuplentes.Recordset.RecordCount > 0 Then
        Respuesta = MsgBox("Está seguro de eliminar la suplencia", vbYesNo, "")
        If Respuesta = vbYes Then
            Conexion.Open
            
            'para SQL Server
            'Set Resultado = Conexion.Execute("DELETE Mesas WHERE Mesas.Numero=" & adoSuplentes.Recordset!Numero & "")
            
            'para Acces
            Set Resultado = Conexion.Execute("DELETE * From Mesas WHERE Mesas.Numero=" & adoSuplentes.Recordset!Numero & "")
            
            Conexion.Close
            adoSuplentes.Refresh
        End If
    End If
End Sub

Private Sub cmdGuardar_Click()
    If dtcTitular.Text = "" Then MsgBox ("Debe especificarse un profesor"): Exit Sub
    Hora = txtHora & ":" & txtMinutos
    
    'esta anda en SQL Server
    'adoControlSuperposicion.RecordSource = "SELECT Mesas.Hora, Carreras.Abreviatura AS Carrera, Materias.Abreviatura AS Materia, Materias.Curso FROM Mesas INNER JOIN (Materias INNER JOIN Carreras ON Materias.Carrera = Carreras.Codigo) ON Mesas.Materia = Materias.Codigo Where (((Mesas.Fecha) = '" & DateValue(dtpFecha.Value) & "') AND ((Mesas.Hora)='" & TimeValue(Hora) & "') And ((Mesas.Titular) = " & dtcTitular.BoundText & ")) Or (((Mesas.Fecha) = '" & DateValue(dtpFecha.Value) & "') AND ((Mesas.Hora)='" & TimeValue(Hora) & "') And ((Mesas.Integrante1) = " & dtcTitular.BoundText & ")) Or (((Mesas.Fecha) = '" & DateValue(dtpFecha.Value) & "') AND ((Mesas.Hora)='" & TimeValue(Hora) & "') And ((Mesas.Integrante2) = " & dtcTitular.BoundText & "))ORDER BY Mesas.Fecha, Mesas.Hora"
    
    'esta anda en acces
    adoControlSuperposicion.RecordSource = "SELECT Mesas.Hora, Carreras.Abreviatura AS Carrera, Materias.Abreviatura AS Materia, Materias.Curso FROM Mesas INNER JOIN (Materias INNER JOIN Carreras ON Materias.Carrera = Carreras.Codigo) ON Mesas.Materia = Materias.Codigo Where (((Mesas.Fecha) = #" & Format(dtpFecha.Value, "mm/dd/yyyy") & "#) AND ((Mesas.Hora)=#" & Format(Hora, "hh:mm:ss") & "#) And ((Mesas.Titular) = " & dtcTitular.BoundText & ")) Or (((Mesas.Fecha) = #" & Format(dtpFecha.Value, "mm/dd/yyyy") & "#) AND ((Mesas.Hora)=#" & Format(Hora, "hh:mm:ss") & "#) And ((Mesas.Integrante1) = " & dtcTitular.BoundText & ")) Or (((Mesas.Fecha) = #" & Format(dtpFecha.Value, "mm/dd/yyyy") & "#) AND ((Mesas.Hora)=#" & Format(Hora, "hh:mm:ss") & "#) And ((Mesas.Integrante2) = " & dtcTitular.BoundText & "))ORDER BY Mesas.Fecha, Mesas.Hora"
    
    adoControlSuperposicion.Refresh
    With adoControlSuperposicion.Recordset
    If .RecordCount > 0 Then
        StringMesas = ""
        For i = 1 To .RecordCount
            StringMesas = StringMesas & "HORA:" & Format(!Hora, "hh:mm") & " CARRERA: " & !Carrera & " Materia: " & !Curso & "º-" & !Materia & Chr(13) & Chr(13)
            .MoveNext
        Next i
        MsgBox ("El profesor " & dtcTitular.Text & " está designado para las siguientes mesas:" & Chr(13) & Chr(13) & StringMesas)
        Exit Sub
    End If
    End With
    Hora = txtHora & ":" & txtMinutos
    Conexion.Open
    Set Resultado = Conexion.Execute("INSERT INTO Mesas(Materia, Turno, Ano, Fecha, Hora, Titular) VALUES (" & dtcCondición.BoundText & ", " & frmMesasArmado.dtcMeses.BoundText & ", " & frmMesasArmado.txtAño & ", '" & DateValue(dtpFecha.Value) & "', '" & TimeValue(Hora) & "', " & dtcTitular.BoundText & " )")
    Conexion.Close
    adoSuplentes.Refresh
    cmdAgregar.Enabled = True
    cmdGuardar.Enabled = False
    cmdCancelar.Enabled = False
    cmdSalir.Enabled = True
    dtpFecha.Enabled = False
    dtcCondición.Enabled = False
    upHora.Enabled = False
    upMinutos.Enabled = False
    dtcTitular.Enabled = False
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub dtgSuplentes_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    MostrarSuplente
End Sub

Private Sub Form_Load()
    Conexion.ConnectionString = ("DSN=Instituto")
    adoSuplentes.RecordSource = "SELECT Mesas.Numero, Mesas.Fecha, Mesas.Hora, Personal.Nombre, Mesas.Titular, Mesas.Materia, Materias.Nombre AS Condicion FROM (Personal INNER JOIN Mesas ON Personal.Codigo = Mesas.Titular) INNER JOIN Materias ON Mesas.Materia = Materias.Codigo Where (((Materias.Carrera) = 16) And ((Mesas.Turno) = " & frmMesasArmado.dtcMeses.BoundText & ") And ((Mesas.Ano) = " & frmMesasArmado.txtAño & " ))ORDER BY Mesas.Fecha, Mesas.Hora, Mesas.Materia"
    adoSuplentes.Refresh
    dtpFecha.Value = Date
    dtcCondición.BoundText = adoCondicion.Recordset!Codigo
End Sub

Private Sub upHora_Change()
    If upHora.Value = 24 Then upHora.Value = 0
    If upHora.Value = -1 Then upHora.Value = 23
    txtHora = Format(upHora.Value, "00")
End Sub

Private Sub upMinutos_Change()
    If upMinutos.Value = 60 Then
        upMinutos.Value = 0
        upHora.Value = upHora.Value + 1
    ElseIf upMinutos.Value = -10 Then
        upMinutos.Value = 50
        upHora.Value = upHora.Value - 1
    Else
        txtMinutos = Format(upMinutos.Value, "00")
    End If
End Sub

Private Function MostrarSuplente()
If adoSuplentes.Recordset.RecordCount > 0 Then
    dtpFecha = adoSuplentes.Recordset!Fecha
    dtcCondición.BoundText = adoSuplentes.Recordset!Materia
    upHora.Value = Format(adoSuplentes.Recordset!Hora, "hh")
    upMinutos.Value = Mid(Format(adoSuplentes.Recordset!Hora, "hh:mm"), Len(Format(adoSuplentes.Recordset!Hora, "hh:mm")) - 1, 2)
    dtcTitular.BoundText = adoSuplentes.Recordset!Titular
End If
End Function
