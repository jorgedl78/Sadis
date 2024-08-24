VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmAlumnosLibres 
   Caption         =   "Alumnos Libres"
   ClientHeight    =   6540
   ClientLeft      =   2550
   ClientTop       =   2670
   ClientWidth     =   5880
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   6540
   ScaleWidth      =   5880
   Begin VB.CommandButton cmdPasarARegular 
      Caption         =   "Pasar a Regular"
      Height          =   495
      Left            =   3240
      TabIndex        =   3
      Top             =   5880
      Width           =   975
   End
   Begin VB.CommandButton cmdHabilitarExamen 
      Caption         =   "Habilitar para Examen Final"
      Height          =   495
      Left            =   480
      TabIndex        =   2
      Top             =   5880
      Width           =   2415
   End
   Begin VB.CommandButton cmdSalir 
      Height          =   550
      Left            =   4560
      Picture         =   "frmAlumnosLibres.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Salir"
      Top             =   5880
      Width           =   550
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frmAlumnosLibres.frx":0442
      Height          =   5535
      Left            =   360
      TabIndex        =   0
      Top             =   120
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   9763
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
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   "Permiso"
         Caption         =   "Permiso"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "Nombre"
         Caption         =   "Nombre"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         MarqueeStyle    =   4
         BeginProperty Column00 
            ColumnWidth     =   720
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   3509,858
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc adoLibres 
      Height          =   375
      Left            =   3120
      Top             =   5040
      Visible         =   0   'False
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   661
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
      RecordSource    =   $"frmAlumnosLibres.frx":045A
      Caption         =   "Libres"
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
End
Attribute VB_Name = "frmAlumnosLibres"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Conexion As New Connection

Private Sub cmdHabilitarExamen_Click()
    Respuesta = MsgBox("Esta seguro de habilitar el permiso para rendir examen final?", vbYesNo, "Habilitar permiso de exámen")
    If Respuesta = vbYes Then
        Me.MousePointer = 11
        Conexion.Open
        Conexion.Execute ("UPDATE Finales SET Finales.Cursada = True, Finales.Asistencia = True WHERE (((Finales.Materia)=" & frmParciales.dtcMaterias.BoundText & ") AND ((Finales.Ano)=" & frmParciales.txtAño & ") AND ((Finales.Libre)=True))")
        Conexion.Close
        Me.MousePointer = 0
    End If
End Sub

Private Sub cmdPasarARegular_Click()
    Respuesta = MsgBox("Va a pasar al alumno " & adoLibres.Recordset!Nombre & " a condiciòn de regular. ¿Continúa?", vbYesNo, "Está seguro")
    If Respuesta = vbYes Then
        Conexion.Open
        Conexion.Execute ("UPDATE Finales SET Libre = False WHERE Finales.Alumno=" & adoLibres.Recordset!Permiso & " AND Finales.Materia=" & frmParciales.dtcMaterias.BoundText & " AND Finales.Ano=" & frmParciales.txtAño)
        Conexion.Close
        frmParciales.VerMatriculados
        Unload Me
    End If
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Conexion.ConnectionString = ("DSN=Instituto")
    Conexion.Open
    'adoLibres.RecordSource = "SELECT Alumnos.Permiso, Alumnos.Nombre FROM Finales INNER JOIN Alumnos ON Finales.Alumno = Alumnos.Permiso Where (((Finales.Materia) = " & frmParciales.dtcMaterias.BoundText & ") And ((Finales.Ano) = " & frmParciales.txtAño & ") And ((Finales.Libre) = True)) ORDER BY Alumnos.Nombre"
    adoLibres.RecordSource = "SELECT Alumnos.Permiso, Alumnos.Nombre FROM Finales INNER JOIN Alumnos ON Finales.Alumno = Alumnos.Permiso Where Finales.Materia = " & frmParciales.dtcMaterias.BoundText & " And Finales.Ano = " & frmParciales.txtAño & " And Finales.Libre = True And Finales.Division = " & frmParciales.cbDivision & " ORDER BY Alumnos.Nombre"
    adoLibres.Refresh
    Conexion.Close
End Sub
