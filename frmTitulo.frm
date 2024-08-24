VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Begin VB.Form frmTitulo 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Titulo"
   ClientHeight    =   7380
   ClientLeft      =   60
   ClientTop       =   300
   ClientWidth     =   12540
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   7380
   ScaleWidth      =   12540
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   5055
      Left            =   0
      TabIndex        =   21
      Top             =   2400
      Width           =   12495
      Begin VB.CommandButton cmdActualizaAsignaturas 
         Caption         =   "Actualizar Asignaturas en SFT"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   6720
         Picture         =   "frmTitulo.frx":0000
         TabIndex        =   38
         Top             =   3840
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.CommandButton cmdEnviaSFT 
         Caption         =   "Enviar al Sistema Federal de Títulos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   3720
         Picture         =   "frmTitulo.frx":0785
         TabIndex        =   37
         Top             =   3840
         Visible         =   0   'False
         Width           =   2655
      End
      Begin Crystal.CrystalReport CrystalReport3 
         Left            =   11760
         Top             =   1560
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   262150
         ReportFileName  =   "titulot2.rpt"
      End
      Begin VB.Frame frTramo 
         Caption         =   "Tramo de Formación Pedagógica"
         Height          =   975
         Left            =   5280
         TabIndex        =   33
         Top             =   960
         Width           =   3255
         Begin VB.CommandButton cmdTituloTramo 
            Height          =   495
            Left            =   2040
            Picture         =   "frmTitulo.frx":0F0A
            Style           =   1  'Graphical
            TabIndex        =   36
            Top             =   240
            Width           =   1095
         End
         Begin VB.OptionButton optNivelII 
            Caption         =   "Nivel II"
            Height          =   255
            Left            =   1080
            TabIndex        =   35
            Top             =   360
            Width           =   855
         End
         Begin VB.OptionButton optNivel1 
            Caption         =   "Nivel I"
            Height          =   255
            Left            =   240
            TabIndex        =   34
            Top             =   360
            Width           =   855
         End
      End
      Begin Crystal.CrystalReport CrystalReport2 
         Left            =   11760
         Top             =   960
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   262150
         ReportFileName  =   "tituloTF.rpt"
      End
      Begin VB.CommandButton cmdImprimir 
         Caption         =   "Imprimir"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   2160
         Picture         =   "frmTitulo.frx":100C
         Style           =   1  'Graphical
         TabIndex        =   32
         Top             =   3840
         Width           =   1095
      End
      Begin VB.CommandButton cmdSalir 
         Height          =   840
         Left            =   8280
         Picture         =   "frmTitulo.frx":1676
         Style           =   1  'Graphical
         TabIndex        =   31
         ToolTipText     =   "Salir"
         Top             =   3840
         Width           =   1080
      End
      Begin VB.ComboBox cbOriginal 
         Height          =   315
         ItemData        =   "frmTitulo.frx":1AB8
         Left            =   2280
         List            =   "frmTitulo.frx":1AC2
         Style           =   2  'Dropdown List
         TabIndex        =   25
         Top             =   1200
         Width           =   1935
      End
      Begin VB.TextBox Text1 
         DataField       =   "Numero"
         DataSource      =   "adoMeses"
         Height          =   285
         Left            =   5880
         TabIndex        =   23
         Text            =   "Text1"
         Top             =   120
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.TextBox txtObservaciones 
         Height          =   1095
         Left            =   120
         MaxLength       =   255
         MultiLine       =   -1  'True
         TabIndex        =   22
         Top             =   2040
         Width           =   10935
      End
      Begin MSAdodcLib.Adodc adoTitulo 
         Height          =   330
         Left            =   2760
         Top             =   1680
         Visible         =   0   'False
         Width           =   1695
         _ExtentX        =   2990
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
         RecordSource    =   "SELECT * FROM Titulo"
         Caption         =   "Titulo"
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
      Begin MSAdodcLib.Adodc adoMeses 
         Height          =   330
         Left            =   4080
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
         RecordSource    =   "SELECT * FROM Meses ORDER BY Numero"
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
      Begin MSDataListLib.DataCombo dtcTitulos 
         Bindings        =   "frmTitulo.frx":1AE3
         Height          =   315
         Left            =   120
         TabIndex        =   24
         Top             =   480
         Width           =   11055
         _ExtentX        =   19500
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "Titulo"
         BoundColumn     =   ""
         Text            =   ""
      End
      Begin Crystal.CrystalReport CrystalReport1 
         Left            =   11760
         Top             =   360
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   262150
         ReportFileName  =   "titulo.rpt"
         WindowLeft      =   50
         WindowWidth     =   600
         WindowHeight    =   400
         WindowTitle     =   "Titulo"
         WindowControlBox=   -1  'True
         WindowMaxButton =   -1  'True
         WindowMinButton =   -1  'True
      End
      Begin MSComCtl2.DTPicker dtFecha 
         Height          =   375
         Left            =   120
         TabIndex        =   26
         Top             =   1200
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         Format          =   42270721
         CurrentDate     =   37672
      End
      Begin MSAdodcLib.Adodc adoTitulos 
         Height          =   330
         Left            =   2280
         Top             =   120
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
         RecordSource    =   "SELECT * FROM TitulosPosibles WHERE TitulosPosibles.Carrera=0"
         Caption         =   "Titulos"
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
      Begin VB.Label Label2 
         Caption         =   "Elija el título correspondiente:"
         Height          =   255
         Left            =   120
         TabIndex        =   30
         Top             =   240
         Width           =   2295
      End
      Begin VB.Label Label3 
         Caption         =   "Fecha de Impresión:"
         Height          =   255
         Left            =   120
         TabIndex        =   29
         Top             =   960
         Width           =   1455
      End
      Begin VB.Label Label4 
         Caption         =   "Certificado:"
         Height          =   255
         Left            =   2280
         TabIndex        =   28
         Top             =   960
         Width           =   975
      End
      Begin VB.Label Label6 
         Caption         =   "Obsevaciones:"
         Height          =   255
         Left            =   120
         TabIndex        =   27
         Top             =   1800
         Width           =   1335
      End
   End
   Begin VB.Frame frAlumnos 
      Caption         =   "Alumno"
      Height          =   2415
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   12495
      Begin VB.TextBox txtNumeroTitulo 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   5400
         TabIndex        =   40
         Top             =   1920
         Width           =   975
      End
      Begin VB.TextBox txtPermiso 
         Height          =   315
         Left            =   120
         MaxLength       =   5
         TabIndex        =   1
         Top             =   480
         Width           =   615
      End
      Begin VB.CommandButton cmdBuscarAlumno 
         Caption         =   "Buscar..."
         Height          =   315
         Left            =   960
         TabIndex        =   11
         Top             =   480
         Width           =   855
      End
      Begin VB.TextBox txtAlumnoNombre 
         DataField       =   "Nombre"
         DataSource      =   "adoAlumnos"
         Enabled         =   0   'False
         Height          =   315
         Left            =   2040
         TabIndex        =   9
         Top             =   480
         Width           =   5415
      End
      Begin VB.TextBox txtAlumnoDocumento 
         DataField       =   "Documento"
         DataSource      =   "adoAlumnos"
         Enabled         =   0   'False
         Height          =   315
         Left            =   8160
         TabIndex        =   8
         Top             =   480
         Width           =   1335
      End
      Begin VB.TextBox txtAlumnoIngreso 
         DataField       =   "Ingreso"
         DataSource      =   "adoCarreras"
         Enabled         =   0   'False
         Height          =   315
         Left            =   11040
         TabIndex        =   7
         Top             =   1200
         Width           =   1095
      End
      Begin VB.TextBox txtAlumnoCondicion 
         DataField       =   "Condicion"
         DataSource      =   "adoCarreras"
         Enabled         =   0   'False
         Height          =   315
         Left            =   120
         TabIndex        =   6
         Top             =   1920
         Width           =   1815
      End
      Begin VB.TextBox txtAlumnoLibro 
         DataField       =   "Libro"
         DataSource      =   "adoCarreras"
         Enabled         =   0   'False
         Height          =   315
         Left            =   3360
         TabIndex        =   5
         Top             =   1920
         Width           =   615
      End
      Begin VB.TextBox txtAlumnoFecha 
         DataField       =   "Fecha"
         DataSource      =   "adoCarreras"
         Enabled         =   0   'False
         Height          =   315
         Left            =   2160
         TabIndex        =   4
         Top             =   1920
         Width           =   1095
      End
      Begin VB.TextBox txtAlumnoFolio 
         DataField       =   "Folio"
         DataSource      =   "adoCarreras"
         Enabled         =   0   'False
         Height          =   315
         Left            =   4440
         TabIndex        =   3
         Top             =   1920
         Width           =   615
      End
      Begin VB.TextBox txtAlumnoTipo 
         DataField       =   "Tipo"
         DataSource      =   "adoAlumnos"
         Enabled         =   0   'False
         Height          =   315
         Left            =   7680
         TabIndex        =   2
         Top             =   480
         Width           =   495
      End
      Begin MSDataListLib.DataCombo dtcCarreras 
         Bindings        =   "frmTitulo.frx":1AFC
         Height          =   315
         Left            =   120
         TabIndex        =   10
         Top             =   1200
         Width           =   10815
         _ExtentX        =   19076
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "Nombre"
         BoundColumn     =   "Carrera"
         Text            =   ""
      End
      Begin MSAdodcLib.Adodc adoAlumnos 
         Height          =   330
         Left            =   120
         Top             =   1080
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
         RecordSource    =   "SELECT Permiso,Nombre,Tipo,Documento FROM Alumnos WHERE Permiso = 0"
         Caption         =   "Alumnos"
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
      Begin MSAdodcLib.Adodc adoCarreras 
         Height          =   330
         Left            =   2160
         Top             =   1080
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
         RecordSource    =   $"frmTitulo.frx":1B16
         Caption         =   "Carreras"
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
      Begin VB.Label Label7 
         Caption         =   "Numero de Título:"
         Height          =   255
         Left            =   5400
         TabIndex        =   39
         Top             =   1680
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "Nº Permiso"
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   240
         Width           =   855
      End
      Begin VB.Label lblFechaAlumno 
         Caption         =   "Fecha:"
         Height          =   255
         Left            =   2160
         TabIndex        =   19
         Top             =   1680
         Width           =   615
      End
      Begin VB.Label lblCondicionAlumno 
         Caption         =   "Condición:"
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   1680
         Width           =   855
      End
      Begin VB.Label lblIngresoAlumno 
         Caption         =   "Ingresó:"
         Height          =   255
         Left            =   11040
         TabIndex        =   17
         Top             =   960
         Width           =   615
      End
      Begin VB.Label Label5 
         Caption         =   "Carrera"
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   960
         Width           =   855
      End
      Begin VB.Label lblFolioAlumno 
         Caption         =   "Folio:"
         Height          =   255
         Left            =   4440
         TabIndex        =   15
         Top             =   1680
         Width           =   495
      End
      Begin VB.Label lblLibroAlumno 
         Caption         =   "Libro:"
         Height          =   255
         Left            =   3360
         TabIndex        =   14
         Top             =   1680
         Width           =   495
      End
      Begin VB.Label Label8 
         Caption         =   "Documento:"
         Height          =   255
         Left            =   7680
         TabIndex        =   13
         Top             =   240
         Width           =   975
      End
      Begin VB.Label lblAlumnoNombre 
         Caption         =   "Nombre"
         Height          =   255
         Left            =   2040
         TabIndex        =   12
         Top             =   240
         Width           =   615
      End
   End
End
Attribute VB_Name = "frmTitulo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Conexion As New Connection
Dim BaseTitulos As New Connection
Dim Resultado As New Recordset
Dim Parametros As New Recordset

Private Sub cmdActualizaAsignaturas_Click()
    Respuesta = MsgBox("¿Actualiza las asignaturas en el SFT? ", vbYesNo, "Actualizar")
    If Respuesta = vbYes Then
        Me.MousePointer = 11
        Conexion.Open
        BaseTitulos.Open
        Set Resultado = Conexion.Execute("SELECT Materias.Codigo, Materias.OrdenPlan & Materias.Codigo AS CodigoNuevo FROM Materias WHERE Materias.Carrera= " & dtcCarreras.BoundText)
        While Resultado.EOF = False
            BaseTitulos.Execute ("UPDATE Asignaturas SET Asignaturas.Codigo = " & Resultado!CodigoNuevo & " WHERE Asignaturas.Codigo=" & Resultado!Codigo)
            BaseTitulos.Execute ("UPDATE Analitico SET Analitico.Cod_Asignatura =  " & Resultado!CodigoNuevo & "  WHERE Analitico.Cod_Asignatura=" & Resultado!Codigo)
            Resultado.MoveNext
        Wend
        Conexion.Close
        BaseTitulos.Close
        Me.MousePointer = 0
    End If
End Sub

Private Sub cmdBuscarAlumno_Click()
    Respuesta = InputBox("Ingrese Nº de Documento", "Buscar Alumno")
    If Respuesta = "" Then Exit Sub
    Conexion.Open
    Set Resultado = Conexion.Execute("SELECT Permiso FROM Alumnos WHERE Documento = " & Respuesta & " AND Eliminado = False")
    If Resultado.EOF = False Then
        txtPermiso = Resultado!Permiso: txtPermiso.SetFocus
    Else
        MsgBox ("El documento no existe"): txtPermiso = ""
    End If
    Conexion.Close
End Sub

Private Sub cmdEnviaSFT_Click()
    Respuesta = MsgBox("¿Confirma la exportación del título? ", vbYesNo, "Confirmar")
    If Respuesta = vbNo Then Exit Sub
    
    Me.MousePointer = 11
    PrepararTitulo
    Conexion.Open
    BaseTitulos.Open
    'levanto parametros del CUE
    Set Parametros = Conexion.Execute("SELECT Parametros.CUE, Parametros.CUE_ANEXO, Parametros.Codigo_establecimiento FROM Parametros")
    Set Resultado = BaseTitulos.Execute("SELECT Carreras.Codigo FROM Carreras WHERE Carreras.Codigo=" & dtcCarreras.BoundText)
    If Resultado.EOF = True Then 'hay que agregar la carrera
       'agrego la carrera
       Set Resultado = Conexion.Execute("SELECT Carreras.Resolucion FROM Carreras WHERE Carreras.Codigo=" & dtcCarreras.BoundText)
       Resolucion = Resultado!Resolucion
       BaseTitulos.Execute ("INSERT INTO Carreras ( CUE, CUE_ANEXO, Codigo, TipoTitulo, Descripcion1, Descripcion2, NormaAprobacion, NormaRatificacion,NroInscripcion, ValidezNacional ) VALUES ('" & Parametros!CUE & "','" & Parametros!CUE_ANEXO & "'," & dtcCarreras.BoundText & ",'TÍTULO','" & dtcTitulos.Text & "', 'EDUCACIÓN SUPERIOR','RES. " & Resolucion & "','---------------------------------------','------------------','------------------')")
       'agrego las materias
       Set Resultado = Conexion.Execute("SELECT Titulo.Curso, Titulo.CodigoMateria, Titulo.Materia FROM Titulo")
       While Resultado.EOF = False
          BaseTitulos.Execute ("INSERT INTO Asignaturas ( CUE, CUE_ANEXO, Año, Cod_Carrera, Codigo, Descripcion ) VALUES ('" & Parametros!CUE & "','" & Parametros!CUE_ANEXO & "'," & Resultado!Curso & "," & dtcCarreras.BoundText & "," & Resultado!CodigoMateria & ",'" & Resultado!Materia & "')")
          Resultado.MoveNext
       Wend
    End If
    'busco si el alumno aun no esta ingresado para este título
    Set Resultado = BaseTitulos.Execute("SELECT Alumnos.NroDocumento, Alumnos.Cod_Carrera FROM Alumnos WHERE (((Alumnos.NroDocumento)='" & Format(txtAlumnoDocumento, "#,###") & "') AND ((Alumnos.Cod_Carrera)=" & dtcCarreras.BoundText & "))")
    If Resultado.EOF = False Then
       MsgBox ("Ya se encuentra ingresado este título en el Sistema Federal de Títulos")
       Conexion.Close
       BaseTitulos.Close
       Me.MousePointer = 0
       Exit Sub
    End If
    MarcarComoRecibido
    'agrego el alumno con la carrera
    Set Resultado = Conexion.Execute("SELECT DISTINCT Alumnos.Nombre, 2 AS Expr1, Alumnos.Documento, UCASE(Alumnos.Lugar) as Lugar, Alumnos.FechaNacimiento, Titulo.Promedio, Titulo.Observaciones, (SELECT Max(Finales.Fecha)FROM Finales INNER JOIN Materias ON Finales.Materia = Materias.Codigo WHERE Materias.Carrera=" & dtcCarreras.BoundText & " AND Materias.Detalle<>5 AND Finales.Habilitada=True AND Finales.Aprobada=True AND Finales.Alumno=" & txtPermiso & ") as Egreso, CarrerasHechas.Libro, CarrerasHechas.Folio, Titulo.FechaDe, Titulo.CodigoCarrera FROM (Titulo INNER JOIN Alumnos ON Titulo.Documento = Alumnos.Documento) INNER JOIN CarrerasHechas ON (Titulo.CodigoCarrera = CarrerasHechas.Carrera) AND (Alumnos.Permiso = CarrerasHechas.Permiso)")
    For i = 1 To Len(txtAlumnoNombre)
    If Mid(txtAlumnoNombre, i, 1) = "," Then
       Apellido = Mid(txtAlumnoNombre, 1, i - 1)
       If Mid(txtAlumnoNombre, i + 1, 1) = " " Then
          Nombre = Mid(txtAlumnoNombre, i + 2, Len(txtAlumnoNombre))
       Else
          Nombre = Mid(txtAlumnoNombre, i + 1, Len(txtAlumnoNombre))
       End If
    End If
    Next i
    If txtObservaciones = "" Then
       txtObservaciones = "----------------------------------------------------------------------------------------------------------------------------------------------------------"
    End If
    BaseTitulos.Execute ("INSERT INTO Alumnos ( CUE, CUE_ANEXO, Apellido, Nombres, TipoDocumento, NroDocumento, LugarNacimiento, FechaNacimiento, PromedioGral, Observaciones, FechaEgreso, NroLibroMatriz, NroFolioLibroMatriz, FechaOtorgAnalitico, Cod_Carrera,Acta, FechaABM ) Values ( '" & Parametros!CUE & "', '" & Parametros!CUE_ANEXO & "', '" & Apellido & "', '" & Nombre & "', 2, '" & Replace(Format(txtAlumnoDocumento, "##,###,###"), ",", ".") & "', '" & Resultado!Lugar & "', '" & Resultado!FechaNacimiento & "', '" & LTrim(Replace(Str(Resultado!Promedio), ".", ",")) & "', '" & txtObservaciones & "', '" & Resultado!Egreso & "', " & Resultado!Libro & ", " & Resultado!Folio & ", '" & dtFecha & "', " & dtcCarreras.BoundText & ",'-','" & DateTime.Now & "' )")
    'levanto las materias aprobadas por el alumno para buscar y agregar calificaciones
    Set Resultado = Conexion.Execute("SELECT Materias.CursoAnalitico, Materias.Carrera,Materias.OrdenPlan & Materias.Codigo as Codigo, Finales.Nota, Finales.Fecha, Finales.Alumno, Finales.Establecimiento, Finales.Equivalencia FROM (Finales INNER JOIN Materias ON Finales.Materia = Materias.Codigo) INNER JOIN Instituciones ON Finales.Establecimiento = Instituciones.Codigo WHERE Materias.Carrera=" & dtcCarreras.BoundText & " AND Finales.Alumno=" & txtPermiso & " AND Materias.Detalle <> 5 AND Finales.Habilitada = True AND ((Finales.Aprobada = True AND Materias.Detalle BETWEEN 1 AND 3) OR Materias.Detalle=4) ORDER BY Materias.OrdenPlan, Materias.CursoAnalitico, Materias.Codigo")
    While Resultado.EOF = False
       If IsNull(Resultado!Fecha) Then 'es un titulo de Espacio
          BaseTitulos.Execute ("INSERT INTO Analitico ( CUE, CUE_ANEXO, TipoDocumento, NroDocumento, Año, Cod_Carrera, Cod_Asignatura, CalifFinal, Cod_Condicion, MesAsignatura, AñoAsignatura, Cod_Establecimiento, NroTitulo ) VALUES ('" & Parametros!CUE & "','" & Parametros!CUE_ANEXO & "', 2, '" & Replace(Format(txtAlumnoDocumento, "##,###,###"), ",", ".") & "', " & Resultado!CursoAnalitico & ", " & Resultado!Carrera & ", " & Resultado!Codigo & ", '----', 5, '--', '----', 3,1)")
       Else 'es una asignatura comun que va con nota y fecha
       If Resultado!Equivalencia = Verdadero Then 'es equivalencia
          Condicion = 2
       Else
          Condicion = 4
       End If
          If Int(Resultado!Nota) / 2 = (Resultado!Nota / 2) Then 'es una nota sin decimales
             Nota = Str(Resultado!Nota)
          Else 'es una nota con decimales
             Nota = Format(Resultado!Nota, "#.00")
             Nota = Replace(Nota, ".", ",")
          End If
          BaseTitulos.Execute ("INSERT INTO Analitico ( CUE, CUE_ANEXO, TipoDocumento, NroDocumento, Año, Cod_Carrera, Cod_Asignatura, CalifFinal, Cod_Condicion, MesAsignatura, AñoAsignatura, Cod_Establecimiento, NroTitulo ) VALUES ('" & Parametros!CUE & "','" & Parametros!CUE_ANEXO & "', 2, '" & Replace(Format(txtAlumnoDocumento, "##,###,###"), ",", ".") & "', " & Resultado!CursoAnalitico & ", " & Resultado!Carrera & ", " & Resultado!Codigo & ", '" & LTrim(Nota) & "'," & Condicion & " ,'" & Format(Month(Resultado!Fecha), "0#") & "', " & Year(Resultado!Fecha) & ", 2,1)")
       End If
       Resultado.MoveNext
    Wend
    BaseTitulos.Close
    Conexion.Close
    Me.MousePointer = 0
    MsgBox ("Título exportado")
End Sub

Private Sub cmdImprimir_Click()
If txtPermiso = "" Then
    MsgBox ("Debe especificar un número de permiso"): txtPermiso.SetFocus: Exit Sub
End If
Conexion.Open
    MarcarComoRecibido
Conexion.Close
    PrepararTitulo
    CrystalReport1.PrintReport
End Sub
Private Sub MarcarComoRecibido()
    Respuesta = MsgBox("¿Desea marcar al alumno como recibido?", vbYesNo, "Atención!")
    If Respuesta = vbNo Then Exit Sub
    Conexion.Execute ("UPDATE CarrerasHechas SET CarrerasHechas.Condición = 2, CarrerasHechas.Fecha = #" & dtFecha & "# WHERE (((CarrerasHechas.Permiso)=" & txtPermiso & ") AND ((CarrerasHechas.Carrera)=" & dtcCarreras.BoundText & "))")
End Sub
Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub cmdTituloTramo_Click()
    If txtPermiso = "" Then
        MsgBox ("Debe especificar un número de permiso"): txtPermiso.SetFocus: Exit Sub
    End If
   adoMeses.Recordset.MoveFirst
    adoMeses.Recordset.Find ("Numero=" & dtFecha.Month)
    Conexion.Open
    Conexion.Execute ("DELETE * FROM Titulo")
    Conexion.Execute ("INSERT INTO Titulo ( Permiso, Nombre, Tipo, Documento, CodigoMateria, Materia, CodigoCarrera, Fecha, Nota, Equivalencia, Curso, Carrera, Medida, Detalle, Caracteristica, FechaDe, Resolucion, Articulo)SELECT [Alumnos].[Permiso], [Alumnos].[Nombre], [Alumnos].[Tipo], [Alumnos].[Documento],[Materias].[Codigo] AS  CodigoMateria,[Materias].[Nombre] AS Materia, Carreras.Codigo AS CodigoCarrera, [Finales].[Fecha], [Finales].[Nota],[Finales].[Equivalencia], [Materias].[CursoAnalitico],[Carreras].[Nombre] AS Carrera, [Modalidad].[Medida], [Materias].[Detalle], [Caracteristica de carrera].[Caracteristica],[Caracteristica de carrera].[FechaDe], [Carreras].[Resolucion],[Caracteristica de carrera].[Articulo] FROM ((((Alumnos INNER JOIN Finales ON [Alumnos].[Permiso]=[Finales].[Alumno]) INNER JOIN Materias ON [Finales].[Materia]=[Materias].[Codigo])INNER JOIN Carreras ON [Materias].[Carrera]=[Carreras].[Codigo])INNER JOIN Modalidad ON [Carreras].[Modalidad]=[Modalidad].[Codigo])" _
    & " INNER JOIN [Caracteristica de carrera] ON [Carreras].[Caracteristica]=[Caracteristica de carrera].[Codigo] Where [Alumnos].[Permiso] = " & txtPermiso & " And [Carreras].[Codigo] = " & dtcCarreras.BoundText & " AND Materias.Detalle <> 5 AND Finales.Habilitada = True AND ((Finales.Aprobada = True AND Materias.Detalle BETWEEN 1 AND 3) OR Materias.Detalle=4) ORDER BY [Materias].[OrdenPlan],[Materias].[CursoAnalitico], [Materias].[Codigo]")
    Conexion.Execute ("UPDATE Titulo SET Certificado = '" & cbOriginal.Text & "'")
    Set Resultado = Conexion.Execute("SELECT NombreInstitucion, Localidad, Provincia FROM Parametros")
    Localidad = Resultado!Localidad
    Provincia = Resultado!Provincia
    PieDePagina = "En fe de lo cual se expide el presente certificado, " & Mid(cbOriginal.Text, 5) & " sin raspaduras y enmiendas, en la ciudad de " & Localidad & " (" & Provincia & ") a los " & dtFecha.Day & " dias del mes de " & adoMeses.Recordset!Nombre & " del año " & dtFecha.Year & ".-"
    Conexion.Execute ("UPDATE Titulo SET PieDePagina = '" & PieDePagina & "'")
    Conexion.Execute ("UPDATE Titulo SET Establecimiento = '" & Resultado!NombreINstitucion & "'")
    Set Resultado = Conexion.Execute("SELECT * FROM Titulo")
    If Resultado.EOF = True Then MsgBox ("El alumno no rindió ninguna materia para esta carrera"): Conexion.Close: Exit Sub
    If optNivel1 = True Then
        CarrerayTitulo = Resultado!Tipo & " Nº " & Resultado!documento & "  ha aprobado las Materias del Campo de la Fundamentación, Campo de la Práctica Docente, Campo de la Subjetividad y la Cultura y los Talleres que, con sus respectivas calificaciones, que abajo se registran, correspondientes al Tramo de Formación Pedagógica para Profesionales y Técnicos Superiores , Resolución Nº 4077/08"
    Else
        CarrerayTitulo = Resultado!Tipo & " Nº " & Resultado!documento & "  ha aprobado las Materias del Campo de la Fundamentación, Campo de la Práctica Docente, Campo de la Subjetividad y la Cultura y los Talleres que, con sus respectivas calificaciones, que abajo se registran, correspondientes al Tramo de Formación Pedagógica para Técnicos de Nivel Medio , Resolución Nº 4077/08"
    End If
    'If Resultado!CodigoCarrera = 51 Then
    '    TituloDeBase = InputBox("Ingrese el título de base")
    '    CarrerayTitulo = CarrerayTitulo & " al " & dtcTitulos.Text & " en conjunción con su título de base: " & TituloDeBase & " que lo habilita, de acuerdo al Nomenclador vigente, en cada nivel mencionado."
    'Else
    '    CarrerayTitulo = CarrerayTitulo & " a la carrera " & Resultado!Carrera & ", haciéndose acreedor/a al título de " & dtcTitulos.Text
    'End If
    'CarrerayTitulo = CarrerayTitulo & " - Resolución " & Resultado!Resolucion & "."
    Conexion.Execute ("UPDATE Titulo SET CarrerayTitulo = '" & CarrerayTitulo & "'")
    Conexion.Execute ("UPDATE Titulo SET Equivalencia = '*' WHERE Equivalencia = '-1'")
    Conexion.Execute ("UPDATE Titulo SET Equivalencia = Null WHERE Equivalencia = '0'")
    adoTitulo.RecordSource = "SELECT * FROM Titulo WHERE Detalle = 1 OR DEtalle = 2"
    adoTitulo.Refresh
    While adoTitulo.Recordset.EOF = False
    Decimo = Int(adoTitulo.Recordset!Nota)
    centesimo = Right(Format(adoTitulo.Recordset!Nota, "0.00"), 3)
    If centesimo = ".00" Then centesimo = "" 'si es entero no se muestra decimales
    If Decimo = 1 Then EnLetras = "UNO" & centesimo
    If Decimo = 2 Then EnLetras = "DOS" & centesimo
    If Decimo = 3 Then EnLetras = "TRES" & centesimo
    If Decimo = 4 Then EnLetras = "CUATRO" & centesimo
    If Decimo = 5 Then EnLetras = "CINCO" & centesimo
    If Decimo = 6 Then EnLetras = "SEIS" & centesimo
    If Decimo = 7 Then EnLetras = "SIETE" & centesimo
    If Decimo = 8 Then EnLetras = "OCHO" & centesimo
    If Decimo = 9 Then EnLetras = "NUEVE" & centesimo
    If Decimo = 10 Then EnLetras = "DIEZ" & centesimo
    Conexion.Execute ("UPDATE Titulo SET EnLetras = '" & EnLetras & "' WHERE CodigoMateria = " & adoTitulo.Recordset!CodigoMateria)
    adoTitulo.Recordset.MoveNext
    Wend
    Set Resultado = Conexion.Execute("SELECT Avg(Nota) AS Promedio FROM Titulo WHERE Detalle = 1 OR Detalle = 2")
    Conexion.Execute ("UPDATE Titulo SET Promedio = " & Format(Resultado!Promedio, "0.00"))
    Decimo = Int(Format(Resultado!Promedio, "0.00"))
    centesimo = Right(Format(Resultado!Promedio, "0.00"), 3)
    If Decimo = 1 Then PromedioLetras = "UNO" & centesimo
    If Decimo = 2 Then PromedioLetras = "DOS" & centesimo
    If Decimo = 3 Then PromedioLetras = "TRES" & centesimo
    If Decimo = 4 Then PromedioLetras = "CUATRO" & centesimo
    If Decimo = 5 Then PromedioLetras = "CINCO" & centesimo
    If Decimo = 6 Then PromedioLetras = "SEIS" & centesimo
    If Decimo = 7 Then PromedioLetras = "SIETE" & centesimo
    If Decimo = 8 Then PromedioLetras = "OCHO" & centesimo
    If Decimo = 9 Then PromedioLetras = "NUEVE" & centesimo
    If Decimo = 10 Then PromedioLetras = "DIEZ" & centesimo
    Conexion.Execute ("UPDATE Titulo SET Fecha = Null, Nota = Null, EnLetras = Null WHERE (Detalle = 3 or Detalle = 4)")
    Conexion.Execute ("UPDATE Titulo SET PromedioLetras = '" & PromedioLetras & "'")
    If txtObservaciones <> "" Then Conexion.Execute ("UPDATE Titulo SET Observaciones = 'Observaciones: " & txtObservaciones & "'")
    Conexion.Close
    
    
    If optNivel1 = True Then
        CrystalReport2.PrintReport
    Else
        CrystalReport3.PrintReport
    End If
End Sub




Private Sub dtcCarreras_Change()
    If dtcCarreras.Text <> "" Then
        adoCarreras.Recordset.MoveFirst
        adoCarreras.Recordset.Find ("Carrera=" & dtcCarreras.BoundText)
        adoTitulos.RecordSource = "SELECT * From TitulosPosibles Where TitulosPosibles.Carrera = " & dtcCarreras.BoundText & " ORDER BY Titulo"
        adoTitulos.Refresh
        If adoTitulos.Recordset.RecordCount = 0 Then
            MsgBox ("El plan seleccionado no contiene su descripción de título"): txtPermiso.SetFocus: Exit Sub
        End If
        dtcTitulos.Text = adoTitulos.Recordset!Titulo
    End If
End Sub

Private Sub Form_Load()
    Conexion.ConnectionString = ("DSN=Instituto")
    BaseTitulos.ConnectionString = ("DSN=Titulos")
    cbOriginal.Text = cbOriginal.List(0)
    dtFecha.Value = Date
End Sub

Private Sub txtPermiso_Click()
    dtcCarreras.Text = ""
    txtPermiso = ""
End Sub

Private Sub txtPermiso_KeyPress(KeyAscii As Integer)

    If InStr(1, "0123456789" & Chr(13) & Chr(8), Chr(KeyAscii)) = 0 Then
        KeyAscii = 0
    End If
    If txtPermiso = "" Then Exit Sub
    If KeyAscii = 13 Then dtcCarreras.Text = "": BuscarDatos
End Sub

Private Function BuscarDatos()
    adoAlumnos.RecordSource = "SELECT Permiso,Nombre,Tipo,Documento FROM Alumnos WHERE Eliminado = 0 AND Permiso = " & txtPermiso
    adoAlumnos.Refresh
    If adoAlumnos.Recordset.RecordCount = 1 Then
        adoCarreras.RecordSource = "SELECT CarrerasHechas.Permiso, CarrerasHechas.Carrera, CarrerasHechas.Ingreso, Condicion.Condicion, CarrerasHechas.Fecha, CarrerasHechas.Libro, CarrerasHechas.Folio, Carreras.Nombre FROM (CarrerasHechas INNER JOIN Carreras ON CarrerasHechas.Carrera = Carreras.Codigo) INNER JOIN Condicion ON CarrerasHechas.Condición = Condicion.Codigo WHERE CarrerasHechas.Permiso=" & txtPermiso
        adoCarreras.Refresh
        If adoCarreras.Recordset.RecordCount <= 0 Then
           MsgBox ("El alumno no tiene ningún plan de estudio asociado"): Exit Function
        End If
        
        dtcCarreras.BoundText = adoCarreras.Recordset!Carrera
    Else
        MsgBox ("El alumno no existe")
        txtPermiso = ""
    End If
End Function

Private Function PrepararTitulo()
    adoMeses.Recordset.MoveFirst
    adoMeses.Recordset.Find ("Numero=" & dtFecha.Month)
    Conexion.Open
    Conexion.Execute ("DELETE * FROM Titulo")
    Conexion.Execute ("INSERT INTO Titulo ( Permiso, Nombre, Tipo, Documento, CodigoMateria, Materia, CodigoCarrera, Fecha, Nota, Equivalencia, Curso, Carrera, Medida, Detalle, Caracteristica, FechaDe, Resolucion, Articulo)SELECT [Alumnos].[Permiso], [Alumnos].[Nombre], [Alumnos].[Tipo], [Alumnos].[Documento],[Materias].[OrdenPlan] & [Materias].[Codigo] AS  CodigoMateria,[Materias].[Nombre] AS Materia, Carreras.Codigo AS CodigoCarrera, [Finales].[Fecha], [Finales].[Nota],[Finales].[Equivalencia], [Materias].[CursoAnalitico],[Carreras].[Nombre] AS Carrera, [Modalidad].[Medida], [Materias].[Detalle], [Caracteristica de carrera].[Caracteristica],[Caracteristica de carrera].[FechaDe], [Carreras].[Resolucion],[Caracteristica de carrera].[Articulo] FROM ((((Alumnos INNER JOIN Finales ON [Alumnos].[Permiso]=[Finales].[Alumno]) INNER JOIN Materias ON [Finales].[Materia]=[Materias].[Codigo])INNER JOIN Carreras ON [Materias].[Carrera]=[Carreras].[Codigo]) " _
    & "INNER JOIN Modalidad ON [Carreras].[Modalidad]=[Modalidad].[Codigo])INNER JOIN [Caracteristica de carrera] ON [Carreras].[Caracteristica]=[Caracteristica de carrera].[Codigo] Where [Alumnos].[Permiso] = " & txtPermiso & " And [Carreras].[Codigo] = " & dtcCarreras.BoundText & " AND Materias.Detalle <> 5 AND Finales.Habilitada = True AND ((Finales.Aprobada = True AND Materias.Detalle BETWEEN 1 AND 3) OR Materias.Detalle=4) ORDER BY [Materias].[OrdenPlan],[Materias].[CursoAnalitico], [Materias].[Codigo]")
    Conexion.Execute ("UPDATE Titulo SET Certificado = '" & cbOriginal.Text & "'")
    Set Resultado = Conexion.Execute("SELECT NombreInstitucion, Localidad, Provincia FROM Parametros")
    Localidad = Resultado!Localidad
    Provincia = Resultado!Provincia
    PieDePagina = "En fe de lo cual se expide el presente certificado, " & Mid(cbOriginal.Text, 5) & " sin raspaduras y enmiendas, en la ciudad de " & Localidad & " (" & Provincia & ") a los " & dtFecha.Day & " dias del mes de " & adoMeses.Recordset!Nombre & " del año " & dtFecha.Year & ".-"
    Conexion.Execute ("UPDATE Titulo SET PieDePagina = '" & PieDePagina & "'")
    Conexion.Execute ("UPDATE Titulo SET Establecimiento = '" & Resultado!NombreINstitucion & "'")
    Set Resultado = Conexion.Execute("SELECT * FROM Titulo")
    If Resultado.EOF = True Then MsgBox ("El alumno no rindió ninguna materia para esta carrera"): Conexion.Close: Return
    CarrerayTitulo = Resultado!Tipo & " Nº " & Resultado!documento & "  ha aprobado " & Resultado!Articulo & " " & Resultado!Caracteristica & " que, con las respectivas calificaciones, que abajo se registran, correspondientes"
    If Resultado!CodigoCarrera = 51 Then
        TituloDeBase = InputBox("Ingrese el título de base")
        CarrerayTitulo = CarrerayTitulo & " al " & dtcTitulos.Text & " en conjunción con su título de base: " & TituloDeBase & " que lo habilita, de acuerdo al Nomenclador vigente, en cada nivel mencionado."
    Else
        CarrerayTitulo = CarrerayTitulo & " a la carrera " & Resultado!Carrera & ", haciéndose acreedor/a al título de " & dtcTitulos.Text
    End If
    CarrerayTitulo = CarrerayTitulo & " - Resolución " & Resultado!Resolucion & "."
    Conexion.Execute ("UPDATE Titulo SET CarrerayTitulo = '" & CarrerayTitulo & "'")
    Conexion.Execute ("UPDATE Titulo SET Equivalencia = '*' WHERE Equivalencia = '-1'")
    Conexion.Execute ("UPDATE Titulo SET Equivalencia = Null WHERE Equivalencia = '0'")
    adoTitulo.RecordSource = "SELECT * FROM Titulo WHERE Detalle = 1 OR DEtalle = 2"
    adoTitulo.Refresh
    While adoTitulo.Recordset.EOF = False
    Decimo = Int(adoTitulo.Recordset!Nota)
    centesimo = Right(Format(adoTitulo.Recordset!Nota, "0.00"), 3)
    If centesimo = ".00" Then centesimo = "" 'si es entero no se muestra decimales
    If Decimo = 1 Then EnLetras = "UNO" & centesimo
    If Decimo = 2 Then EnLetras = "DOS" & centesimo
    If Decimo = 3 Then EnLetras = "TRES" & centesimo
    If Decimo = 4 Then EnLetras = "CUATRO" & centesimo
    If Decimo = 5 Then EnLetras = "CINCO" & centesimo
    If Decimo = 6 Then EnLetras = "SEIS" & centesimo
    If Decimo = 7 Then EnLetras = "SIETE" & centesimo
    If Decimo = 8 Then EnLetras = "OCHO" & centesimo
    If Decimo = 9 Then EnLetras = "NUEVE" & centesimo
    If Decimo = 10 Then EnLetras = "DIEZ" & centesimo
    Conexion.Execute ("UPDATE Titulo SET EnLetras = '" & EnLetras & "' WHERE CodigoMateria = " & adoTitulo.Recordset!CodigoMateria)
    adoTitulo.Recordset.MoveNext
    Wend
    Set Resultado = Conexion.Execute("SELECT Avg(Nota) AS Promedio FROM Titulo WHERE Detalle = 1 OR Detalle = 2")
    Conexion.Execute ("UPDATE Titulo SET Promedio = " & Replace(Resultado!Promedio, ",", "."))
    Decimo = Int(Format(Resultado!Promedio, "0.00"))
    centesimo = Right(Format(Resultado!Promedio, "0.00"), 3)
    If Decimo = 1 Then PromedioLetras = "UNO" & centesimo
    If Decimo = 2 Then PromedioLetras = "DOS" & centesimo
    If Decimo = 3 Then PromedioLetras = "TRES" & centesimo
    If Decimo = 4 Then PromedioLetras = "CUATRO" & centesimo
    If Decimo = 5 Then PromedioLetras = "CINCO" & centesimo
    If Decimo = 6 Then PromedioLetras = "SEIS" & centesimo
    If Decimo = 7 Then PromedioLetras = "SIETE" & centesimo
    If Decimo = 8 Then PromedioLetras = "OCHO" & centesimo
    If Decimo = 9 Then PromedioLetras = "NUEVE" & centesimo
    If Decimo = 10 Then PromedioLetras = "DIEZ" & centesimo
    Conexion.Execute ("UPDATE Titulo SET Fecha = Null, Nota = Null, EnLetras = Null WHERE (Detalle = 3 or Detalle = 4)")
    Conexion.Execute ("UPDATE Titulo SET PromedioLetras = '" & PromedioLetras & "'")
    If txtObservaciones <> "" Then Conexion.Execute ("UPDATE Titulo SET Observaciones = 'Observaciones: " & txtObservaciones & "'")
    Conexion.Execute ("UPDATE Titulo SET Libro= '" & txtAlumnoLibro & "', Folio='" & txtAlumnoFolio & "', Numero='" & txtNumeroTitulo & "'")
    Conexion.Close
End Function

