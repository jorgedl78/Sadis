VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Begin VB.Form frmInformePorcentajeDeAvance 
   Caption         =   "Control de avance de porcentajes por Carrera"
   ClientHeight    =   3105
   ClientLeft      =   3435
   ClientTop       =   5400
   ClientWidth     =   10335
   LinkTopic       =   "Form1"
   ScaleHeight     =   3105
   ScaleWidth      =   10335
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   360
      TabIndex        =   15
      Text            =   "Text1"
      Top             =   2880
      Visible         =   0   'False
      Width           =   9495
   End
   Begin VB.Frame Frame1 
      Height          =   2895
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   10095
      Begin VB.ComboBox cmCondicion 
         Height          =   315
         Left            =   4800
         TabIndex        =   12
         Text            =   "Seleccione condición"
         Top             =   1680
         Width           =   2055
      End
      Begin Crystal.CrystalReport CrystalReport1 
         Left            =   2040
         Top             =   1080
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   262150
         ReportFileName  =   "avancepo.rpt"
         WindowTitle     =   "OListado de Porcentaje de Avance"
         WindowControlBox=   -1  'True
         WindowMaxButton =   -1  'True
         WindowMinButton =   -1  'True
      End
      Begin VB.Frame Frame2 
         Caption         =   "Parametros de Porcentaje"
         Height          =   1455
         Left            =   2160
         TabIndex        =   5
         Top             =   1200
         Width           =   5415
         Begin VB.TextBox txtHasta 
            Height          =   285
            Left            =   1200
            TabIndex        =   8
            Top             =   840
            Width           =   975
         End
         Begin VB.TextBox txtDesde 
            Height          =   285
            Left            =   1200
            TabIndex        =   7
            Top             =   480
            Width           =   975
         End
         Begin MSComCtl2.DTPicker dateFechaFactura 
            Height          =   375
            Left            =   3240
            TabIndex        =   13
            Top             =   840
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   661
            _Version        =   393216
            Format          =   88932353
            CurrentDate     =   42366
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Desde:"
            Height          =   195
            Left            =   2640
            TabIndex        =   14
            Top             =   960
            Width           =   510
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Desde: %"
            Height          =   195
            Left            =   360
            TabIndex        =   11
            Top             =   480
            Width           =   735
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Hasta: %"
            Height          =   195
            Left            =   360
            TabIndex        =   6
            Top             =   840
            Width           =   630
         End
      End
      Begin VB.CommandButton cmdMostrar 
         Caption         =   "Procesar Informe"
         Height          =   975
         Left            =   7680
         Picture         =   "frmInformePorcentajeDeAvance.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   1560
         Width           =   1215
      End
      Begin VB.CommandButton cmdSalir 
         Height          =   960
         Left            =   9000
         Picture         =   "frmInformePorcentajeDeAvance.frx":066A
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Salir"
         Top             =   1560
         Width           =   960
      End
      Begin MSAdodcLib.Adodc adoCarreras 
         Height          =   330
         Left            =   3960
         Top             =   240
         Visible         =   0   'False
         Width           =   2295
         _ExtentX        =   4048
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
         RecordSource    =   $"frmInformePorcentajeDeAvance.frx":0AAC
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
      Begin MSDataListLib.DataCombo dtcCarreras 
         Bindings        =   "frmInformePorcentajeDeAvance.frx":0B4D
         Height          =   315
         Left            =   120
         TabIndex        =   3
         Top             =   720
         Width           =   9855
         _ExtentX        =   17383
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "Nombre"
         BoundColumn     =   "Codigo"
         Text            =   ""
      End
      Begin VB.Label lblTotal 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   480
         TabIndex        =   10
         Top             =   1920
         Width           =   975
      End
      Begin VB.Label Label4 
         Caption         =   "Total de materias del plan"
         Height          =   495
         Left            =   120
         TabIndex        =   9
         Top             =   1440
         Width           =   1935
      End
      Begin VB.Label Label1 
         Caption         =   "Carreras Vigentes:"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   360
         Width           =   1455
      End
   End
End
Attribute VB_Name = "frmInformePorcentajeDeAvance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Conexion As New Connection
Dim Materias As New Recordset
Dim Correlativas As New Recordset
Dim Matriculados As New Recordset
Dim Final As New Recordset

Private Sub cmdMostrar_Click()
If cmCondicion.ListIndex < 0 Then MsgBox ("Debe elejir una condición"): Exit Sub
Respuesta = MsgBox("Este proceso puede tardar algunos minutos" & Chr(13) & "¿Desea continuar?", vbYesNo, "Atención")
If Respuesta = vbNo Then Exit Sub
Me.MousePointer = 11
Conexion.Open
Conexion.Execute ("DELETE * FROM rptPorcentajeAvance")
'Set Materias = Conexion.Execute("INSERT INTO rptPorcentajeAvance ( Permiso, Nombre, Tipo, Documento, Aprobadas, Porcentaje, Carrera, TotalPlan, Desde, Hasta ) SELECT Alumnos.Permiso, Alumnos.Nombre, Alumnos.Tipo, Alumnos.Documento, Count(Finales.Aprobada) AS Aprobadas, ((Count(Finales.Aprobada)*100)/" & lblTotal & ") AS Porcentaje, '" & dtcCarreras & "', " & lblTotal & ", " & txtDesde & "," & txtHasta & _
'" FROM Finales  INNER JOIN ((Alumnos INNER JOIN CarrerasHechas ON Alumnos.Permiso = CarrerasHechas.Permiso) INNER JOIN Materias ON CarrerasHechas.Carrera = Materias.Carrera) ON (Finales.Materia = Materias.Codigo) AND (Finales.Alumno = Alumnos.Permiso) Where (((CarrerasHechas.Carrera) =" & dtcCarreras.BoundText & ") And ((Finales.Aprobada) = True) And ((Materias.Detalle) = 1 Or (Materias.Detalle) = 2) And ((Materias.Eliminada) = 0)) GROUP BY Alumnos.Permiso, Alumnos.Nombre, Alumnos.Tipo, Alumnos.Documento HAVING ((((Count([Finales].[Aprobada])*100)/'" & lblTotal & "') Between " & Val(txtDesde) & " And " & Val(txtHasta) & ") AND ((CarrerasHechas.Condición)=1) AND ((CarrerasHechas.Fecha)>=#1/1/2010#)")
'Text1.Text = "INSERT INTO rptPorcentajeAvance ( Permiso, Nombre, Tipo, Documento, Aprobadas, Porcentaje, Carrera, TotalPlan, Desde, Hasta, Fecha, Ingreso ) SELECT Alumnos.Permiso, Alumnos.Nombre, Alumnos.Tipo, Alumnos.Documento, Count(Finales.Aprobada) AS Aprobadas, ((Count(Finales.Aprobada)*100)/" & lblTotal & ") AS Porcentaje, '" & dtcCarreras & "', " & lblTotal & ", " & txtDesde & "," & txtHasta & ", CarrerasHechas.Fecha, CarrerasHechas.Ingreso FROM Finales INNER JOIN ((Alumnos INNER JOIN CarrerasHechas ON Alumnos.Permiso = CarrerasHechas.Permiso) INNER JOIN Materias ON CarrerasHechas.Carrera = Materias.Carrera) ON (Finales.Alumno = Alumnos.Permiso) AND (Finales.Materia = Materias.Codigo) Where (((CarrerasHechas.Carrera) = " & dtcCarreras.BoundText & ") And ((Finales.Aprobada) = True) And ((Materias.Detalle) = 1 Or (Materias.Detalle) = 2) And ((Materias.Eliminada) = 0))" & _
'" GROUP BY Alumnos.Permiso, Alumnos.Nombre, Alumnos.Tipo, Alumnos.Documento, CarrerasHechas.Condición, CarrerasHechas.Fecha, CarrerasHechas.Ingreso HAVING ((((Count([Finales].[Aprobada])*100)/'" & lblTotal & "') Between " & Val(txtDesde) & " And " & Val(txtHasta) & ") AND ((CarrerasHechas.Condición)=" & cmCondicion.ItemData(cmCondicion.ListIndex) & ") AND ((CarrerasHechas.Fecha)>#" & Format(dateFechaFactura, "dd/mm/yyyy") & "#))"

' consulta con promedio agregado
'SELECT A.Permiso, A.Nombre, A.Tipo, A.Documento, Count(F.Aprobada) AS Aprobadas, ((Count(F.Aprobada)*100)/50) AS Porcentaje, 74 AS Expr1, 100 AS Expr2, 20 AS Expr3, 100 AS Expr4, CarrerasHechas.Fecha, CarrerasHechas.Ingreso, (SELECT Avg(Nota) AS Promedio
'FROM  Finales AS FF INNER JOIN Materias ON FF.Materia = Materias.codigo
'Where (((Materias.Detalle) = 1 Or (Materias.Detalle) = 2) And ((FF.Aprobada) = True) And ((FF.Alumno) = A.Permiso) And ((Materias.Carrera) = 74))
') as Promedio
'FROM Finales AS F INNER JOIN ((Alumnos AS A INNER JOIN CarrerasHechas ON A.Permiso = CarrerasHechas.Permiso) INNER JOIN Materias ON CarrerasHechas.Carrera = Materias.Carrera) ON (F.Materia = Materias.Codigo) AND (F.Alumno = A.Permiso)
'Where (((CarrerasHechas.Carrera) = 74) And ((F.Aprobada) = True) And ((Materias.Detalle) = 1 Or (Materias.Detalle) = 2) And ((Materias.Eliminada) = 0))
'GROUP BY A.Permiso, A.Nombre, A.Tipo, A.Documento, CarrerasHechas.Fecha, CarrerasHechas.Ingreso, CarrerasHechas.Condición
'HAVING (((CarrerasHechas.Fecha)>=#1/1/2019#) AND ((CarrerasHechas.Condición)=1) AND (((Count([F].[Aprobada])*100)/20) Between 20 And 80));



Conexion.Execute ("INSERT INTO rptPorcentajeAvance ( Permiso, Nombre, Tipo, Documento, Aprobadas, Porcentaje, Carrera, TotalPlan, Desde, Hasta, Fecha, Ingreso ) SELECT Alumnos.Permiso, Alumnos.Nombre, Alumnos.Tipo, Alumnos.Documento, Count(Finales.Aprobada) AS Aprobadas, ((Count(Finales.Aprobada)*100)/" & lblTotal & ") AS Porcentaje, '" & dtcCarreras & "', " & lblTotal & ", " & txtDesde & "," & txtHasta & ", CarrerasHechas.Fecha, CarrerasHechas.Ingreso FROM Finales INNER JOIN ((Alumnos INNER JOIN CarrerasHechas ON Alumnos.Permiso = CarrerasHechas.Permiso) INNER JOIN Materias ON CarrerasHechas.Carrera = Materias.Carrera) ON (Finales.Alumno = Alumnos.Permiso) AND (Finales.Materia = Materias.Codigo) Where (((CarrerasHechas.Carrera) = " & dtcCarreras.BoundText & ") And ((Finales.Aprobada) = True) And ((Materias.Detalle) = 1 Or (Materias.Detalle) = 2) And ((Materias.Eliminada) = 0))" & _
" GROUP BY Alumnos.Permiso, Alumnos.Nombre, Alumnos.Tipo, Alumnos.Documento, CarrerasHechas.Condición, CarrerasHechas.Fecha, CarrerasHechas.Ingreso HAVING ((((Count([Finales].[Aprobada])*100)/'" & lblTotal & "') Between " & Val(txtDesde) & " And " & Val(txtHasta) & ") AND ((CarrerasHechas.Condición)=" & cmCondicion.ItemData(cmCondicion.ListIndex) & ") AND ((CarrerasHechas.Fecha)>=#" & Format(dateFechaFactura, "mm/dd/yyyy") & "#))")
Conexion.Close
Me.MousePointer = 0
CrystalReport1.PrintReport
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub


Private Sub Command1_Click()

End Sub

Private Sub dtcCarreras_Change()
   Conexion.Open
   Set Materias = Conexion.Execute("SELECT Count(Codigo) As Total FROM Materias WHERE (Detalle=1 Or Detalle=2) AND Carrera=" & dtcCarreras.BoundText)
   lblTotal = Materias!Total
   Conexion.Close
End Sub

Private Sub Form_Load()
    Conexion.ConnectionString = ("DSN=Instituto")
    txtAño = Format(Date, "yyyy")
    dtcCarreras.BoundText = adoCarreras.Recordset!Codigo
    
   Conexion.Open
   Set rs = Conexion.Execute("SELECT Codigo, Condicion from Condicion")
    Do While rs.EOF = False
        cmCondicion.AddItem (rs!Condicion)
        cmCondicion.ItemData(cmCondicion.NewIndex) = rs!Codigo
        rs.MoveNext
    Loop
    Conexion.Close
    dateFechaFactura = Date

    
On Error GoTo hErr
   Conexion.Open
   Conexion.Execute ("Create Table rptPorcentajeAvance (Permiso int,Nombre Text (50),Tipo Text (3),Documento int, Aprobadas int, Porcentaje int, Carrera Text (120), TotalPlan int, Desde int, Hasta int, Fecha, Ingreso)")
   Conexion.Close
   Exit Sub
hErr:
   'MsgBox Err.Number & " " & Err.Description
   Conexion.Close
   Exit Sub
End Sub




Private Sub txtAño_Change()

End Sub

