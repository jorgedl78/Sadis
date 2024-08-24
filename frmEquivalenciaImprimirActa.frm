VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "msadodc.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmEquivalenciaImprimirActa 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3105
   ClientLeft      =   6495
   ClientTop       =   3840
   ClientWidth     =   5070
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3105
   ScaleWidth      =   5070
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Height          =   2655
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   4815
      Begin VB.CommandButton cmdAceptar 
         Height          =   615
         Left            =   1800
         Picture         =   "frmEquivalenciaImprimirActa.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Aceptar"
         Top             =   1800
         Width           =   615
      End
      Begin MSComCtl2.DTPicker dtpFechaDeLaMesa 
         DragIcon        =   "frmEquivalenciaImprimirActa.frx":0442
         Height          =   285
         Left            =   2040
         TabIndex        =   3
         Top             =   480
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   503
         _Version        =   393216
         Format          =   52428801
         CurrentDate     =   37550
         MinDate         =   -36522
      End
      Begin MSDataListLib.DataCombo dtcProfesor 
         Bindings        =   "frmEquivalenciaImprimirActa.frx":0884
         Height          =   315
         Left            =   360
         TabIndex        =   5
         Top             =   1200
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "Nombre"
         BoundColumn     =   "Codigo"
         Text            =   ""
      End
      Begin MSAdodcLib.Adodc adoProfesor 
         Height          =   330
         Left            =   2040
         Top             =   1080
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
         RecordSource    =   "SELECT * FROM Personal ORDER BY Nombre"
         Caption         =   "Profesor"
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
      Begin VB.Label Label3 
         Caption         =   "Profesor que otrga las equivalencias:"
         Height          =   255
         Left            =   360
         TabIndex        =   4
         Top             =   960
         Width           =   2895
      End
      Begin VB.Label Label2 
         Caption         =   "Fecha del acta:"
         Height          =   255
         Left            =   840
         TabIndex        =   2
         Top             =   480
         Width           =   1215
      End
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Imprimir Acta de Equivalencias Otorgadas"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   4815
   End
End
Attribute VB_Name = "frmEquivalenciaImprimirActa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdAceptar_Click()
    frmEquivalencias.FechaDeLaMesa = DateValue(dtpFechaDeLaMesa.Value)
    frmEquivalencias.NombreProfesor = dtcProfesor
    frmEquivalencias.ProfesorCodigo = dtcProfesor.BoundText
    Unload Me
End Sub

Private Sub Form_Load()
    dtpFechaDeLaMesa.Value = Date
    dtcProfesor.BoundText = frmEquivalencias.dtcProfesor.BoundText
End Sub
