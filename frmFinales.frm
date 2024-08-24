VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmFinales 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Control de Exámenes Finales"
   ClientHeight    =   10515
   ClientLeft      =   -3105
   ClientTop       =   -3150
   ClientWidth     =   15030
   ControlBox      =   0   'False
   DrawMode        =   11  'Not Xor Pen
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   16551.39
   ScaleMode       =   0  'User
   ScaleWidth      =   15030
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame frFinales 
      Caption         =   "Cursadas aprobadas con sus correspondientes exámenes finales"
      Height          =   8415
      Left            =   7440
      TabIndex        =   16
      Top             =   1920
      Width           =   7455
      Begin TabDlg.SSTab SSTab1 
         Height          =   2295
         Left            =   240
         TabIndex        =   37
         Top             =   5760
         Width           =   7095
         _ExtentX        =   12515
         _ExtentY        =   4048
         _Version        =   393216
         Tabs            =   2
         Tab             =   1
         TabsPerRow      =   2
         TabHeight       =   520
         TabCaption(0)   =   "Cursada y Asistencia"
         TabPicture(0)   =   "frmFinales.frx":0000
         Tab(0).ControlEnabled=   0   'False
         Tab(0).Control(0)=   "Label14"
         Tab(0).Control(1)=   "Label10"
         Tab(0).Control(2)=   "frCursada"
         Tab(0).ControlCount=   3
         TabCaption(1)   =   "Final"
         TabPicture(1)   =   "frmFinales.frx":001C
         Tab(1).ControlEnabled=   -1  'True
         Tab(1).Control(0)=   "Label18"
         Tab(1).Control(0).Enabled=   0   'False
         Tab(1).Control(1)=   "Label17"
         Tab(1).Control(1).Enabled=   0   'False
         Tab(1).Control(2)=   "Label13"
         Tab(1).Control(2).Enabled=   0   'False
         Tab(1).Control(3)=   "Label9"
         Tab(1).Control(3).Enabled=   0   'False
         Tab(1).Control(4)=   "Label15"
         Tab(1).Control(4).Enabled=   0   'False
         Tab(1).Control(5)=   "Label16"
         Tab(1).Control(5).Enabled=   0   'False
         Tab(1).Control(6)=   "Label25"
         Tab(1).Control(6).Enabled=   0   'False
         Tab(1).Control(7)=   "dtpFecha"
         Tab(1).Control(7).Enabled=   0   'False
         Tab(1).Control(8)=   "txtActa"
         Tab(1).Control(8).Enabled=   0   'False
         Tab(1).Control(9)=   "txtMesa"
         Tab(1).Control(9).Enabled=   0   'False
         Tab(1).Control(10)=   "txtCantidadMesas"
         Tab(1).Control(10).Enabled=   0   'False
         Tab(1).Control(11)=   "chkAprobada"
         Tab(1).Control(11).Enabled=   0   'False
         Tab(1).Control(12)=   "txtNota"
         Tab(1).Control(12).Enabled=   0   'False
         Tab(1).Control(13)=   "txtFolio"
         Tab(1).Control(13).Enabled=   0   'False
         Tab(1).Control(14)=   "frEquivalencia"
         Tab(1).Control(14).Enabled=   0   'False
         Tab(1).Control(15)=   "chkHabilitada"
         Tab(1).Control(15).Enabled=   0   'False
         Tab(1).Control(16)=   "txtLibro"
         Tab(1).Control(16).Enabled=   0   'False
         Tab(1).ControlCount=   17
         Begin VB.TextBox txtLibro 
            Enabled         =   0   'False
            Height          =   285
            Left            =   4200
            TabIndex        =   85
            Top             =   780
            Width           =   615
         End
         Begin VB.CheckBox chkHabilitada 
            Alignment       =   1  'Right Justify
            Caption         =   "Habilitado para imprimir en título:"
            Enabled         =   0   'False
            Height          =   195
            Left            =   2160
            TabIndex        =   57
            Top             =   1200
            Width           =   2655
         End
         Begin VB.Frame frEquivalencia 
            Height          =   615
            Left            =   120
            TabIndex        =   53
            Top             =   1440
            Width           =   6855
            Begin VB.TextBox txtEstablecimiento 
               Enabled         =   0   'False
               Height          =   285
               Left            =   3120
               MultiLine       =   -1  'True
               TabIndex        =   55
               Top             =   240
               Width           =   3615
            End
            Begin VB.CheckBox chkEquivalencia 
               Caption         =   "Equivalencia:"
               Enabled         =   0   'False
               Height          =   255
               Left            =   120
               TabIndex        =   54
               Top             =   240
               Width           =   1455
            End
            Begin VB.Label Label11 
               Caption         =   "Establecimiento"
               Height          =   255
               Left            =   1920
               TabIndex        =   56
               Top             =   240
               Width           =   1215
            End
         End
         Begin VB.TextBox txtFolio 
            Enabled         =   0   'False
            Height          =   285
            Left            =   4920
            TabIndex        =   46
            Top             =   780
            Width           =   615
         End
         Begin VB.TextBox txtNota 
            Enabled         =   0   'False
            Height          =   285
            Left            =   240
            TabIndex        =   45
            Top             =   780
            Width           =   615
         End
         Begin VB.CheckBox chkAprobada 
            Alignment       =   1  'Right Justify
            Caption         =   "Aprobo final:"
            Enabled         =   0   'False
            Height          =   195
            Left            =   5640
            TabIndex        =   44
            Top             =   840
            Width           =   1215
         End
         Begin VB.TextBox txtCantidadMesas 
            Enabled         =   0   'False
            Height          =   285
            Left            =   1560
            TabIndex        =   42
            Top             =   1200
            Width           =   375
         End
         Begin VB.TextBox txtMesa 
            Enabled         =   0   'False
            Height          =   285
            Left            =   2520
            TabIndex        =   41
            Top             =   780
            Width           =   855
         End
         Begin VB.TextBox txtActa 
            Enabled         =   0   'False
            Height          =   285
            Left            =   3480
            TabIndex        =   40
            Top             =   780
            Width           =   615
         End
         Begin MSComCtl2.DTPicker dtpFecha 
            DragIcon        =   "frmFinales.frx":0038
            Height          =   285
            Left            =   1080
            TabIndex        =   43
            Top             =   780
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   503
            _Version        =   393216
            Enabled         =   0   'False
            Format          =   119603201
            CurrentDate     =   37550
            MinDate         =   -36522
         End
         Begin VB.Frame frCursada 
            Height          =   1875
            Left            =   -74760
            TabIndex        =   58
            Top             =   360
            Width           =   6615
            Begin VB.TextBox txtComentario 
               Height          =   285
               Left            =   1200
               MaxLength       =   30
               TabIndex        =   84
               Top             =   1440
               Width           =   5295
            End
            Begin VB.TextBox txtPorcentaje 
               Enabled         =   0   'False
               Height          =   285
               Left            =   6120
               TabIndex        =   81
               Top             =   360
               Width           =   375
            End
            Begin VB.CheckBox chkAsistencia 
               Alignment       =   1  'Right Justify
               Caption         =   "Asistencia:"
               Height          =   250
               Left            =   4560
               TabIndex        =   80
               Top             =   360
               Width           =   1095
            End
            Begin VB.TextBox txtDivision 
               Enabled         =   0   'False
               Height          =   285
               Left            =   2520
               TabIndex        =   79
               Top             =   1080
               Width           =   375
            End
            Begin VB.TextBox txtAnoCursada 
               Enabled         =   0   'False
               Height          =   285
               Left            =   1200
               TabIndex        =   78
               Top             =   1080
               Width           =   615
            End
            Begin VB.TextBox txtCursoCon 
               DataField       =   "CursoCon"
               DataSource      =   "adoFinales"
               Enabled         =   0   'False
               Height          =   285
               Left            =   3960
               TabIndex        =   77
               Top             =   840
               Width           =   2535
            End
            Begin VB.CheckBox chkCursadaAprobada 
               Alignment       =   1  'Right Justify
               Caption         =   "Aprobada"
               Enabled         =   0   'False
               Height          =   195
               Left            =   3120
               TabIndex        =   66
               Top             =   360
               Width           =   1215
            End
            Begin VB.TextBox txtParcial1 
               Enabled         =   0   'False
               Height          =   300
               Left            =   1200
               TabIndex        =   65
               Top             =   360
               Width           =   495
            End
            Begin VB.TextBox txtRecuperatorio3 
               Enabled         =   0   'False
               Height          =   300
               Left            =   2400
               TabIndex        =   64
               Top             =   720
               Width           =   495
            End
            Begin VB.TextBox txtParcial3 
               Enabled         =   0   'False
               Height          =   300
               Left            =   2400
               TabIndex        =   63
               Top             =   360
               Width           =   495
            End
            Begin VB.TextBox txtRecuperatorio2 
               Enabled         =   0   'False
               Height          =   300
               Left            =   1800
               TabIndex        =   62
               Top             =   720
               Width           =   495
            End
            Begin VB.TextBox txtParcial2 
               Enabled         =   0   'False
               Height          =   300
               Left            =   1800
               TabIndex        =   61
               Top             =   360
               Width           =   495
            End
            Begin VB.TextBox txtRecuperatorio1 
               Enabled         =   0   'False
               Height          =   300
               Left            =   1200
               TabIndex        =   60
               Top             =   720
               Width           =   495
            End
            Begin VB.CheckBox chkPromociono 
               Alignment       =   1  'Right Justify
               Caption         =   "Promocionó"
               Enabled         =   0   'False
               Height          =   195
               Left            =   3120
               TabIndex        =   59
               Top             =   600
               Width           =   1215
            End
            Begin VB.Label Label24 
               Alignment       =   1  'Right Justify
               Caption         =   "Comentario:"
               Height          =   255
               Left            =   120
               TabIndex        =   83
               Top             =   1440
               Width           =   975
            End
            Begin VB.Label Label19 
               Caption         =   "Cursó con:"
               Height          =   255
               Left            =   3120
               TabIndex        =   76
               Top             =   840
               Width           =   855
            End
            Begin VB.Label lblPerdioCursada 
               Caption         =   "Cursada Vencida"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000C0&
               Height          =   255
               Left            =   3000
               TabIndex        =   75
               Top             =   120
               Visible         =   0   'False
               Width           =   1575
            End
            Begin VB.Label Label2 
               Caption         =   "Parcial:"
               Height          =   255
               Left            =   600
               TabIndex        =   74
               Top             =   360
               Width           =   615
            End
            Begin VB.Label Label3 
               Caption         =   "Recuperatorio:"
               Height          =   255
               Left            =   120
               TabIndex        =   73
               Top             =   720
               Width           =   1095
            End
            Begin VB.Label Label4 
               Caption         =   "1º"
               Height          =   255
               Left            =   1320
               TabIndex        =   72
               Top             =   120
               Width           =   255
            End
            Begin VB.Label Label6 
               Caption         =   "2º"
               Height          =   255
               Left            =   1920
               TabIndex        =   71
               Top             =   120
               Width           =   255
            End
            Begin VB.Label Label7 
               Caption         =   "3º"
               Height          =   255
               Left            =   2520
               TabIndex        =   70
               Top             =   120
               Width           =   255
            End
            Begin VB.Label Label21 
               Alignment       =   1  'Right Justify
               Caption         =   "Año:"
               Height          =   255
               Left            =   600
               TabIndex        =   69
               Top             =   1080
               Width           =   495
            End
            Begin VB.Label Label22 
               Caption         =   "%"
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
               Left            =   5880
               TabIndex        =   68
               Top             =   360
               Width           =   255
            End
            Begin VB.Label Label23 
               Caption         =   "División:"
               Height          =   255
               Left            =   1920
               TabIndex        =   67
               Top             =   1080
               Width           =   615
            End
         End
         Begin VB.Label Label25 
            Caption         =   "Libro"
            Height          =   255
            Left            =   4200
            TabIndex        =   86
            Top             =   540
            Width           =   375
         End
         Begin VB.Label Label16 
            Caption         =   "Fecha:"
            Height          =   255
            Left            =   1080
            TabIndex        =   52
            Top             =   540
            Width           =   615
         End
         Begin VB.Label Label15 
            Caption         =   "Folio:"
            Height          =   255
            Left            =   4920
            TabIndex        =   51
            Top             =   540
            Width           =   375
         End
         Begin VB.Label Label9 
            Caption         =   "Nota:"
            Height          =   255
            Left            =   240
            TabIndex        =   50
            Top             =   540
            Width           =   495
         End
         Begin VB.Label Label13 
            Caption         =   "Turnos rendidos:"
            Height          =   255
            Left            =   240
            TabIndex        =   49
            Top             =   1200
            Width           =   1335
         End
         Begin VB.Label Label17 
            Caption         =   "Nº Mesa"
            Height          =   255
            Left            =   2520
            TabIndex        =   48
            Top             =   540
            Width           =   735
         End
         Begin VB.Label Label18 
            Caption         =   "Acta"
            Height          =   255
            Left            =   3480
            TabIndex        =   47
            Top             =   540
            Width           =   375
         End
         Begin VB.Label Label10 
            Caption         =   "Division:"
            Height          =   255
            Left            =   -70800
            TabIndex        =   39
            Top             =   1020
            Width           =   615
         End
         Begin VB.Label Label14 
            Caption         =   "Cursada:"
            Height          =   255
            Left            =   -71640
            TabIndex        =   38
            Top             =   1020
            Width           =   735
         End
      End
      Begin VB.Frame frModificaciones 
         BorderStyle     =   0  'None
         Height          =   2175
         Left            =   240
         TabIndex        =   32
         Top             =   6120
         Width           =   7095
      End
      Begin MSAdodcLib.Adodc adoFinales 
         Height          =   330
         Left            =   3360
         Top             =   960
         Visible         =   0   'False
         Width           =   1800
         _ExtentX        =   3175
         _ExtentY        =   582
         ConnectMode     =   1
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
         RecordSource    =   $"frmFinales.frx":047A
         Caption         =   "Finales"
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
      Begin MSDataGridLib.DataGrid dtgCursadas 
         Bindings        =   "frmFinales.frx":0890
         Height          =   5055
         Left            =   120
         TabIndex        =   18
         Top             =   240
         Width           =   7215
         _ExtentX        =   12726
         _ExtentY        =   8916
         _Version        =   393216
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
         ColumnCount     =   5
         BeginProperty Column00 
            DataField       =   "Curso"
            Caption         =   ""
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
         BeginProperty Column01 
            DataField       =   "Codigo"
            Caption         =   "Codigo"
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
         BeginProperty Column02 
            DataField       =   "Nombre"
            Caption         =   "Nombre"
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
            DataField       =   "Fecha"
            Caption         =   "Fecha"
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
         BeginProperty Column04 
            DataField       =   "Nota"
            Caption         =   "Nota"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "0.00"
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
            BeginProperty Column00 
               Alignment       =   2
               ColumnWidth     =   390,047
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   750,047
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   3855,118
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   929,764
            EndProperty
            BeginProperty Column04 
               ColumnWidth     =   689,953
            EndProperty
         EndProperty
      End
      Begin VB.Label lblTotalAprobadas 
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
         TabIndex        =   35
         Top             =   5400
         Width           =   4455
      End
      Begin VB.Label lblPorcentajeAprobado 
         Caption         =   "%"
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
         Left            =   4800
         TabIndex        =   36
         Top             =   5400
         Width           =   735
      End
   End
   Begin VB.Frame frBotones 
      BackColor       =   &H8000000C&
      BorderStyle     =   0  'None
      Height          =   8415
      Left            =   6720
      TabIndex        =   15
      Top             =   1920
      Width           =   615
      Begin VB.CommandButton cmdCancelarFinal 
         Enabled         =   0   'False
         Height          =   600
         Left            =   0
         Picture         =   "frmFinales.frx":08A9
         Style           =   1  'Graphical
         TabIndex        =   23
         ToolTipText     =   "Cancelar"
         Top             =   5400
         Width           =   600
      End
      Begin VB.CommandButton cmdEliminarFinal 
         Enabled         =   0   'False
         Height          =   600
         Left            =   0
         Picture         =   "frmFinales.frx":0CEB
         Style           =   1  'Graphical
         TabIndex        =   22
         ToolTipText     =   "Borrar"
         Top             =   4320
         Width           =   600
      End
      Begin VB.CommandButton cmdGuardarFinal 
         Enabled         =   0   'False
         Height          =   600
         Left            =   0
         Picture         =   "frmFinales.frx":112D
         Style           =   1  'Graphical
         TabIndex        =   21
         ToolTipText     =   "Guardar"
         Top             =   3240
         Width           =   600
      End
      Begin VB.CommandButton cmdModificarFinal 
         Enabled         =   0   'False
         Height          =   600
         Left            =   0
         Picture         =   "frmFinales.frx":156F
         Style           =   1  'Graphical
         TabIndex        =   20
         ToolTipText     =   "Modificar"
         Top             =   2160
         Width           =   600
      End
      Begin VB.CommandButton cmdAgregarFinal 
         Enabled         =   0   'False
         Height          =   600
         Left            =   0
         Picture         =   "frmFinales.frx":19B1
         Style           =   1  'Graphical
         TabIndex        =   19
         ToolTipText     =   "Agregar"
         Top             =   1080
         Width           =   600
      End
      Begin VB.CommandButton cmdSalir 
         Height          =   600
         Left            =   0
         Picture         =   "frmFinales.frx":1DF3
         Style           =   1  'Graphical
         TabIndex        =   17
         ToolTipText     =   "Salir"
         Top             =   6480
         Width           =   600
      End
   End
   Begin VB.Frame frPlanCompleto 
      Caption         =   "Plan Completo"
      Height          =   8535
      Left            =   -120
      TabIndex        =   13
      Top             =   1800
      Width           =   6735
      Begin MSAdodcLib.Adodc adoPlanCompleto 
         Height          =   330
         Left            =   1080
         Top             =   960
         Visible         =   0   'False
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   582
         ConnectMode     =   0
         CursorLocation  =   3
         IsolationLevel  =   -1
         ConnectionTimeout=   15
         CommandTimeout  =   30
         CursorType      =   2
         LockType        =   1
         CommandType     =   8
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
         RecordSource    =   $"frmFinales.frx":2235
         Caption         =   "PlanCompleto"
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
      Begin MSDataGridLib.DataGrid dtgPlanCompleto 
         Bindings        =   "frmFinales.frx":2310
         Height          =   7695
         Left            =   120
         TabIndex        =   14
         Top             =   240
         Width           =   6495
         _ExtentX        =   11456
         _ExtentY        =   13573
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
            DataField       =   "Curso"
            Caption         =   ""
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
         BeginProperty Column01 
            DataField       =   "Codigo"
            Caption         =   "Codigo"
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
         BeginProperty Column02 
            DataField       =   "Abreviatura"
            Caption         =   "Abreviatura"
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
            DataField       =   "Detalle"
            Caption         =   "Detalle"
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
            BeginProperty Column00 
               ColumnWidth     =   374,74
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   824,882
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   3420,284
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   1260,284
            EndProperty
         EndProperty
      End
      Begin VB.Label lblTotalPlan 
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
         TabIndex        =   34
         Top             =   8040
         Width           =   3735
      End
      Begin VB.Label Label20 
         Caption         =   "Label20"
         Height          =   15
         Left            =   360
         TabIndex        =   33
         Top             =   6480
         Width           =   615
      End
   End
   Begin VB.Frame frAlumnos 
      Caption         =   "Alumno"
      Height          =   1695
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   14775
      Begin VB.TextBox txtAlumnoTipo 
         DataField       =   "Tipo"
         DataSource      =   "adoAlumnos"
         Enabled         =   0   'False
         Height          =   315
         Left            =   7800
         TabIndex        =   31
         Top             =   480
         Width           =   495
      End
      Begin VB.TextBox txtAlumnoFolio 
         DataField       =   "Folio"
         DataSource      =   "adoCarreras"
         Enabled         =   0   'False
         Height          =   315
         Left            =   13920
         TabIndex        =   30
         Top             =   1200
         Width           =   615
      End
      Begin VB.TextBox txtAlumnoFecha 
         DataField       =   "Fecha"
         DataSource      =   "adoCarreras"
         Enabled         =   0   'False
         Height          =   315
         Left            =   12000
         TabIndex        =   29
         Top             =   1200
         Width           =   1095
      End
      Begin VB.TextBox txtAlumnoLibro 
         DataField       =   "Libro"
         DataSource      =   "adoCarreras"
         Enabled         =   0   'False
         Height          =   315
         Left            =   13200
         TabIndex        =   28
         Top             =   1200
         Width           =   615
      End
      Begin VB.TextBox txtAlumnoCondicion 
         DataField       =   "Condicion"
         DataSource      =   "adoCarreras"
         Enabled         =   0   'False
         Height          =   315
         Left            =   10440
         TabIndex        =   27
         Top             =   1200
         Width           =   1455
      End
      Begin VB.TextBox txtAlumnoIngreso 
         DataField       =   "Ingreso"
         DataSource      =   "adoCarreras"
         Enabled         =   0   'False
         Height          =   315
         Left            =   9840
         TabIndex        =   26
         Top             =   480
         Width           =   1095
      End
      Begin VB.TextBox txtAlumnoDocumento 
         DataField       =   "Documento"
         DataSource      =   "adoAlumnos"
         Enabled         =   0   'False
         Height          =   315
         Left            =   8280
         TabIndex        =   25
         Top             =   480
         Width           =   1335
      End
      Begin VB.TextBox txtAlumnoNombre 
         DataField       =   "Nombre"
         DataSource      =   "adoAlumnos"
         Enabled         =   0   'False
         Height          =   315
         Left            =   2040
         TabIndex        =   24
         Top             =   480
         Width           =   5415
      End
      Begin MSDataListLib.DataCombo dtcCarreras 
         Bindings        =   "frmFinales.frx":232E
         Height          =   315
         Left            =   120
         TabIndex        =   12
         Top             =   1200
         Width           =   10215
         _ExtentX        =   18018
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "Nombre"
         BoundColumn     =   "Carrera"
         Text            =   ""
      End
      Begin VB.CommandButton cmdBuscarAlumno 
         Caption         =   "Buscar..."
         Height          =   315
         Left            =   960
         TabIndex        =   3
         Top             =   480
         Width           =   855
      End
      Begin VB.TextBox txtPermiso 
         Height          =   315
         Left            =   120
         MaxLength       =   5
         TabIndex        =   1
         Top             =   480
         Width           =   615
      End
      Begin MSAdodcLib.Adodc adoAlumnos 
         Height          =   330
         Left            =   120
         Top             =   1440
         Visible         =   0   'False
         Width           =   2055
         _ExtentX        =   3625
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
         Top             =   1440
         Visible         =   0   'False
         Width           =   2055
         _ExtentX        =   3625
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
         RecordSource    =   $"frmFinales.frx":2348
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
      Begin VB.Label lblAlumnoNombre 
         Caption         =   "Nombre"
         Height          =   255
         Left            =   2040
         TabIndex        =   11
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label8 
         Caption         =   "Documento:"
         Height          =   255
         Left            =   7800
         TabIndex        =   10
         Top             =   240
         Width           =   975
      End
      Begin VB.Label lblLibroAlumno 
         Caption         =   "Libro:"
         Height          =   255
         Left            =   13200
         TabIndex        =   9
         Top             =   960
         Width           =   495
      End
      Begin VB.Label lblFolioAlumno 
         Caption         =   "Folio:"
         Height          =   255
         Left            =   13920
         TabIndex        =   8
         Top             =   960
         Width           =   495
      End
      Begin VB.Label Label5 
         Caption         =   "Carrera"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   960
         Width           =   855
      End
      Begin VB.Label lblIngresoAlumno 
         Caption         =   "Ingresó:"
         Height          =   255
         Left            =   9840
         TabIndex        =   6
         Top             =   240
         Width           =   615
      End
      Begin VB.Label lblCondicionAlumno 
         Caption         =   "Condición:"
         Height          =   255
         Left            =   10440
         TabIndex        =   5
         Top             =   960
         Width           =   855
      End
      Begin VB.Label lblFechaAlumno 
         Caption         =   "Fecha:"
         Height          =   255
         Left            =   12000
         TabIndex        =   4
         Top             =   960
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "Nº Permiso"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.Label Label12 
      Caption         =   "Año de cursada:"
      Height          =   375
      Left            =   0
      TabIndex        =   82
      Top             =   0
      Width           =   3735
   End
End
Attribute VB_Name = "frmFinales"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Conexion As New Connection
Dim Estado As String
Dim Esta As String 'para saber si la materia ya esta cargada al intentar agregar
Dim RegistroActual As String
Dim CantidadPlan As New Recordset
Dim CantidadAprobadas As New Recordset
Dim BanderaMostrarCursada As String
Dim TotalPlan As Integer
Dim TotalAprobadas As Integer
Dim CursorEn As Double

Private Sub chkCursadaAprobada_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    If chkCursadaAprobada.Value = 1 Then
'       chkCursadaAprobada.Value = 0
'       Respuesta = MsgBox("Esta cursada se enuentra vencida", vbCritical, "Imposible aprobar cursada")
'       Exit Sub
'    End If
End Sub

Private Sub chkEquivalencia_Click()
    If chkEquivalencia.Value = 0 Then
        'txtEstablecimiento.Enabled = False
    Else
        'txtEstablecimiento.Enabled = True
    End If
End Sub

Private Sub cmdAgregarFinal_Click()
    Estado = "Agregando"
    Esta = "No"
    BanderaMostrarCursada = "No"
    VerExistencia
    BanderaMostrarCursada = "Si"
    If Esta = "Si" Then Exit Sub
    RegistroActual = adoPlanCompleto.Recordset!Codigo
    cmdAgregarFinal.Enabled = False
    cmdModificarFinal.Enabled = False
    cmdGuardarFinal.Enabled = True
    cmdEliminarFinal.Enabled = False
    cmdCancelarFinal.Enabled = True
    cmdSalir.Enabled = False
    LimpiarCursadas
    txtAnoCursada = 0
    txtDivision = 1
    HabilitarCursadas
    chkAsistencia.Value = 1: chkAsistencia.Enabled = True
    If adoPlanCompleto.Recordset!CodigoDetalle = 3 Or adoPlanCompleto.Recordset!CodigoDetalle = 4 Then chkHabilitada.Value = 1: cmdGuardarFinal_Click
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

Private Sub cmdCancelarFinal_Click()
    If Estado = "Agregando" Then
        If adoFinales.Recordset.RecordCount > 0 Then
            adoFinales.Recordset.MoveFirst
            adoFinales.Recordset.Find ("Codigo=" & CursorEn)
        End If
    End If
    DeshabilitarCursadas
    cmdAgregarFinal.Enabled = True
    cmdModificarFinal.Enabled = True
    cmdGuardarFinal.Enabled = False
    cmdEliminarFinal.Enabled = True
    cmdCancelarFinal.Enabled = False
    cmdSalir.Enabled = True
    MostrarCursadas
End Sub

Private Sub cmdEliminarFinal_Click()
    MsgBox ("No se pueden eliminar materias desde este formulario"): Exit Sub
    Respuesta = MsgBox("¿Está seguro de eliminar la materia " & adoFinales.Recordset!Nombre & "?", vbYesNo, "Eliminar Final")
    If Respuesta = vbYes Then
        Conexion.Open
        Conexion.Execute ("DELETE * FROM Finales WHERE Alumno = " & adoAlumnos.Recordset!Permiso & " AND Materia=" & adoFinales.Recordset!Codigo & " AND Ano=" & adoFinales.Recordset!Ano)
        Conexion.Close
        adoFinales.Refresh
    End If
End Sub

Private Sub cmdGuardarFinal_Click()
    Me.MousePointer = 11
    cmdAgregarFinal.Enabled = True
    cmdModificarFinal.Enabled = True
    cmdGuardarFinal.Enabled = False
    cmdEliminarFinal.Enabled = True
    cmdCancelarFinal.Enabled = False
    cmdSalir.Enabled = True
    DeshabilitarCursadas
    PasarCursadas
    Me.MousePointer = 0
    dtgCursadas.SetFocus
End Sub

Private Sub cmdModificarFinal_Click()
    If Val(txtMesa) > 0 Then
        Respuesta = MsgBox("Solo se puede hacer desde el Ingreso de Actas", , "Imposible modificar desde este formulario")
        chkHabilitada.Enabled = True
    Else
        If adoFinales.Recordset!Detalle = 4 Then
            chkHabilitada.Enabled = True
        Else
           HabilitarCursadas
        End If
    End If
    RegistroActual = adoFinales.Recordset!Codigo
    cmdAgregarFinal.Enabled = False
    cmdModificarFinal.Enabled = False
    cmdGuardarFinal.Enabled = True
    cmdEliminarFinal.Enabled = False
    cmdCancelarFinal.Enabled = True
    cmdSalir.Enabled = False
    Estado = "Modificando"
End Sub

Private Sub cmdSalir_Click()
    Unload Me
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

Private Sub dtcCarreras_Change()
    If dtcCarreras.Text <> "" Then
        adoCarreras.Recordset.MoveFirst
        adoCarreras.Recordset.Find ("Carrera=" & dtcCarreras.BoundText)
        adoPlanCompleto.RecordSource = "SELECT Materias.Curso, Materias.Codigo, Materias.Abreviatura, Materias.Detalle AS CodigoDetalle, Detalles.Detalle FROM Materias INNER JOIN Detalles ON Materias.Detalle = Detalles.Codigo WHERE Eliminada=0 AND Materias.Carrera=" & dtcCarreras.BoundText & " ORDER BY Materias.Codigo"
        adoPlanCompleto.Refresh
        Conexion.Open
        Set CantidadPlan = Conexion.Execute("SELECT Count(Materias.Nombre) AS Cantidad From Materias WHERE Materias.Carrera=" & dtcCarreras.BoundText & " AND (Materias.Detalle=1 OR Materias.Detalle=2) AND Eliminada=0")
        lblTotalPlan = "Total del materias del plan: " & CantidadPlan!Cantidad
        TotalPlan = CantidadPlan!Cantidad
        Conexion.Close
        If adoPlanCompleto.Recordset.RecordCount > 0 And frmIdentificacion.Permisos!ModificarFinales = True Then
            cmdAgregarFinal.Enabled = True
        Else
            cmdAgregarFinal.Enabled = False
        End If
        cmdModificarFinal.Enabled = False
        cmdEliminarFinal.Enabled = False
        adoFinales.RecordSource = "SELECT [Materias].[Curso], [Finales].[Materia] AS Codigo, [Materias].[Nombre], [Materias].[Detalle], [Personal].[Nombre] AS CursoCon, [Finales].[Ano], [Finales].[Division], [Finales].[Parcial1], [Finales].[Parcial2], [Finales].[Totalizador], [Finales].[Recuperatorio1], [Finales].[Recuperatorio2], [Finales].[Cursada], [Finales].[Asistencia], [Finales].[AsistenciaPorcentaje], [Finales].[Promocion], [Finales].[Nota], [Finales].[Fecha] , [Finales].[Libro], [Finales].[Folio], [Finales].[Mesa], [Finales].[Acta], [Finales].[Equivalencia], [Finales].[Establecimiento], [Instituciones].[Institucion], [Finales].[Aprobada], [Finales].[PerdioTurno], [Finales].[Habilitada], [Finales].[Profesor], [Finales].[CantidadMesas], [Finales].[PerdioCursada], [Finales].[Comentario] FROM ((Finales INNER JOIN Materias ON [Finales].[Materia]=[Materias].[Codigo]) INNER JOIN Personal ON [Finales].[Profesor]=[Personal].[Codigo]) " _
        & "  INNER JOIN Instituciones ON [Finales].[Establecimiento]=[Instituciones].[Codigo] Where ((([Finales].[Alumno]) = " & adoAlumnos.Recordset!Permiso & ") And (([Materias].[Carrera]) = " & dtcCarreras.BoundText & ") And (([Materias].[Eliminada]) = 0)) ORDER BY [Finales].[Materia]"
        adoFinales.Refresh
        Conexion.Open
        Set CantidadAprobadas = Conexion.Execute("SELECT Count(Materias.Detalle) As Cantidad FROM Finales INNER JOIN Materias ON Finales.Materia = Materias.Codigo WHERE (((Materias.Detalle)=1 Or (Materias.Detalle)=2) AND ((Finales.Alumno)=" & adoAlumnos.Recordset!Permiso & ") AND ((Materias.Carrera)=" & dtcCarreras.BoundText & ") AND Materias.Eliminada=0 AND ((Finales.Aprobada)=1))")
        lblTotalAprobadas = "Total de Materias Aprobadas: " & CantidadAprobadas!Cantidad
        TotalAprobadas = CantidadAprobadas!Cantidad
        'lblPorcentajeAprobado = "% " & Str(Format((TotalAprobadas / TotalPlan) * 100, "00"))
        Conexion.Close
    End If
End Sub

Private Sub dtgCursadas_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    If BanderaMostrarCursada = "Si" Then MostrarCursadas
End Sub

Private Sub Form_Load()
    Conexion.ConnectionString = ("DSN=Instituto")
    BanderaMostrarCursada = "Si"
End Sub



Private Sub txtNota_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then dtpFecha.SetFocus
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
Private Function MostrarCursadas()
    If adoFinales.Recordset.RecordCount > 0 Then
        If frmIdentificacion.Permisos!ModificarFinales = True Then
            cmdModificarFinal.Enabled = True
            cmdEliminarFinal.Enabled = True
        Else
            cmdModificarFinal.Enabled = False
            cmdEliminarFinal.Enabled = False
        End If
        
        txtParcial1 = Format(adoFinales.Recordset!Parcial1, "0.00")
        txtParcial2 = Format(adoFinales.Recordset!Parcial2, "0.00")
        txtRecuperatorio1 = Format(adoFinales.Recordset!Recuperatorio1, "0.00")
        txtRecuperatorio2 = Format(adoFinales.Recordset!Recuperatorio2, "0.00")
        If adoFinales.Recordset!Cursada = "Verdadero" Then
            chkCursadaAprobada.Value = 1
        Else
            chkCursadaAprobada.Value = 0
        End If
        If adoFinales.Recordset!Promocion = "Verdadero" Then
            chkPromociono.Value = 1
        Else
            chkPromociono.Value = 0
        End If
        txtNota = Format(adoFinales.Recordset!Nota, "0.00")
        If Not adoFinales.Recordset!Fecha Then
            dtpFecha.Value = adoFinales.Recordset!Fecha
        Else
       
        End If
        If IsNull(adoFinales.Recordset!Libro) Then
            txtLibro = 0
        Else
            txtLibro = adoFinales.Recordset!Libro
        End If
        txtFolio = adoFinales.Recordset!Folio
        txtAnoCursada = adoFinales.Recordset!Ano
        txtDivision = adoFinales.Recordset!division
        txtCantidadMesas = adoFinales.Recordset!CantidadMesas
        txtMesa = adoFinales.Recordset!Mesa
        txtActa = adoFinales.Recordset!Acta
        If adoFinales.Recordset!Aprobada = "Verdadero" Then
            chkAprobada.Value = 1
        Else
           chkAprobada.Value = 0
        End If
        If adoFinales.Recordset!Asistencia = "Verdadero" Then
            chkAsistencia.Value = 1
        Else
            chkAsistencia.Value = 0
        End If
        txtPorcentaje = adoFinales.Recordset!AsistenciaPorcentaje
        If adoFinales.Recordset!Equivalencia = "Verdadero" Then
            chkEquivalencia.Value = 1
        Else
            chkEquivalencia.Value = 0
        End If
        If adoFinales.Recordset!Establecimiento <> 0 Then
            txtEstablecimiento = adoFinales.Recordset!Institucion
        Else
            txtEstablecimiento = ""
        End If
        If adoFinales.Recordset!Habilitada = "Verdadero" Then
            chkHabilitada.Value = 1
        Else
            chkHabilitada.Value = 0
        End If
        If adoFinales.Recordset!PerdioCursada = True Then
            lblPerdioCursada.Visible = True
        Else
            lblPerdioCursada.Visible = False
        End If
        If adoFinales.Recordset!Comentario <> 0 Then
            txtComentario = adoFinales.Recordset!Comentario
        Else
            txtComentario = ""
        End If

        Conexion.Open
        Set CantidadAprobadas = Conexion.Execute("SELECT Count(Materias.Detalle) As Cantidad FROM Finales INNER JOIN Materias ON Finales.Materia = Materias.Codigo WHERE (((Materias.Detalle)=1 Or (Materias.Detalle)=2) AND ((Finales.Alumno)=" & adoAlumnos.Recordset!Permiso & ") AND ((Materias.Carrera)=" & dtcCarreras.BoundText & ") AND ((Finales.Aprobada)=1))")
        lblTotalAprobadas = "Total de Materias Aprobadas: " & CantidadAprobadas!Cantidad
        TotalAprobadas = CantidadAprobadas!Cantidad
        lblPorcentajeAprobado = "% " & Str(Format((TotalAprobadas / TotalPlan) * 100, "00"))
        Conexion.Close
    End If
End Function

Private Function LimpiarCursadas()
    txtParcial1 = ""
    txtParcial2 = ""
    txtParcial3 = ""
    txtRecuperatorio1 = ""
    txtRecuperatorio2 = ""
    txtRecuperatorio3 = ""
    chkCursadaAprobada.Value = 0
    chkPromociono.Value = 0
    txtNota = ""
    dtpFecha.Value = Date
    txtLibro = ""
    txtFolio = ""
    txtAnoCursada = 0
    txtDivision = ""
    txtCantidadMesas = ""
    txtMesa = ""
    txtActa = ""
    chkAprobada.Value = 0
    chkAsistencia.Value = 0
    txtPorcentaje = ""
    chkEquivalencia.Value = 0
    txtEstablecimiento = ""
    chkHabilitada.Value = 0
End Function

Private Function HabilitarCursadas()
    'txtParcial1
    'txtParcial2
    'txtParcial2
    'txtRecuperatorio1
    'txtRecuperatorio2
    'txtRecuperatorio3
    chkCursadaAprobada.Enabled = True
    chkPromociono.Enabled = True
    frAlumnos.Enabled = False
    frPlanCompleto.Enabled = False
    dtgCursadas.Enabled = False
    txtNota.Enabled = True
    dtpFecha.Enabled = True
    txtLibro.Enabled = True
    txtFolio.Enabled = True
    txtAnoCursada.Enabled = True
    txtDivision.Enabled = True
    txtCantidadMesas.Enabled = True
    txtMesa.Enabled = True
    txtActa.Enabled = True
    chkAprobada.Enabled = True
    'chkAsistencia.Enabled = True
    'txtPorcentaje.Enabled = True
    chkEquivalencia.Enabled = True
    'If chkEquivalencia.Value = 1 Then txtEstablecimiento.Enabled = True
    chkHabilitada.Enabled = True
End Function

Private Function DeshabilitarCursadas()
    'txtParcial1
    'txtParcial2
    'txtParcial2
    'txtRecuperatorio1
    'txtRecuperatorio2
    'txtRecuperatorio3
    chkCursadaAprobada.Enabled = False
    chkPromociono.Enabled = False
    frAlumnos.Enabled = True
    frPlanCompleto.Enabled = True
    dtgCursadas.Enabled = True
    txtNota.Enabled = False
    dtpFecha.Enabled = False
    txtLibro.Enabled = False
    txtFolio.Enabled = False
    txtAnoCursada.Enabled = False
    txtDivision.Enabled = False
    txtCantidadMesas.Enabled = False
    txtMesa.Enabled = False
    txtActa.Enabled = False
    chkAprobada.Enabled = False
    chkAsistencia.Enabled = False
    txtPorcentaje.Enabled = False
    chkEquivalencia.Enabled = False
    'txtEstablecimiento.Enabled = False
    chkHabilitada.Enabled = False
End Function

Private Function PasarCursadas()
    Alumno = txtPermiso
    AnoCursada = txtAnoCursada
    If Estado = "Agregando" Then
        Materia = adoPlanCompleto.Recordset!Codigo
        Conexion.Open
        Conexion.Execute ("INSERT INTO Finales(Alumno,Materia, Ano, Profesor, Establecimiento) VALUES (" & Alumno & "," & Materia & ", " & AnoCursada & ", 0, 0)")
        Conexion.Close
    Else
        Materia = adoFinales.Recordset!Codigo
        AnoCursada = adoFinales.Recordset!Ano
    End If
    Conexion.Open
    Conexion.Execute ("UPDATE Finales SET Parcial1 = " & Replace(Val(txtParcial1), ",", ".") & ",Parcial2 = " & Replace(Val(txtParcial2), ",", ".") & ", Recuperatorio1 = " & Replace(Val(txtRecuperatorio1), ",", ".") & ",Recuperatorio2 = " & Replace(Val(txtRecuperatorio2), ",", ".") & ",Cursada = " & chkCursadaAprobada.Value & ",Promocion = " & chkPromociono.Value & ",Ano = " & Val(txtAnoCursada) & ",Division = " & Val(txtDivision) & ",CantidadMesas = " & Val(txtCantidadMesas) & ",Aprobada = " & chkAprobada.Value & ",Asistencia =" & chkAsistencia.Value & ",AsistenciaPorcentaje=" & Replace(Val(txtPorcentaje), ",", ".") & ",Equivalencia = " & chkEquivalencia.Value & ",Habilitada = " & chkHabilitada.Value & ",Comentario = '" & txtComentario.Text & "' WHERE Alumno= " & Alumno & " AND Materia = " & Materia & " AND Ano =" & AnoCursada)
    If chkAprobada.Value = 0 Then
        Conexion.Execute ("UPDATE Finales SET Nota = 0,Fecha = Null,Libro=0, Folio=0,Mesa=0,Acta=0 WHERE Alumno= " & Alumno & " AND Materia = " & Materia & " AND Ano =" & AnoCursada)
    Else
        Conexion.Execute ("UPDATE Finales SET Nota = " & Replace(Val(txtNota), ",", ".") & ",Fecha = '" & DateValue(dtpFecha.Value) & "',Libro= " & Val(txtLibro) & ",Folio= " & Val(txtFolio) & ",Mesa= " & Val(txtMesa) & ",Acta= " & Val(txtActa) & " WHERE Alumno= " & Alumno & " AND Materia = " & Materia & " AND Ano =" & AnoCursada)
    End If
    Conexion.Close
    BanderaMostrarCursada = "No" 'para que no se ejecute MostrarCursadas cuando se llama al dtgCursadas_RowColChange
    adoFinales.Recordset.Requery
    adoFinales.Recordset.Find ("Codigo=" & RegistroActual)
    BanderaMostrarCursada = "Si"
    adoFinales.Refresh
    MostrarCursadas
End Function

Private Function VerExistencia()
If adoFinales.Recordset.RecordCount > 0 Then
    CursorEn = adoFinales.Recordset!Codigo
    adoFinales.Recordset.MoveFirst
    For i = 1 To adoFinales.Recordset.RecordCount
        If adoPlanCompleto.Recordset!Codigo = adoFinales.Recordset!Codigo Then
            MsgBox ("La materia ya está cargada")
            Esta = "Si"
            adoFinales.Recordset.MoveFirst
            adoFinales.Recordset.Find ("Codigo=" & CursorEn)
            Exit Function
        Else
            adoFinales.Recordset.MoveNext
        End If
    Next i
End If
End Function
