VERSION 5.00
Begin VB.MDIForm frmServicios 
   BackColor       =   &H8000000C&
   Caption         =   "I.S.F.D. y T. Nº 20 - Certificación de Servicios"
   ClientHeight    =   9045
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   11850
   Icon            =   "Servicios.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   Begin VB.Menu mnuCarreras 
      Caption         =   "Carreras"
      Begin VB.Menu mnuCarrerasPlanes 
         Caption         =   "Planes"
      End
   End
   Begin VB.Menu mnuPersonal 
      Caption         =   "Personal"
      Begin VB.Menu mnuPersonalPersonal 
         Caption         =   "Personal"
      End
      Begin VB.Menu mnuPersonalMovimientos 
         Caption         =   "Movimientos"
      End
   End
End
Attribute VB_Name = "frmServicios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub mnuCarrerasPlanes_Click()
    frmCarreras.Show
End Sub

Private Sub mnuPersonalMovimientos_Click()
    frmMovimientos.Show
End Sub

Private Sub mnuPersonalPersonal_Click()
    frmPersonal.Show
End Sub
