Attribute VB_Name = "Module1"
Dim Conectar As New Connection
Public cn As New Connection
Public Anio As Integer
Public ReporteFichaPermiso As Integer
Public ReporteFichaCarrera As Integer
Public numero_mesa As Long
Public numero_acta As Integer
Public codigo_materia As Long 'para imprimir la planila de parciales
Public division As Integer 'para imprimir la planilla de parciales



Public Saltar As Integer
Private Const WM_COMMAND As Long = &H111
Private Const CBN_SELCHANGE As Long = &H1
Private Const CB_SETCURSEL As Long = &H14E
Private Declare Function GetParent Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function GetDlgCtrlID Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Private Declare Function SendNotifyMessage Lib "user32" Alias "SendNotifyMessageA" (ByVal hWnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Const SPI_SCREENSAVERRUNNING = 97&
Private Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, lpvParam As Any, ByVal fuWinIni As Long) As Long

'Para usar archivos ini
Declare Function GetPrivateProfileInt Lib "kernel32" Alias "GetPrivateProfileIntA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal nDefault As Long, ByVal lpFileName As String) As Long
Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long

Public Enum enmZoom
    zoom200 = &H0
    zoom150 = &H1
    zoom100 = &H2
    zoom75 = &H3
    zoom50 = &H4
    zoom25 = &H5
    zoom10 = &H6
    zoomAjustar = &H7
End Enum
Sub Main()
    'chequear configuracion regional
    SeparadorDecimal = Format(0.1, "#. #")
    SeparadorDecimal = IIf(InStr(SeparadorDecimal, ","), ",", ".")
    If SeparadorDecimal = "," Then
       'MsgBox ("La configuracion regional no es la recomendada" & Chr(13) & "Debe configurar el punto para separador decimal y la coma para separador de miles") : Exit Sub
    End If
    
    'agrego nuevos campos por si no estan en la base de datos
    Conectar.ConnectionString = ("DSN=Instituto")
    cn.ConnectionString = ("DSN=Instituto")
    

    
    On Error Resume Next
        Conectar.Open
        Conectar.Execute ("ALTER TABLE Alumnos ADD BloquearAutogestion bit ")
        Conectar.Execute ("ALTER TABLE Finales ADD Comentario text(30)")
        Conectar.Execute ("ALTER TABLE Personal ADD Contrasena text(25)")
        Conectar.Execute ("ALTER TABLE Mesas ADD LimiteInscriptos Integer")
        Conectar.Execute ("ALTER TABLE Divisiones ADD LimiteMatriculados Integer")
        Conectar.Execute ("ALTER TABLE Parametros ADD PathLogoReporte text(100)")
        Conectar.Execute ("ALTER TABLE Parametros ADD Domicilio text(100)")
        Conectar.Execute ("ALTER TABLE Parametros ADD Telefono text(100)")
        Conectar.Execute ("ALTER TABLE Titulo ADD Numero text(10)")
        Conectar.Execute ("ALTER TABLE Titulo ADD Numero text(10)")
        Conectar.Execute ("ALTER TABLE rpt_Planilla_analitica ADD Division Integer")
        Conectar.Execute ("ALTER TABLE Finales ADD Libro Integer")
        Conectar.Execute ("ALTER TABLE Actas ADD Libro Integer")
        Conectar.Execute ("ALTER TABLE Equivalencias ADD Libro Integer")
        Conectar.Execute ("ALTER TABLE EquivalenciasResumen ADD Libro Integer")
        Conectar.Execute ("ALTER TABLE Titulo ADD LibroFinal Integer")
        Conectar.Execute ("ALTER TABLE Titulo ADD FolioFinal Integer")
        Conectar.Execute ("ALTER TABLE Personal ADD Usuario text(20)")
        Conectar.Execute ("ALTER TABLE Correlativas ADD PorFinal bit ")
        
        'Creo campo contrasena y paso los valores del campo contraseña
        Conectar.Execute ("ALTER TABLE Alumnos ADD Contrasena text(6)")
        'Conectar.Execute ("UPDATE Alumnos set Contrasena = Contraseña")

        
    frmIdentificacion.Show
       

     
End Sub
Public Sub DisableKeys(blnState As Boolean)
    Dim lngRet As Long
    Dim blnOld As Boolean
    If blnState = True Then
        lngRet = SystemParametersInfo(SPI_SCREENSAVERRUNNING, True, blnOld, 0&)
    Else
        lngRet = SystemParametersInfo(SPI_SCREENSAVERRUNNING, False, blnOld, 0&)
    End If
End Sub

Public Sub EstablecerZoom(ByVal hWnd As Long, ByVal Zoom As enmZoom)
    hWnd = FindWindowEx(hWnd, 0, "MSDataReportRunTimeWndClass6.0", vbNullString)
    If hWnd Then
        hWnd = FindWindowEx(hWnd, 0, "REPT_DISPLAYFRAME", vbNullString)
        If hWnd Then
            hWnd = FindWindowEx(hWnd, 0, "#32770", vbNullString)
            If hWnd Then
                hWnd = FindWindowEx(hWnd, 0, "ComboBox", vbNullString)
                If hWnd Then
                    SendMessage hWnd, CB_SETCURSEL, Zoom, vbNullString
                    SendNotifyMessage GetParent(hWnd), WM_COMMAND, (&H10000 * CBN_SELCHANGE) Or GetDlgCtrlID(hWnd), hWnd
               End If
           End If
       End If
   End If
End Sub



