Attribute VB_Name = "Main"
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
Private Const WM_COMMAND As Long = &H111
Private Const CBN_SELCHANGE As Long = &H1
Private Const CB_SETCURSEL As Long = &H14E
Private Declare Function GetParent Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function GetDlgCtrlID Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Private Declare Function SendNotifyMessage Lib "user32" Alias "SendNotifyMessageA" (ByVal hWnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

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

