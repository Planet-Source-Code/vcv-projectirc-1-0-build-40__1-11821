Attribute VB_Name = "modWindows"
Public Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Public Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long


Public Const SRCCOPY = &HCC0020 ' (DWORD) dest = source

Public Type POINTAPI
        X As Long
        Y As Long
End Type

Global Const ICON_SIZE = 16
Global bSBL As Boolean
Sub HideWin(intWhich As Integer)
    Dim i As Integer, cnt As Integer

    If intWhich = 1 Then
        If Status.Visible = False Then Status.Visible = True
        Status.Visible = False
        Exit Sub
    End If
    If intWhich = 2 Then
        If BuddyList.Visible = False Then BuddyList.Visible = True
        BuddyList.Visible = False
        Exit Sub
    End If
    
    cnt = 3
    For i = 1 To intChannels
        If Channels(i).strName <> "" Then
            If cnt = intWhich Then
                Channels(i).Visible = False
                Exit Sub
            End If
            cnt = cnt + 1
        End If
    Next i
    
    For i = 1 To intQueries
        On Error Resume Next
        If Queries(i).strNick <> "" Then
            If cnt = intWhich Then
                Queries(i).Visible = False
                Exit Sub
            End If
            cnt = cnt + 1
        End If
    Next i
    
    
    'WindowCount = cnt + 1 'add 1 for status window

End Sub

Public Sub StayOnTop(frmForm As Form, fOnTop As Boolean)
    
    Const HWND_TOPMOST = -1
    Const HWND_NOTOPMOST = -2
    
    Dim lState As Long
    Dim iLeft As Integer, iTop As Integer, iWidth As Integer, iHeight As Integer


    With frmForm
        iLeft = .Left / Screen.TwipsPerPixelX
        iTop = .Top / Screen.TwipsPerPixelY
        iWidth = .Width / Screen.TwipsPerPixelX
        iHeight = .Height / Screen.TwipsPerPixelY
    End With
    


    If fOnTop Then
        lState = HWND_TOPMOST
    Else
        lState = HWND_NOTOPMOST
    End If
    Call SetWindowPos(frmForm.hwnd, lState, iLeft, iTop, iWidth, iHeight, 0)
End Sub
Function GetWindowIndex(strCaption As String)
    Dim i As Integer
    For i = 1 To WindowCount
        If GetWindowTitle(i) = strCaption Then
            GetWindowIndex = i
            Exit Function
        End If
    Next i
    GetWindowIndex = -1
End Function

Function GetWindowTitle(intWhich As Integer) As String
    Dim i As Integer, cnt As Integer
    If intWhich = 1 Then GetWindowTitle = "Status": Exit Function
    If intWhich = 2 Then GetWindowTitle = "Friend Tracker": Exit Function
    
    cnt = 2
    For i = 1 To intChannels
        If Channels(i).strName <> "" Then cnt = cnt + 1
        If cnt = intWhich Then
            GetWindowTitle = Channels(i).strName
            Exit Function
        End If
    Next i
    
    For i = 1 To intQueries
        If Queries(i).strNick <> "" Then cnt = cnt + 1
        If cnt = intWhich Then
            On Error Resume Next
            GetWindowTitle = Queries(i).strNick
            Exit Function
        End If
    Next i
    
    GoTo final
    
    For i = 1 To intDCCChats
        If cnt = intWhich Then
            'GetWindowTitle = DCCChats(i).Caption
            Exit Function
        End If
        cnt = cnt + 1
    Next i
    
    For i = 1 To intDCCSends
        If cnt = intWhich Then
            'GetWindowTitle = DCCSends(i).Caption
            Exit Function
        End If
        cnt = cnt + 1
    Next i
final:
    GetWindowTitle = ""
End Function
Sub SetWinFocus(intWhich As Integer)
    Dim i As Integer, cnt As Integer

    If intWhich = 1 Then
        If Status.Visible = False Then Status.Visible = True
        Status.SetFocus
        Status.Visible = True
        Exit Sub
    End If
    If intWhich = 2 Then
        If BuddyList.Visible = False Then BuddyList.Visible = True
        BuddyList.SetFocus
        BuddyList.Visible = True
        Exit Sub
    End If
    
    cnt = 2
    For i = 1 To intChannels
        If Channels(i).strName <> "" Then
            cnt = cnt + 1
            If cnt = intWhich Then
                Channels(i).Visible = True
                Channels(i).SetFocus
                Exit Sub
            End If
        End If
    Next i
    
    For i = 1 To intQueries
        'On Error Resume Next
        If Queries(i).strNick <> "" Then
            cnt = cnt + 1
            If cnt = intWhich Then
                'On Error Resume Next
                'MsgBox Queries(i).strNick
                Queries(i).Visible = True
                Queries(i).DataOut.SelStart = 0
                Exit Sub
            End If
        End If
    Next i
    
    
    'WindowCount = cnt + 1 'add 1 for status window

End Sub

Function TaskCenter(intActual As Integer, strText As String) As Integer
'    MsgBox strText
    Dim intRet As Integer
    intRet = (((intActual - ICON_SIZE) - Client.picTask.TextWidth(strText)) / 2) + (ICON_SIZE / 2)
    If Right(strText, 1) = "." Then intRet = intRet + 8
    TaskCenter = intRet
End Function


Function TaskText(intWidth As Integer, strText As String) As String
    'MsgBox intWidth & ".." & Client.picTask.TextWidth(strText) & ".."
    Dim lastWidth As Integer, i As Integer, strBuf As String
    Dim inttemp As Integer
    
    For i = 1 To Len(strText)
        strBuf = Left(strText, i) & "..."
        inttemp = Client.picTask.TextWidth(strBuf) ' + 2 + ICON_SIZE
        
        If inttemp >= intWidth - 2 - ICON_SIZE Then
            TaskText = Left(strText, i - 1) & "..."
            Exit Function
        End If
    Next i
    TaskText = strText
        
End Function

Function WindowCount() As Integer
    Dim cnt As Integer, i As Integer
    
    For i = 1 To intChannels
        If Channels(i).strName <> "" Then cnt = cnt + 1
    Next i
    
    For i = 1 To intQueries
        On Error Resume Next
        If Queries(i).strNick <> "" Then cnt = cnt + 1
    Next i
    
    
    WindowCount = cnt + 2 'add 2 for status window and buddy list
    
    '* Add DCC and stuff here

End Function


Function WindowNewBuffer(intWhich As Integer) As String
    Dim i As Integer, cnt As Integer
    If intWhich = 1 Then
        If Status.newBuffer = True Then
            WindowNewBuffer = True
        Else
            WindowNewBuffer = False
        End If
        Exit Function
    End If
    If intWhich = 1 Then
        If BuddyList.newBuffer = True Then
            WindowNewBuffer = True
        Else
            WindowNewBuffer = False
        End If
        Exit Function
    End If
    
    cnt = 3
    For i = 1 To intChannels
        If cnt = intWhich Then
            If Channels(i).newBuffer = True Then
                WindowNewBuffer = True
                Exit Function
            Else
                WindowNewBuffer = False
                Exit Function
            End If
        End If
        cnt = cnt + 1
    Next i
    
    
    For i = 1 To intQueries
        If cnt = intWhich Then
            On Error Resume Next
            If Queries(i).newBuffer = True Then
                WindowNewBuffer = True
                Exit Function
            Else
                WindowNewBuffer = False
                Exit Function
            End If
        End If
        cnt = cnt + 1
    Next i
    
    GoTo final
    
    For i = 1 To intDCCChats
        If cnt = intWhich Then
'            If dccchats(i).newBuffer = True Then
                WindowNewBuffer = True
                Exit Function
'            Else
                WindowNewBuffer = False
                Exit Function
'            End If
        End If
        cnt = cnt + 1
    Next i
    
    For i = 1 To intDCCSends
        If cnt = intWhich Then
'            If dccsends(i).newBuffer = True Then
                WindowNewBuffer = True
                Exit Function
'            Else
                WindowNewBuffer = False
                Exit Function
'            End If
        End If
        cnt = cnt + 1
    Next i
final:
    WindowNewBuffer = False ' ""

End Function


