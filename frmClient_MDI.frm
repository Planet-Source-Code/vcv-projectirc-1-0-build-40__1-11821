VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Begin VB.MDIForm Client 
   Appearance      =   0  'Flat
   BackColor       =   &H8000000C&
   Caption         =   "projectIRC"
   ClientHeight    =   7065
   ClientLeft      =   4425
   ClientTop       =   1695
   ClientWidth     =   10965
   Icon            =   "frmClient_MDI.frx":0000
   LinkTopic       =   "MDIForm1"
   Begin VB.PictureBox picTask 
      Align           =   2  'Align Bottom
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   0
      ScaleHeight     =   22
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   731
      TabIndex        =   3
      Top             =   6735
      Width           =   10965
      Begin VB.PictureBox picTaskBuffer 
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   525
         ScaleHeight     =   23
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   732
         TabIndex        =   4
         Top             =   540
         Visible         =   0   'False
         Width           =   10980
      End
      Begin VB.Timer tmrTask 
         Enabled         =   0   'False
         Interval        =   1
         Left            =   1755
         Top             =   30
      End
      Begin VB.Shape shpTask 
         BorderColor     =   &H00808080&
         Height          =   315
         Left            =   0
         Top             =   15
         Visible         =   0   'False
         Width           =   30
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   3015
      Top             =   1035
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   7
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClient_MDI.frx":08CA
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClient_MDI.frx":307E
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClient_MDI.frx":5832
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClient_MDI.frx":7FE6
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClient_MDI.frx":8582
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClient_MDI.frx":8B1E
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClient_MDI.frx":90BA
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picToolMain 
      Align           =   1  'Align Top
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   0
      ScaleHeight     =   28
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   731
      TabIndex        =   0
      Top             =   0
      Width           =   10965
      Begin ComCtl3.CoolBar CoolBar1 
         Height          =   420
         Left            =   0
         TabIndex        =   1
         Top             =   0
         Width           =   10980
         _ExtentX        =   19368
         _ExtentY        =   741
         BandCount       =   1
         ForeColor       =   -2147483628
         FixedOrder      =   -1  'True
         _CBWidth        =   10980
         _CBHeight       =   420
         _Version        =   "6.0.8169"
         MinHeight1      =   24
         Width1          =   728
         UseCoolbarColors1=   0   'False
         NewRow1         =   0   'False
         BandStyle1      =   1
         AllowVertical1  =   0   'False
         Begin VB.PictureBox picIcon 
            AutoRedraw      =   -1  'True
            Height          =   390
            Left            =   4110
            ScaleHeight     =   330
            ScaleWidth      =   315
            TabIndex        =   5
            Top             =   0
            Visible         =   0   'False
            Width           =   375
         End
         Begin MSComctlLib.Toolbar Toolbar1 
            Height          =   330
            Left            =   75
            TabIndex        =   2
            Top             =   45
            Width           =   5310
            _ExtentX        =   9366
            _ExtentY        =   582
            ButtonWidth     =   609
            ButtonHeight    =   582
            Style           =   1
            ImageList       =   "ImageList1"
            _Version        =   393216
            BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
               NumButtons      =   5
               BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  ImageIndex      =   1
               EndProperty
               BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  ImageIndex      =   2
               EndProperty
               BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  ImageIndex      =   3
               EndProperty
               BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Style           =   3
               EndProperty
               BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Object.ToolTipText     =   "Click here to join a channel"
                  ImageIndex      =   4
               EndProperty
            EndProperty
         End
      End
   End
   Begin MSWinsockLib.Winsock IDENT 
      Left            =   1155
      Top             =   1365
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      RemotePort      =   113
      LocalPort       =   113
   End
   Begin MSWinsockLib.Winsock sock 
      Left            =   1575
      Top             =   1365
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Menu mnu_file 
      Caption         =   "&File"
      Begin VB.Menu mnu_File_Connect 
         Caption         =   "&Connect"
         Shortcut        =   {F2}
      End
      Begin VB.Menu mnu_File_Disconnect 
         Caption         =   "&Disconnect"
         Enabled         =   0   'False
         Shortcut        =   {F3}
      End
      Begin VB.Menu mnu_File_LB01 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_File_Options 
         Caption         =   "&Options"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnu_File_LB02 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_File_Quit 
         Caption         =   "&Quit"
         Shortcut        =   {F12}
      End
   End
   Begin VB.Menu mnu_View 
      Caption         =   "&View"
      Begin VB.Menu mnu_View_BuddyList 
         Caption         =   "&Buddy List"
      End
      Begin VB.Menu mnu_View_Debug 
         Caption         =   "&Debug Window"
      End
      Begin VB.Menu mnu_View_Status 
         Caption         =   "&Status Window"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnu_view_lb01 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_view_TBTop 
         Caption         =   "Taskbar on &Top"
      End
      Begin VB.Menu mnu_View_TBBot 
         Caption         =   "Taskbar on &Bottom"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnu_View_TBrTop 
         Caption         =   "Toolbar on Top"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnu_View_TBrBot 
         Caption         =   "Toolbar on Bottom"
      End
   End
   Begin VB.Menu mnu_Window 
      Caption         =   "&Window"
      WindowList      =   -1  'True
      Begin VB.Menu mnu_Window_Cascade 
         Caption         =   "&Cascade"
      End
      Begin VB.Menu mnu_Window_TileH 
         Caption         =   "&Tile Horizontally"
      End
      Begin VB.Menu mnu_Tile_Vertically 
         Caption         =   "&Tile Vertically"
      End
   End
   Begin VB.Menu mnu_Help 
      Caption         =   "&Help"
      Begin VB.Menu mnu_Help_About 
         Caption         =   "&About"
      End
   End
   Begin VB.Menu mnu_nicks 
      Caption         =   "nicks"
      Visible         =   0   'False
      Begin VB.Menu mnu_nicks_op 
         Caption         =   "&Op"
      End
      Begin VB.Menu mnu_nicks_halfop 
         Caption         =   "&HalfOp"
      End
      Begin VB.Menu mnu_nicks_voice 
         Caption         =   "&Voice"
      End
   End
End
Attribute VB_Name = "Client"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const clrSep = &H80000010

Public intActive As Integer, sngLastX As Integer
Public intHover As Integer, bReDraw As Boolean, intLast As Integer

Const buttonShadow = &H80000010
Const buttonHilight = &H80000014

Dim Modes()     As typMode
Public intModes As Integer

Public strDataBuffer As String, minSize As Long

Sub ClearModes()
    For i = 1 To intModes
        Modes(i).bPos = False
        Modes(i).mode = ""
    Next i
    intModes = 0
End Sub


Public Function ModeString() As String
    If intModes = 0 Then Exit Function
    Dim strFinal As String, bWhich As Boolean, i As Integer
    If Modes(1).bPos = True Then bWhich = True
    
    If bWhich Then strFinal = strFinal & "+" Else strFinal = strFinal & "-"
    
    For i = 1 To intModes
        If Modes(i).bPos <> bWhich Then
            bWhich = Not bWhich
            If bWhich Then strFinal = strFinal & "+" Else strFinal = strFinal & "-"
        End If
        strFinal = strFinal & Modes(i).mode
    Next i
    ModeString = strFinal
End Function

Public Sub AddMode(strMode As String, bPlus As Boolean)
    Dim i As Integer
    For i = 1 To intModes
        If Modes(i).mode = strMode Then Exit Sub
    Next i
    
    intModes = intModes + 1
    ReDim Preserve Modes(1 To intModes) As typMode
    
    With Modes(intModes)
        .bPos = True
        .mode = strMode
    End With
    Status.Update
End Sub

Public Sub RemoveMode(strMode As String)
    Dim i As Integer, j As Integer
    
    For i = 1 To intModes
        If Modes(i).mode = strMode Then
            Modes(i).mode = ""
            For j = i To intModes - 1
                Modes(j) = Modes(j + 1)
            Next j
            intModes = intModes - 1
            ReDim Preserve Modes(1 To intModes) As typMode
            Status.Update
            Exit Sub
        End If
    Next i
    Status.Update
End Sub

Sub CTCPReply(strNick As String, strReply As String)
    Client.SendData "NOTICE " & strNick & " :" & strAction & strReply & strAction
End Sub

Sub DrawToolbar()
    Dim intSeps As Integer, CenX As Integer, j As Integer, realWidth As Long
    Dim strTitle As String, intWidth As Integer, i As Integer, intBegin As Integer
    Dim intEnd As Integer, strDrawText As String, lngRet As Long, iconY As Integer, iconX As Integer
    bReDraw = True
    
    intSeps = WindowCount - 1
    
    If bStretchButtons Then
        realWidth = Client.ScaleWidth / 15
        minSize = (intSeps + 1) * (ICON_SIZE + 40)
        If realWidth < minSize Then realWidth = minSize Else minSize = realWidth
    Else
        realWidth = (intSeps + 1) * intButtonWidth
        minSize = realWidth
    End If
    
    picTaskBuffer.Width = realWidth
    
    intWidth = realWidth / (intSeps + 1)
    picTaskBuffer.Cls
    
    picTaskBuffer.CurrentY = 2
    '* If one window open, draw only single thing..
    If intSeps = 0 Then
        strTitle = GetWindowTitle(1)
        strDrawText = TaskText(intWidth, strTitle)
        CenX = (realWidth - picTaskBuffer.TextWidth(strDrawText)) / 2
        picTaskBuffer.CurrentX = CenX
        picTaskBuffer.ForeColor = vbRed
        picTaskBuffer.Print strDrawText;
        Exit Sub
    End If
    
    picTaskBuffer.ForeColor = clrSep
    For i = 1 To intSeps + 5
        picTaskBuffer.CurrentX = intWidth * i
        For j = 3 To picTaskBuffer.ScaleHeight - 4 Step 2
            picTaskBuffer.PSet (picTaskBuffer.CurrentX, j)
        Next j
    Next i
    
    picTaskBuffer.ForeColor = vbBlack
    For i = 1 To intSeps + 1
        picTaskBuffer.CurrentY = 2
        picTaskBuffer.CurrentX = intWidth * (i - 1)
        intBegin = picTaskBuffer.CurrentX + 1
        intEnd = intBegin + intWidth - 2
        
        strTitle = GetWindowTitle(Int(i))
        'MsgBox strTitle & "~" & i
        If strTitle Like "Status" Then
            picIcon.Picture = ImageList1.ListImages.Item(6).Picture
        ElseIf strTitle Like "Friend Tracker" Then
            picIcon.Picture = ImageList1.ListImages.Item(7).Picture
        ElseIf strTitle Like "[#]*" Then
            picIcon.Picture = ImageList1.ListImages.Item(4).Picture
        Else
            picIcon.Picture = ImageList1.ListImages.Item(5).Picture
        End If
        picIcon.Picture = picIcon.Image
        
        intBegin = intBegin + 1
        intEnd = intEnd - 1
        
        If intActive = i Then iconY = 4 Else iconY = 3
        If intActive = i Then iconX = intBegin + 3 Else iconX = intBegin + 2
        
        lngRet = BitBlt(picTaskBuffer.hDC, iconX, iconY, ICON_SIZE, ICON_SIZE, picIcon.hDC, 0, 0, SRCCOPY)
        
        If intActive = i Then
            picTaskBuffer.ForeColor = vbRed
            '* Active window, let's draw a inset bevel
            picTaskBuffer.Line (intBegin, 2)-(intEnd, 2), buttonShadow
            picTaskBuffer.Line (intBegin, 2)-(intBegin, picTaskBuffer.ScaleHeight - 2), buttonShadow
            picTaskBuffer.Line (intEnd, 2)-(intEnd, picTaskBuffer.ScaleHeight - 2), buttonHilight
            picTaskBuffer.Line (intBegin, picTaskBuffer.ScaleHeight - 2)-(intEnd, picTaskBuffer.ScaleHeight - 2), buttonHilight
            picTaskBuffer.CurrentY = 3
            If intHover = i Then picTask.Tag = intHover
        ElseIf intHover = i Then
            picTask.Tag = intHover
            picTaskBuffer.ForeColor = vbBlack
            picTaskBuffer.Line (intBegin, 2)-(intEnd, 2), buttonHilight
            picTaskBuffer.Line (intBegin, 2)-(intBegin, picTaskBuffer.ScaleHeight - 2), buttonHilight
            picTaskBuffer.Line (intEnd, 2)-(intEnd, picTaskBuffer.ScaleHeight - 2), buttonShadow
            picTaskBuffer.Line (intBegin, picTaskBuffer.ScaleHeight - 2)-(intEnd, picTaskBuffer.ScaleHeight - 2), buttonShadow
            picTaskBuffer.CurrentY = 2
        ElseIf WindowNewBuffer(i) Then
            Dim clr As Long
            clr = vbRed
            picTaskBuffer.Line (intBegin, 2)-(intEnd, 2), clr
            picTaskBuffer.Line (intBegin, 2)-(intBegin, picTaskBuffer.ScaleHeight - 2), clr
            picTaskBuffer.Line (intEnd, 2)-(intEnd, picTaskBuffer.ScaleHeight - 2), clr
            picTaskBuffer.Line (intBegin, picTaskBuffer.ScaleHeight - 2)-(intEnd, picTaskBuffer.ScaleHeight - 2), clr
            picTaskBuffer.ForeColor = vbBlack
            picTaskBuffer.CurrentY = 2
        Else
            If picTaskBuffer.ForeColor <> vbBlack Then picTaskBuffer.ForeColor = vbBlack
            picTaskBuffer.CurrentY = 2
        End If
        
        picTaskBuffer.CurrentX = intWidth * (i - 1)
        intBegin = picTaskBuffer.CurrentX
        intEnd = intBegin + intWidth
        strDrawText = TaskText(intWidth, strTitle)
'        MsgBox strDrawText
        picTaskBuffer.CurrentX = TaskCenter(intWidth, strDrawText) + picTaskBuffer.CurrentX
        
        picTaskBuffer.CurrentY = picTaskBuffer.CurrentY + 2
        If intActive = i Then
            picTaskBuffer.CurrentX = picTaskBuffer.CurrentX + 1
        End If
        
        picTaskBuffer.Print strDrawText;
    Next i
    picTask.Picture = picTaskBuffer.Image
    'picTask.Refresh
        
End Sub

Public Sub HandleCTCP(strNick As String, strData As String)
    strData = RightR(strData, 1)
    strData = LeftR(strData, 1)
    
    Dim strCom As String, strParam As String, inttemp As Integer
    Seperate strData, " ", strCom, strParam
    
    Select Case LCase(strCom)
        Case "version"
            PutData Status.DataIn, strColor & "05" & strBold & strNick & strBold & " has just requested your client version"
            CTCPReply strNick, "VERSION projectIRC for Windows"
        Case "ping"
            PutData Status.DataIn, strColor & "05" & strBold & strNick & strBold & " has just pinged you"
            CTCPReply strNick, "PING 0 seconds, cause projectIRC is elite"
        Case "action"
            inttemp = GetQueryIndex(strNick)
            If inttemp = -1 Then Exit Sub
            PutText Queries(inttemp).DataIn, strColor & "06" & strNick & " " & strParam
    End Select
End Sub


Sub interpret(strData As String)
    Dim parsed As ParsedData, inttemp As Integer
    Dim i As Integer, strChan As String, strTemp As String
    
    strData = Replace(strData, Chr(10), "")
    ParseData strData, parsed
    DoEvents
    If Len(parsed.strCommand) = 0 Then Exit Sub
    
    '* I will put these in alphanumeric order one day, dont worry
    
    Select Case LCase(parsed.strCommand)
        Case "001"
            strMyNick = params(parsed, 1, 1)
            PutData Status.DataIn, "* " & params(parsed, 2, -1)
            Status.Update
            If bInvisible Then Client.SendData "MODE " & strMyNick & " +i"
            Exit Sub
        Case "002", "003", "004", "005"
            PutData Status.DataIn, "* " & params(parsed, 2, -1)
            Exit Sub
        Case "ping"
            SendData "PONG :" & params(parsed, 1, 1)
            'PutData Status.DataIn, strColor & "03Ping? Pong! [" & params(parsed, 1, 1) & "]"
            Exit Sub
        Case "join"
            If LCase(parsed.strNick) = LCase(strMyNick) Then
                inttemp = NewChannel(params(parsed, 1, -1))
                Client.SendData "MODE " & params(parsed, 1, -1)
            Else
                inttemp = GetChanIndex(params(parsed, 1, 1))
                If inttemp = -1 Then Exit Sub
                Channels(inttemp).AddNick parsed.strNick
                PutData Channels(inttemp).DataIn, strColor & "03" & strBold & parsed.strNick & strBold & " has joined " & strBold & Channels(inttemp).strName
            End If
            Exit Sub
        Case "privmsg"
            strChan = params(parsed, 1, 1)
            If Left(strChan, 1) = "#" Or Left(strChan, 1) = "&" Then  'privmsg to channel
                inttemp = GetChanIndex(strChan)
                If inttemp <> -1 Then Channels(inttemp).PutText parsed.strNick, params(parsed, 2, -1)               '
            ElseIf parsed.strNick = strMyNick Then
                If params(parsed, 2, 2) = strAction & "VERSION" & strAction Then    'version
                    'Client.SendData "CTCP REPLY " & strChan & " VERSION :jIRC for Windows9x"
                    Client.SendData "NOTICE " & parsed.strNick & " :VERSION projectIRC for Win32"
                End If
                GoTo msg
            
            Else    'send to query window
msg:
                strTemp = params(parsed, 2, -1)
                If Left(strTemp, 1) = strAction Then
                    HandleCTCP parsed.strNick, strTemp
                    Exit Sub
                End If
                
                'msgbox parsed.strNick
                
                If QueryExists(parsed.strNick) Then
                    'MsgBox "exists"
                    inttemp = GetQueryIndex(parsed.strNick)
                    If inttemp = -1 Then Exit Sub
                    
                    If Queries(inttemp).strHost <> parsed.strFullHost Then
                        Queries(inttemp).strHost = RightOf(parsed.strFullHost, "!")
                        Queries(inttemp).lblHost = RightOf(parsed.strFullHost, "!")
                        
                    End If
                    Queries(inttemp).Caption = parsed.strNick
                    Queries(inttemp).strNick = parsed.strNick
                    Queries(inttemp).lblNick = parsed.strNick
                    Queries(inttemp).PutText parsed.strNick, strTemp
                Else
                    'MsgBox "doesnt"
                    NewQuery parsed.strNick, parsed.strFullHost
                    inttemp = GetQueryIndex(parsed.strNick)
                    If inttemp = -1 Then Exit Sub
                    Queries(inttemp).Caption = parsed.strNick
                    Queries(inttemp).strNick = parsed.strNick
                    Queries(inttemp).lblNick = parsed.strNick
                    Queries(inttemp).PutText parsed.strNick, strTemp
                End If
            End If
            Exit Sub
        Case "nick"
            If parsed.strNick = strMyNick Then
                strMyNick = params(parsed, 1, 1)
                PutData Status.DataIn, strColor & "03Your nick is now " & strBold & strMyNick
                ChangeNick parsed.strNick, params(parsed, 1, -1)
                Status.Update
            Else
                ChangeNick parsed.strNick, params(parsed, 1, 1)
            End If
            Exit Sub
        Case "part"
            If parsed.strNick = strMyNick Then Exit Sub
            inttemp = GetChanIndex(parsed.strParams(1))
            If inttemp = -1 Then Exit Sub
            Channels(inttemp).RemoveNick parsed.strNick
            PutData Channels(inttemp).DataIn, strColor & "03" & strBold & parsed.strNick & strBold & " has left " & strBold & Channels(inttemp).strName
            If parsed.strNick = strMyNick Then Unload Channels(inttemp)
            Exit Sub
        Case "353" 'nick list!
            'MsgBox parsed.strParams(3)
            inttemp = GetChanIndex(parsed.strParams(3))
            If inttemp = -1 Then Exit Sub
            Dim strNicks() As String
            strNicks = Split(params(parsed, 4, -1), " ")
            For i = LBound(strNicks) To UBound(strNicks)
                Channels(inttemp).AddNick strNicks(i)
            Next i
            Exit Sub
        Case "mode"     'set mode
            inttemp = GetChanIndex(params(parsed, 1, 1))
            If inttemp = -1 And params(parsed, 1, 1) <> strMyNick Then Exit Sub
            strTemp = parsed.strNick
            If strTemp = "" Then strTemp = parsed.strFullHost
            
            If inttemp <> -1 Then PutData Channels(inttemp).DataIn, strColor & "03" & strBold & strTemp & strBold & " sets mode: " & params(parsed, 2, -1)
            If params(parsed, 1, 1) = strMyNick Then PutData Status.DataIn, strColor & "03" & strBold & strTemp & strBold & " sets mode: " & params(parsed, 2, -1)
            
            ParseMode params(parsed, 1, 1), params(parsed, 2, -1)
            Exit Sub
        Case "quit"     'quit
            NickQuit parsed.strNick, params(parsed, 1, -1)
            Exit Sub
        Case "kick"     'kick
            inttemp = GetChanIndex(params(parsed, 1, 1))
            If inttemp = -1 Then Exit Sub
            PutData Channels(inttemp).DataIn, strColor & "03" & strBold & params(parsed, 2, 2) & strBold & " was kicked from " & strBold & params(parsed, 1, 1) & strBold & " by " & strBold & parsed.strNick & strBold & " [ " & params(parsed, 3, -1) & " ]"
            Channels(inttemp).RemoveNick params(parsed, 2, 2)
            
            '* If user, close channel
            If params(parsed, 2, 2) = strMyNick Then
                Channels(inttemp).Tag = "NOPART"
                Unload Channels(inttemp)
                PutData Status.DataIn, strColor & "03" & "You were kicked from " & strBold & params(parsed, 1, 1) & strBold & " by " & strBold & parsed.strNick & strBold & " [ " & params(parsed, 3, -1) & " ]"
            End If
            
            Exit Sub
        Case "332"  'topic!
            inttemp = GetChanIndex(params(parsed, 2, 2))
            If inttemp = -1 Then Exit Sub
            Channels(inttemp).rtbTopic.Text = ""
            PutData Channels(inttemp).rtbTopic, params(parsed, 3, -1)
            Channels(inttemp).rtbTopic.SelStart = 0
            Channels(inttemp).rtbTopic.SelLength = 1
            Channels(inttemp).rtbTopic.SelText = ""
            PutData Channels(inttemp).DataIn, strColor & "03Topic is """ & strColor & params(parsed, 3, -1) & strColor & "03"""
            Channels(inttemp).rtbTopic.SelStart = 0
            Channels(inttemp).rtbTopic.Tag = "locked"
            Exit Sub
        Case "topic"    'change in topic!
            inttemp = GetChanIndex(params(parsed, 1, 1))
            If inttemp = -1 Then Exit Sub
            Channels(inttemp).rtbTopic.Text = ""
            PutData Channels(inttemp).rtbTopic, params(parsed, 2, -1)
            Channels(inttemp).rtbTopic.SelStart = 0
            Channels(inttemp).rtbTopic.SelLength = 1
            Channels(inttemp).rtbTopic.SelText = ""
            PutData Channels(inttemp).DataIn, strColor & "03Topic changed by " & strBold & parsed.strNick & strBold & " : " & params(parsed, 2, -1)
            Exit Sub
        Case "333"  'topic on param2 set by param3, on param4
            inttemp = GetChanIndex(params(parsed, 2, 2))
            If inttemp = -1 Then Exit Sub
            PutData Channels(inttemp).DataIn, strColor & "03Topic set by " & strBold & params(parsed, 3, 3) & strBold
            '* gotta add when it was SET!!
            Exit Sub
        Case "366"  'end of names list
            Exit Sub
        Case "324"  'set channel modes
            ParseMode params(parsed, 2, 2), params(parsed, 3, -1)
            Exit Sub
        Case "notice"
            'On Error Resume Next
            If parsed.strNick = "" Then
                PutData Status.DataIn, strColor & "05" & params(parsed, 2, -1)
            Else
                If Client.ActiveForm.hwnd = BuddyList.hwnd Then
                    PutData Status.DataIn, strColor & "05" & strBold & "NOTICE" & strBold & strColor & " " & strBold & parsed.strNick & strBold & ":" & Chr(9) & params(parsed, 2, -1)
                Else
                    PutData Client.ActiveForm.DataIn, strColor & "05" & strBold & "NOTICE" & strBold & strColor & " " & strBold & parsed.strNick & strBold & ":" & Chr(9) & params(parsed, 2, -1)
                End If
            End If
            
            Exit Sub
        Case "433"  'nick name already in use
            If params(parsed, 2, 2) = strMyNick Then
                strTemp = strMyNick
                strMyNick = strOtherNick
                strOtherNick = strMyNick
                Client.SendData "NICK " & strMyNick
            End If
            PutData Status.DataIn, strColor & "04* " & params(parsed, 2, -1)
            Exit Sub
        Case "372"  'MOTD
            PutData Status.DataIn, params(parsed, 2, -1)
            Exit Sub
        Case "375"  'start of MOTD
            PutData Status.DataIn, strColor & "02" & params(parsed, 2, -1)
            Exit Sub
        Case "376"  'end of MOTD
            Exit Sub
        Case "251", "252", "253", "254", "255", _
            "265", "266" 'server info, users, ops, channels, clients
            PutData Status.DataIn, strColor & "06" & params(parsed, 2, -1)
            Exit Sub
        Case "303"  'users on!
            BuddyList.AddUsers params(parsed, 2, -1)
            Exit Sub
        Case "329"  'date created for channel, $1 = channel, $2 = when
            '*  ack i gotta fix THIS!
            Exit Sub
        Case "484", "482", "461", "412", "403", "421", "401", "481", "402" 'error msgs
        
            PutData Status.DataIn, strColor & "04* " & params(parsed, 2, -1)
            Exit Sub
        Case "321"  'blah channel list
            'ChannelsList.lvChannels.ListItems.Clear
            ChannelsList.Visible = True
            Exit Sub
        Case "322"          'channel list! ACK!!
            ChannelsList.AddChannel params(parsed, 2, 2), params(parsed, 3, 3), params(parsed, 4, -1)
            ChannelsList.Caption = "Channels List : " & ChannelsList.lvChannels.ListItems.Count & " channels"
            Exit Sub
        Case "323"  'end of channe list!
            'ChannelsList.bNeedClear = True
            'ChannelsList.lvChannels.ListItems.Clear
            Exit Sub
        Case "328"
            inttemp = GetChanIndex(params(parsed, 2, 2))
            If inttemp = -1 Then Exit Sub
            PutData Channels(inttemp).DataIn, strColor & "03Channel URL" & strColor & strUnderline & ":" & strUnderline & " " & strColor & "12" & params(parsed, 3, -1) & strColor
            Exit Sub
        Case "301"      'WHOIS! away
            strTemp = params(parsed, 2, 2)
            If bWhoisInQuery Then
                inttemp = GetQueryIndex(strTemp)
                If inttemp = -1 Then
                    PutData Status.DataIn, strColor & "05" & strBold & strTemp & strBold & " is away: " & params(parsed, 3, -1)
                Else
                    PutData Queries(inttemp).DataIn, strColor & "05" & strBold & strTemp & strBold & " is away: " & params(parsed, 3, -1)
                End If
            Else
                PutData Status.DataIn, strColor & "05" & strBold & strTemp & strBold & " is away: " & params(parsed, 3, -1)
            End If
            Exit Sub
        Case "307"
            strTemp = params(parsed, 2, 2)
            If bWhoisInQuery Then
                inttemp = GetQueryIndex(strTemp)
                If inttemp = -1 Then
                    PutData Status.DataIn, strColor & "05" & strBold & strTemp & strBold & " " & params(parsed, 3, -1)
                Else
                    PutData Queries(inttemp).DataIn, strColor & "05" & strBold & strTemp & strBold & " " & params(parsed, 3, -1)
                End If
            Else
                PutData Status.DataIn, strColor & "05" & strBold & strTemp & strBold & " " & params(parsed, 3, -1)
            End If
            Exit Sub
        Case "311"
            strTemp = params(parsed, 2, 2)
            If bWhoisInQuery Then
                inttemp = GetQueryIndex(strTemp)
                If inttemp = -1 Then
                    PutData Status.DataIn, " "
                    PutData Status.DataIn, strColor & "05" & strBold & strTemp & strBold & " is " & params(parsed, 3, 3) & "@" & params(parsed, 4, -1)
                Else
                    PutData Queries(inttemp).DataIn, " "
                    PutData Queries(inttemp).DataIn, strColor & "05" & strBold & strTemp & strBold & " is " & params(parsed, 3, 3) & "@" & params(parsed, 4, -1)
                End If
            Else
                PutData Status.DataIn, " "
                PutData Status.DataIn, strColor & "05" & strBold & strTemp & strBold & " is " & params(parsed, 3, 3) & "@" & params(parsed, 4, -1)
            End If
            Exit Sub
        Case "312"
            strTemp = params(parsed, 2, 2)
            If bWhoisInQuery Then
                inttemp = GetQueryIndex(strTemp)
                If inttemp = -1 Then
                    PutData Status.DataIn, strColor & "05" & strBold & strTemp & strBold & " is using: " & params(parsed, 3, 3) & " ( " & params(parsed, 4, -1) & " )"
                Else
                    PutData Queries(inttemp).DataIn, strColor & "05" & strBold & strTemp & strBold & " is using: " & params(parsed, 3, 3) & " ( " & params(parsed, 4, -1) & " )"
                End If
            Else
                PutData Status.DataIn, strColor & "05" & strBold & strTemp & strBold & " is using: " & params(parsed, 3, 3) & " ( " & params(parsed, 4, -1) & " )"
            End If
            Exit Sub
        Case "313"
            strTemp = params(parsed, 2, 2)
            If bWhoisInQuery Then
                inttemp = GetQueryIndex(strTemp)
                If inttemp = -1 Then
                    PutData Status.DataIn, strColor & "05" & strBold & strTemp & strBold & " " & params(parsed, 3, -1)
                Else
                    PutData Queries(inttemp).DataIn, strColor & "05" & strBold & strTemp & strBold & " " & params(parsed, 3, -1)
                End If
            Else
                PutData Status.DataIn, strColor & "05" & strBold & strTemp & strBold & " " & params(parsed, 3, -1)
            End If
            Exit Sub
        Case "317"
            strTemp = params(parsed, 2, 2)
            If bWhoisInQuery Then
                inttemp = GetQueryIndex(strTemp)
                If inttemp = -1 Then
                    PutData Status.DataIn, strColor & "05" & strBold & strTemp & strBold & " " & params(parsed, 3, 3) & " seconds idle / " & params(parsed, 4, 4) & " sign-on time"
                Else
                    PutData Queries(inttemp).DataIn, strColor & "05" & strBold & strTemp & strBold & " " & params(parsed, 3, 3) & " seconds idle / " & params(parsed, 4, 4) & " sign-on time"
                End If
            Else
                PutData Status.DataIn, strColor & "05" & strBold & strTemp & strBold & " " & params(parsed, 3, 3) & " seconds idle / " & params(parsed, 4, 4) & " sign-on time"
            End If
            Exit Sub
        Case "318"
            Exit Sub
        Case "319"
            strTemp = params(parsed, 2, 2)
            If bWhoisInQuery Then
                inttemp = GetQueryIndex(strTemp)
                If inttemp = -1 Then
                    PutData Status.DataIn, strColor & "05" & strBold & strTemp & strBold & " is on: " & params(parsed, 3, -1)
                Else
                    PutData Queries(inttemp).DataIn, strColor & "05" & strBold & strTemp & strBold & " is on: " & params(parsed, 3, -1)
                End If
            Else
                PutData Status.DataIn, strColor & "05" & strBold & strTemp & strBold & " is on: " & params(parsed, 3, -1)
            End If
            Exit Sub
        Case "320"
            strTemp = params(parsed, 2, 2)
            Seperate strTemp, " ", strChan, strTemp
            If bWhoisInQuery Then
                inttemp = GetQueryIndex(strChan)
                'MsgBox bWhoisInQuery & "~" & inttemp & "~" & strTemp & "~"
                If inttemp = -1 Then
                    PutData Status.DataIn, strColor & "05" & strBold & strChan & strBold & " " & strTemp
                Else
                    PutData Queries(inttemp).DataIn, strColor & "05" & strBold & strChan & strBold & " " & strTemp
                End If
            Else
                PutData Status.DataIn, strColor & "05" & strBold & strChan & strBold & " " & strTemp
            End If
            Exit Sub
    End Select
    PutData Status.DataIn, "*** " & strBold & parsed.strCommand & strBold & " " & params(parsed, 1, -1) ' & " [" & parsed.strFullHost & "]"
End Sub


Sub SendData(strData As String)
    On Error Resume Next
    If sock.State = 0 Then Exit Sub
    sock.SendData strData & Chr(10)
    If DebugWin.Visible Then
        DebugWin.txtDataIn = DebugWin.txtDataIn & "<< " & strData & vbCrLf
        DebugWin.txtDataIn.SelStart = Len(DebugWin.txtDataIn)
    End If
    
End Sub


Private Sub Command1_Click()
    Clipboard.SetText Status.DataIn.SelRTF
End Sub

Private Sub IDENT_ConnectionRequest(ByVal requestID As Long)
    IDENT.Close
    IDENT.Accept requestID
    IDENT.SendData IDENT.LocalPort & ", " & IDENT.RemotePort & " : USERID : UNIX : " & strMyIdent & vbCrLf
    
    Dim i As Integer
    For i = 1 To 500
        DoEvents
    Next i
    
    IDENT.Close
End Sub

Private Sub IDENT_DataArrival(ByVal bytesTotal As Long)
    Dim dat As String
    IDENT.GetData dat, vbString
    
    If dat Like "*, *" Then
        dat = LeftR(dat, 2)
        PutData Status.DataIn, "*** IDENT : " & dat
        dat = dat & " : USERID : UNIX : " & strMyIdent
        IDENT.SendData dat
        PutData Status.DataIn, "*** IDENT reply : " & dat
        'MsgBox "~" & dat & "~"
        Dim i As Integer
        IDENT.Close
    End If
End Sub

Private Sub IDENT_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    PutData Status.DataIn, Chr(Color) & "04IDENT Error " & strColor & Description
End Sub

Private Sub MDIForm_Activate()
    DrawToolbar
End Sub

Private Sub MDIForm_Load()
        
    intButtonWidth = 100
    sngLastX = 1
    DoEvents
    Me.Visible = True
    
    Load LoadSettings
    LoadSettings.Visible = True
    DoEvents
    StayOnTop LoadSettings, True
        
    mnu_Help.Caption = Chr$(8) & mnu_Help.Caption
    DrawToolbar
    
    '* INI stuff
    path = App.path
    If Right(App.path, 1) <> "\" Then path = path & "\"
    INI = path & "settings.ini"
    '/if doesnt exist, create
    If Not FileExists(path & "settings.ini") Then
        Open INI For Output As #1
            Print #1, ""
        Close #1
    End If
    
    '/* Connection settings
    strServer = ReadINI("connect", "server", "irc.dal.net")
    strMyNick = ReadINI("connect", "nick", "pIRCu")
    strOtherNick = ReadINI("connect", "altnick", "OtherNick")
    strFullName = ReadINI("connect", "fullname", "projectIRC user")
    strMyIdent = ReadINI("connect", "ident", "projectIRC")
    lngPort = CLng(ReadINI("connect", "port", "6667"))
    bConOnLoad = CBool(ReadINI("connect", "connonload", "true"))
    bReconnect = CBool(ReadINI("connect", "reconnect", "true"))
    bInvisible = CBool(ReadINI("connect", "invisible", "true"))
    bRetry = CBool(ReadINI("connect", "retry", "true"))
    intRetry = CInt(ReadINI("connect", "retrynum", "99"))
    
    '/* Display settings
    lngBackColor = CLng(ReadINI("display", "backcolor", CStr(RGB(255, 255, 255))))
    lngForeColor = CLng(ReadINI("display", "forecolor", CStr(RGB(0, 0, 0))))
    lngLeftColor = CLng(ReadINI("display", "leftcolor", CStr(&H800000)))
    lngRightColor = CLng(ReadINI("display", "rightcolor", CStr(&H8000000F)))
    
    '* Window/Interface Settings
    bStretchButtons = CBool(ReadINI("windows", "stretch", "true"))
    
    '* IRC Settings
    bWhoisInQuery = CBool(ReadINI("irc", "whoisquery", "false"))
    MAX_TEXT_HISTORY = CInt(ReadINI("irc", "maxtexthistory", "30"))
    
    TimeOut 0.1
    Unload LoadSettings
    
    
End Sub

 
Private Sub MDIForm_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    intHover = 0
End Sub

Private Sub MDIForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Cancel = 1
    Unload BuddyList
    Unload Status
    Unload DebugWin
    Client.SendData "QUIT :using projectIRC, closed"
    
    Dim i As Integer
    
    For i = 1 To intChannels
        If Channels(i).strName <> "" Then Unload Channels(i)
    Next
    For i = 1 To intQueries
        If Queries(i).strNick <> "" Then Unload Queries(i)
    Next i
    
    For i = 0 To 15
        DoEvents
    Next i
    Unload Me
    'Cancel = 0
    End
End Sub

Private Sub MDIForm_Resize()
    CoolBar1.Width = (Me.ScaleWidth / 15) + 1
    shpTask.Width = (Me.Width / 15) - 8
    SetWinFocus intActive
    DrawToolbar
    
    If bSBL = True Then
        DoEvents
        If BuddyList.Visible = True Then BuddyList.Move Client.Width - BuddyList.Width - 130, 0
    End If
    
End Sub

Private Sub MDIForm_Terminate()
'    Client.SendData "QUIT :Client closed, using projectIRC"
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
'    Client.SendData "QUIT :Client closed, using projectIRC"
End Sub

Sub mnu_File_Connect_Click()
    Select Case mnu_File_Connect.Caption
        Case "&Connect"
            '* Connect
            sock.Close
            mnu_File_Connect.Caption = "&Cancel"
            sock.RemoteHost = strServer
            sock.RemotePort = lngPort
            IDENT.Close
            On Error Resume Next
            IDENT.Listen
            sock.Connect
            PutData Status.DataIn, strColor & "02Connecting to " & strBold & strServer & strBold & " port " & strBold & lngPort
        Case "&Cancel"
            '* Cancel
            IDENT.Close
            sock.Close
            mnu_File_Connect.Caption = "&Connect"
            PutData Status.DataIn, strColor & "05Connection attempt cancelled"
    End Select
End Sub


Sub mnu_File_Disconnect_Click()
    mnu_File_Connect.Enabled = True
    mnu_File_Connect.Caption = "&Connect"
    mnu_File_Disconnect.Enabled = False
    PutData Status.DataIn, strColor & "05Disconnected from " & sock.RemoteHost
    sock.Close
    ClearModes
    Status.lblServer = "not connected"
    Status.Update
    BuddyList.lstNicks.Clear
End Sub


Private Sub mnu_File_Options_Click()
    Options.Show 1
End Sub

Private Sub mnu_File_Quit_Click()
    Client.SendData "QUIT :Using projectIRC, closed"
    IDENT.Close
    sock.Close
    Dim i As Integer
    For i = 1 To 1000
        DoEvents
    Next i
    Unload Me
End Sub

Private Sub mnu_Help_About_Click()
    About.Show vbModal
End Sub

Sub mnu_nicks_halfop_Click()
    Dim strChan As String, strNick As String
    strChan = Client.ActiveForm.strName
    strNick = RealNick(Client.ActiveForm.lstNicks.List(Client.ActiveForm.lstNicks.ListIndex))
    
    If mnu_nicks_halfop.Caption = "&HalfOp" Then
        Client.SendData "MODE " & strChan & " +h " & strNick
    Else
        Client.SendData "MODE " & strChan & " -h " & strNick
    End If
        
End Sub

 Sub mnu_nicks_op_Click()
    Dim strChan As String, strNick As String
    strChan = Client.ActiveForm.strName
    strNick = RealNick(Client.ActiveForm.lstNicks.List(Client.ActiveForm.lstNicks.ListIndex))
    
    If mnu_nicks_op.Caption = "&Op" Then
        Client.SendData "MODE " & strChan & " +o " & strNick
    Else
        Client.SendData "MODE " & strChan & " -o " & strNick
    End If
End Sub

Sub mnu_nicks_voice_Click()
    Dim strChan As String, strNick As String
    strChan = Client.ActiveForm.strName
    strNick = RealNick(Client.ActiveForm.lstNicks.List(Client.ActiveForm.lstNicks.ListIndex))
    
    If mnu_nicks_halfop.Caption = "&Voice" Then
        Client.SendData "MODE " & strChan & " +v " & strNick
    Else
        Client.SendData "MODE " & strChan & " -v " & strNick
    End If
End Sub

Private Sub mnu_Tile_Vertically_Click()
    Client.Arrange vbTileVertical
End Sub

Private Sub mnu_View_BuddyList_Click()
    With mnu_View_BuddyList
        .Checked = Not .Checked
        BuddyList.Visible = .Checked
    End With
End Sub

Private Sub mnu_View_Debug_Click()
    With mnu_View_Debug
        .Checked = Not .Checked
        DebugWin.Visible = .Checked
    End With

End Sub


Private Sub mnu_View_Status_Click()
    mnu_View_Status.Checked = Not mnu_View_Status.Checked
    Status.Visible = mnu_View_Status.Checked
End Sub

Private Sub mnu_viewTBTop_Click()
    mnu_view_TBTop.Checked = True
End Sub

Private Sub mnu_View_TBBot_Click()
    mnu_view_TBTop.Checked = False
    mnu_View_TBBot.Checked = True
    picTask.Align = 2
    If picToolMain.Align = 2 Then
        picToolMain.Align = 1
        picToolMain.Align = 2
    End If
End Sub

Private Sub mnu_View_TBrBot_Click()
    mnu_View_TBrTop.Checked = False
    mnu_View_TBrBot.Checked = True
    
    picToolMain.Align = 2
    If picTask.Align = 2 Then
        picTask.Align = 1
        picTask.Align = 2
    End If
End Sub

Private Sub mnu_View_TBrTop_Click()
    mnu_View_TBrTop.Checked = True
    mnu_View_TBrBot.Checked = False
    picToolMain.Align = 1
    If picTask.Align = 1 Then
        picTask.Align = 2
        picTask.Align = 1
    End If
End Sub


Private Sub mnu_view_TBTop_Click()
    mnu_view_TBTop.Checked = True
    mnu_View_TBBot.Checked = False
    picTask.Align = 1
    'If picToolMain.Align = 1 Then
    '    picToolMain.Align = 2
    '    picToolMain.Align = 1
    'End If
End Sub


Private Sub mnu_Window_Cascade_Click()
    Client.Arrange vbCascade
End Sub

Private Sub mnu_Window_TileH_Click()
    Client.Arrange vbTileHorizontal
End Sub


Private Sub picTask_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If x > minSize Then Exit Sub
    Dim which As Integer, wincnt As Integer, wid As Integer
    wincnt = WindowCount
    
    wid = minSize \ wincnt
    which = Int((x \ wid) + 0.5) + 1
    sngLastX = x
    
    If intActive = which Then
        intActive = 0
        HideWin which
        DoEvents
        picTask.Tag = ""
        intHover = which
        DrawToolbar
        Exit Sub
    End If
    '* Which now contains which button was clicked
    intActive = which
    'DrawToolbar
    SetWinFocus which
    DrawToolbar
End Sub


Private Sub picTask_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If x > minSize Then
        If intHover > 0 Then
            intHover = 0
            Client.DrawToolbar
            picTask.Tag = 0
        End If
        Exit Sub
    End If
    
    Dim which As Integer, wincnt As Integer, wid As Integer
    wincnt = WindowCount
    
    wid = minSize \ wincnt
    which = Int((x \ wid) + 0.5) + 1
    'sngLastX = X
    intHover = which
    tmrTask.Enabled = True
    If Val(picTask.Tag) = intHover Then Exit Sub
    Client.DrawToolbar
    
        
    'bReDraw = False
End Sub


Private Sub sock_Close()
    PutData Status.DataIn, strColor & "02Disconnected by SERVER from " & strServer
    sock.Close
    ClearModes
    IDENT.Close
    mnu_File_Connect.Enabled = True
    mnu_File_Connect.Caption = "&Connect"
    mnu_File_Disconnect.Enabled = False
    Status.lblServer = "not connected"
    Status.Update
    BuddyList.lstNicks.Clear

End Sub

Private Sub sock_Connect()
    '* Let's close all open windows
    Dim i As Integer
    For i = 1 To intChannels
        Channels(i).Tag = "NOPART"
        Unload Channels(i)
    Next i
    intChannels = 0
    
    For i = 1 To intQueries
        Unload Queries(i)
    Next i
    intQueries = 0
    mnu_File_Connect.Enabled = False
    mnu_File_Disconnect.Enabled = True
    mnu_File_Connect.Caption = "&Connect"
    PutData Status.DataIn, strColor & "03Connected to " & strServer
    
    SendData "PASS password"
    TimeOut 0.01
    SendData "NICK " & strMyNick
    TimeOut 0.01
    SendData "USER mynick local irc :" & strFullName & vbCrLf
    TimeOut 0.01
    
    Status.lblServer = sock.RemoteHost
    Status.Update
    
    '* Buddy List
    If mnu_View_BuddyList.Checked Then
        Dim strGet As String
        For i = 1 To BuddyList.lstSetup.ListCount
            strGet = strGet & BuddyList.lstSetup.List(i - 1) & " "
        Next i
        TimeOut 5
        If Trim(strGet) <> "" Then Client.SendData "ISON " & strGet

    End If
    
    
End Sub

Private Sub sock_DataArrival(ByVal bytesTotal As Long)
    Dim dat As String, AllParams As String
    Dim strData() As String, i As Integer
    
    On Error Resume Next
    If sock.State <> 0 Then sock.GetData dat, vbString
    If DebugWin.Visible Then
        If Len(DebugWin.txtDataIn.Text) + Len(dat) > 60000 Then DebugWin.txtDataIn.Text = ""
        On Error Resume Next
        DebugWin.txtDataIn = DebugWin.txtDataIn & ">> " & dat & vbCrLf
        DebugWin.txtDataIn.SelStart = Len(DebugWin.txtDataIn)
    End If
    
    strDataBuffer = strDataBuffer & dat
    
    If Right(strDataBuffer, 1) <> Chr(13) And Right(strDataBuffer, 1) <> Chr(10) Then
        Exit Sub
    Else
        dat = strDataBuffer
        strDataBuffer = ""
    End If
    
    strData = Split(dat, Chr(10))
    
    For i = LBound(strData) To UBound(strData)
        interpret strData(i)
    Next i
    
End Sub


Private Sub sock_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    PutData Status.DataIn, strColor & "04ERROR : " & Description

End Sub


Private Sub tmrTask_Timer()
    Dim pt As POINTAPI, lngRet As Long, hwnd As Long
    lngRet = GetCursorPos(pt)
    
    hwnd = WindowFromPoint(pt.x, pt.y)
    If hwnd <> picTask.hwnd Then
        intHover = -1
        Client.DrawToolbar
        picTask.Tag = -1
        tmrTask.Enabled = False
    End If
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
        Case 1
            mnu_File_Connect_Click
        Case 2
            mnu_File_Disconnect_Click
        Case 3
            mnu_File_Options_Click
        Case 5
            Dim strInput As String
            strInput = InputBox("Enter channel(s) to join, seperated each by a comma (EX: ""#irc"" or ""#irc, #mIRC"").  Also if you don't prefix # or &, I will for you (EX: ""irc,help"")", "Join channels")
            strInput = GetAlias("", "JOIN " & strInput)
    End Select
End Sub


