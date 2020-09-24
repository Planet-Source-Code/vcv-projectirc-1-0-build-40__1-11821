Attribute VB_Name = "modIRC"
'* projectIRC version 1.0
'* By Matt C, sappy@adelphia.net
'* Feel free to EMail me with any questions you may have.

'* Module that handles many of the Client's procedures,
'* Feel free to use in your project as long as you give me proper credit

'/me is using projectIRC $version $+ , on $server port $port me = $me on channel $chan
Global path As String
'* Channels and Queries
Global Const MAX_CHANNELS = 30
Global Const MAX_QUERIES = 30
Global Const MAX_DCCCHATS = 30
Global Const MAX_DCCSENDS = 30
Public Channels(1 To MAX_CHANNELS)  As Channel
Public Queries(1 To MAX_QUERIES)    As Query
'Public DCCChats(1 To MAX_DCCCHATS)  As DCCChat
'Public DCCSends(1 To MAX_DCCSENDS) As dccsend
Public intChannels  As Integer
Public intQueries   As Integer
Public intDCCChats  As Integer
Public intDCCSends  As Integer

'* Variables for incoming commands
Type ParsedData
    bHasPrefix   As Boolean
    strParams()  As String
    intParams    As Integer
    strFullHost  As String
    strCommand   As String
    strNick      As String
    strIdent     As String
    strHost      As String
    AllParams    As String
End Type

'* ANSI Formatting character values
Global Const BOLD = 2
Global Const UNDERLINE = 31
Global Const Color = 3
Global Const REVERSE = 22
Global Const ACTION = 1

'* ANSI Formatting characters
Global strBold As String
Global strUnderline As String
Global strColor As String
Global strReverse As String
Global strAction As String

'Nick storage for nick list inchannels
Type Nick
    Nick    As String
    op      As Boolean
    voice   As Boolean
    helper  As Boolean
    host    As String
    IDENT   As String
End Type

'Mode storage for each channel
Type typMode
    mode    As String
    bPos    As Boolean
End Type


Private Const EM_CHARFROMPOS& = &HD7
Private Type POINTAPI
    X As Long
    Y As Long
End Type

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Sub PutText(RTF As RichTextBox, strData As String)
    Dim i As Long, length As Integer, strChar As String
    Dim strBuffer As String, j As Long
    strData = " " & strData
    length = Len(strData)
    i = 1
    RTF.SelStart = Len(RTF.Text)
    RTF.SelColor = lngForeColor
    RTF.SelBold = False
    RTF.SelUnderline = False
    RTF.SelStrikeThru = False
    
    Do
        strChar = Mid(strData, i, 1)
        Select Case strChar
            Case strBold
                RTF.SelStart = Len(RTF.Text)
                RTF.SelText = strBuffer
                strBuffer = ""
                RTF.SelBold = Not RTF.SelBold
                i = i + 1
            Case strUnderline
                RTF.SelStart = Len(RTF.Text)
                RTF.SelText = strBuffer
                strBuffer = ""
                RTF.SelUnderline = Not RTF.SelUnderline
                i = i + 1
            Case strReverse
                RTF.SelStart = Len(RTF.Text)
                RTF.SelText = strBuffer
                strBuffer = ""
                RTF.SelStrikeThru = Not RTF.SelStrikeThru
                i = i + 1
            Case strColor
                RTF.SelStart = Len(RTF.Text)
                RTF.SelText = strBuffer
                strBuffer = ""
                i = i + 1
                Do Until Not ValidColorCode(strBuffer) Or i > length
                    strBuffer = strBuffer & Mid(strData, i, 1)
                    i = i + 1
                Loop
                strBuffer = LeftR(strBuffer, 1)
                RTF.SelStart = Len(RTF.Text)
                If strBuffer = "" Then
                    RTF.SelColor = vbBlack
                Else
                    RTF.SelColor = AnsiColor(LeftOf(strBuffer, ","))
                End If
                i = i - 1
                strBuffer = ""
            Case Else
                strBuffer = strBuffer & strChar
                i = i + 1
        End Select
    Loop Until i > length
    If strBuffer <> "" Then
            RTF.SelStart = Len(RTF.Text)
            RTF.SelText = strBuffer
            strBuffer = ""
    End If
    RTF.SelBold = False
    RTF.SelUnderline = False
    RTF.SelStrikeThru = False
    RTF.SelStart = Len(RTF.Text)
    RTF.SelText = vbCrLf
End Sub

Public Function RichWordOver(rch As RichTextBox, X As Single, Y As Single) As String
    Dim pt As POINTAPI
    Dim pos As Long
    Dim start_pos As Integer
    Dim end_pos As Integer
    Dim ch As String
    Dim txt As String
    Dim txtlen As Integer
    Dim i As Integer, j As Integer

    ' Convert the position to pixels.
    pt.X = X \ Screen.TwipsPerPixelX
    pt.Y = Y \ Screen.TwipsPerPixelY

    ' Get the character number
    On Error Resume Next
    pos = SendMessage(rch.hwnd, EM_CHARFROMPOS, 0&, pt)
    'Client.Caption = "~" & pos & "~" & Len(rch.Text)
    
    'Exit Function
    If pos <= 0 Or pos >= Len(rch.Text) Then
        RichWordOver = ""
        Exit Function
    End If
    
    txt = ""
    For i = pos To 1 Step -1
        ch = Mid(rch.Text, i, 1)
        If ch = " " Or _
           ch = "," Or _
           ch = "(" Or _
           ch = ")" Or _
           ch = "]" Or _
           ch = "[" Or _
           ch = "{" Or _
           ch = """" Or _
           ch = "'" Or _
           ch = Chr(9) Or _
           ch = "}" Then
            start_pos = i
            GoTo haha
        End If
    Next i
haha:
    txt = ""
    For i = pos To Len(rch.Text)
        ch = Mid(rch.Text, i, 1)
        If ch = " " Or _
           ch = "," Or _
           ch = "(" Or _
           ch = ")" Or _
           ch = "]" Or _
           ch = "[" Or _
           ch = "{" Or _
           ch = "}" Or _
           ch = """" Or _
           ch = "'" Or _
           ch = Chr(9) Then
            end_pos = i
            Exit For
        End If
    Next i
    
    If end_pos > Len(rch.Text) Or end_pos <= 0 Then end_pos = Len(rch.Text)
    
    RichWordOver = RightR(Replace(Mid(rch.Text, start_pos, end_pos - start_pos), Chr(13), ""), 1)
End Function



Function AnsiColor(intColNum As Integer) As Long
    Select Case intColNum
        Case 0: AnsiColor = RGB(255, 255, 255)
        Case 1: AnsiColor = RGB(0, 0, 0)
        Case 2: AnsiColor = RGB(0, 0, 127)
        Case 3: AnsiColor = RGB(0, 127, 0)
        Case 4: AnsiColor = RGB(255, 0, 0)
        Case 5: AnsiColor = RGB(127, 0, 0)
        Case 6: AnsiColor = RGB(127, 0, 127)
        Case 7: AnsiColor = RGB(255, 127, 0)
        Case 8: AnsiColor = RGB(255, 255, 0)
        Case 9: AnsiColor = RGB(0, 255, 0)
        Case 10: AnsiColor = RGB(0, 0, 0)
        Case 11: AnsiColor = RGB(0, 255, 255)
        Case 12: AnsiColor = RGB(0, 0, 255)
        Case 13: AnsiColor = RGB(255, 0, 255)
        Case 14: AnsiColor = RGB(92, 92, 92)
        Case 15: AnsiColor = RGB(184, 184, 184)
        Case Else: AnsiColor = RGB(0, 0, 0)
    End Select
End Function



Sub ChangeNick(strOldNick As String, strNewNick As String)
    Dim i As Integer, bChangedQuery As Boolean, inttemp As Integer
    
    For i = 1 To intChannels
    
        
        If Channels(i).InChannel(strOldNick) Then
            
            'change in queries :)
            If Not bChangedQuery Then
                inttemp = GetQueryIndex(strOldNick)
                If inttemp <> -1 Then
                    Queries(inttemp).lblNick = strNewNick
                    Queries(inttemp).strNick = strNewNick
                    Queries(inttemp).Caption = strNewNick
                    bChangedQuery = True
                End If
            End If
            
            'change in channel :)
            If Channels(i).strName <> "" Then
                Channels(i).ChangeNck strOldNick, strNewNick
            End If
        End If
    Next i
End Sub

Function Combine(arrItems() As String, intStart As Integer, intEnd As Integer) As String
    '* This returns the given parameter range as one string
    '* -1 specified as strEnd means from intStart to the last parameter
    '* Ex: strParams(1) is "#mIRC", then params 2 thru 16 are the nicks
    '*     of the users in the channel, simply do something like this
    '*     strNames = Params(2, -1)
    '*     that would return all the nicks into one string, similar to mIRC's
    '*      $#- identifier ($2-) in this case
    
    '* Declare variables
    Dim strFinal As String, intLast As Integer, i As Integer
    
    '* check for bad parameters
    If intStart < 1 Or intEnd > UBound(arrItems) + 1 Then Exit Function
    
    '* if intEnd = -1 then set it to param count
    If intEnd = -1 Then intLast = UBound(arrItems) + 1 Else intLast = intEnd
        
    For i = intStart To intLast
        strFinal = strFinal & arrItems(i - 1)
        If i <> intLast Then strFinal = strFinal & " "
    Next i
    
    Combine = Replace(strFinal, vbCr, "")
    strFinal = ""
End Function

Function DisplayNick(nckNick As Nick) As String
    Dim strPre As String
    If nckNick.voice Then strPre = "+"
    If nckNick.helper Then strPre = "%"
    If nckNick.op Then strPre = "@"
    DisplayNick = strPre & nckNick.Nick
End Function

Sub DoMode(strChannel As String, bAdd As Boolean, strMode As String, strParam As String)
    
    If strChannel = strMyNick Then
        If bAdd Then
            Client.AddMode strMode, bAdd
        Else
            Client.RemoveMode strMode
        End If
        Exit Sub
    End If
    
    Dim intX As Integer, i As Integer
    intX = GetChanIndex(strChannel)
    If intX = -1 Then Exit Sub
    
    Select Case strMode
        Case "v"
            Channels(intX).SetVoice strParam, bAdd
            
        Case "o"
            Channels(intX).SetOp strParam, bAdd
            If bAdd Then
                If strParam = strMyNick Then Channels(intX).rtbTopic.Tag = ""
            Else
                If strParam = strMyNick Then Channels(intX).rtbTopic.Tag = "locked"
            End If
        Case "h"
            Channels(intX).SetHelper strParam, bAdd
        Case "b"
        Case "k"
            If bAdd = True Then
                Channels(intX).strKey = strParam
                Channels(intX).AddMode strMode, bAdd
            Else
                Channels(intX).strKey = ""
                Channels(intX).RemoveMode strMode
            End If
        Case "l"
            If bAdd = True Then
                Channels(intX).intLimit = CInt(strParam)
                Channels(intX).AddMode strMode, bAdd
            Else
                Channels(intX).intLimit = 0
                Channels(intX).RemoveMode strMode
            End If
        Case Else
            If bAdd = True Then
                Channels(intX).AddMode strMode, bAdd
            Else
                Channels(intX).RemoveMode strMode
            End If
    End Select
End Sub

Function GetAlias(strChan As String, strData As String) As String
    Dim arrParams() As String, i As Integer, strP As String, strCom As String
    Dim strFinal As String, strAdd As String, bSpace As Boolean, inttemp As Integer
    Dim strTemp As String, strNck As String
    
    Seperate strData, " ", strCom, strData
    arrParams = Split(strData, " ")
    bSpace = True
    DoEvents
    
    For i = LBound(arrParams) To UBound(arrParams)
        strP = arrParams(i)
        'MsgBox strP & ":" & i
        strAdd = ""
        If strP = "$+" Then
            strFinal = LeftR(strFinal, 1)
            bSpace = False
        ElseIf Left(strP, 1) = "$" Then
            strAdd = GetVar(strChan, RightR(strP, 1))
        Else
            strAdd = strP
        End If
        
        strFinal = strFinal & strAdd
        If bSpace Then
            strFinal = strFinal & " "
        Else
            bSpace = True
        End If
    Next i
    
    DoEvents
    
    If Len(strFinal) > 0 Then strFinal = LeftR(strFinal, 1)
    
    ReDim arrParams(1) As String
    arrParams = Split(strFinal, " ")
    
    Dim r As String 'return
    Select Case LCase(strCom)
        Case "query"
            strTemp = Combine(arrParams, 2, -1)
            strNck = Combine(arrParams, 1, 1)
            'MsgBox strNck & "~"
            If QueryExists(strNck) Then
                inttemp = GetQueryIndex(strNck)
                Queries(inttemp).PutText strMyNick, strTemp
                r = "PRIVMSG " & strNck & " :" & strTemp
            Else
                inttemp = NewQuery(strNck, "")
                If UBound(arrParams) > 0 Then
                    Queries(inttemp).PutText strMyNick, strTemp
                    r = "PRIVMSG " & strNck & " :" & strTemp
                End If
            End If
                
        Case "msg"
            r = "PRIVMSG " & Combine(arrParams, 1, 1) & " :" & Combine(arrParams, 2, -1)
        Case "me"
            strTemp = Combine(arrParams, 1, -1)
            r = "PRIVMSG " & strChan & " :" & strAction & "ACTION " & strTemp & strAction
            If Left(strChan, 1) = "#" Then
                inttemp = GetChanIndex(strChan)
                If inttemp = -1 Then Exit Function
                PutData Channels(inttemp).DataIn, strColor & "06" & strMyNick & " " & strTemp
            Else
                inttemp = GetQueryIndex(strChan)
                If inttemp = -1 Then Exit Function
                PutData Queries(inttemp).DataIn, strColor & "06" & strMyNick & " " & strTemp
            End If
        Case "quit"
            r = "QUIT :" & Combine(arrParams, 1, -1)
        Case "notice"
            r = "NOTICE " & Combine(arrParams, 1, 1) & " :" & Combine(arrParams, 2, -1)
        Case "raw"
            r = Combine(arrParams, 1, -1)
        Case "nick"
            If Client.sock.State = 0 Then
                strMyNick = Combine(arrParams, 1, 1)
            Else
                r = "NICK " & Combine(arrParams, 1, 1)
            End If
        Case "id"   'identify with nickserv
            r = "PRIVMSG NickServ :IDENTIFY " & Combine(arrParams, 1, 1)
        Case "part"
            strTemp = Combine(arrParams, 1, -1)
            If UBound(arrParams) = 0 Then
                r = "PART " & strChan
                strTemp = strTemp
            Else
                r = "PART " & Combine(arrParams, 1, 1) & " :" & Combine(arrParams, 2, -1)
                strTemp = LeftOf(strTemp, " ")
            End If
            
            inttemp = GetChanIndex(strTemp)
            
            
            On Error Resume Next
            Channels(inttemp).Tag = "PARTNOW"
            
'            MsgBox r
        Case "server"
            strServer = Combine(arrParams, 1, 1)
            If UBound(arrParams) > 0 Then
                strport = Int(Combine(arrParams, 2, 2))
            End If
            Call Client.mnu_File_Disconnect_Click
            TimeOut 0.1
            Call Client.mnu_File_Connect_Click
        Case "join"
            'ok here's how it is, you can type /join #blah (key), #blah2, #blah3, #blah4 (key),
            'so we need special equiptment to handle this
            Dim strChans() As String
            strChans = Split(strData, ",")
            
            For inttemp = LBound(strChans) To UBound(strChans)
                Dim prefix As String
                prefix = ""
                If Left(strChans(inttemp), 1) <> "#" And Left(strChans(inttemp), 1) <> "&" Then prefix = "#"
                Client.SendData "JOIN " & prefix & strChans(inttemp)
                TimeOut 0.8
            Next inttemp
        Case "connect"
            If Combine(arrParams, 1, 1) <> "" Then
                strServer = Combine(arrParams, 1, 1)
            End If
            If UBound(arrParams) > 0 Then
                strport = Int(Combine(arrParams, 2, 2))
            End If
            Call Client.mnu_File_Disconnect_Click
            Call Client.mnu_File_Connect_Click
        Case "disconnect"
            Call Client.mnu_File_Disconnect_Click
        Case "bl"
            If BuddyList.Visible Then
                Unload BuddyList
            Else
                Load BuddyList
            End If
        Case "kill"
            r = "KILL " & Combine(arrParams, 1, 1) & " :" & Combine(arrParams, 2, -1)
        Case "es"
            PutText Status.DataIn, Combine(arrParams, 1, -1)
        Case "list"
            ChannelsList.lvChannels.ListItems.Clear
            If Combine(arrParams, 2, 2) = "" Then
                r = "LIST >0"
            Else
                r = "LIST " & Combine(arrParams, 2, -1)
            End If
        Case Else
            r = strCom & " " & Combine(arrParams, 1, -1)
    End Select
    
    GetAlias = r
    
End Function
Sub TimeOut(duration)
    StartTime = Timer


    Do While Timer - StartTime < duration
        X = DoEvents()
    Loop
End Sub

Function GetChanIndex(strName As String) As Integer
    Dim i As Integer

    For i = 1 To intChannels
        If LCase(Channels(i).strName) = Replace(LCase(strName), Chr(13), "") Then
            GetChanIndex = i
            Exit Function
        End If
    Next i
    GetChanIndex = -1
End Function

Function GetQueryIndex(strNick As String) As Integer
    Dim i As Integer

    For i = 1 To intQueries
        If LCase(Queries(i).strNick) = LCase(strNick) Then
            GetQueryIndex = i
            Exit Function
        End If
    Next i
    GetQueryIndex = -1
End Function

Function GetVar(strChan As String, strName As String)
    Dim r As String     'r is the return value
    Dim inttemp As String
    
    On Error Resume Next
    Select Case LCase(strName)
        Case "version"
            r = App.Major & "." & App.Minor & App.Revision
        Case "chan", "channel", "ch"
            r = strChan
        Case "me"
            r = strMyNick
        Case "server"
            r = Client.sock.RemoteHost
        Case "port"
            r = Client.sock.RemotePort
        Case "randnick"
            inttemp = GetChanIndex(strChan)
            If Left(strChan, "1") = "#" Then
                With Channels(inttemp)
                    Randomize
                    r = .GetNick(Int(Rnd * .intNicks) + 1)
                End With
            End If
        Case "date"
            r = Date
        Case "time"
            r = Time
    End Select
    GetVar = r
End Function

Function LeftOf(strData As String, strDelim As String) As String
    Dim intPos As Integer
    
    intPos = InStr(strData, strDelim)
    If intPos Then
        LeftOf = Left(strData, intPos - 1)
    Else
        LeftOf = strData
    End If
End Function

Function LeftR(strData As String, intMin As Integer)
    
    On Error Resume Next
    LeftR = Left(strData, Len(strData) - intMin)
End Function

Function NewChannel(strName As String) As Integer
    Dim i As Integer

    For i = 1 To intChannels
        If Channels(i).strName = "" Then
            Channels(i).Caption = strName
            Channels(i).lblName = strName
            Channels(i).strName = strName
            Channels(i).Visible = True
            Channels(i).Tag = i
            NewChannel = i
            Exit Function
        End If
    Next i
    intChannels = intChannels + 1
    Set Channels(intChannels) = New Channel
    Channels(intChannels).strName = strName
    Channels(intChannels).lblName = strName
    Channels(intChannels).Caption = strName
    Channels(intChannels).Visible = True
    Channels(intChannels).Tag = intChannels
    NewChannel = intChannels
End Function
Function NewQuery(strNick As String, strHost As String) As Integer
    Dim i As Integer, strHostX As String
    strHostX = RightOf(strHost, "!")
    
    i = GetQueryIndex(strNick)
    If i <> -1 Then
        Queries(i).SetFocus
        Exit Function
    End If

    For i = 1 To intQueries
        If Queries(i).strNick = "" Then
            Queries(i).Caption = strNick
            Queries(i).lblNick = strNick
            Queries(i).strNick = strNick
            Queries(i).strHost = strHostX
            Queries(i).lblHost = strHostX
            Queries(i).Visible = True
            Queries(i).Tag = i
            NewQuery = i
            Exit Function
        End If
    Next i
    
    intQueries = intQueries + 1
    Set Queries(intQueries) = New Query
    Queries(intQueries).strNick = strNick
    Queries(intQueries).lblNick = strNick
    Queries(intQueries).Caption = strNick
    Queries(intQueries).lblHost = strHostX
    Queries(intQueries).strHost = strHostX
    Queries(intQueries).Visible = True
    Queries(intQueries).Tag = intQueries
    NewQuery = intQueries
End Function

Sub NickQuit(strNick As String, strMsg As String)
    For i = 1 To intChannels
        If Channels(i).InChannel(strNick) And Channels(i).strName <> "" Then
            Channels(i).RemoveNick strNick
            PutData Channels(i).DataIn, strColor & "02" & strBold & strNick & strBold & " has Quit IRC [ " & strMsg & " ]"
            Exit For
        End If
    Next i

    For i = 1 To intQueries
        If LCase(Queries(i).strNick) = LCase(strNick) Then
            PutData Queries(i).DataIn, strColor & "02" & strBold & strNick & strBold & " has Quit IRC [ " & strMsg & " ]"
            Exit Sub
        End If
    Next i
End Sub

Function params(parsed As ParsedData, intStart As Integer, intEnd As Integer) As String
    '* This returns the given parameter range as one string
    '* -1 specified as strEnd means from intStart to the last parameter
    '* Ex: strParams(1) is "#mIRC", then params 2 thru 16 are the nicks
    '*     of the users in the channel, simply do something like this
    '*     strNames = Params(2, -1)
    '*     that would return all the nicks into one string, similar to mIRC's
    '*      $#- identifier ($2-) in this case
    
    '* Declare variables
    Dim strFinal As String, intLast As Integer, i As Integer
    
    '* check for bad parameters
    If intStart < 1 Or intEnd > parsed.intParams Then Exit Function
    
    '* if intEnd = -1 then set it to param count
    If intEnd = -1 Then intLast = parsed.intParams Else intLast = intEnd
        
    For i = intStart To intLast
        strFinal = strFinal & parsed.strParams(i)
        If i <> intLast Then strFinal = strFinal & " "
    Next i
    
    params = Replace(strFinal, vbCr, "")
    strFinal = ""
End Function

Sub ParseData(ByVal strData As String, ByRef parsed As ParsedData)

    '* Declare variables
    Dim strTMP As String, i As Integer
    
    '* Reset variables
    bHasPrefix = False
    parsed.strNick = ""
    parsed.strIdent = ""
    parsed.strHost = ""
    parsed.strCommand = ""
    parsed.intParams = 1
    ReDim parsed.strParams(1 To 1) As String
    
    '* Check for prefix, if so, parse nick, ident and host (or just host)
    If Left(strData, 1) = ":" Then
        bHasPrefix = True
        strData = Right(strData, Len(strData) - 1)
        '* Put data left of " " in strHost, data right of " "
        '* into strData
        Seperate strData, " ", parsed.strHost, strData
        parsed.strFullHost = parsed.strHost
        
        '* Check to see if client host name
        If InStr(parsed.strHost, "!") Then
            Seperate parsed.strHost, "!", parsed.strNick, parsed.strHost
            Seperate parsed.strHost, "@", parsed.strIdent, parsed.strHost
        End If
    End If
    
    '* If any params, parse
    If InStr(strData, " ") Then
        Seperate strData, " ", parsed.strCommand, strData
        
        parsed.AllParams = strData
       '* Let's parse all the parameters.. yummy
Begin: '* OH NO I USED A LABEL!

        '* If begginning of param is :, indicates that its the last param
        If Left(strData, 1) = ":" Then
            parsed.strParams(parsed.intParams) = Right(strData, Len(strData) - 1)
            GoTo Finish
        End If
        '* If there is a space still, there is more params
        If InStr(strData, " ") Then
            Seperate strData, " ", parsed.strParams(parsed.intParams), strData
            'If Left(parsed.strParams(1), 1) = "#" Then MsgBox parsed.strParams(parsed.intParams) & "~~"
            parsed.intParams = parsed.intParams + 1
            ReDim Preserve parsed.strParams(1 To parsed.intParams) As String
            GoTo Begin
        Else
            parsed.strParams(parsed.intParams) = strData
        End If
    Else
        '* No params, strictly command
        parsed.intParams = 0
        parsed.strCommand = strData
    End If
Finish:
End Sub

Sub ParseMode(strChannel As String, strData As String)
    Dim strModes() As String, strChar As String
    Dim i As Integer, intParam As Integer
    Dim bAdd As Boolean
    
    bAdd = True
    strModes = Split(strData, " ")
    For i = 1 To Len(strModes(0))
        strChar = Mid(strModes(0), i, 1)
        Select Case strChar
            Case "+"
                bAdd = True
            Case "-"
                bAdd = False
            Case "v", "b", "o", "h", "k", "l"
                intParam = intParam + 1
                DoMode strChannel, bAdd, strChar, strModes(intParam)
            Case Else
                DoMode strChannel, bAdd, strChar, ""
        End Select
    Next i
End Sub

Sub PutData(RTF As RichTextBox, strData As String)
    If strData = "" Then Exit Sub
    DoEvents
    Dim i As Long, length As Integer, strChar As String, strBuffer As String
    strData = " " & strData
    length = Len(strData)
    i = 1
    RTF.SelStart = Len(RTF.Text)
    RTF.SelColor = lngForeColor
    RTF.SelBold = False
    RTF.SelUnderline = False
    RTF.SelStrikeThru = False
    
    Do
        strChar = Mid(strData, i, 1)
        Select Case strChar
            Case strBold
                Randomize
                If Int(Rnd * 3) = 1 Then DoEvents
                RTF.SelStart = Len(RTF.Text)
                RTF.SelText = strBuffer
                strBuffer = ""
                RTF.SelBold = Not RTF.SelBold
                i = i + 1
            Case strUnderline
                Randomize
                If Int(Rnd * 3) = 1 Then DoEvents
                RTF.SelStart = Len(RTF.Text)
                RTF.SelText = strBuffer
                strBuffer = ""
                RTF.SelUnderline = Not RTF.SelUnderline
                i = i + 1
            Case strReverse
                Randomize
                If Int(Rnd * 3) = 1 Then DoEvents
                RTF.SelStart = Len(RTF.Text)
                RTF.SelText = strBuffer
                strBuffer = ""
                RTF.SelStrikeThru = Not RTF.SelStrikeThru
                i = i + 1
            Case strColor
                Randomize
                If Int(Rnd * 3) = 1 Then DoEvents
                RTF.SelStart = Len(RTF.Text)
                RTF.SelText = strBuffer
                strBuffer = ""
                i = i + 1
                If i > length Then GoTo TheEnd
                Do Until Not ValidColorCode(strBuffer)
                    strBuffer = strBuffer & Mid(strData, i, 1)
                    i = i + 1
                Loop
                strBuffer = LeftR(strBuffer, 1)
                RTF.SelStart = Len(RTF.Text)
                If strBuffer = "" Then
                    RTF.SelColor = lngForeColor
                Else
                    RTF.SelColor = AnsiColor(LeftOf(strBuffer, ","))
                End If
                i = i - 1
                strBuffer = ""
            Case Else
                strBuffer = strBuffer & strChar
                i = i + 1
        End Select
    Loop Until i > length
    If strBuffer <> "" Then
            RTF.SelStart = Len(RTF.Text)
            RTF.SelText = strBuffer
            strBuffer = ""
    End If
TheEnd:
    RTF.SelBold = False
    RTF.SelUnderline = False
    RTF.SelStrikeThru = False
    RTF.SelStart = Len(RTF.Text)
    RTF.SelText = vbCrLf
End Sub
Function QueryExists(strNick As String) As Boolean
    Dim i As Integer

    For i = 1 To intQueries
        If LCase(Queries(i).strNick) = LCase(strNick) Then
            QueryExists = True
            Exit Function
        End If
    Next i
    QueryExists = False
End Function

Function RealNick(strNick As String) As String
    strNick = Replace(strNick, "@", "")
    strNick = Replace(strNick, "%", "")
    strNick = Replace(strNick, "+", "")
    RealNick = strNick
End Function

Sub RefreshList(lstBox As ListBox)
    'lstBox.AddItem "", 0
    'lstBox.RemoveItem 0
End Sub

Function RightOf(strData As String, strDelim As String) As String
    Dim intPos As Integer
    intPos = InStr(strData, strDelim)
    
    If intPos Then
        RightOf = Mid(strData, intPos + 1, Len(strData) - intPos)
    Else
        RightOf = strData
    End If
End Function


Function RightR(strData As String, intMin As Integer)
    On Error Resume Next
    RightR = Right(strData, Len(strData) - intMin)
End Function

Sub Seperate(strData As String, strDelim As String, ByRef strLeft As String, ByRef strRight As String)
    '* Seperates strData into 2 variables based on strDelim
    '* Ex: strData is "Bill Clinton"
    '*     Dim strFirstName As String, strLastName As String
    '*     Seperate strData, " ", strFirstName, strLastName
    
    Dim intPos As Integer
    intPos = InStr(strData, strDelim)
    
    If intPos Then
        strLeft = Left(strData, intPos - 1)
        strRight = Mid(strData, intPos + 1, Len(strData) - intPos)
    Else
        strLeft = strData
        strRight = strData
    End If
End Sub


Function ValidColorCode(strCode As String) As Boolean
    'MsgBox strCode
    Dim c1 As Integer, c2 As Integer
    If strCode Like "" Or _
       strCode Like "#" Or _
       strCode Like "##" Or _
       strCode Like "#,#" Or _
       strCode Like "##,#" Or _
       strCode Like "#,##" Or _
       strCode Like "#," Or _
       strCode Like "##," Or _
       strCode Like "##,##" Or _
       strCode Like ",#" Or _
       strCode Like ",##" Then
        Dim strCol() As String
        strCol = Split(strCode, ",")
        If UBound(strCol) = -1 Then
            ValidColorCode = True
        ElseIf UBound(strCol) = 0 Then
            If strCol(0) = "" Then strCol(0) = 0
            If CInt(strCol(0)) >= 0 And CInt(strCol(0)) < 16 Then
                ValidColorCode = True
            Else
                ValidColorCode = False
            End If
        Else
            If strCol(0) = "" Then strCol(0) = lngForeColor
            If strCol(1) = "" Then strCol(1) = 0
            c1 = CInt(strCol(0))
            c2 = CInt(strCol(1))
            If c2 < 0 Or c2 > 16 Then
                ValidColorCode = False
            Else
                ValidColorCode = True
            End If
        End If
        ValidColorCode = True
    Else
        ValidColorCode = False
    End If
End Function


