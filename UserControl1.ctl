VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.UserControl VOIP 
   BackColor       =   &H80000001&
   ClientHeight    =   2700
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2715
   ScaleHeight     =   2700
   ScaleWidth      =   2715
   Begin VB.TextBox Text1 
      Height          =   345
      Left            =   3900
      TabIndex        =   3
      Text            =   "66"
      Top             =   1440
      Width           =   90
   End
   Begin VB.ComboBox txtServer 
      Height          =   315
      ItemData        =   "UserControl1.ctx":0000
      Left            =   3840
      List            =   "UserControl1.ctx":0002
      Sorted          =   -1  'True
      TabIndex        =   2
      Top             =   1140
      Width           =   330
   End
   Begin VB.ListBox ConnectionList 
      Height          =   1815
      ItemData        =   "UserControl1.ctx":0004
      Left            =   60
      List            =   "UserControl1.ctx":0006
      TabIndex        =   1
      Top             =   750
      Width           =   2595
   End
   Begin VB.CommandButton cmdTalk 
      Caption         =   "&Talk"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   150
      TabIndex        =   0
      Top             =   120
      Width           =   1065
   End
   Begin MSWinsockLib.Winsock TCPSocket 
      Index           =   0
      Left            =   3570
      Top             =   480
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
End
Attribute VB_Name = "VOIP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

Option Explicit
Public CLOSINGAPPLICATION As Boolean                    ' Application status flag
Public wStream As Object
Public voiceport As Long
Public voiceport2 As String
'--------------------------------------------------------------
'--------------------------------------------------------------
'--------------------------------------------------------------
Private Sub cmdTalk_Click()                             ' Activates Audio PlayBack
    On Error Resume Next
    '--------------------------------------------------------------
    Dim rc As Long                                      ' Return Code Variable
    Dim iPort As Integer                                ' Local Port
    Dim itm As Integer                                  ' Current listitem
    '--------------------------------------------------------------
    If (Not wStream.Playing And wStream.PlayDeviceFree And Not wStream.Recording And wStream.RecDeviceFree) Then ' Validate Audio Device Status
        wStream.Playing = True                          ' Turn Playing Status On
        cmdTalk.Caption = "&Playing"                    ' Modify Button Status Caption
        Screen.MousePointer = vbHourglass               ' Set Pointer To HourGlass
        iPort = wStream.StreamInQueue
        Do While (iPort <> NULLPORTID)                  ' While socket ports have data to playback
            ' inLight.Picture = ImgIcons.ListImages(speakON).Picture ' Flash playback image
            'inLight.Refresh                             ' Repaint picture image
            For itm = 0 To ConnectionList.ListCount - 1 ' Search for listitem currently playing sound data
                If (ConnectionList.ItemData(itm) = iPort) Then ' If a match is found...
                    ConnectionList.TopIndex = itm       ' Set that listitem to top of listbox
                    ConnectionList.Selected(itm) = True ' Select listitem to show who is currently talking...
                    Exit For                            ' Quit listitem search
                End If
            Next                                        ' Check next listitem
            rc = wStream.PlayWave(UserControl.hWND, iPort)       ' Play wave data in iPort...
            Call wStream.RemoveStreamFromQueue(iPort)   ' Remove PortID From PlayWave Queue
            iPort = wStream.StreamInQueue
            'inLight.Picture = ImgIcons.ListImages(speakOFF).Picture ' Show done talking image...
            'inLight.Refresh                             ' Repaint image...
        Loop                                            ' Search for next socket in playback queue
        ConnectionList.TopIndex = 0                     ' Reset top image...
        If (ConnectionList.ListCount > 0) Then
            ConnectionList.Selected(0) = True           ' Deselect previously listitem
            ConnectionList.Selected(0) = False          ' Deselect currently selected listitem
        End If
        Screen.MousePointer = vbDefault                 ' Set Pointer To Normal
        cmdTalk.Caption = "&Talk"                       ' Modify Button Status Caption
        wStream.Playing = False                         ' Turn Playing Status Off
    End If
    '--------------------------------------------------------------
End Sub
'--------------------------------------------------------------
'--------------------------------------------------------------
Private Sub cmdTalk_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    ' Activates Audio Recording...
    '--------------------------------------------------------------
    Dim rc As Long                                      ' Return Code Variable
    '--------------------------------------------------------------
    If (Not wStream.Playing And Not wStream.Recording And wStream.RecDeviceFree And wStream.PlayDeviceFree) Then          ' Check Audio Device Status
        wStream.Recording = True                                ' Set Recording Flag
        cmdTalk.Caption = "&Talking"                    ' Update Button Status To "Talking"
        Screen.MousePointer = vbHourglass               ' Set Hourglass
        'outLight.Picture = ImgIcons.ListImages(mikeON).Picture ' Show outgoing message image
        'outLight.Refresh                                ' Repaint image
        rc = wStream.RecordWave(UserControl.hWND, TCPSocket)     ' Record voice and send to all connected sockets
        'outLight.Picture = ImgIcons.ListImages(mikeOFF).Picture ' Show done image
        'outLight.Refresh                                ' Repaint image
        Screen.MousePointer = vbDefault                 ' Reset Mouse Pointer
        cmdTalk.Caption = "&Talk"                       ' Reset Button Status
        If Not wStream.Playing And wStream.PlayDeviceFree And wStream.RecDeviceFree Then               ' Is Audio Device Free?
            Call cmdTalk_Click                          ' Active Playback Of Any Inbound Messages...
        End If
    End If
    '--------------------------------------------------------------
End Sub
'--------------------------------------------------------------
Private Sub cmdTalk_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    wStream.Recording = False                           ' Stop Recording
End Sub

 
'--------------------------------------------------------------
Private Sub TCPSocket_Connect(Index As Integer)
    On Error Resume Next
    ' TCP Connection Has Been Accepted And Is Open...
    '--------------------------------------------------------------
    Call AddConnectionToList(TCPSocket(Index), ConnectionList) ' Add New Connection To List
    ' imgStatus = ImgIcons.ListImages(phoneRingIng).Picture   ' Show Phone Ringing Icon
    Call ResPlaySound(RingOutId)
    'imgStatus = ImgIcons.ListImages(phoneAnswered).Picture  ' Show Phone Answered Icon
    cmdTalk.Enabled = True                                  ' Enabled For Connection...
    ' Tools.Buttons(tbHANGUP).Enabled = (ConnectionList.Text <> "")
    '--------------------------------------------------------------
End Sub
'--------------------------------------------------------------
'--------------------------------------------------------------
Private Sub TCPSocket_ConnectionRequest(Index As Integer, ByVal requestID As Long)
    On Error Resume Next
    ' Accepting Inbound TCP Connection Request...
    '--------------------------------------------------------------
    Dim rc As Long
    Dim Idx As Long
    Dim RemHost As String
    '--------------------------------------------------------------
    If (TCPSocket(Index).RemoteHost <> "") Then
        RemHost = UCase(TCPSocket(Index).RemoteHost)
    Else
        RemHost = UCase(TCPSocket(Index).RemoteHostIP)
    End If
    ' If (Tools.Buttons(tbAUTOANSWER).Value = tbrUnpressed) Then
    '    rc = MsgBox("Incomming call from [" & RemHost & "]..." & vbCrLf & "Do you wish to answer?", vbYesNo)          ' Prompt user to answer...
    'Else
    rc = vbYes
    ' End If
    If (rc = vbYes) Then
        Idx = InstanceTCP(TCPSocket)                            ' Instance TCP Control...
        If (Idx > 0) Then                                       ' Validate that control instance was created...
            TCPSocket(Idx).LocalPort = 0                        ' Set local port to 0, in order to get next available port.
            Call TCPSocket(Idx).Accept(requestID)               ' Accept connection
            Call AddConnectionToList(TCPSocket(Idx), ConnectionList) ' Add New Connection To List
            'imgStatus = ImgIcons.ListImages(phoneRingIng).Picture  ' Show Phone Ringing Icon
            Call ResPlaySound(RingInId)
            'imgStatus = ImgIcons.ListImages(phoneAnswered).Picture ' Show Phone Answered Icon
            cmdTalk.Enabled = True                                 ' Enabled For Connection...
            ' Tools.Buttons(tbHANGUP).Enabled = (ConnectionList.Text <> "")
        End If
    End If
    '--------------------------------------------------------------
End Sub
'--------------------------------------------------------------
'--------------------------------------------------------------
Private Sub TCPSocket_DataArrival(Index As Integer, ByVal BytesTotal As Long)
    ' Incomming Buffer On...
    '--------------------------------------------------------------
    On Error Resume Next                               ' Return Code Variable
    Dim WaveData() As Byte                              ' Byte array of wave data
    Static ExBytes(MAXTCP) As Long                      ' Extra bytes in frame buffer
    Static ExData(MAXTCP) As Variant                    ' Extra bytes from frame buffer
    '--------------------------------------------------------------
    With wStream
        If (TCPSocket(Index).BytesReceived > 0) Then        ' Validate that bytes where actually received
            Do While (TCPSocket(Index).BytesReceived > 0)   ' While data available...
                If (ExBytes(Index) = 0) Then                ' Was there leftover data from last time
                    If (.waveChunkSize <= TCPSocket(Index).BytesReceived) Then ' Can we get and entire wave buffer of data
                        Call TCPSocket(Index).GetData(WaveData, vbByte + vbArray, .waveChunkSize) ' Get 1 wave buffer of data
                        Call .SaveStreamBuffer(Index, WaveData) ' Save wave data to buffer
                        Call .AddStreamToQueue(Index)       ' Queue current stream for playback
                    Else
                        ExBytes(Index) = TCPSocket(Index).BytesReceived ' Save Extra bytes
                        Call TCPSocket(Index).GetData(ExData(Index), vbByte + vbArray, ExBytes(Index)) ' Get Extra data
                    End If
                Else
                    Call TCPSocket(Index).GetData(WaveData, vbByte + vbArray, .waveChunkSize - ExBytes(Index)) ' Get leftover bits
                    ExData(Index) = MidB(ExData(Index), 1) & MidB(WaveData, 1) ' Sync wave bits...
                    Call .SaveStreamBuffer(Index, ExData(Index)) ' Save the current wave data to the wave buffer
                    Call .AddStreamToQueue(Index)           ' Queue the current wave stream
                    ExBytes(Index) = 0                      ' Clear Extra byte count
                    ExData(Index) = ""                      ' Clear Extra data buffer
                End If
            Loop                                            ' Look for next Data Chunk
            If (Not .Playing And .PlayDeviceFree And Not .Recording And .RecDeviceFree) Then     ' Check Audio Device Status
                Call cmdTalk_Click                          ' Start PlayBack...
            End If
        End If
    End With
    '--------------------------------------------------------------
End Sub
'--------------------------------------------------------------
Private Sub TCPSocket_Error(Index As Integer, ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    On Error Resume Next
    TCPSocket(Index).Close                                  ' Close down socket
    Debug.Print "TCPSocket_Error: Number:", Number
    Debug.Print "TCPSocket_Error: Scode:", Hex(Scode)
    Debug.Print "TCPSocket_Error: Source:", Source
    Debug.Print "TCPSocket_Error: HelpFile:", HelpFile
    Debug.Print "TCPSocket_Error: HelpContext:", HelpContext
    Debug.Print "TCPSocket_Error: Description:", Description
    Call DebugSocket(TCPSocket(Index))
End Sub
'--------------------------------------------------------------
Private Sub TCPSocket_Close(Index As Integer)
    ' Closing Current TCP Connection...
    On Error Resume Next '--------------------------------------------------------------
    Call RemoveConnectionFromList(TCPSocket(Index), ConnectionList) ' Remove Connection From List
    Call Disconnect(TCPSocket(Index))                           ' Close Port Connection...
    cmdTalk.Enabled = (ConnectionList.ListCount > 0)            ' Enable/Disable Talk Button...
    If Not cmdTalk.Enabled Then
        'inLight.Picture = ImgIcons.ListImages(speakNO).Picture
        ' outLight.Picture = ImgIcons.ListImages(mikeNO).Picture
    End If
    ' Tools.Buttons(tbHANGUP).Enabled = (ConnectionList.Text <> "")
    If cmdTalk.Enabled Then
        '  imgStatus = ImgIcons.ListImages(phoneHungUp).Picture    ' Show Phone HungUp Icon...
    End If
    Unload TCPSocket(Index)                                     ' Destroy socket instance
    '--------------------------------------------------------------
End Sub

'--------------------------------------------------------------
Private Sub connectionlist_Click()
    'Tools.Buttons(tbHANGUP).Enabled = True
End Sub
'--------------------------------------------------------------
Private Sub ConnectionList_DblClick()
    On Error Resume Next '--------------------------------------------------------------
    Dim MemberID As String                              ' (Server)(TCPidx)(RemoteIP)
    Dim Idx As Long                                     ' TCP idx
    '--------------------------------------------------------------
    If (ConnectionList.Text = "") Then Exit Sub
    MemberID = ConnectionList.List(ConnectionList.ListIndex) ' Get The Conversation MemberID String From List Box
    Call GetIdxFromMemberID(TCPSocket, MemberID, Idx)  ' Get TCP idx From Member ID
    Call RemoveConnectionFromList(TCPSocket(Idx), ConnectionList) ' Clear ListBox Entry(s)...
    Call Disconnect(TCPSocket(Idx))                     ' Disconnect Socket Connection
    Unload TCPSocket(Idx)                               ' Destroy socket instance
    cmdTalk.Enabled = (ConnectionList.ListCount > 0)    ' Enable/Disable Talk Button...
    'Tools.Buttons(tbHANGUP).Enabled = (ConnectionList.Text <> "")
    If Not cmdTalk.Enabled Then
        ' inLight.Picture = ImgIcons.ListImages(speakNO).Picture
        ' outLight.Picture = ImgIcons.ListImages(mikeNO).Picture
    End If
    '--------------------------------------------------------------
End Sub
'--------------------------------------------------------------
'--------------------------------------------------------------
Private Sub Form_Unload(Cancel As Integer)
    '--------------------------------------------------------------
    On Error Resume Next                                  ' TCP socket index
    Dim Socket As Winsock                                   ' TCP socket control
    '--------------------------------------------------------------
    CLOSINGAPPLICATION = True                           ' Set status flag to closing...
    For Each Socket In TCPSocket                        ' For each socket instance
        Call Disconnect(Socket)                         ' Close connection/listen
    Next                                                ' Next Cntl
    Set wStream = Nothing
    ' End Program
    '--------------------------------------------------------------
End Sub
'--------------------------------------------------------------
'--------------------------------------------------------------
'--------------------------------------------------------------
'Private Sub txtServer_KeyPress(KeyAscii As Integer)
'    On Error Resume Next '--------------------------------------------------------------
'    Dim Conn As Long                                        ' Index counter
'    '--------------------------------------------------------------
'    If (KeyAscii = vbKeyReturn) Then                        ' If Return Key Was Pressed...
'        For Conn = 0 To txtServer.ListCount                 ' Search Each Entry In ListBox
'            If (UCase(txtServer.Text) = UCase(txtServer.List(Conn))) Then Exit Sub
'        Next                                                ' If Found Then Exit
'        txtServer.AddItem UCase(txtServer.Text)             ' Add Server To List
'    End If
'    '--------------------------------------------------------------
'End Sub
'--------------------------------------------------------------
Private Sub UserControl_Initialize()
    'frmChat2.Visible = True
    '--------------------------------------------------------------
    On Error Resume Next                                   ' Return Code Variable
     voiceport2 = "701"
     voiceport = "701"
    
    ' Current TCP idx variable
    ' Newly created TCP idx value
    '--------------------------------------------------------------
    CLOSINGAPPLICATION = False                          ' Set status to not closing
    Call InitServerList(txtServer)                      ' Get Common Servers List
    txtServer.Text = txtServer.List(0)                  ' Display First Name In The List
    ' imgStatus = ImgIcons.ListImages(phoneHungUp).Picture ' Change Icon To Phone HungUp
    Set wStream = New WaveStream
    Call wStream.InitACMCodec(WAVE_FORMAT_GSM610, TIMESLICE)
    '   Call wStream.InitACMCodec(WAVE_FORMAT_ADPCM, TIMESLICE)
    '   Call wStream.InitACMCodec(WAVE_FORMAT_MSN_AUDIO, TIMESLICE)
    '   Call wStream.InitACMCodec(WAVE_FORMAT_PCM, TIMESLICE)
    cmdTalk.Enabled = False                             ' Disable Until Connect
    '  Tools.Buttons(tbHANGUP).Enabled = (ConnectionList.Text <> "")
    'inLight.Picture = ImgIcons.ListImages(speakNO).Picture
    'outLight.Picture = ImgIcons.ListImages(mikeNO).Picture
    
End Sub

Public Function connectit(TheServer As String)
    On Error Resume Next                                 ' Return Code Variable
    Dim Idx As Long
    ' TCP Socket control index
  txtServer.Text = TheServer
    ' LocalPort Setting
    Idx = InstanceTCP(TCPSocket)                        ' Instance TCP Control...
    ' If (Idx > 0) Then                                   ' Did control instance get created???
    'Button.Enabled = False                          ' Disable Connect Button
    ConnectionList.Enabled = False                  ' Disable connection list box
    On Error Resume Next
    If Not Connect(TCPSocket(Idx), TheServer, voiceport) Then ' Attempt to connect
        '    Unload TCPSocket(Idx)                       ' Connect failed unload control instance
    End If
    ConnectionList.Enabled = True                   ' Renable connection list box
    ' Button.Enabled = True
End Function
Public Function setports(voiceport2_ As String, voiceport_ As Long)
       voiceport = voiceport_
 
       Call Listen(TCPSocket(0), voiceport2_)
End Function

