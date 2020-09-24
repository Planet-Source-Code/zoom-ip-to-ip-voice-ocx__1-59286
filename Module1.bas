Attribute VB_Name = "Module1"

Option Explicit
Public Sub InitServerList(ServerList As ComboBox)
    ' Populate Server List Box...
    ServerList.AddItem "localhost"
    ServerList.AddItem "127.0.0.1"
End Sub
Public Sub DebugSocket(TCPSocket As Winsock)
    Debug.Print "TCPSocket.RemoteHost", TCPSocket.RemoteHost
    Debug.Print "TCPSocket.RemoteHostIP", TCPSocket.RemoteHostIP
    Debug.Print "TCPSocket.RemotePort", TCPSocket.RemotePort
    Debug.Print "TCPSocket.LocalHostName", TCPSocket.LocalHostName
    Debug.Print "TCPSocket.LocalIP", TCPSocket.LocalIP
    Debug.Print "TCPSocket.LocalPort", TCPSocket.LocalPort
    Debug.Print "TCPSocket.State", TCPSocket.State
    Debug.Print "====================================================="
End Sub
Public Sub ResPlaySound(ResourceId As Long)
    ' Uses Sound Play Sound To Play Back PreRecorded WaveFiles
    Dim sndBuff As String
'    sndBuff = StrConv(LoadResData(ResourceId, "WAVE"), vbUnicode)
    Call sndPlaySound(sndBuff, SND_SYNC Or SND_MEMORY)
End Sub
Public Sub AddConnectionToList(Socket As Winsock, ConnList As ListBox)
    ' Adds A Connection Reference To A ListBox - [(Server)(LocalPort)(RemotePort)]
    Dim MemberID As String                      ' Connection Reference Variable
    ' Create MemberID From HostName and RemoteIP
    MemberID = Socket.RemoteHostIP & "  [" & Format(Socket.RemotePort, "0") & "] - [" & Format(Socket.LocalPort, "0") & "]"
    ConnList.AddItem MemberID                   ' Add New Member To List
    ConnList.ItemData(ConnList.NewIndex) = Socket.Index
End Sub
Public Sub RemoveConnectionFromList(Socket As Winsock, ConnList As ListBox)
    ' Removes A Connection Reference From A ListBox
    Dim Conn As Long                                ' Connection Array Element Variable
    Dim MemberID As String                          ' Connection Reference Variable
    ' Create MemberID From HostName and RemoteIP
    MemberID = Socket.RemoteHostIP & "  [" & Format(Socket.RemotePort, "0") & "] - [" & Format(Socket.LocalPort, "0") & "]"
    For Conn = 0 To ConnList.ListCount - 1          ' Search Each Member In List
        If (ConnList.List(Conn) = MemberID) Then    ' Look For MemberID In List
            ConnList.RemoveItem Conn                ' Remove MemberID From List
        End If
    Next                                            ' Next Connection
End Sub
Public Sub GetIdxFromMemberID(Sockets As Variant, MemberID As String, Index As Long)
                                 ' Socket cntl index
    Dim LocPortID As Long                           ' Local Port ID
    Dim RemPortID As Long                           ' Remote Port ID
    Dim RemoteIP As String                          ' Remote Host IP address
    Dim sStart As Long                              ' Substring begin position
    Dim sEnd As Long                                ' Substring end postition
    Dim Socket As Winsock                               ' Winsock socket
    sStart = 1
    sEnd = InStr(1, MemberID, " ") - 1              ' Get end of remote ip address
    If (sEnd > 1) Then
        RemoteIP = Mid(MemberID, sStart, sEnd)      ' Get remote host ip address
        sStart = InStr(sEnd, MemberID, "[") + 1     ' Get start of remote port
        If (sStart > 1) Then                        ' If Start found
            sEnd = InStr(sStart, MemberID, "]") - 1 ' Get end of remote port
            If (sEnd > 2) Then                      ' If end found
                RemPortID = Val(Mid(MemberID, sStart, sEnd)) ' Get RemotePort
                sStart = InStr(sEnd, MemberID, "[") + 1      ' Get start of local port
                If (sStart > 1) Then
                    sEnd = InStr(sStart, MemberID, "]") - 1  ' Get end of local port
                    If (sEnd > 2) Then                       ' If End Found
                        LocPortID = Val(Mid(MemberID, sStart, sEnd)) ' Extract local port
                        For Each Socket In Sockets
                            If ((Socket.RemoteHostIP = RemoteIP) And (Socket.RemotePort = RemPortID) And (Socket.LocalPort = LocPortID) And (Socket.Index > 0)) Then    ' Was a match found???
                                Index = Socket.Index        ' Save and return index
                                Exit Sub                    ' Done... exit
                            End If
                        Next
                    End If
                End If
            End If
        End If
    End If
End Sub

