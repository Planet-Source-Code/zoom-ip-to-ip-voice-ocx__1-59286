Attribute VB_Name = "Module3"

Option Explicit
Public Function InstanceTCP(TCPArray As Variant) As Long
    Dim Ind As Long                                 ' Array Index Var...
    InstanceTCP = -1                                ' Set Default Value
    On Error GoTo InitControl                       ' IF Error Then Control Is Available
    For Ind = MINTCP To MAXTCP                      ' For Each Member In TCPArray() + 1
        If (TCPArray(Ind).Name = "") Then           ' If Control Is Not Valid Then..
        End If                                      ' ..A Runtime Error Will Occure
    Next                                            ' Search Next Item In Array
InitControl:                                                    ' Initialize New Control
    On Error GoTo ErrorHandler                      ' Enable Error Handling...
    If ((Ind >= MINTCP) And (Ind <= MAXTCP)) Then   ' Check to make sure index value is with in range
        Load TCPArray(Ind)                          ' Create New Member In TCPArray
        InstanceTCP = Ind                           ' Return New TCPctl Index
    End If
    Exit Function                                   ' Exit
ErrorHandler:                                           ' Handler
    Debug.Print Err.Number, Err.Description         ' Debug Errors
    Resume Next                                     ' Ignore Error And Continue
End Function
Public Function Connect(Socket As Winsock, RemHost As String, RemPort As Long) As Boolean
    Connect = False                             ' Set default return code
    Call CloseListen(Socket)                    ' Stop Listening On LocalPort
    Socket.LocalPort = 0                        ' Not necessary, but done just in case
    Call Socket.Connect(RemHost, RemPort)       ' Connect To Server
    Do While ((Socket.State = sckConnecting) Or (Socket.State = sckConnectionPending) Or (Socket.State = sckResolvingHost) Or (Socket.State = sckHostResolved) Or (Socket.State = sckOpen))         ' Attempting To Connect...
        DoEvents                                ' Post Events
    Loop                                        ' Keep Waiting
    Connect = (Socket.State = sckConnected)     ' Did Socket Connect On Port...
End Function
Public Function Listen(Socket As Winsock, theport As String) As Long
    If (Socket.State <> sckListening) Then      ' Is Socket Already Listening
        If (Socket.LocalPort = 0) Then          ' If local port is not initialized then...
            Socket.LocalPort = theport        ' Set standard application port
        End If
        Call Socket.Listen                      ' Listen On Local Port...
    End If
End Function
Public Function CloseListen(Socket As Winsock) As Long
    If (Socket.State = sckListening) Then       ' Is Socket Listening?
        Socket.Close                            ' Close Listen
    End If
End Function
Public Sub Disconnect(Socket As Winsock)
    If (Socket.State <> sckClosed) Then         ' Is Socket Already Closed?
        Socket.Close                            ' Close Socket
    End If
End Sub

