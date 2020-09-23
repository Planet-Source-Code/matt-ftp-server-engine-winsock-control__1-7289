VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Server"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'''''''''''''''''''''''''''''''''''''''''''
'Server configuration
'''''''''''''''''''''''''''''''''''''''''''
'Port server listens on
Private Port As Long

'Allows server admin to set max connections,
'must be MAX_CONNECTIONS <= the server constant
Private MaxClients As Integer

'Specifies the packet size (in bytes) to send data
'Shouldn't really be over 4096 bytes (4K)
Private TransferBufferSize As Long

'Every time a new clients connectes ClientCounter is
'incremented by 1, that new number becomes the new
'clients unique ID number.
Private ClientCounter As Long

'Every time a new client connects ConnectedClients is
'incremented by 1, every time a client disconnects,
'the number is decremented by 1.
Private ConnectedClients As Long

'This gets set to True when the server has been started
'and gets set back to False when it is shutdown.
'It should be False by default.
Private ServerActive As Boolean

'''''''''''''''''''''''''''''''''''''''''''
'Server types
'''''''''''''''''''''''''''''''''''''''''''

'Client actions
Private Enum ClientStatus

    'CLIENT STATUS CONSTANTS
    stat_IDLE = 0
    stat_LOGGING_IN = 1
    stat_GETTING_DIR_LIST = 2
    stat_UPLOADING = 3
    stat_DOWNLOADING = 4

End Enum

Private Enum ConnectModes
    
    'CONNECTION MODE CONSTANTS
    cMode_NORMAL = 0
    cMode_PASV = 1

End Enum

Private Type ftpClient

    inUse As Boolean                'Identifies if this slot is being used.
    ID As Long                      'Unique number to identify a client.
    UserName As String              'User name client is is logged in as.
    IPAddress As String             'IP address of the client.
    DataPort As Long                'Port number open on the client for the server to connect to.
    ConnectedAt As String           'Time the client first connected.
    IdleSince As String             'Last recorded time the client sent a command to the server.
    TotalBytesUploaded As Long      'Total bytes uploaded by client from the current session.
    TotalBytesDownloaded As Long    'Total bytes downloaded by client from the current session.
    TotalFilesUploaded As Long      'Total files uploaded by client from the current session.
    TotalFilesDownloaded As Long    'Total files downloaded by client from the current session.
    CurrentFile As String           'Current file being transfer, if any.
    cFileTotalBytes As Long         'Total number of bytes of the file being transfered.
    cTotalBytesXfer As Long         'Total bytes of the current file that has been transfered.
    fFile As Long                   'Reference number to an open file on the server, if any.
    ConnectMode As ConnectModes     'If the client uses PASV mode or not.
    HomeDir As String               'Initial directory client starts in when they first connect.
    CurrentDir As String            'Current directory.
    Status As ClientStatus          'What the client is currently doing.

End Type

'''''''''''''''''''''''''''''''''''''''''''
'Server constants
'''''''''''''''''''''''''''''''''''''''''''
'900 seconds (15 mins.) of idle time will disconnect the client
'without regard to specific user settings
Private Const MAX_IDLE_TIME = 900

'500 simultaneous connections for the server
Private Const MAX_CONNECTIONS = 500

'Array that holds client information for every client.
Private client(MAX_CONNECTIONS) As ftpClient

'''''''''''''''''''''''''''''''''''''''''''
'Server events
'''''''''''''''''''''''''''''''''''''''''''
'ServerErrorOccurred() event fires when an error occurs
'in a sub/function that wasnt called outside of the object.
Public Event ServerErrorOccurred(ByVal errNumber As Long)

'ServerStarted() event fires after successful startup of Winsock
'and the specified port is successfully opened, server is ready to accept clients.
Public Event ServerStarted()

'ServerStopped() event fires after all connected clients
'have been disconnected, Winsock is shutdown and other
'misc. variables are reset.
Public Event ServerStopped()

'NewClient() event fires when a new connection is accepted.
Public Event NewClient(ByVal ClientID As Long)

'ClientLoggedIn() event fires when a connected client
'successfully logs in with a correct user name and password.
Public Event ClientLoggedIn(ByVal ClientID As Long)

'ClientLoggedOut() event fires when a client
'disconnects or is disconnected by the server.
'When this event fires, the information on that
'client will no longer be available.
Public Event ClientLoggedOut(ByVal ClientID As Long)

'ClientSentCommand() event fires when a client
'sends a FTP command to the server.
Public Event ClientSentCommand(ByVal ClientID As Long, Command As String, Args As String)

'ClientStatusChanged() event fires when a clients
'status has changed.  This event will most likely
'follow after another.  eg. client starts a file
'transfer...
Public Event ClientStatusChanged(ByVal ClientID As Long)


'================================= MISC APIS/TYPES =================================


'These functions are for generating a list of files and directories.
Private Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" ( _
    ByVal lpFileName As String, _
    lpFindFileData As WIN32_FIND_DATA _
    ) As Long

Private Declare Function FindNextFile Lib "kernel32" Alias "FindNextFileA" ( _
    ByVal hFindFile As Long, _
    lpFindFileData As WIN32_FIND_DATA _
    ) As Long

Private Declare Function FileTimeToSystemTime Lib "kernel32" ( _
    lpFileTime As FILETIME, _
    lpSystemTime As SYSTEMTIME _
    ) As Long

Private Type FILETIME
    dwLowDateTime As Long
    dwHighDateTime As Long
End Type

Private Declare Function FindClose Lib "kernel32" ( _
    ByVal hFindFile As Long _
    ) As Long

Private Const MAX_PATH = 260

Private Type WIN32_FIND_DATA
    dwFileAttributes As Long
    ftCreationTime As FILETIME
    ftLastAccessTime As FILETIME
    ftLastWriteTime As FILETIME
    nFileSizeHigh As Long
    nFileSizeLow As Long
    dwReserved0 As Long
    dwReserved1 As Long
    cFileName As String * MAX_PATH
    cAlternate As String * 14
End Type

Private Type SYSTEMTIME
    wYear As Integer
    wMonth As Integer
    wDayOfWeek As Integer
    wDay As Integer
    wHour As Integer
    wMinute As Integer
    wSecond As Integer
    wMilliseconds As Long
End Type

'''''''''''''''''''''''''''''''''''''''''''
'Server public properties
'''''''''''''''''''''''''''''''''''''''''''
'Use these properties to get information about the server
'outside of this class module.
Public Property Get ListeningPort() As Long
    ListeningPort = Port
End Property

Public Property Let ListeningPort(NewPort As Long)
    If Port = 0 Then
        Port = NewPort
    End If
End Property

Public Property Get ServerMaxClients() As Integer
    ServerMaxClients = MaxClients
End Property

Public Property Let ServerMaxClients(Max As Integer)
    If Max >= 0 Then
        MaxClients = Max
    End If
End Property

Public Property Get TransBufferSize() As Long
    TransBufferSize = TransferBufferSize
End Property

Public Property Let TransBufferSize(BuffSize As Long)
    If BuffSize > 0 Then
        TransferBufferSize = BuffSize
    End If
End Property

Public Property Get CurrentConnectedClients() As Long
    CurrentConnectedClients = ConnectedClients
End Property

Public Property Get CurrentClientCounter() As Long
    CurrentClientCounter = ClientCounter
End Property

'Use these properties to get information about a specific client
'outside of this class module.
Public Property Get GetClientConnectedAt(ClientID As Long) As String
    GetClientConnectedAt = client(GetClientArrayLocByID(ClientID)).ConnectedAt
End Property

Public Property Get GetClientConnectMode(ClientID As Long) As String
    GetClientConnectMode = client(GetClientArrayLocByID(ClientID)).ConnectMode
End Property

Public Property Get GetClientcTotalBytesXfer(ClientID As Long) As Long
    GetClientcTotalBytesXfer = client(GetClientArrayLocByID(ClientID)).cTotalBytesXfer
End Property

Public Property Get GetClientcFileTotalBytes(ClientID As Long) As Long
    GetClientcFileTotalBytes = client(GetClientArrayLocByID(ClientID)).cFileTotalBytes
End Property

Public Property Get GetClientCurrentDir(ClientID As Long) As String
    GetClientCurrentDir = client(GetClientArrayLocByID(ClientID)).CurrentDir
End Property

Public Property Get GetClientCurrentFile(ClientID As Long) As String
    GetClientCurrentFile = client(GetClientArrayLocByID(ClientID)).CurrentFile
End Property

Public Property Get GetClientDataPort(ClientID As Long) As Long
    GetClientDataPort = client(GetClientArrayLocByID(ClientID)).DataPort
End Property

Public Property Get GetClientfFile(ClientID As Long) As Long
    GetClientfFile = client(GetClientArrayLocByID(ClientID)).fFile
End Property

Public Property Get GetClientHomeDir(ClientID As Long) As String
    GetClientHomeDir = client(GetClientArrayLocByID(ClientID)).HomeDir
End Property

Public Property Get GetClientIdleSince(ClientID As Long) As Long
    GetClientIdleSince = client(GetClientArrayLocByID(ClientID)).IdleSince
End Property

Public Property Get GetClientIPAddress(ClientID As Long) As String
    GetClientIPAddress = client(GetClientArrayLocByID(ClientID)).IPAddress
End Property

Public Property Get GetClientStatus(ClientID As Long) As String
    GetClientStatus = ServerGetClientStatusDescription(client(GetClientArrayLocByID(ClientID)).Status)
End Property

Public Property Get GetClientTotalBytesDownloaded(ClientID As Long) As Long
    GetClientTotalBytesDownloaded = client(GetClientArrayLocByID(ClientID)).TotalBytesDownloaded
End Property

Public Property Get GetClientTotalBytesUploaded(ClientID As Long) As Long
    GetClientTotalBytesUploaded = client(GetClientArrayLocByID(ClientID)).TotalBytesUploaded
End Property

Public Property Get GetClientTotalFilesDownloaded(ClientID As Long) As Long
    GetClientTotalFilesDownloaded = client(GetClientArrayLocByID(ClientID)).TotalFilesDownloaded
End Property

Public Property Get GetClientTotalFilesUploaded(ClientID As Long) As Long
    GetClientTotalFilesUploaded = client(GetClientArrayLocByID(ClientID)).TotalFilesUploaded
End Property

Public Property Get GetClientUserName(ClientID As Long) As String
    GetClientUserName = client(GetClientArrayLocByID(ClientID)).UserName
End Property

'''''''''''''''''''''''''''''''''''''''''''
'Server functions
'''''''''''''''''''''''''''''''''''''''''''
Public Function StartServer() As Long

    'Make sure you dont start the server again.
    If ServerActive = True Then
        StartServer = 1001
        Exit Function
    End If

    'Make sure the server has been givin a port before doing anything.
    If Port < 1 Then
        StartServer = 1002
        Exit Function
    End If

    'If no TransferBufferSize was assigned then default to 4096 (4K)
    If TransferBufferSize < 1 Then TransferBufferSize = 4096

    'Start listening using the intial instance of the Winsock control.
    With frmWinsock.CommandSock(0)
        .LocalPort = Port
        .Listen
    End With

    ServerActive = True
    'Server started with out error and is ready to accept clients.
    RaiseEvent ServerStarted

End Function

Public Sub NewClient(requestID As Long)

    Dim tmpID As Long
    Dim i As Integer

    'Increment counters
    ConnectedClients = ConnectedClients + 1
    ClientCounter = ClientCounter + 1

    'This just helps to make sure that things dont get
    'screwed up incase two clients connected at nearly the exact same time.
    tmpID = ClientCounter

    'Find an open slot in the client() array.
    Do
        i = i + 1
    Loop Until client(i).inUse = False

    With client(i)
        .inUse = True

        'Create new instances of the Winsock control for the
        'new client.  This is to keep the first instance, CommandSock(0)
        'listening for more clients.
        Load frmWinsock.CommandSock(i)
        Load frmWinsock.DataSock(i)
    
        'Accept the new connection with the new instance of the Winsock control.
        frmWinsock.CommandSock(i).Accept requestID

        'Fill client data
        .ConnectedAt = Now
        .ID = tmpID
        .Status = stat_LOGGING_IN
        .IdleSince = Now
        .IPAddress = frmWinsock.CommandSock(i).RemoteHostIP
    End With

    'Raise event to tell the outside program that a new client has connected.
    RaiseEvent NewClient(client(i).ID)

    'If max clients has been reached then disconnect the client.
    If ((ConnectedClients > MaxClients) And (MaxClients <> 0)) Or (ConnectedClients > MAX_CONNECTIONS) Then
        SendResponse i, "421 Too many users - try again later."
        LogoutClient , i
        Exit Sub
    End If

    SendResponse i, "220 P1mp FTP Engine version " & App.Major & ".0" & App.Minor & " build " & App.Revision

End Sub

Private Sub SendResponse(sckArrayLoc As Integer, data As String)

    'Send a command or response to the client, a carriage return and line feed
    'must be added to the end of the string to singnify the end of this command/response.
    frmWinsock.CommandSock(sckArrayLoc).SendData data & vbCrLf
    DoEvents

End Sub

Private Sub SendData(sckArrayLoc As Integer, data As String)

    frmWinsock.DataSock(sckArrayLoc).SendData data

End Sub

Public Sub SendComplete(sckArrayLoc As Integer)

    'After sending data on a data connection to the client
    'via the DataSock control on frmWinsock, this sub will
    'run after all the data has been sent.

    With client(sckArrayLoc)
        Select Case .Status
            Case stat_GETTING_DIR_LIST
                'Send complete means the directory list is done sending
                'close the data connection and alert the client.
                frmWinsock.DataSock(sckArrayLoc).Close
                SendResponse sckArrayLoc, "226 Transfer complete."
    
                'Change client status
                .Status = stat_IDLE
                RaiseEvent ClientStatusChanged(.ID)
    
            Case stat_DOWNLOADING
                If .cFileTotalBytes = .cTotalBytesXfer Then
                    'Reset variables used for the file transfer.

                    'Close the open file.
                    Close #.fFile

                    'Close the data connection.
                    frmWinsock.DataSock(sckArrayLoc).Close
                    .DataPort = 0

                    'Tell the client the transfer is complete.
                    SendResponse sckArrayLoc, "226 Transfer complete."

                    .cFileTotalBytes = 0
                    .cTotalBytesXfer = 0

                    'Change client status
                    .Status = stat_IDLE
                    RaiseEvent ClientStatusChanged(.ID)
                Else
                    SendFile sckArrayLoc
                End If
        End Select
    End With

End Sub

Private Sub LoginClient(cArrayLoc As Integer, Password As String)

    'Right now the server only allows anonymous connections,
    'because of this we will allow any password to work.

    'Fill in some dummy data until i put in a user system.
    'When i do, validate passwords here, etc.
    With client(cArrayLoc)
        .HomeDir = "D:\"
        .CurrentDir = .HomeDir
        SendResponse cArrayLoc, "230 User logged in, proceed."
        .Status = stat_IDLE
    End With

    RaiseEvent ClientLoggedIn(ByVal client(cArrayLoc).ID)
    RaiseEvent ClientStatusChanged(ByVal client(cArrayLoc).ID)

End Sub

Public Sub LogoutClient(Optional ByVal ID As Long, Optional cArrayLoc As Integer)

    'Either pass ID or cArrayLoc of the client to log them out.
    'This Sub will disconnect the client,
    'and free system resources used by this client.

    On Error Resume Next

    If ID = 0 And cArrayLoc = 0 Then Exit Sub
    
    Dim ArrayPos As Integer
    Dim tmp As Long
    
    If ID = 0 Then
        ArrayPos = cArrayLoc
    Else
        ArrayPos = GetClientArrayLocByID(ID)
    End If

    'More precautionary code.
    If client(ArrayPos).ID = 0 Then Exit Sub
    If ArrayPos < 1 Then Exit Sub

    With client(ArrayPos)
        'Close the connection
        frmWinsock.CommandSock(ArrayPos).Close
        frmWinsock.DataSock(ArrayPos).Close

        Unload frmWinsock.CommandSock(ArrayPos)
        Unload frmWinsock.DataSock(ArrayPos)

        If .fFile <> 0 Then Close #.fFile

        'Reset all client info for next use.
        .ConnectedAt = ""
        .ConnectMode = 0
        .cTotalBytesXfer = 0
        .cFileTotalBytes = 0
        .CurrentDir = ""
        .CurrentFile = ""
        .DataPort = 0
        .fFile = 0
        .HomeDir = ""
        tmp = .ID 'Save the ID for the ClientLoggedOut event.
        .ID = 0
        .IdleSince = ""
        .IPAddress = ""
        .Status = stat_IDLE
        .TotalBytesDownloaded = 0
        .TotalBytesUploaded = 0
        .TotalFilesDownloaded = 0
        .TotalFilesUploaded = 0
        .UserName = ""
        .inUse = False
    End With

    If ConnectedClients > 0 Then ConnectedClients = ConnectedClients - 1

    'Client is now officially logged out.
    RaiseEvent ClientLoggedOut(ByVal tmp)

End Sub

Private Function GetClientArrayLocByID(ByVal ID As Long) As Integer

    'This function returns the array number where the information
    'of the client of the specified ID is located in the client() array.
    Dim i As Integer

    For i = 0 To UBound(client)
        If client(i).ID = ID Then
            GetClientArrayLocByID = i
            Exit Function
        End If
    Next

    'ID not found in client array, return -1 error
    GetClientArrayLocByID = -1

End Function

Public Sub ProcFTPCommand(ByVal sckArrayLoc As Integer, ByRef raw_data As String)

    'This sub processes FTP commands sent by the client.

    Dim data
    Dim ftpCommand As String
    Dim ftpArgs As String

    data = Replace$(raw_data, vbCrLf, "")    'Remove carriage return and line feed.

    If InStr(data, " ") = 0 Then
        ftpCommand = data 'Command has no arguments...
    Else
        ftpCommand = Left$(data, (InStr(data, " ") - 1))          'Get the command.
        ftpArgs = Right$(data, (Len(data) - InStr(data, " ")))    'Get the command arguments.
    End If

    RaiseEvent ClientSentCommand(client(sckArrayLoc).ID, ftpCommand, ftpArgs)
    client(sckArrayLoc).IdleSince = Now

    Select Case UCase$(ftpCommand)

        Case "USER"
            'For now i am simply hardcoding an anonymous account until
            'i get around to putting in a user system.
            If ftpArgs = "anonymous" Then
                client(sckArrayLoc).UserName = ftpArgs
                SendResponse sckArrayLoc, "331 User name ok, need password."
            Else
                SendResponse sckArrayLoc, "530 Not logged in: No such account " & ftpArgs
            End If

        Case "PASS"
            'Now have user name and password, attempt to login the client.
            LoginClient sckArrayLoc, ftpArgs
        '    OLD CODE FROM OLD FTP SERVER
        '    If ftpArgs = "" Then
        '        Dim i As Integer
        '        Dim msgArray() As String
        '            msgArray = Split(Server.WelcomeMsg, vbCrLf)
        '        For i = 0 To UBound(msgArray)
        '            msgArray(i) = "230-" & msgArray(i)
        '            SendResponse SocketHandle, msgArray(i)
        '        Next
        '        login_client Socket
        '        SendResponse SocketHandle, "230 User logged in, proceed."
        '    Else
        '        SendResponse SocketHandle, "530 Not logged in"
        '        logout_client Socket
        '    End If

        Case "TYPE"
            SendResponse sckArrayLoc, "200 Type set to " & ftpArgs

        Case "REST"
            SendResponse sckArrayLoc, "350 Restarting at " & ftpArgs & " - send STORE or RETRIEVE to initiate transfer."

        Case "PWD"
            SendResponse sckArrayLoc, "257 " & Chr(34) _
                & ConvPathToRelative(client(sckArrayLoc).HomeDir, client(sckArrayLoc).CurrentDir) _
                & Chr(34) & " is current directory."

        Case "PORT"
            Dim tmpArray() As String 'Six slots required (0 - 5)
            tmpArray = Split(ftpArgs, ",")
            client(sckArrayLoc).DataPort = tmpArray(4) * 256 Or tmpArray(5)
            SendResponse sckArrayLoc, "200 Port command successful."

        Case "LIST"
            SendResponse sckArrayLoc, "150 Opening ASCII mode data connection for /bin/ls."
            
            client(sckArrayLoc).Status = stat_GETTING_DIR_LIST
            RaiseEvent ClientStatusChanged(client(sckArrayLoc).ID)
            
            GetDirectoryList sckArrayLoc

        Case "RETR"
            GetFileToSend sckArrayLoc, ftpArgs

        Case "CWD"
            ChangeDirectory sckArrayLoc, ftpArgs

        Case "CDUP"
            Dim tmp As String

            tmp = client(sckArrayLoc).CurrentDir
            If isRootDir(sckArrayLoc, tmp) = False Then
                'Hack off the last directory
                If Right$(tmp, 1) = "\" Then tmp = Left$(tmp, Len(tmp) - 1)
                tmp = Left$(tmp, InStrRev(tmp, "\"))
            End If

            ChangeDirectory sckArrayLoc, ConvPathToRelative(client(sckArrayLoc).HomeDir, tmp)

        Case "PASV"
            client(sckArrayLoc).ConnectMode = cMode_PASV
            SendResponse sckArrayLoc, "227 Entering Passive Mode (" _
                & Replace(frmWinsock.CommandSock(0).LocalIP, ".", ",") & OpenLocalDataPort(sckArrayLoc) & ")"

        Case "NOOP"
            SendResponse sckArrayLoc, "200 NOOP command successful."

        Case Else
            SendResponse sckArrayLoc, "502 Command not implemented."

    End Select

End Sub

Private Sub GetDirectoryList(cArrayLoc As Integer)

    'This code tends to be very very close when it
    'comes across a directory with lots of files to list.
    'When trying to list my C:\WinNT\System32 directory on my computer
    'it took nearly 10 seconds!  :(
    'If anyone has a faster routine to do this please tell me! =)

    Dim hFile As Long
    Dim r As Long
    Dim fname As String
    Dim WFD As WIN32_FIND_DATA
    Dim dirList As String
    Dim permissions As String

    'Get the first file in the directory (it will usually return ".")
    hFile = FindFirstFile(client(cArrayLoc).CurrentDir & "*.*" + Chr$(0), WFD)
    
    'If nothing returned, bail out.
    'If hFile < 0 Then
    '    frmMain.dataSock(socket).Close
    '    Exit Sub
    'End If

    If Left$(WFD.cFileName, InStr(WFD.cFileName, vbNullChar) - 1) <> "." And Left$(WFD.cFileName, InStr(WFD.cFileName, vbNullChar) - 1) <> ".." Then
        If (WFD.dwFileAttributes And vbDirectory) Then
            permissions = "drwx------"
        Else
            permissions = "-rwx------"
        End If
        
        dirList = permissions _
            & " 1 user group " _
            & WFD.nFileSizeLow _
            & get_date(WFD.ftLastWriteTime) _
            & Left$(WFD.cFileName, InStr(WFD.cFileName, vbNullChar) - 1) _
            & vbCrLf
    End If

    While FindNextFile(hFile, WFD)
        If Left$(WFD.cFileName, InStr(WFD.cFileName, vbNullChar) - 1) <> "." And Left$(WFD.cFileName, InStr(WFD.cFileName, vbNullChar) - 1) <> ".." Then
            If (WFD.dwFileAttributes And vbDirectory) Then
                permissions = "drwx------"
            Else
                permissions = "-rwx------"
            End If
            dirList = dirList _
                & permissions _
                & " 1 user group " _
                & WFD.nFileSizeLow _
                & get_date(WFD.ftLastWriteTime) _
                & Left$(WFD.cFileName, InStr(WFD.cFileName, vbNullChar) - 1) _
                & vbCrLf
        End If

        DoEvents

    Wend

    'Close the search handle
    r = FindClose(hFile)
    
    MakeDataConnection cArrayLoc

    If dirList = "" Then
        'No files in this directory.
        frmWinsock.DataSock(cArrayLoc).Close
        SendResponse cArrayLoc, "226 Transfer complete."

        'Change client status
        client(cArrayLoc).Status = stat_IDLE
        RaiseEvent ClientStatusChanged(client(cArrayLoc).ID)
        Exit Sub
    End If

    SendData cArrayLoc, dirList

End Sub

Private Function MakeDataConnection(sckArrayLoc As Integer) As Long

    'In Normal mode the server will connect to the client.
    'In PASV mode the client will connect to an open port
    'on the server for data transfers.

    If client(sckArrayLoc).ConnectMode = cMode_NORMAL Then
        'Connect to the clients open data port so data transfer can begin.
        frmWinsock.DataSock(sckArrayLoc).RemoteHost = client(sckArrayLoc).IPAddress
        frmWinsock.DataSock(sckArrayLoc).RemotePort = client(sckArrayLoc).DataPort
        frmWinsock.DataSock(sckArrayLoc).Connect
    End If

    'Make sure the socket is connected and ready before continuing.
    Do
        DoEvents
    Loop Until frmWinsock.DataSock(sckArrayLoc).State = sckConnected

End Function

Private Function OpenLocalDataPort(sckArrayLoc As Integer) As String

    'In PASV mode the server creates a listing socket for the
    'client to connect to.  In Normal mode the client creates
    'the listening socket and the server will connect but if
    'the client is behind a firewall this can cause problems.

    Dim Nr1 As Integer
    Dim Nr2 As Integer

    Randomize Timer
    Nr1 = Int(Rnd * 12) + 5
    Nr2 = Int(Rnd * 254) + 1

    frmWinsock.DataSock(sckArrayLoc).Close
    frmWinsock.DataSock(sckArrayLoc).LocalPort = (Nr1 * 256) Or Nr2
    frmWinsock.DataSock(sckArrayLoc).Listen

    'In PASV mode you will not make the connection here, you first need to
    'tell the client what the new open port is.

    OpenLocalDataPort = "," & Nr1 & "," & Nr2

End Function

Private Function isRootDir(cArrayLoc As Integer, strDir As String) As Boolean

    If client(cArrayLoc).HomeDir = strDir Then isRootDir = True

End Function

Private Sub ChangeDirectory(cArrayLoc As Integer, ChangeTo As String)

    If Left$(ChangeTo, 1) = "/" Then
        'Client specified an absolute path.
        If FileExists(ConvPathToLocal(client(cArrayLoc).HomeDir, ChangeTo)) = True Then
            client(cArrayLoc).CurrentDir = ConvPathToLocal(client(cArrayLoc).HomeDir, ChangeTo)
        Else
            SendResponse cArrayLoc, "550 " & ChangeTo & ": No such file or directory."
            Exit Sub
        End If
    Else
        'Client specified a path relative to their current path.
        If FileExists(ConvPathToLocal(client(cArrayLoc).CurrentDir, ChangeTo)) = True Then
            client(cArrayLoc).CurrentDir = ConvPathToLocal(client(cArrayLoc).CurrentDir, ChangeTo)
        Else
            SendResponse cArrayLoc, "550 " & ChangeTo & ": No such file or directory."
            Exit Sub
        End If
    End If

    SendResponse cArrayLoc, "250 Directory changed to " & ConvPathToRelative(client(cArrayLoc).HomeDir, client(cArrayLoc).CurrentDir)

End Sub

Private Sub GetFileToSend(cArrayLoc As Integer, File As String)

    With client(cArrayLoc)
        'Make sure file exists
        If FileExists(.CurrentDir & File) = False Then
            SendResponse cArrayLoc, "550 " & File & ": No such file or directory."
            Exit Sub
        End If

        'Save the total size of the file
        .cFileTotalBytes = FileLen(.CurrentDir & File)

        'Save the file name and location.
        .CurrentFile = .CurrentDir & File

        'Tell the client that we are now going to start the transfer.
        SendResponse cArrayLoc, "150 Opening BINARY mode data connection for " & File & " (" & .cFileTotalBytes & " bytes)"

        'Open the requested file for reading.
        .fFile = FreeFile
        Open .CurrentDir & File For Binary Access Read As #.fFile

        .Status = stat_DOWNLOADING
        RaiseEvent ClientStatusChanged(.ID)
    End With

    'You now need to open a data connection to use for the file transfer.
    MakeDataConnection cArrayLoc
    
    'Once a connection has been made, start the transfer.
    SendFile cArrayLoc

End Sub

Private Sub SendFile(cArrayLoc As Integer)
    
    Dim BlockSize As Integer
    Dim DataToSend As String
    
    BlockSize = TransferBufferSize

    With client(cArrayLoc)
        'Determine the proper buffer size.
        If BlockSize > (.cFileTotalBytes - .cTotalBytesXfer) Then
            BlockSize = (.cFileTotalBytes - .cTotalBytesXfer)
        End If

        DataToSend = Space$(BlockSize) 'allocate space to store data.
        Get #.fFile, , DataToSend 'get data

        .cTotalBytesXfer = .cTotalBytesXfer + BlockSize
        .TotalBytesDownloaded = .TotalBytesDownloaded + BlockSize
    End With

    'Send the data chunk to the client.
    SendData cArrayLoc, DataToSend

End Sub

Public Function ShutdownServer() As Long

    'Close the listening socket
    frmWinsock.CommandSock(0).Close

    ServerActive = False
    RaiseEvent ServerStopped

End Function

Private Function ConvPathToLocal(ByVal StartPath As String, ByVal CurrentPath As String) As String

    'This function takes a relative path and converts to to a local one.
    'eg. home directory is "/", locally "/" equates to "C:\Files\".
    'Their current directory is "/Uploads/Stuff".
    'If passed through this function you will get
    '"C:\Files\Uploads\Stuff\"

    'StartPath must be a local path.  "C:\Files\" identifying the home directory.
    'CurrentPath is the clients current relative path.  "/Uploads/Stuff"

    Dim result As String

    If Right$(StartPath, 1) <> "\" Then StartPath = StartPath & "\"

    If Left$(CurrentPath, 1) = "/" Then CurrentPath = Right$(CurrentPath, Len(CurrentPath) - 1)
    CurrentPath = Replace$(CurrentPath, "/", "\")

    result = StartPath & CurrentPath

    If Right$(result, 1) <> "\" Then result = result & "\"

    ConvPathToLocal = result

End Function

Private Function ConvPathToRelative(ByVal StartPath As String, ByVal CurrentPath As String) As String

    'This function takes an absolute path and converts it to a relative unix style one.
    'eg. home directory is "C:\MyFiles\Music\" and the client is currently in
    '"C:\MyFiles\Music\Alternative\", pass these two paths into this function
    'will give you the result /Alternative

    'If Right$(tmp, 1) = "\" Then tmp = Left$(tmp, Len(tmp) - 1)
    'tmp = Left$(tmp, InStrRev(tmp, "\"))

    If Right$(StartPath, 1) <> "\" Then StartPath = StartPath & "\"
    If Right$(CurrentPath, 1) <> "\" Then CurrentPath = CurrentPath & "\"

    Dim strRelPath As String

    If StartPath = CurrentPath Then
        strRelPath = "/"
    Else
        strRelPath = Replace$(CurrentPath, StartPath, "/")
        strRelPath = Replace$(strRelPath, "\", "/")
    
        If Right$(strRelPath, 1) = "/" Then strRelPath = Left$(strRelPath, Len(strRelPath) - 1)
    End If

    ConvPathToRelative = strRelPath

End Function

Public Function ServerGetClientStatusDescription(ByVal stat As Integer) As String

    Select Case stat
        Case stat_IDLE: ServerGetClientStatusDescription = "Idle"
        Case stat_LOGGING_IN: ServerGetClientStatusDescription = "Connecting..."
        Case stat_GETTING_DIR_LIST: ServerGetClientStatusDescription = "Downloading list of files"
        Case stat_UPLOADING: ServerGetClientStatusDescription = "Uploading"
        Case stat_DOWNLOADING: ServerGetClientStatusDescription = "Downloading"
        Case Else: ServerGetClientStatusDescription = "Unknown status"
    End Select

End Function

Public Function ServerGetErrorDescription(ByVal errCode As Long) As String

    Select Case errCode

        Case 1001: ServerGetErrorDescription = "Server is already running."
        Case 1002: ServerGetErrorDescription = "Server failed to start becuase no port or invalid port was specified."
        
        Case Else: ServerGetErrorDescription = "Unknown error " & errCode

    End Select

End Function


'================================= MISC CODE =================================


Private Function get_date(FT As FILETIME) As String

    'convert the FILETIME (FT) to a SYSTEMTIME (ST) type
    Dim ST As SYSTEMTIME
    Dim r As Long
    Dim ds As String
    
    r = FileTimeToSystemTime(FT, ST)

    ds = DateSerial(ST.wYear, ST.wMonth, ST.wDay)

    If DateDiff("d", ds, Date) > 365 Then
        get_date = Format$(ds, " mmm dd yyyy ")
    Else
        get_date = Format$(ds & " " & ST.wHour & ":" & ST.wMinute, " mmm dd hh:mm ")
    End If

End Function

Private Function FileExists(FileName As String) As Boolean

    Dim hFindFile As Long
    Dim FileData As WIN32_FIND_DATA
    If Right(FileName, 1) = "\" Then
        FileName = FileName & "*.*"
    End If

    hFindFile = FindFirstFile(FileName, FileData)
    If hFindFile = -1 Then
        FileExists = False
    Else
        FileExists = True
    End If

    FindClose hFindFile

End Function
