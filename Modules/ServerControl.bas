Attribute VB_Name = "ServerControl"
Option Explicit

Public Sub StartServer()

    'Variable to store the result of functions
    Dim r As Long

    'Before you can actually start the server you must set
    'the proper settings first.
    With frmMain
        'Tell the server object which port to listen on.
        .FTPServer.ListeningPort = 21

        'Total max clients
        .FTPServer.ServerMaxClients = 1
    
        'Start the FTP server.
        r = .FTPServer.StartServer()

        If r <> 0 Then  'Problem starting server
            MsgBox .FTPServer.ServerGetErrorDescription(r), vbCritical
        End If
    End With

End Sub

Public Sub StopServer()

    frmMain.FTPServer.ShutdownServer

End Sub
