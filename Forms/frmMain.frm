VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "FTP Engine"
   ClientHeight    =   3855
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   7695
   LinkTopic       =   "Form1"
   ScaleHeight     =   3855
   ScaleWidth      =   7695
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtSvrLog 
      Height          =   3855
      Left            =   0
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Text            =   "frmMain.frx":0000
      Top             =   0
      Width           =   7695
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuServer 
      Caption         =   "&Server"
      Begin VB.Menu mnuStartSvr 
         Caption         =   "&Start Server"
      End
      Begin VB.Menu mnuStopSvr 
         Caption         =   "S&top  Server"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Demo program to show off my FTP engine.
'Author: Matt Thomas (mthomas@aspire.com)
'
'Description: I set out to create a easy to use FTP server class module
'that you could take and drop into any program to make it an FTP server.
'Another reason was that I wanted to have an Event driven FTP server,
'working with Events makes thing soooo much easier.
'
'This is just a start, its not nearly complete but I wanted to show
'you all what I was working on and get some feedback.
'
'You should be able to use the FTP server without having to understand
'much of how it works internally.
'The two files REQUIRED for the FTP server to run are
'Server.cls and frmWinsock.frm
'The rest of the files are just small parts of this demo program.
'
'If you find any bugs/problems in the server code please let me know.
'Sorry for the lack of comments, I will try to add more as time goes on.
'
'This server seems to take a lot of memory, if anyone can see ANYWAY
'to lessen the amount of memory consumed by this program please let me know.
'The program starts off using about 2MB of ram or so.
'Ive tested it with 10 clients banging away at it, with 10 clients
'the server takes about 4MB of ram.
'
'To test, start the server, and login to it as anonymous
'
'This code/program has ONLY been tested in Visual Basic 6.
'This code/program has ONLY been tested using CuteFTP as the FTP client.
'This code/program has ONLY been tested on a Windows 2000 system.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Option Explicit

'''''''''''''''''''''''''''''''''''''''''''
'Server object, events and variables
'''''''''''''''''''''''''''''''''''''''''''
'Declare FTPServer as the Server object
'WithEvents so it gets the events from the Server object.
Public WithEvents FTPServer As Server
Attribute FTPServer.VB_VarHelpID = -1

'''''''''''''''''''''''''''''''''''''''''''
'Form events
'''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_Load()

    'Create new instance of the server
    Set FTPServer = New Server
    
    'When you create an instance of the Server you need to also
    'set the server instance on the frmWinsock form to the instance
    'that you created here so that it can work with the instance of the
    'server that you created.
    Set frmWinsock.FTPServer = FTPServer

End Sub

Public Sub Form_Resize()

    On Error Resume Next

    txtSvrLog.Width = (frmMain.Width - 120)
    txtSvrLog.Height = (frmMain.Height - 690)

End Sub

Public Sub Form_UnLoad(Cancel As Integer)

    'Shutdown server, close Winsock
    StopServer

    'Remove the FTP server object from memory
    Set FTPServer = Nothing
    Set frmWinsock.FTPServer = Nothing

    'Remove the forms from memory
    Unload frmWinsock
    Unload Me

    Set frmWinsock = Nothing
    Set frmMain = Nothing
    End

End Sub

Private Sub mnuStartSvr_Click()

    StartServer

End Sub

Private Sub mnuStopSvr_Click()

    StopServer

End Sub

'''''''''''''''''''''''''''''''''''''''''''
'FTPServer events
'''''''''''''''''''''''''''''''''''''''''''
Private Sub FTPServer_ServerStarted()

    'Once the server has successfully started with out errors
    'and is ready to accept clients, this event sub will run.

    WriteToLogWindow "Server started!", True

End Sub

Private Sub FTPServer_ServerStopped()

    'ServerStopped() event fires after all connected clients
    'have been disconnected, Winsock is shutdown and other
    'misc. variables are reset.

    WriteToLogWindow "Server stopped!", True

End Sub

Private Sub FTPServer_ServerErrorOccurred(ByVal errNumber As Long)

    'Event fires when an error in the server occurs.

    MsgBox FTPServer.ServerGetErrorDescription(errNumber), vbInformation, "Error occured!"

End Sub

Private Sub FTPServer_NewClient(ByVal ClientID As Long)

    'Event fires when a new client successfully connects.

    WriteToLogWindow "Client " & ClientID & " connected! (" & FTPServer.GetClientIPAddress(ClientID) & ")", True

End Sub

Private Sub FTPServer_ClientSentCommand(ByVal ClientID As Long, Command As String, Args As String)

    'Event fires when a connected client sends a FTP command to the server.

    WriteToLogWindow "Client " & ClientID & " sent: " & Command & " " & Args, True

End Sub

Private Sub FTPServer_ClientStatusChanged(ByVal ClientID As Long)

    'Event fires when the clients status has been changed.

    WriteToLogWindow "Client " & ClientID & " Status: " & FTPServer.GetClientStatus(ClientID), True

End Sub

Private Sub FTPServer_ClientLoggedOut(ByVal ClientID As Long)

    'Event fires when a connected client disconnects/is disconnected.
    
    WriteToLogWindow "Client " & ClientID & " logged out!", True

End Sub
