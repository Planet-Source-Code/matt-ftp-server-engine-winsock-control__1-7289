VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmWinsock 
   Caption         =   "Form1"
   ClientHeight    =   420
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3240
   Enabled         =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   420
   ScaleWidth      =   3240
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin MSWinsockLib.Winsock DataSock 
      Index           =   0
      Left            =   480
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock CommandSock 
      Index           =   0
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
End
Attribute VB_Name = "frmWinsock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'The purpose of this form is to basically host the Winsock controls, events etc.,
'I will probably put a Timer control on here as well.

Public WithEvents FTPServer As Server
Attribute FTPServer.VB_VarHelpID = -1

'''''''''''''''''''''''''''''''''''''''''''
'Winsock Events
'''''''''''''''''''''''''''''''''''''''''''
Private Sub CommandSock_ConnectionRequest(Index As Integer, ByVal requestID As Long)

    DoEvents
    FTPServer.NewClient requestID

End Sub

Private Sub DataSock_ConnectionRequest(Index As Integer, ByVal requestID As Long)

    'A connection should only be requested by a client when they are working
    'in PASV mode where the server creates an open port for the client to
    'connect to for data transfers.

    DoEvents
    DataSock(Index).Close
    DataSock(Index).Accept requestID

End Sub

Private Sub CommandSock_DataArrival(Index As Integer, ByVal bytesTotal As Long)

    DoEvents
    Dim raw_data As String
    CommandSock(Index).GetData raw_data

    FTPServer.ProcFTPCommand Index, raw_data

End Sub

Private Sub DataSock_SendComplete(Index As Integer)

    DoEvents
    FTPServer.SendComplete Index

End Sub

Private Sub CommandSock_Close(Index As Integer)

    'This event may be called because the client has been logged out by the server.
    'There is a small piece of code in the LogoutClient routine
    'to catch this.
    DoEvents
    FTPServer.LogoutClient , Index

End Sub
