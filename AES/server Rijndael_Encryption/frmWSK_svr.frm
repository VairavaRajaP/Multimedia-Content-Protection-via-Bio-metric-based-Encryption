VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmWSK 
   Caption         =   "Form1"
   ClientHeight    =   840
   ClientLeft      =   2730
   ClientTop       =   3555
   ClientWidth     =   1590
   LinkTopic       =   "Form1"
   ScaleHeight     =   840
   ScaleWidth      =   1590
   Begin MSWinsockLib.Winsock tcpServer 
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
End
Attribute VB_Name = "frmWSK"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    ' connect to the port
    tcpServer.LocalPort = Port
    ' Listen for incoming data
    tcpServer.Listen
    
    bInconnection = False
    
    Status "Listening.... (Not Connected)"
End Sub

Private Sub tcpServer_DataArrival(ByVal bytesTotal As Long)
    '
    Dim Command      As String
    Dim NewArrival   As String
    Dim Data         As String
    Dim FileSize    As String
    Static DataCnt   As Long
        Dim pk As String
   ' tcpServer.PeekData pk
    'Debug.Print pk
    tcpServer.GetData NewArrival$, vbString
    
    
    If NewArrival$ = "NextChunk" Then
    NextChunk = True
    Exit Sub
    End If
    ' Extract the command from the Left
    ' of the comma
    Command = Split(NewArrival$, ",")(0)

    
    ' execute according to command sent
    Select Case Command$
                  
        Case "OpenFile"  ' open the file
           Dim Fname As String
               ' extract the data being sent from the
             ' right of the comma
                Data$ = Split(NewArrival$, ",")(1)
                FileSize$ = Split(NewArrival$, ",")(2)
               frmServer.ProgressBar1.Max = CLng(FileSize$)
           ' the file name only should've been sent
           Fname$ = App.Path & "\" & Data$
           Open Fname$ For Binary As #1
           ' file now opened to recieve input
           Status "File opened.... " & Data$
                Status "Recieving Data... "
        Case "CloseFile" ' close the file
           ' all data has been sent, close the file
           Close #1
           Status "File Transfer complete..."
           Pause 3000
           Status "Listening... (Connected)"
        
        Case Else ' a 4169 byte string of incoming data
           ' write the incoming chunk of data to the
           ' opened file
           Put #1, , NewArrival$
           
           SendData "NextChunk"

frmServer.ProgressBar1.Value = Seek(1) - 1
          
        
    End Select
    
End Sub
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
'  end  G E N E R A L   W I N S O C K   P R O C W D U R E S  end
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\








'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
'   G E N E R A L   W I N S O C K   P R O C W D U R E S
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
Private Sub tcpServer_Close()
    '
    'Socket got a close call so close it if it's not already closed
    If tcpServer.State <> sckClosed Then tcpServer.Close
    Form_Load      ' resume listening
    
End Sub


Private Sub tcpServer_ConnectionRequest(ByVal requestID As Long)
    '
     On Error GoTo IDERROR
     If tcpServer.State <> sckClosed Then tcpServer.Close ' close Connection
     tcpServer.Accept requestID    'Make the connection
     
     bInconnection = True
     Status "Listening... Connected."
     SendData "Accepted,"
     Exit Sub
     
IDERROR:
     MsgBox Err.Description, vbCritical
End Sub

