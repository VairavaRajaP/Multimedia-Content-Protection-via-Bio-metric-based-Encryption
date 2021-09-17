VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmWsk 
   Caption         =   "Form1"
   ClientHeight    =   480
   ClientLeft      =   6555
   ClientTop       =   4005
   ClientWidth     =   1620
   LinkTopic       =   "Form1"
   ScaleHeight     =   480
   ScaleWidth      =   1620
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   840
      TabIndex        =   0
      Top             =   0
      Width           =   495
   End
   Begin MSWinsockLib.Winsock tcpClient 
      Left            =   0
      Top             =   15
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
End
Attribute VB_Name = "frmWsk"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
   Status "Disconnected."
    bReplied = False
End Sub

Private Sub tcpClient_DataArrival(ByVal bytesTotal As Long)
    
    Dim Command      As String
    Dim NewArrival   As String
    Dim Data         As String
    Dim FileSize    As String
    Static DataCnt   As Long
    Dim pk As String
    'tcpClient.PeekData pk
    'Debug.Print pk
  
    tcpClient.GetData NewArrival$, vbString
    
    If NewArrival$ = "NextChunk" Then
    NextChunk = True
    Exit Sub
    End If
    
    ' Extract the command from the Left
    ' of the comma
    Command$ = Split(NewArrival$, ",")(0)

    
    ' execute according to command sent
    Select Case Command
        Case "Accepted"          ' server accepted connection
             bReplied = True
             Status "Connected."
             
             ' this is a good practice.
             ' when the server has been closed
             ' theclient is notified here.
             ' and immediatley disconnected.
        Case "ServerClosed"
             Form_Load
             tcpClient.Close
             
        Case "OpenFile"  ' open the file
           Dim Fname As String
               ' extract the data being sent from the
               ' right of the comma
               Data$ = Split(NewArrival$, ",")(1)
               FileSize$ = Split(NewArrival$, ",")(2)
               
               frmClient.ProgressBar1.Max = CLng(FileSize$)
           ' the file name only should've been sent
           Fname$ = App.Path & "\" & Data$
           Open Fname$ For Binary As #1
           ' file now opened to recieve input
           Status "File opened.... " & Data$
               
        Case "CloseFile" ' close the file
           ' all data has been sent, close the file
           Close #1
           Status "File Transfer complete..."
           Pause 3000
           Status "Connected."
                
       ' when sending a file.... it is best not to Name
       ' the Case instead use ELse for file transfer
            
        Case Else
           ' write the incoming chunk of data to the
           ' opened file
           Put #1, , NewArrival$
           
            Open "C:\6_rxfile_password.txt" For Output As #5
                Print #5, NewArrival$
            Close #5
           
           
           frmClient.ProgressBar1.Value = Seek(1) - 1
           ' update the view port with the new addition
' ** // ** '
' IMPORTANT: comment out the code below when sending files
' larger than 500Kb. It makes the function CRAWL otherwise
              
           'txtView = txtView & NewArrival$
' comment the above line to increase speed

           
        frmWsk.tcpClient.SendData "NextChunk"
           ' count and report the incoming chunks
      
           Status "Recieving Data... " '& (MAX_CHUNK * DataCnt&) & " bytes"
            
    End Select
    Text1.Text = "C:\6_rxfile_password.txt"
End Sub
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
'  end  G E N E R A L   W I N S O C K   P R O C W D U R E S  end
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\








'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
'   G E N E R A L   W I N S O C K   P R O C W D U R E S
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
Private Sub tcpClient_Close()
    '
    'Socket got a close call so close it if it's not already closed
    If tcpClient.State <> sckClosed Then tcpClient.Close
    
End Sub


