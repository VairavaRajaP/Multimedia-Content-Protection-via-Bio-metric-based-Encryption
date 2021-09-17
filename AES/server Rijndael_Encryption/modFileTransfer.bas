Attribute VB_Name = "modFileTransfer"
Option Explicit


Declare Function GetTickCount Lib "kernel32" () As Long
Public NextChunk As Boolean

Public Const Port = 1256                ' Port to listen on
Public MAX_CHUNK As Long  'Max buffer length

Public bInconnection     As Boolean     ' True if connected




' --- a function for pausing

Sub Pause(HowLong As Long)
    Dim u%, tick As Long
    tick = GetTickCount()
    
    Do
      u% = DoEvents
    Loop Until tick + HowLong < GetTickCount
End Sub

' --- SendFile() Function
'
' Sends a file from one computer to another via WinSock

Sub SendFile(Fname As String)
    Dim DataChunk As String


    
    ' open the file to be sent
    Open Fname$ For Binary As #1
       frmServer.ProgressBar1.Max = LOF(1)
        SendData "OpenFile," & Fname$ & "," & LOF(1)
    ' pause to give app time to get ready
    Pause 200
       Status "Transfering... "
       
        Do While Not EOF(1)
          ' get some of the file data
          DataChunk$ = Input(MAX_CHUNK, #1)
          ' send it to the server
          NextChunk = False
          SendData DataChunk$
          ' report status
  
                 Dim Timeout As Long
          Timeout = 0
          Do Until NextChunk = True Or Timeout = 300000
          DoEvents
          Timeout = Timeout + 1
          Loop
          
          If Timeout = 300000 Then Debug.Print "Timeout on file send"
          
          frmServer.ProgressBar1.Value = Seek(1) - 1
        Loop ' loop until all data is sent
        
        ' transfer done, notify the server to close the file
        SendData "CloseFile,"

        ' re-init byte counter and update status
        Status "Listening..... (Connected)"
   
    Close #1
            Kill Fname
End Sub

' --- send data function this is merely a better way to access
' the winsock "SendData" function. does it's own error
' checking

Sub SendData(sData As String)
    On Error GoTo ErrH

    Dim Timeout As Long
    
    frmWSK.tcpServer.SendData sData
    
    Do Until (frmWSK.tcpServer.State = 0) Or (Timeout < 10000)
        DoEvents
        Timeout = Timeout + 1
        If Timeout > 10000 Then Exit Do
    Loop
    
ErrH:
    Exit Sub
End Sub


' GetFileName()
'
' Extract the file name and extension only from
' the full path.

Function GetFileName(Fname As String) As String
    ' return the filename given the path
    Dim i As Integer
    Dim tempStr As String
    
    For i% = 1 To Len(Fname$)
       ' look for the "\"
       tempStr$ = Right$(Fname$, i%)
       
       If Left$(tempStr$, 1) = "\" Then
         GetFileName$ = Mid$(tempStr$, 2, Len(tempStr$))
         Exit Function
       End If
    Next i
End Function



' Status message procedure
Public Sub Status(Msg As String)
   frmServer.lblStatus = " Status : " & Msg$
End Sub



