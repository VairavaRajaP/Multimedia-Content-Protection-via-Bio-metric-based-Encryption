Attribute VB_Name = "modFileTransferClient"
Option Explicit

Declare Function GetTickCount Lib "kernel32" () As Long
Public NextChunk As Boolean
Public MAX_CHUNK As Long 'Max Buffer size/length in bytes
Public Const Port = 6000                ' Port
Public bReplied          As Boolean     ' True if server replied
Public lTIme             As Long        ' varible to track connection time.




' --- SendFile() Function
'
' Sends a file from one computer to another via WinSock

Sub SendFile(Fname As String)
    Dim DataChunk As String
    Dim passes As Long
    
    '
    ' send over the filename so the Server knows where
    ' to store the file.

    
    ' open the file to be sent
    Open Fname$ For Binary As #1 ' this mode works well with any file
       Status "Transfering... "
           SendData "OpenFile," & Fname$ & "," & LOF(1)
    ' pause to give Server time to open
    Pause 200
       frmClient.ProgressBar1.Max = LOF(1)
       
        Do While Not EOF(1)

          DataChunk$ = Input(MAX_CHUNK, #1)
          ' send it to the server
          NextChunk = False
          SendData DataChunk$
          ' report status
          

          ' information
          Dim Timeout As Long
          Timeout = 0
          Do Until NextChunk = True Or Timeout = 300000
          DoEvents
          Timeout = Timeout + 1
          Loop
          
          If Timeout = 300000 Then Debug.Print "Timeout on file send"
          
          frmClient.ProgressBar1.Value = Seek(1) - 1
          
        Loop ' loop until all data is sent
        
        ' transfer done, notify the server to close the file
        SendData "CloseFile,"
        ' re-init byte counter and update status
        Status "Connected."
        passes& = 0
    Close #1
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
   frmClient.lblStatus = " Status : " & Msg$
End Sub


'--- SendData() This function merely sends the data to the Server and handles
'--- it's own Errors.
Function SendData(sData As String) As Boolean
    On Error GoTo ErrH
    Dim Timeout As Long
    
    ' no reply.... nothing sent yet
    bReplied = False
    ' send data
    frmWsk.tcpClient.SendData sData
    
    ' check for timeout or closed socket
    Do Until (frmWsk.tcpClient.State = 0) Or (Timeout < 100000)
        DoEvents
        Timeout = Timeout + 1
        If Timeout > 100000 Then Exit Do
    Loop
    ' ok.... sent
    SendData = True
    Exit Function
    
ErrH:
    SendData = False
   Debug.Print Err.Description, 16, "Error #" & Err.Number
    Status "Disconnected."
End Function

' --- a function for pausing, the same effect can be obtained
' using the sleep API function

Sub Pause(HowLong As Long)
    Dim u%, tick As Long
    tick = GetTickCount()
    
    Do
      u% = DoEvents
    Loop Until tick + HowLong < GetTickCount
End Sub
