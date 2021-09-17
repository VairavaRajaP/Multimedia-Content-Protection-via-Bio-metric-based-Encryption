VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form frmClient 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "File Transfer (Client)"
   ClientHeight    =   2370
   ClientLeft      =   4335
   ClientTop       =   4260
   ClientWidth     =   5265
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmClient.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2370
   ScaleWidth      =   5265
   Begin MSComctlLib.Slider Slider1 
      Height          =   330
      Left            =   150
      TabIndex        =   11
      Top             =   1365
      Width           =   5040
      _ExtentX        =   8890
      _ExtentY        =   582
      _Version        =   393216
      Min             =   1
      SelStart        =   1
      TickStyle       =   1
      Value           =   1
   End
   Begin VB.TextBox Text1 
      Height          =   330
      Left            =   3150
      TabIndex        =   9
      Text            =   "127.0.0.1"
      Top             =   945
      Width           =   2070
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   360
      Left            =   0
      TabIndex        =   8
      Top             =   1725
      Width           =   5265
      _ExtentX        =   9287
      _ExtentY        =   635
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.CommandButton cmdConnect 
      Caption         =   "Connect"
      Height          =   252
      Left            =   135
      TabIndex        =   7
      Top             =   645
      Width           =   972
   End
   Begin VB.CommandButton cmdDisconnect 
      Caption         =   "Disconnect"
      Height          =   252
      Left            =   3135
      TabIndex        =   6
      Top             =   660
      Width           =   972
   End
   Begin VB.CommandButton cmdSend 
      Caption         =   "&Send"
      Height          =   252
      Left            =   1185
      TabIndex        =   5
      Top             =   645
      Width           =   972
   End
   Begin VB.CommandButton cmdBrowse 
      Caption         =   "&Browse..."
      Height          =   252
      Left            =   4170
      TabIndex        =   4
      Top             =   360
      Width           =   972
   End
   Begin VB.TextBox txtFileName 
      Height          =   300
      Left            =   120
      TabIndex        =   2
      Top             =   315
      Width           =   3990
   End
   Begin MSComDlg.CommonDialog cdOpen 
      Left            =   990
      Top             =   720
      _ExtentX        =   688
      _ExtentY        =   688
      _Version        =   393216
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Height          =   252
      Left            =   4170
      TabIndex        =   0
      Top             =   660
      Width           =   972
   End
   Begin VB.Label lblbuff 
      BackStyle       =   0  'Transparent
      Caption         =   "Buffer lenght"
      Height          =   240
      Left            =   135
      TabIndex        =   12
      Top             =   1155
      Width           =   1425
   End
   Begin VB.Label Label1 
      Caption         =   "IP Address:"
      Height          =   195
      Left            =   2250
      TabIndex        =   10
      Top             =   1005
      Width           =   885
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "File to send:"
      Height          =   210
      Index           =   0
      Left            =   120
      TabIndex        =   3
      Top             =   75
      Width           =   840
   End
   Begin VB.Label lblStatus 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Status : Disconnected"
      Height          =   255
      Left            =   -15
      TabIndex        =   1
      Top             =   2100
      Width           =   5280
   End
End
Attribute VB_Name = "frmClient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdClose_Click()
    End
End Sub

'Private Sub cmdConnect_Click()
'
'    'try to make a connection to the Server.
'    bReplied = False
'    frmWsk.tcpClient.Connect Text1.Text, 1256
'
'    lTIme = 0
'
'    While (Not bReplied) And (lTIme < 100000)
'        DoEvents
'        lTIme = lTIme + 1
'    Wend
'
'
'    If lTIme >= 100000 Then
'        'Didn't reply or timed out. close the connection
'        MsgBox "Unable to connect to remote server", vbCritical, "Connection Error"
'
'        frmWsk.tcpClient.Close
'        Exit Sub
'    End If
'
'End Sub



Private Sub cmdDisconnect_Click()
    frmWsk.tcpClient.Close
    Form_Load
End Sub

Private Sub cmdSend_Click()
    Dim FName_Only As String
    
    If txtFileName = "" Then
       MsgBox "No file selected to send...", vbCritical
    Else ' send the file, if connected
       If frmWsk.tcpClient.State <> sckClosed Then
          ' send only the file name because it will
          ' be stored in another area than the source
          FName_Only$ = GetFileName(txtFileName)
          SendFile FName_Only$
       End If
    End If
End Sub

Private Sub Form_Load()
       
    'try to make a connection to the Server.
    bReplied = False
    frmWsk.tcpClient.Connect Text1.Text, 1256
    
    lTIme = 0
    
    While (Not bReplied) And (lTIme < 100000)
        DoEvents
        lTIme = lTIme + 1
    Wend
    
    
    If lTIme >= 100000 Then
        'Didn't reply or timed out. close the connection
        MsgBox "Unable to connect to remote server", vbCritical, "Connection Error"
        
        frmWsk.tcpClient.Close
        Exit Sub
    End If
Load frmWsk
Slider1.Value = 4
End Sub



Private Sub cmdBrowse_Click()
    ' show the Open Dialog for the user to select a file.
    cdOpen.ShowOpen
    
    If Not vbCancel Then
       txtFileName = cdOpen.FileName
    End If
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim frm As Form

For Each frm In Forms
Unload frm
Set frm = Nothing
Next
End
End Sub



Private Sub Slider1_Change()
Slider1_Scroll
End Sub

Private Sub Slider1_Scroll()
MAX_CHUNK = Slider1.Value * 1024


lblbuff.Caption = "Buffer=" & Slider1.Value * 1024 & " bytes"
End Sub


