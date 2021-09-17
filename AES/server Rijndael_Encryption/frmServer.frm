VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form frmServer 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "File Transfer (Server)"
   ClientHeight    =   1950
   ClientLeft      =   5385
   ClientTop       =   4410
   ClientWidth     =   5400
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmServer.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1950
   ScaleWidth      =   5400
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   360
      Left            =   75
      TabIndex        =   4
      Top             =   1290
      Width           =   5280
      _ExtentX        =   9313
      _ExtentY        =   635
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.TextBox txtFileName 
      Height          =   300
      Left            =   60
      TabIndex        =   2
      Top             =   285
      Width           =   3960
   End
   Begin MSComDlg.CommonDialog cdOpen 
      Left            =   4800
      Top             =   120
      _ExtentX        =   688
      _ExtentY        =   688
      _Version        =   393216
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Height          =   252
      Left            =   4125
      TabIndex        =   0
      Top             =   630
      Width           =   972
   End
   Begin MSComctlLib.Slider Slider1 
      Height          =   315
      Left            =   60
      TabIndex        =   5
      Top             =   915
      Width           =   5040
      _ExtentX        =   8890
      _ExtentY        =   556
      _Version        =   393216
      Min             =   1
      SelStart        =   1
      TickStyle       =   1
      Value           =   1
   End
   Begin VB.Label lblbuff 
      BackStyle       =   0  'Transparent
      Caption         =   "Buffer lenght"
      Height          =   240
      Left            =   75
      TabIndex        =   6
      Top             =   720
      Width           =   1425
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "File to send:"
      Height          =   210
      Index           =   0
      Left            =   60
      TabIndex        =   3
      Top             =   45
      Width           =   840
   End
   Begin VB.Label lblStatus 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Status : Listening......."
      Height          =   255
      Left            =   60
      TabIndex        =   1
      Top             =   1680
      Width           =   5295
   End
End
Attribute VB_Name = "frmServer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdClose_Click()
    Unload Me
End Sub



'    Private Sub cmdSend_Click()
'        Dim FName_Only As String
'
'        If txtFileName = "" Then
'           MsgBox "No file selected to send...", vbCritical
'        Else ' send the file, if connected
'           If frmWSK.tcpServer.State <> sckClosed Then
'              ' send only the file name because it will
'              ' be stored in another area than the source
'              FName_Only$ = GetFileName(txtFileName)
'              SendFile FName_Only$
'           End If
'        End If
'    End Sub

Private Sub Form_Load()
Load frmWSK
Slider1.Value = 4
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ' alert the client the server has been disconnected
    On Error Resume Next
    SendData "ServerClosed,"
    Pause 500
    frmWSK.tcpServer.Close
    Dim frm As Form

For Each frm In Forms
Unload frm
Set frm = Nothing
Next
End
End Sub


'Private Sub cmdBrowse_Click()
'    ' show the Open Dialog for the user to select a file.
'    cdOpen.ShowOpen
'
'    If Not vbCancel Then
'       txtFileName = cdOpen.FileName
'    End If
'
'End Sub

Private Sub Slider1_Change()
Slider1_Scroll
End Sub

Private Sub Slider1_Scroll()
MAX_CHUNK = Slider1.Value * 1024


lblbuff.Caption = "Buffer=" & Slider1.Value * 1024 & " bytes"
End Sub

