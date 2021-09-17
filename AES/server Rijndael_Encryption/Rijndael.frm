VERSION 5.00
Begin VB.Form fRijndael 
   Caption         =   "Server"
   ClientHeight    =   1950
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5880
   LinkTopic       =   "Form1"
   ScaleHeight     =   1950
   ScaleWidth      =   5880
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   10
      Top             =   0
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text3 
      Height          =   495
      Left            =   3000
      TabIndex        =   9
      Text            =   "Finger Print Data"
      Top             =   1200
      Width           =   2415
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   240
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   8
      Text            =   "Rijndael.frx":0000
      Top             =   0
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CommandButton cmdSend 
      Caption         =   "&Send"
      Height          =   495
      Left            =   1680
      TabIndex        =   7
      Top             =   1200
      Width           =   972
   End
   Begin VB.CommandButton cmdFileEncrypt 
      BackColor       =   &H8000000D&
      Caption         =   "Encrypt File"
      Height          =   495
      Left            =   480
      TabIndex        =   6
      Top             =   1200
      Width           =   1095
   End
   Begin VB.ComboBox cboKeySize 
      Height          =   315
      Left            =   1800
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   3240
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.ComboBox cboBlockSize 
      Height          =   315
      Left            =   1800
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   2880
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.TextBox txtPassword 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   2160
      MaxLength       =   16
      PasswordChar    =   "*"
      TabIndex        =   0
      Text            =   "password"
      Top             =   600
      Width           =   3495
   End
   Begin VB.Label Label3 
      Caption         =   "Key Size:"
      Height          =   255
      Left            =   360
      TabIndex        =   5
      Top             =   3240
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label Label2 
      Caption         =   "Block Size:"
      Height          =   255
      Left            =   360
      TabIndex        =   4
      Top             =   2880
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Enter Key (16 Bytes)"
      Height          =   195
      Left            =   240
      TabIndex        =   3
      Top             =   600
      Width           =   1440
   End
End
Attribute VB_Name = "fRijndael"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_Rijndael As New cRijndael
'Used to display what the program is doing in the Form's caption
Public Property Let Status(TheStatus As String)
    If Len(TheStatus) = 0 Then
        Me.Caption = App.Title
    Else
        Me.Caption = App.Title & " - " & TheStatus
    End If
    Me.Refresh
End Property
'Reverse of HexDisplay.  Given a String containing Hex values, convert to byte array data()
'Returns number of bytes n in data(0 ... n-1)
Private Function HexDisplayRev(TheString As String, data() As Byte) As Long
    Dim i As Long
    Dim j As Long
    Dim c As Long
    Dim d As Long
    Dim n As Long
    Dim data2() As Byte

    n = 2 * Len(TheString)
    data2 = TheString

    ReDim data(n \ 4 - 1)

    d = 0
    i = 0
    j = 0
    Do While j < n
        c = data2(j)
        Select Case c
        Case 48 To 57    '"0" ... "9"
            If d = 0 Then   'high
                d = c
            Else            'low
                data(i) = (c - 48) Or ((d - 48) * 16&)
                i = i + 1
                d = 0
            End If
        Case 65 To 70   '"A" ... "F"
            If d = 0 Then   'high
                d = c - 7
            Else            'low
                data(i) = (c - 55) Or ((d - 48) * 16&)
                i = i + 1
                d = 0
            End If
        Case 97 To 102  '"a" ... "f"
            If d = 0 Then   'high
                d = c - 39
            Else            'low
                data(i) = (c - 87) Or ((d - 48) * 16&)
                i = i + 1
                d = 0
            End If
        End Select
        j = j + 2
    Loop
    n = i
    If n = 0 Then
        Erase data
    Else
        ReDim Preserve data(n - 1)
    End If
    HexDisplayRev = n

End Function

'Returns a byte array containing the password in the txtPassword TextBox control.
'If "Plaintext is hex" is checked, and the TextBox contains a Hex value the correct
'length for the current KeySize, the Hex value is used.  Otherwise, ASCII values
'of the txtPassword characters are used.
Private Function GetPassword() As Byte()
    Dim data() As Byte

        If HexDisplayRev(txtPassword.Text, data) <> (cboKeySize.ItemData(cboKeySize.NewIndex) \ 8) Then
            data = StrConv(txtPassword.Text, vbFromUnicode)
            ReDim Preserve data(31)
        End If
    
    GetPassword = data
End Function


Private Sub cmdFileEncrypt_Click()
    Dim FileName  As String
    Dim FileName1 As String
    Dim FileName2 As String
    Dim FileName3 As String
    Dim pass()    As Byte
    Dim KeyBits  As Long
    Dim BlockBits As Long
    Dim AESFile As String

    If Len(txtPassword.Text) = 0 Then
        MsgBox "No Password"
    Else
        FileName = FileDialog(Me, False, "File to Encrypt", "*.*|*.*")
        If Len(FileName) <> 0 Then
            FileName1 = FileDialog(Me, True, "Save Encrypted Data As ...", "*.aes|*.aes|*.*|*.*", FileName & ".aes")

            If Len(FileName1) <> 0 Then
       
                RidFile FileName1
                KeyBits = cboKeySize.ItemData(cboKeySize.NewIndex)
                BlockBits = cboBlockSize.ItemData(cboBlockSize.NewIndex)
                pass = GetPassword

                Status = "Encrypting File"
                
                m_Rijndael.SetCipherKey pass, KeyBits
               m_Rijndael.FileEncrypt FileName, FileName1
               

               Open FileName1 For Binary As #1
                    Text1 = Input$(LOF(1), 1)
                    Close #1

               Text2 = Text1 + "Biokey" + Text3
               
               Open "c:\AESFile" For Output As #1
        Print #1, Text2.Text
    Close #1
    FileName3 = "C:\filename3.txt"
               FileCopy "c:\AESFile", FileName3
                 Kill "c:\AESFile"
                 Kill FileName1
            
            FileName2 = FileDialog(Me, True, "Save Encrypted Data As ...", "*.aes.aes|*.aes.aes|*.aes*|*.aes*", FileName3 & ".aes.aes")

                RidFile FileName2
                KeyBits = cboKeySize.ItemData(cboKeySize.NewIndex)
                BlockBits = cboBlockSize.ItemData(cboBlockSize.NewIndex)
                pass = GetPassword

               Status = "Encrypting File"
                
               m_Rijndael.SetCipherKey pass, KeyBits
               m_Rijndael.FileEncrypt FileName3, FileName2
          
               frmServer.Slider1 = 4
               frmServer.txtFileName = FileName2
                Status = ""
            End If
        End If
    End If
End Sub

Private Sub cmdSend_Click()
 Dim FName_Only As String
    
    If frmServer.txtFileName = "" Then
       MsgBox "No file selected to send...", vbCritical
    Else ' send the file, if connected
       If frmWSK.tcpServer.State <> sckClosed Then
          ' send only the file name because it will
          ' be stored in another area than the source
          FName_Only$ = GetFileName(frmServer.txtFileName)
          SendFile FName_Only$
       End If
    End If
End Sub

Private Sub Form_Initialize()
    cboBlockSize.AddItem "128 Bit"
    cboBlockSize.ItemData(cboBlockSize.NewIndex) = 128
    cboBlockSize.Enabled = False
    
    cboKeySize.AddItem "128 Bit"
    cboKeySize.ItemData(cboKeySize.NewIndex) = 128
    cboKeySize.Enabled = False

    txtPassword = "1234567891234567"
    Status = ""
End Sub

Private Sub Form_Load()
Load frmServer
End Sub
