VERSION 5.00
Begin VB.Form fRijndael 
   Caption         =   "Form1"
   ClientHeight    =   1965
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6180
   LinkTopic       =   "Form1"
   ScaleHeight     =   1965
   ScaleWidth      =   6180
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdFileDecrypt 
      Caption         =   "Decrypt File"
      Height          =   495
      Left            =   2160
      TabIndex        =   6
      Top             =   1080
      Width           =   1095
   End
   Begin VB.ComboBox cboKeySize 
      Height          =   315
      Left            =   2160
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   2280
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.ComboBox cboBlockSize 
      Height          =   315
      Left            =   2160
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   1920
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
      Text            =   "Password"
      Top             =   600
      Width           =   3495
   End
   Begin VB.Label Label3 
      Caption         =   "Key Size:"
      Height          =   255
      Left            =   720
      TabIndex        =   5
      Top             =   2280
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label Label2 
      Caption         =   "Block Size:"
      Height          =   255
      Left            =   720
      TabIndex        =   4
      Top             =   1920
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "Enter Key (16 Bytes)"
      Height          =   255
      Left            =   480
      TabIndex        =   3
      Top             =   600
      Width           =   1455
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
Private Function HexDisplayRev(TheString As String, Data() As Byte) As Long
    Dim i As Long
    Dim j As Long
    Dim c As Long
    Dim d As Long
    Dim n As Long
    Dim data2() As Byte

    n = 2 * Len(TheString)
    data2 = TheString

    ReDim Data(n \ 4 - 1)

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
                Data(i) = (c - 48) Or ((d - 48) * 16&)
                i = i + 1
                d = 0
            End If
        Case 65 To 70   '"A" ... "F"
            If d = 0 Then   'high
                d = c - 7
            Else            'low
                Data(i) = (c - 55) Or ((d - 48) * 16&)
                i = i + 1
                d = 0
            End If
        Case 97 To 102  '"a" ... "f"
            If d = 0 Then   'high
                d = c - 39
            Else            'low
                Data(i) = (c - 87) Or ((d - 48) * 16&)
                i = i + 1
                d = 0
            End If
        End Select
        j = j + 2
    Loop
    n = i
    If n = 0 Then
        Erase Data
    Else
        ReDim Preserve Data(n - 1)
    End If
    HexDisplayRev = n
End Function


'Returns a byte array containing the password in the txtPassword TextBox control.
'If "Plaintext is hex" is checked, and the TextBox contains a Hex value the correct
'length for the current KeySize, the Hex value is used.  Otherwise, ASCII values
'of the txtPassword characters are used.
Private Function GetPassword() As Byte()
    Dim Data() As Byte

        If HexDisplayRev(txtPassword.Text, Data) <> (cboKeySize.ItemData(cboKeySize.NewIndex) \ 8) Then
            Data = StrConv(txtPassword.Text, vbFromUnicode)
            ReDim Preserve Data(31)
        End If
    GetPassword = Data
End Function

Private Sub cmdFileDecrypt_Click()
    Dim FileName  As String
    Dim FileName2 As String
    Dim FileName4 As String
    Dim pass()    As Byte
    Dim KeyBits   As Long
    Dim BlockBits As Long
    Dim a, b, ext, fingerprintdata As String
    Dim filename2data As String
    

    If Len(txtPassword.Text) = 0 Then
        MsgBox "No Password"
    Else
        'FileName = FileDialog(Me, False, "File to Decrypt", "*.aes|*.aes|*.*|*.*")
        FileName = frmWsk.Text1
        a = frmWsk.Text1
        b = InStr(1, a, ".aes") - InStr(1, a, ".")
        ext = Mid(a, InStr(1, a, ".") + 1, b - 1)
        If Len(FileName) <> 0 Then
            If InStrRev(FileName, ".aes") = Len(FileName) - 3 Then FileName2 = Left$(FileName, Len(FileName) - 4)
            FileName2 = FileDialog(Me, True, "Save Decrypted Data As ...", "*.ext|*.ext", FileName2)
            If Len(FileName2) <> 0 Then
                RidFile FileName2
                KeyBits = cboKeySize.ItemData(cboKeySize.NewIndex)
                BlockBits = cboBlockSize.ItemData(cboBlockSize.NewIndex)
                pass = GetPassword

                Status = "Decrypting File"
               m_Rijndael.SetCipherKey pass, KeyBits
              m_Rijndael.FileDecrypt FileName2, FileName
              Kill FileName
              
              Open FileName2 For Binary As #1
                        filename2data = Input(LOF(1), 1)
            Close #1
              
                fingerprintdata = Mid(filename2data, InStr(1, filename2data, "Biokey") + 6)
                filename2data = CStr(Left(filename2data, InStr(1, filename2data, "Biokey") - 1))
                Open FileName2 For Output As 2
                Print #2, filename2data
                Close #2
                
              
              FileName4 = FileDialog(Me, True, "Save Decrypted Data As ...", "*.ext|*.ext", FileName4)
                RidFile FileName4
                KeyBits = cboKeySize.ItemData(cboKeySize.NewIndex)
                BlockBits = cboBlockSize.ItemData(cboBlockSize.NewIndex)
                pass = GetPassword

                Status = "Decrypting File"
               m_Rijndael.SetCipherKey pass, KeyBits
              m_Rijndael.FileDecrypt FileName4, FileName2



               Status = ""
            End If
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
Load frmClient
End Sub
