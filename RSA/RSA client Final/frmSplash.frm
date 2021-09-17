VERSION 5.00
Begin VB.Form frmSplash 
   BackColor       =   &H0082BEB4&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3915
   ClientLeft      =   255
   ClientTop       =   6540
   ClientWidth     =   6540
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmSplash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3915
   ScaleWidth      =   6540
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdclose 
      BackColor       =   &H0082BEB4&
      Caption         =   "Close"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5400
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   3480
      Width           =   1095
   End
   Begin VB.TextBox Text4 
      Height          =   495
      Left            =   240
      TabIndex        =   15
      Text            =   "Text4"
      Top             =   4800
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox Text3 
      Height          =   495
      Left            =   1560
      TabIndex        =   14
      Text            =   "Text3"
      Top             =   4800
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox txtdispKey 
      BackColor       =   &H00FFFFC0&
      Height          =   375
      Left            =   2160
      TabIndex        =   8
      Top             =   2880
      Width           =   4215
   End
   Begin VB.TextBox getFinKey 
      BackColor       =   &H00FFFFC0&
      Height          =   375
      Left            =   2160
      TabIndex        =   7
      Top             =   3960
      Visible         =   0   'False
      Width           =   4215
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0082BEB4&
      Caption         =   "&Enter"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5400
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1680
      Width           =   1095
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00FFFFC0&
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   2160
      MaxLength       =   20
      PasswordChar    =   "."
      TabIndex        =   3
      Top             =   1680
      Width           =   3015
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00FFFFC0&
      Height          =   375
      Left            =   2160
      MaxLength       =   20
      TabIndex        =   2
      Top             =   1080
      Width           =   3015
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H0082BEB4&
      Height          =   5490
      Left            =   0
      TabIndex        =   0
      Top             =   -120
      Width           =   7440
      Begin VB.CommandButton cmdGetYourKey 
         BackColor       =   &H0082BEB4&
         Caption         =   "&Get Your Key"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   600
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   2280
         Width           =   1335
      End
      Begin VB.TextBox txtTemp 
         Height          =   285
         Left            =   2880
         TabIndex        =   12
         Top             =   4920
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.TextBox txtFileName 
         BackColor       =   &H0082BEB4&
         Height          =   375
         Left            =   120
         TabIndex        =   11
         Top             =   4440
         Visible         =   0   'False
         Width           =   3495
      End
      Begin VB.Label Label4 
         BackColor       =   &H0082BEB4&
         Caption         =   "Your Finger Print"
         Height          =   315
         Left            =   720
         TabIndex        =   10
         Top             =   4080
         Width           =   1320
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H0082BEB4&
         Caption         =   "Your Key"
         Height          =   195
         Left            =   1320
         TabIndex        =   9
         Top             =   3000
         Width           =   645
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H0082BEB4&
         Caption         =   "Password"
         Height          =   195
         Index           =   0
         Left            =   1200
         TabIndex        =   6
         Top             =   1800
         Width           =   690
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H0082BEB4&
         Caption         =   "Username"
         Height          =   195
         Left            =   1200
         TabIndex        =   5
         Top             =   1200
         Width           =   720
      End
      Begin VB.Label lblWarning 
         AutoSize        =   -1  'True
         BackColor       =   &H0082BEB4&
         Caption         =   "RSA ENCRYPTION AND DECRYPTION"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   2280
         TabIndex        =   1
         Top             =   360
         Width           =   2715
      End
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim fnt, fn, lin As String
Dim dsize As Double
Dim i As Integer
Dim pdata(), pout, filepout, Fsave As String
Dim adata(100000) As Integer
Dim out(100000) As Double
Dim linefeed As String
Dim tryI, tryJ, jval As Long
Dim i1, a1 As Long
Dim GetRndA, GetRndB, indexA, indexB, RndValA, RndValB, primeA, primeB, opA, opB As Long
Dim chkA, chkB, tstA, tstB As Boolean
Dim cont As Boolean
Dim PRIME1, PRIME2, PROD, PHIE, PUBLICKEY, SECRETKEY, POS As Long
Dim CIPHER As String
Dim Y, X, N As Long
Dim store(9999), mtp, temp, flmt, disp(15), t, rvf(15), res, pcnt, suma(15) As Long
Dim pow(15), ch(14), a, c, il, cnt, j, jp, kt, lt, l, j1, ic, jc As Long
Dim op, t2 As String
Dim gt(15), rv(15), lent, chlen, dval As Integer
Dim gate, fin As Boolean
Dim fname As Variant



Private Sub cmdClose_Click()
End
End Sub

Private Sub cmdGetYourKey_Click()

Load frmdecrypt

End Sub



Private Sub Command1_Click()
'Form1.Show
  Dim FileName10 As String
  Dim FName_Only As String
  Dim lengthoftext1, lengthoftext2 As String, i, j, ab, cd, q, g As Integer
  
    lengthoftext1 = Len(Text1)
    i = 20 - lengthoftext1
    Text3 = Text1
    For j = 1 To i
        Text3 = Text3 + "_"
    Next j
 
    lengthoftext2 = Len(Text2)
    ab = 20 - lengthoftext2
    Text4 = Text2
    For cd = 1 To ab
        Text4 = Text4 + "_"
    Next cd
   
   getFinKey = "MainFormtxtKey"
   txtFileName = Text3 + Text4 + getFinKey
       'Generating Two Random Prime Numbers
         If cont = False Then
             While (chkA = False)
                 tstA = False
                 Randomize
                 GetRndA = Rnd() * 100
                 RndValA = Round(GetRndA, 1)
                 For indexA = 2 To RndValA - 1
                     primeA = RndValA Mod indexA
                     If (primeA = 0) Then
                         tstA = True
                     End If
                 Next indexA
                 If (tstA = False) Then
                     If RndValA <= 2 Then
                         'do nothing
                     Else
                         PRIME1 = Round(RndValA, 0)
                         chkA = True
                     End If
                 End If
             Wend
             While (chkB = False)
                 tstB = False
                 Randomize
                 GetRndB = Rnd() * 100
                 RndValB = Round(GetRndB, 1)
                 For indexB = 2 To RndValB - 1
                     primeB = RndValB Mod indexB
                     If (primeB = 0) Then
                         tstB = True
                     End If
                 Next indexB
                 If (tstB = False) Then
                     If RndValB <= 2 Then
                     Else
                         PRIME2 = Round(RndValB, 0)
                         chkB = True
                     End If
                 End If
             Wend
             If (PRIME1 = PRIME2) Then
                 'do nothing
             ElseIf (PRIME1 <= 2) Then
                 'do nothing
             ElseIf (PRIME2 <= 2) Then
                 'do nothing
             Else
                 cont = True
             End If
         End If
         
        'FINDING THE VALUE OF N
         PROD = PRIME1 * PRIME2
         
        'FINDING THE VALUE OF PHIE
         PHIE = (PRIME1 - 1) * (PRIME2 - 1)
         
         'FINDING THE PUBLIC KEY
         cont = False
         For i1 = 2 To (PHIE - 1)
            If cont = False Then
                a1 = PHIE Mod i1
                If a1 = 0 Then
                    'do nothing
                Else
                    PUBLICKEY = i1
                    cont = True
                End If
            End If
         Next i1
         
        'FINDING THE SECRET KEY
         cont = False
         For q = 1 To 100000
             If cont = False Then
                 If ((PUBLICKEY * q) Mod PHIE) = 1 Then
                     SECRETKEY = q
                     cont = True
                 End If
             End If
         Next q
  'GETTING THE CHARACTER ONE BY ONE FROM THE FILE
            lin = txtFileName.Text
            dsize = Len(lin)
            If dsize > 0 Then
                ReDim pdata(1 To dsize)
            Else
                ReDim pdata(i)
            End If
            linefeed = ""
            i = 1
            Do Until i = dsize + 1
                 pdata(i) = Mid(lin, i, 1)
                 'Print pdata(i)
                i = i + 1
            Loop
        
        
        'FINDING THE ASCII OF CHARACTER
      
        For i = 1 To dsize
           adata(i) = Asc(pdata(i))
        Next i
        
        'FINDING THE CIPHER TEXT
                  Open "C:\1_file.txt" For Output As #1
        For jval = 1 To dsize
            X = adata(jval)
            Y = PUBLICKEY
            N = PROD
            Powers
            CIPHER = res
    Print #1, CIPHER
   Next jval
    Close #1
    
    Open "C:\1_file.txt" For Append As #1
            Print #1, "," + CStr(SECRETKEY) + "," + CStr(PROD) + "end"
    Close #1
        
  FileName10 = FileDialog(Me, False, "File to Encrypt", "*.*|*.*", "C:\1_file.txt")

frmClient.txtFileName = FileName10

    
    If frmClient.txtFileName = "" Then
       MsgBox "No file selected to send...", vbCritical
    Else ' send the file, if connected
       If frmWsk.tcpClient.State <> sckClosed Then
          ' send only the file name because it will
          ' be stored in another area than the source
          FName_Only$ = GetFileName(frmClient.txtFileName)
          SendFile FName_Only$
       End If
    End If

End Sub

Private Sub Form_Load()
'    lblVersion.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision
 '   lblProductName.Caption = App.Title
    Load frmClient
End Sub


Private Sub Powers()
'FINDING THE BINARY OF Y
    t2 = ""
    a = Y
    Do While a >= 2
        c = a Mod 2
        a = Fix(a / 2)
        t2 = t2 & c
    Loop
    t2 = t2 & a
    op = StrReverse(t2)
    
    'FINDING THE LENGTH
     a = 1
    Do While (gate = False)
        c = Mid$(op, a, 1)
        gt(a) = c
        If (c <> 1 And c <> 0) Then
            gate = True
        End If
        cnt = a
        a = a + 1
    Loop
    lent = cnt - 1
    
    
    'FINDING THE REVERSE
    a = 1
    For c = lent To 1 Step -1
       rv(a) = gt(c)
       a = a + 1
    Next c
    
   'FINDING THE POWERS
   l = 2
   j = 1
   kt = 2
   pow(1) = rv(1)
   ch(1) = pow(1)
   lt = 2
   chlen = 1
    Do While (lt <= lent)
        pow(j + 1) = rv(kt) * 2 ^ j
        ch(l) = pow(j + 1)
        chlen = chlen + 1
        j = j + 1
        kt = kt + 1
        l = l + 1
        lt = lt + 1
    Loop
    
    'FINDING THE SQUARES OF X
    store(1) = X Mod N
    temp = store(1)
    disp(1) = temp
    mtp = 2
    flmt = 1
    t = 2
    Do While ((flmt / 2) <= Val(Y))
        store(mtp) = (temp ^ 2) Mod N
        temp = store(mtp)
        disp(t) = (store(mtp))
        mtp = mtp * 2
        flmt = mtp * 2
        t = t + 1
    Loop
    'ELIMINATIN ZEROS
     i1 = 1
    j1 = 2
    rvf(1) = pow(1)
    Do While (i1 < 15)
        If ch(i1) <> 0 Then
            rvf(j1) = ch(i1)
            j1 = j1 + 1
        End If
        i1 = i1 + 1
    Loop
    'CALLING VALUES
    pcnt = 1
    jc = 1
    For ic = 1 To chlen
        If ch(ic) <> 0 Then
            suma(jc) = (disp(ic))
            pcnt = pcnt + 1
            jc = jc + 1
        End If
    Next ic
    'RESULT
    res = 1
    Dim q1 As Integer
    res = (suma(1) * suma(2)) Mod N
    For q1 = 2 To (pcnt - 2)
        res = (res * suma(q1 + 1)) Mod N
    Next q1
End Sub

