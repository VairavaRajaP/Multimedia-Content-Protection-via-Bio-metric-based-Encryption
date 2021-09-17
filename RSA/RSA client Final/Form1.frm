VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H0082BEB4&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "RSA Encryption and Decryption"
   ClientHeight    =   6030
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8640
   ForeColor       =   &H00000080&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6030
   ScaleWidth      =   8640
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      BackColor       =   &H0082BEB4&
      Height          =   855
      Left            =   600
      TabIndex        =   6
      Top             =   600
      Width           =   4335
      Begin VB.CommandButton cmdDecrypt 
         BackColor       =   &H0082BEB4&
         Caption         =   "&Decrypt"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2760
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H0082BEB4&
      Height          =   615
      Left            =   240
      TabIndex        =   4
      Top             =   1320
      Visible         =   0   'False
      Width           =   3495
      Begin VB.Label Label2 
         BackColor       =   &H0082BEB4&
         Caption         =   "The Selected File:    "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   360
         TabIndex        =   5
         Top             =   5040
         Width           =   1575
      End
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   2640
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Top             =   2160
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   1920
      TabIndex        =   2
      Top             =   2160
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   600
      TabIndex        =   1
      Top             =   2160
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   195
      Left            =   600
      TabIndex        =   0
      Top             =   2760
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Timer Timer2 
      Interval        =   1000
      Left            =   1080
      Top             =   2520
   End
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   1560
      Top             =   2520
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Dim fnt, fn, lin As String
'Dim dsize As Double
'Dim i As Integer
'Dim pdata(), pout, filepout, Fsave As String
'Dim adata(100000) As Integer
'Dim out(100000) As Double
'Dim linefeed As String
'Dim tryI, tryJ, jval As Long
'Dim i1, a1 As Long
'Dim GetRndA, GetRndB, indexA, indexB, RndValA, RndValB, primeA, primeB, opA, opB As Long
'Dim chkA, chkB, tstA, tstB As Boolean
'Dim cont As Boolean
'Dim PRIME1, PRIME2, PROD, PHIE, PUBLICKEY, SECRETKEY, POS As Long
'Dim CIPHER As String
'Dim Y, X, N As Long
'Dim store(9999), mtp, temp, flmt, disp(15), t, rvf(15), res, pcnt, suma(15) As Long
'Dim pow(15), ch(14), a, c, il, cnt, j, jp, kt, lt, l, j1, ic, jc As Long
'Dim op, t2 As String
'Dim gt(15), rv(15), lent, chlen, dval As Integer
'Dim gate, fin As Boolean
'Dim fname As Variant
Private Sub Command2_Click()
    Form2.Show
    Unload Me
    Dim d As String
    MsgBox ("")
    d = (UCase(Right(Label6.Caption, 3)))
    MsgBox (d)
End Sub
Private Sub Command3_Click()
    End
End Sub


'Private Sub cmdDecrypt_Click()
'txtTemp.Text = "c:\file.txt"
'fn = txtTemp.Text
'
'    SECRETKEY = InputBox("Enter The Secret Key", "Secret Key")
'    PROD = InputBox("Enter The Value Of Phi", "Phi")
'    Open fn For Input As #1
'        While Not EOF(1)
'            Line Input #1, lin
'            dsize = Len(lin)
'            If dsize > 0 Then
'                ReDim pdata(1 To dsize)
'            Else
'                ReDim pdata(1)
'            End If
'            rlinefeed = ""
'            j = ""
'            i = 1
'            Do Until i = dsize + 1
'                j = Mid(lin, i, 1)
'                rlinefeed = rlinefeed & j
'                i = i + 1
'            Loop
'            X = rlinefeed
'            Y = SECRETKEY
'            N = PROD
'
'           'FINDING THE BINARY OF Y
'              t2 = ""
'              a = Y
'              Do While a >= 2
'                  c = a Mod 2
'                  a = Fix(a / 2)
'                  t2 = t2 & c
'              Loop
'              t2 = t2 & a
'              op = StrReverse(t2)
'
'            'FINDING THE LENGTH
'             a = 1
'            Do While (gate = False)
'                c = Mid$(op, a, 1)
'                gt(a) = c
'                If (c <> 1 And c <> 0) Then
'                    gate = True
'                End If
'                cnt = a
'                a = a + 1
'            Loop
'            lent = cnt - 1
'
'            'FINDING THE REVERSE
'            a = 1
'            For c = lent To 1 Step -1
'               rv(a) = gt(c)
'               a = a + 1
'            Next c
'
'            'FINDING THE POWERS
'            l = 2
'            j = 1
'            kt = 2
'            pow(1) = rv(1)
'            ch(1) = pow(1)
'            lt = 2
'            chlen = 1
'            Do While (lt <= lent)
'                 pow(j + 1) = rv(kt) * 2 ^ j
'                 ch(l) = pow(j + 1)
'                 chlen = chlen + 1
'                 j = j + 1
'                 kt = kt + 1
'                 l = l + 1
'                 lt = lt + 1
'           Loop
'
'           'FINDING THE SQUARES OF X
'            store(1) = X Mod N
'            temp = store(1)
'            disp(1) = temp
'            mtp = 2
'            flmt = 1
'            t = 2
'            Do While ((flmt / 2) <= Val(Y))
'                store(mtp) = (temp ^ 2) Mod N
'                temp = store(mtp)
'                disp(t) = (store(mtp))
'                mtp = mtp * 2
'                flmt = mtp * 2
'                t = t + 1
'            Loop
'
'            'ELIMINATIN ZEROS
'             i1 = 1
'            j1 = 2
'            rvf(1) = pow(1)
'            Do While (i1 < 15)
'                If ch(i1) <> 0 Then
'                    rvf(j1) = ch(i1)
'                    j1 = j1 + 1
'                End If
'                i1 = i1 + 1
'            Loop
'
'            'CALLING VALUES
'            pcnt = 1
'            jc = 1
'            For ic = 1 To chlen
'                If ch(ic) <> 0 Then
'                    suma(jc) = (disp(ic))
'                    pcnt = pcnt + 1
'                    jc = jc + 1
'                End If
'            Next ic
'
'            'RESULT
'            res = 1
'            Dim q1 As Integer
'            res = (suma(1) * suma(2)) Mod N
'            For q1 = 2 To (pcnt - 2)
'                res = (res * suma(q1 + 1)) Mod N
'            Next q1
'            pout = Chr(res)
'            Dim loc As Integer
'            filepout = filepout & pout
'            X = Y = N = res = 0
'        Wend
'         Close #1
'        Open "c:\defile.txt" For Append As #2
'                Print #2, filepout
'        Close #2
'
'    MsgBox ("Decryption Is Completed Successfully")
'    End
'End Sub
    




'Private Sub cmdEncrypt_Click()
'
'  Dim lengthoftext1, lengthoftext2 As String, i, j, ab, cd As Integer
'
'    lengthoftext1 = Len(frmSplash.Text1)
'    i = 20 - lengthoftext1
'    For j = 1 To i
'        frmSplash.Text1 = frmSplash.Text1 + "0"
'    Next j
'
'    lengthoftext2 = Len(frmSplash.Text2)
'    ab = 20 - lengthoftext2
'    For cd = 1 To ab
'        frmSplash.Text2 = frmSplash.Text2 + "0"
'    Next cd
'
'   frmSplash.Text3 = MainForm.txtKey
'   txtFileName = frmSplash.Text1 + frmSplash.Text2 + frmSplash.Text3
'       'Generating Two Random Prime Numbers
'         If cont = False Then
'             While (chkA = False)
'                 tstA = False
'                 Randomize
'                 GetRndA = Rnd() * 100
'                 RndValA = Round(GetRndA, 1)
'                 For indexA = 2 To RndValA - 1
'                     primeA = RndValA Mod indexA
'                     If (primeA = 0) Then
'                         tstA = True
'                     End If
'                 Next indexA
'                 If (tstA = False) Then
'                     If RndValA <= 2 Then
'                         'do nothing
'                     Else
'                         PRIME1 = Round(RndValA, 0)
'                         chkA = True
'                     End If
'                 End If
'             Wend
'             While (chkB = False)
'                 tstB = False
'                 Randomize
'                 GetRndB = Rnd() * 100
'                 RndValB = Round(GetRndB, 1)
'                 For indexB = 2 To RndValB - 1
'                     primeB = RndValB Mod indexB
'                     If (primeB = 0) Then
'                         tstB = True
'                     End If
'                 Next indexB
'                 If (tstB = False) Then
'                     If RndValB <= 2 Then
'                     Else
'                         PRIME2 = Round(RndValB, 0)
'                         chkB = True
'                     End If
'                 End If
'             Wend
'             If (PRIME1 = PRIME2) Then
'                 'do nothing
'             ElseIf (PRIME1 <= 2) Then
'                 'do nothing
'             ElseIf (PRIME2 <= 2) Then
'                 'do nothing
'             Else
'                 cont = True
'             End If
'         End If
'
'        'FINDING THE VALUE OF N
'         PROD = PRIME1 * PRIME2
'
'        'FINDING THE VALUE OF PHIE
'         PHIE = (PRIME1 - 1) * (PRIME2 - 1)
'
'         'FINDING THE PUBLIC KEY
'         cont = False
'         For i1 = 2 To (PHIE - 1)
'            If cont = False Then
'                a1 = PHIE Mod i1
'                If a1 = 0 Then
'                    'do nothing
'                Else
'                    PUBLICKEY = i1
'                    cont = True
'                End If
'            End If
'         Next i1
'
'        'FINDING THE SECRET KEY
'         cont = False
'         For q = 1 To 100000
'             If cont = False Then
'                 If ((PUBLICKEY * q) Mod PHIE) = 1 Then
'                     SECRETKEY = q
'                     cont = True
'                 End If
'             End If
'         Next q
'  'GETTING THE CHARACTER ONE BY ONE FROM THE FILE
'            lin = txtFileName.Text
'            dsize = Len(lin)
'            If dsize > 0 Then
'                ReDim pdata(1 To dsize)
'            Else
'                ReDim pdata(i)
'            End If
'            linefeed = ""
'            i = 1
'            Do Until i = dsize + 1
'                 pdata(i) = Mid(lin, i, 1)
'                 'Print pdata(i)
'                i = i + 1
'            Loop
'
'
'        'FINDING THE ASCII OF CHARACTER
'
'        For i = 1 To dsize
'           adata(i) = Asc(pdata(i))
'        Next i
'
'        'FINDING THE CIPHER TEXT
'                  Open "C:\file.txt" For Output As #1
'        For jval = 1 To dsize
'            X = adata(jval)
'            Y = PUBLICKEY
'            N = PROD
'            Powers
'            CIPHER = res
'    Print #1, CIPHER
'   Next jval
'
'    Close #1
'
'    MsgBox "File Saved"
'
'      g = MsgBox("Encryption Is Complete", vbInformation, "Encrption")
'      Form3.lblSecret.Caption = SECRETKEY
'      Form3.lblN.Caption = PROD
'      Unload Me
'      Form3.Show
'End Sub
 
 Private Sub Form_Load()
    Dim g As Integer
    txtTemp.Visible = False
End Sub

'Private Sub Powers()
''FINDING THE BINARY OF Y
'    t2 = ""
'    a = Y
'    Do While a >= 2
'        c = a Mod 2
'        a = Fix(a / 2)
'        t2 = t2 & c
'    Loop
'    t2 = t2 & a
'    op = StrReverse(t2)
'
'    'FINDING THE LENGTH
'     a = 1
'    Do While (gate = False)
'        c = Mid$(op, a, 1)
'        gt(a) = c
'        If (c <> 1 And c <> 0) Then
'            gate = True
'        End If
'        cnt = a
'        a = a + 1
'    Loop
'    lent = cnt - 1
'
'
'    'FINDING THE REVERSE
'    a = 1
'    For c = lent To 1 Step -1
'       rv(a) = gt(c)
'       a = a + 1
'    Next c
'
'   'FINDING THE POWERS
'   l = 2
'   j = 1
'   kt = 2
'   pow(1) = rv(1)
'   ch(1) = pow(1)
'   lt = 2
'   chlen = 1
'    Do While (lt <= lent)
'        pow(j + 1) = rv(kt) * 2 ^ j
'        ch(l) = pow(j + 1)
'        chlen = chlen + 1
'        j = j + 1
'        kt = kt + 1
'        l = l + 1
'        lt = lt + 1
'    Loop
'
'    'FINDING THE SQUARES OF X
'    store(1) = X Mod N
'    temp = store(1)
'    disp(1) = temp
'    mtp = 2
'    flmt = 1
'    t = 2
'    Do While ((flmt / 2) <= Val(Y))
'        store(mtp) = (temp ^ 2) Mod N
'        temp = store(mtp)
'        disp(t) = (store(mtp))
'        mtp = mtp * 2
'        flmt = mtp * 2
'        t = t + 1
'    Loop
'    'ELIMINATIN ZEROS
'     i1 = 1
'    j1 = 2
'    rvf(1) = pow(1)
'    Do While (i1 < 15)
'        If ch(i1) <> 0 Then
'            rvf(j1) = ch(i1)
'            j1 = j1 + 1
'        End If
'        i1 = i1 + 1
'    Loop
'    'CALLING VALUES
'    pcnt = 1
'    jc = 1
'    For ic = 1 To chlen
'        If ch(ic) <> 0 Then
'            suma(jc) = (disp(ic))
'            pcnt = pcnt + 1
'            jc = jc + 1
'        End If
'    Next ic
'    'RESULT
'    res = 1
'    Dim q1 As Integer
'    res = (suma(1) * suma(2)) Mod N
'    For q1 = 2 To (pcnt - 2)
'        res = (res * suma(q1 + 1)) Mod N
'    Next q1
'End Sub

