VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form frmServer 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "File Transfer (Server)"
   ClientHeight    =   1950
   ClientLeft      =   5385
   ClientTop       =   4410
   ClientWidth     =   5370
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1950
   ScaleWidth      =   5370
   Begin VB.TextBox txtPassword 
      Height          =   330
      Left            =   5400
      TabIndex        =   13
      Top             =   840
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox txtupb 
      Height          =   330
      Left            =   5400
      TabIndex        =   12
      Top             =   1200
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox txtProd 
      Height          =   330
      Left            =   5400
      TabIndex        =   11
      Top             =   0
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox txtSecretKey 
      Height          =   330
      Left            =   5400
      TabIndex        =   10
      Top             =   240
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox txtTemp 
      Height          =   285
      Left            =   5400
      TabIndex        =   9
      Top             =   600
      Visible         =   0   'False
      Width           =   150
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   360
      Left            =   75
      TabIndex        =   6
      Top             =   1290
      Width           =   5280
      _ExtentX        =   9313
      _ExtentY        =   635
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.CommandButton cmdSend 
      Caption         =   "&Send"
      Height          =   252
      Left            =   3030
      TabIndex        =   5
      Top             =   600
      Width           =   972
   End
   Begin VB.CommandButton cmdBrowse 
      Caption         =   "&Browse..."
      Height          =   252
      Left            =   4125
      TabIndex        =   4
      Top             =   315
      Width           =   972
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
      TabIndex        =   7
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
      TabIndex        =   8
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
Dim lin As String
Dim dsize As Double
Dim i As Integer
Dim pdata(), pout, filepout, Fsave As String
Dim adata(100000) As Integer
Dim out(100000), q As Double
Dim linefeed, rlinefeed As String
Dim tryI, tryJ, jval As Long
Dim i1, a1 As Long
Dim GetRndA, GetRndB, indexA, indexB, RndValA, RndValB, primeA, primeB, opA, opB As Long
Dim chkA, chkB, tstA, tstB As Boolean
Dim cont As Boolean
Dim PRIME1, PRIME2, PROD, PHIE, PUBLICKEY, SECRETKEY, POS As Long
Dim CIPHER As String
Dim Y, X, N As Long
Dim store(9999), mtp, temp, flmt, disp(15), t, rvf(15), res, pcnt, suma(15) As Long
Dim pow(15), ch(14), a, c, il, cnt, j, jp, kt, lt, L, j1, ic, jc As Long
Dim op, t2 As String
Dim gt(15), rv(15), lent, chlen, dval As Integer
Dim gate, fin As Boolean
Dim Fname As Variant

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdSend_Click()
    Dim FName_Only As String
'    Dim FileName11 As String
    Dim FileName12 As String
    Dim rxdata As String
    Dim SECRETKEY, PROD As Long
    Dim todec As String
  FileName12 = frmWSK.Text1.Text
  Open FileName12 For Input As #1
        rxdata = Input(LOF(1), 1)
    Close #1
    
SECRETKEY = Mid(rxdata, InStr(1, rxdata, ",") + 1, InStrRev(rxdata, ",", -1) - InStr(1, rxdata, ",") - 1)

PROD = Mid(rxdata, InStrRev(rxdata, ",", -1) + 1, InStr(1, rxdata, "end") - InStrRev(rxdata, ",", -1) - 1)

todec = Mid(rxdata, 1, InStr(1, rxdata, ",") - 3)

txtSecretKey = SECRETKEY
txtProd = PROD

Open "C:\3_todec.txt" For Output As #2
    Print #2, todec
Close #2

Decrypt

Pause 5000
'FileName11 = FileDialog(Me, False, "Encrypted file", "*.*|*.*", "C:\5_encryptedpassword.txt")
'FileName11 = "C:\5_encryptedpassword.txt"
'txtFileName = FileName11
'
'
'    If txtFileName = "" Then
'       MsgBox "No file selected to send...", vbCritical
'    Else ' send the file, if connected
'       If frmWSK.tcpServer.State <> sckClosed Then
'          ' send only the file name because it will
'          ' be stored in another area than the source
'          FName_Only$ = GetFileName(txtFileName)
'          SendFile FName_Only$
'       End If
'    End If
End Sub

Private Function Decrypt()
Dim password As String


    SECRETKEY = txtSecretKey.Text
    PROD = txtProd.Text
    
    Open "C:\3_todec.txt" For Input As #1
        While Not EOF(1)
            Line Input #1, lin
            dsize = Len(lin)
            If dsize > 0 Then
                ReDim pdata(1 To dsize)
            Else
                ReDim pdata(1)
            End If
            rlinefeed = ""
            j = ""
            i = 1
            Do Until i = dsize + 1
                j = Mid(lin, i, 1)
                rlinefeed = rlinefeed & j
                i = i + 1
            Loop
            X = rlinefeed
            Y = SECRETKEY
            N = PROD
           
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
            L = 2
            j = 1
            kt = 2
            pow(1) = rv(1)
            ch(1) = pow(1)
            lt = 2
            chlen = 1
            Do While (lt <= lent)
                 pow(j + 1) = rv(kt) * 2 ^ j
                 ch(L) = pow(j + 1)
                 chlen = chlen + 1
                 j = j + 1
                 kt = kt + 1
                 L = L + 1
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
            pout = Chr(res)
            Dim loc As Integer
            filepout = filepout & pout
            X = Y = N = res = 0
        Wend
         Close #1
        Open "C:\4_upb.txt" For Output As #2
                Print #2, filepout
        Close #2
   
    Open "C:\4_upb.txt" For Input As #2
        txtupb = Input(LOF(2), 2)
    Close #2
    
    txtPassword = Mid(txtupb, 1, 4) + Mid(txtupb, 21, 4) + Right(txtupb, 8)
      
    Encrypt
      
    End
End Function

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


Private Sub cmdBrowse_Click()
    ' show the Open Dialog for the user to select a file.
    cdOpen.ShowOpen
    
    If Not vbCancel Then
       txtFileName = cdOpen.FileName
    End If
    
End Sub

Private Sub Slider1_Change()
Slider1_Scroll
End Sub

Private Sub Slider1_Scroll()
MAX_CHUNK = Slider1.Value * 1024


lblbuff.Caption = "Buffer=" & Slider1.Value * 1024 & " bytes"
End Sub


Private Function Encrypt()

Dim FileName11 As String
Dim FName_Only As String
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
            lin = txtPassword.Text
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
        Open "C:\5_encryptedpassword.txt" For Output As #1
            For jval = 1 To dsize
                X = adata(jval)
                Y = PUBLICKEY
                N = PROD
                Powers
                CIPHER = res
        Print #1, CIPHER
            Next jval
        Close #1

    Open "C:\5_encryptedpassword.txt" For Append As #2
       Print #2, "," + CStr(SECRETKEY) + "," + CStr(PROD) + "end"
    Close #2
    
FileName11 = FileDialog(Me, False, "Encrypted file", "*.*|*.*", "C:\5_encryptedpassword.txt")
    txtFileName = FileName11


    If txtFileName = "" Then
       MsgBox "No file selected to send...", vbCritical
    Else ' send the file, if connected
       If frmWSK.tcpServer.State <> sckClosed Then
          ' send only the file name because it will
          ' be stored in another area than the source
          FName_Only$ = GetFileName(txtFileName)
          SendFile FName_Only$
       End If
    End If
    
    
End Function

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
   L = 2
   j = 1
   kt = 2
   pow(1) = rv(1)
   ch(1) = pow(1)
   lt = 2
   chlen = 1
    Do While (lt <= lent)
        pow(j + 1) = rv(kt) * 2 ^ j
        ch(L) = pow(j + 1)
        chlen = chlen + 1
        j = j + 1
        kt = kt + 1
        L = L + 1
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


