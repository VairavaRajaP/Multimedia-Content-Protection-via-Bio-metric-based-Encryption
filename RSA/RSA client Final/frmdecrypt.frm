VERSION 5.00
Begin VB.Form frmdecrypt 
   Caption         =   "Form1"
   ClientHeight    =   3135
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3135
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "frmdecrypt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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

Private Sub Form_Load()

 Dim FileName22 As String
    Dim rxdata1 As String
    Dim todec1 As String
    Dim dispkey As String
    diskey = 1
    
  Open "C:\6_rxfile_password.txt" For Input As #6
        rxdata1 = Input(LOF(6), 6)
    Close #6

SECRETKEY = Mid(rxdata1, InStr(1, rxdata1, ",") + 1, InStrRev(rxdata1, ",", -1) - InStr(1, rxdata1, ",") - 1)

PROD = Mid(rxdata1, InStrRev(rxdata1, ",", -1) + 1, InStr(1, rxdata1, "end") - InStrRev(rxdata1, ",", -1) - 1)

todec1 = Mid(rxdata1, 1, InStr(1, rxdata1, ",") - 3)


Open "C:\7_todec_password.txt" For Output As #7
    Print #7, todec1
Close #7

'FileName22 = FileDialog(Me, False, "En pwd to De", "*.*|*.*", "C:\7_todec_password.txt")

FileName22 = "C:\7_todec_password.txt"


Open FileName22 For Input As #1
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
            pout = Chr(res)
            Dim loc As Integer
            filepout = filepout & pout
            X = Y = N = res = 0
        Wend
         Close #1
        Open "c:\pwd.txt" For Append As #2
                Print #2, filepout
        Close #2
'   dispkey = dispkey + filepout
      End
      
'      dispkey = Mid(dispkey, 2)
'      frmSplash.txtdispKey.Text = dispkey
End Sub
