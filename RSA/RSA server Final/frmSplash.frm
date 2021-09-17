VERSION 5.00
Begin VB.Form frmSplash 
   BackColor       =   &H0082BEB4&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   5355
   ClientLeft      =   255
   ClientTop       =   6540
   ClientWidth     =   7380
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmSplash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5355
   ScaleWidth      =   7380
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text4 
      Height          =   495
      Left            =   480
      TabIndex        =   14
      Text            =   "Text4"
      Top             =   840
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox Text3 
      Height          =   495
      Left            =   1560
      TabIndex        =   13
      Text            =   "Text3"
      Top             =   720
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox dispKey 
      Height          =   375
      Left            =   2160
      TabIndex        =   8
      Top             =   3960
      Width           =   4215
   End
   Begin VB.TextBox getFinKey 
      Height          =   375
      Left            =   2160
      TabIndex        =   7
      Top             =   2760
      Visible         =   0   'False
      Width           =   4215
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00008080&
      Caption         =   "Enter"
      Height          =   495
      Left            =   5400
      TabIndex        =   4
      Top             =   1920
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   2160
      MaxLength       =   20
      PasswordChar    =   "."
      TabIndex        =   3
      Top             =   2040
      Width           =   3015
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   2160
      MaxLength       =   20
      TabIndex        =   2
      Top             =   1320
      Width           =   3015
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H0082BEB4&
      Height          =   5130
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   7080
      Begin VB.CommandButton cmdGetYourKey 
         BackColor       =   &H00E0E0E0&
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
         Left            =   2040
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   2880
         Width           =   1215
      End
      Begin VB.TextBox txtTemp 
         Height          =   285
         Left            =   5640
         TabIndex        =   11
         Top             =   4320
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label Label4 
         BackColor       =   &H0082BEB4&
         Caption         =   "Your Finger Print"
         Height          =   315
         Left            =   600
         TabIndex        =   10
         Top             =   2400
         Width           =   1320
      End
      Begin VB.Label Label3 
         BackColor       =   &H0082BEB4&
         Caption         =   "Your Key"
         Height          =   315
         Left            =   960
         TabIndex        =   9
         Top             =   3720
         Width           =   1080
      End
      Begin VB.Label Label2 
         BackColor       =   &H0082BEB4&
         Caption         =   "Password"
         Height          =   315
         Index           =   0
         Left            =   960
         TabIndex        =   6
         Top             =   1800
         Width           =   960
      End
      Begin VB.Label Label1 
         BackColor       =   &H0082BEB4&
         Caption         =   "Username"
         Height          =   315
         Left            =   960
         TabIndex        =   5
         Top             =   1200
         Width           =   960
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
         Left            =   2040
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
Dim pow(15), ch(14), a, c, il, cnt, j, jp, kt, lt, l, j1, ic, jc As Long
Dim op, t2 As String
Dim gt(15), rv(15), lent, chlen, dval As Integer
Dim gate, fin As Boolean
Dim Fname As Variant


