VERSION 5.00
Begin VB.Form Form2 
   BackColor       =   &H0082BEB4&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Decryption"
   ClientHeight    =   6810
   ClientLeft      =   4935
   ClientTop       =   1875
   ClientWidth     =   7125
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6810
   ScaleWidth      =   7125
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H0082BEB4&
      Height          =   2895
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   4095
      Begin VB.TextBox txtN 
         BackColor       =   &H0082BEB4&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2760
         TabIndex        =   7
         Top             =   1200
         Width           =   1095
      End
      Begin VB.CommandButton Command3 
         BackColor       =   &H0082BEB4&
         Caption         =   "Exit"
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
         Left            =   2640
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   2160
         Width           =   1215
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H0082BEB4&
         Caption         =   "Clear"
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
         Left            =   1440
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   2160
         Width           =   1215
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H0082BEB4&
         Caption         =   "Ok"
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
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   2160
         Width           =   1215
      End
      Begin VB.TextBox txtSecret 
         BackColor       =   &H0082BEB4&
         Height          =   495
         Left            =   2760
         TabIndex        =   2
         Top             =   480
         Width           =   1095
      End
      Begin VB.Label Label3 
         BackColor       =   &H0082BEB4&
         Caption         =   "Enter The Secret Key 2"
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
         Left            =   120
         TabIndex        =   6
         Top             =   1320
         Width           =   2055
      End
      Begin VB.Label Label2 
         BackColor       =   &H0082BEB4&
         Caption         =   "Enter The Secret Key 1"
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
         Left            =   120
         TabIndex        =   1
         Top             =   600
         Width           =   2055
      End
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
   Form1.txtFileName.Text = Form1.txtTemp.Text
   Form1.Text1.Text = Form2.txtSecret.Text
   Form1.Text2.Text = Form2.txtN.Text
   Unload Me
   Form1.Show
End Sub

Private Sub Command2_Click()
    Text1.Text = ""
    
End Sub

Private Sub Command3_Click()
    End
    
End Sub

Private Sub Text1_Change()
    Text1.FontSize = "12"
    Text1.FontBold = True
End Sub

