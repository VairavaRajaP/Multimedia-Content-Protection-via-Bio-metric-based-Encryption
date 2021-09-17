VERSION 5.00
Begin VB.Form Form4 
   Caption         =   "Select file for Decryption"
   ClientHeight    =   5640
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   5775
   LinkTopic       =   "Form4"
   ScaleHeight     =   5640
   ScaleWidth      =   5775
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   7000
      Left            =   2280
      Top             =   2640
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H0082BEB4&
      Height          =   5655
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5775
      Begin VB.DriveListBox Drive1 
         BackColor       =   &H0082BEB4&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   3240
         TabIndex        =   4
         Top             =   960
         Width           =   2295
      End
      Begin VB.DirListBox Dir1 
         BackColor       =   &H0082BEB4&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Left            =   3240
         TabIndex        =   3
         Top             =   1560
         Width           =   2295
      End
      Begin VB.FileListBox File1 
         BackColor       =   &H0082BEB4&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1650
         Left            =   3240
         TabIndex        =   2
         Top             =   3000
         Width           =   2295
      End
      Begin VB.TextBox txtFileName 
         BackColor       =   &H0082BEB4&
         Height          =   375
         Left            =   2160
         TabIndex        =   1
         Top             =   4920
         Width           =   3495
      End
      Begin VB.Label Label3 
         BackColor       =   &H0082BEB4&
         Caption         =   "     Decryption"
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
         Left            =   1800
         TabIndex        =   9
         Top             =   240
         Width           =   1575
      End
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
         Index           =   0
         Left            =   360
         TabIndex        =   8
         Top             =   5040
         Width           =   1575
      End
      Begin VB.Label Label3 
         BackColor       =   &H0082BEB4&
         Caption         =   "Select your Drive"
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
         Index           =   0
         Left            =   360
         TabIndex        =   7
         Top             =   960
         Width           =   1575
      End
      Begin VB.Label Label4 
         BackColor       =   &H0082BEB4&
         Caption         =   "Select Your Folder"
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
         Left            =   480
         TabIndex        =   6
         Top             =   1800
         Width           =   1695
      End
      Begin VB.Label Label5 
         BackColor       =   &H0082BEB4&
         Caption         =   "Select Your File"
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
         Left            =   480
         TabIndex        =   5
         Top             =   3120
         Width           =   1455
      End
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Dir1_Change()
    File1.Path = Dir1
End Sub

Private Sub Drive1_Change()
    Dir1.Path = Drive1
End Sub

Private Sub File1_Click()
    txtFileName.Text = Dir1 & "\" & File1
    txtFileName.FontBold = True
    txtFileName.FontSize = "8"
    fn = txtFileName.Text
    fnt = Dir1 & "\" & "temp.tmp"
End Sub
