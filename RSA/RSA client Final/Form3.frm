VERSION 5.00
Begin VB.Form Form3 
   BackColor       =   &H0082BEB4&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Keys"
   ClientHeight    =   4740
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5640
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4740
   ScaleWidth      =   5640
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdEnd 
      BackColor       =   &H0082BEB4&
      Caption         =   "&End"
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
      Left            =   1560
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   2160
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H0082BEB4&
      Caption         =   "}"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   7
      Left            =   4080
      TabIndex        =   6
      Top             =   1560
      Width           =   135
   End
   Begin VB.Label lblN 
      Alignment       =   2  'Center
      BackColor       =   &H0082BEB4&
      Caption         =   "1256"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3480
      TabIndex        =   5
      Top             =   1560
      Width           =   495
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H0082BEB4&
      Caption         =   ","
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   5
      Left            =   3240
      TabIndex        =   4
      Top             =   1560
      Width           =   135
   End
   Begin VB.Label lblSecret 
      Alignment       =   2  'Center
      BackColor       =   &H0082BEB4&
      Caption         =   "1256"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2640
      TabIndex        =   3
      Top             =   1560
      Width           =   495
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H0082BEB4&
      Caption         =   "{"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   3
      Left            =   2400
      TabIndex        =   2
      Top             =   1560
      Width           =   135
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H0082BEB4&
      Caption         =   "The Key Pair Is "
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   240
      TabIndex        =   1
      Top             =   1560
      Width           =   2055
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H0082BEB4&
      Caption         =   "Please Note The Secret keys"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   1
      Left            =   600
      TabIndex        =   0
      Top             =   600
      Width           =   3495
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdEnd_Click()
    End
End Sub


