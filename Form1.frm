VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "MS NET Downloader"
   ClientHeight    =   1590
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6765
   LinkTopic       =   "Form1"
   ScaleHeight     =   1590
   ScaleWidth      =   6765
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Text            =   "http://www.planet-source-code.com/vb/"
      Top             =   480
      Width           =   6495
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Download"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5280
      TabIndex        =   0
      Top             =   960
      Width           =   1215
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Type URL Address"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   1740
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function DoFileDownload Lib "shdocvw.dll" (ByVal lpszFile As String) As Long

Private Sub Command1_Click()

DoFileDownload StrConv(Text1.Text, vbUnicode)

End Sub
