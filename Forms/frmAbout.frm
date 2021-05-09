VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "About Pinger"
   ClientHeight    =   4395
   ClientLeft      =   2340
   ClientTop       =   1935
   ClientWidth     =   5460
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3033.508
   ScaleMode       =   0  'User
   ScaleWidth      =   5127.223
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame frmeMain 
      Height          =   3735
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   5175
      Begin VB.Label Label5 
         Caption         =   $"frmAbout.frx":0000
         Height          =   495
         Left            =   120
         TabIndex        =   9
         Top             =   2880
         Width           =   4935
      End
      Begin VB.Label Label4 
         Caption         =   "Future"
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
         Left            =   120
         TabIndex        =   8
         Top             =   2640
         Width           =   3015
      End
      Begin VB.Label Label3 
         Caption         =   $"frmAbout.frx":0087
         Height          =   855
         Left            =   120
         TabIndex        =   7
         Top             =   1680
         Width           =   4935
      End
      Begin VB.Label Label2 
         Caption         =   "Why I 've written this little tool"
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
         Left            =   120
         TabIndex        =   6
         Top             =   1440
         Width           =   2655
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "2-BYTE 2001-2002"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   135
         Left            =   1080
         TabIndex        =   5
         Top             =   3480
         Width           =   3015
      End
      Begin VB.Label lblProductName 
         Caption         =   $"frmAbout.frx":018C
         Height          =   735
         Index           =   2
         Left            =   120
         TabIndex        =   4
         Top             =   480
         Width           =   4935
      End
      Begin VB.Label lblProductName 
         Caption         =   "Version 1.0"
         Height          =   255
         Index           =   1
         Left            =   720
         TabIndex        =   3
         Top             =   240
         Width           =   855
      End
      Begin VB.Label lblProductName 
         Caption         =   "Pinger"
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
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   615
      End
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Height          =   360
      Left            =   3720
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   3960
      Width           =   1530
   End
   Begin VB.Label Label8 
      Caption         =   "to_byte@hotmail.com?subject=Pinger v1.0 (vb)"
      Height          =   255
      Left            =   600
      TabIndex        =   12
      Top             =   5280
      Width           =   3135
   End
   Begin VB.Label Label7 
      Caption         =   " to_byte@hotmail.com"
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   720
      TabIndex        =   11
      Top             =   3960
      Width           =   1575
   End
   Begin VB.Label Label6 
      Caption         =   "E-mail:"
      Height          =   255
      Left            =   240
      TabIndex        =   10
      Top             =   3960
      Width           =   495
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function SendMessage Lib "User32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub Label7_Click()
ShellExecute 0, "Open", "mailto:" & Label8.Caption, "", "", vbNormalFocus
End Sub
