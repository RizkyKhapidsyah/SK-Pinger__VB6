VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmOptions 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "HTM options"
   ClientHeight    =   4155
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   5130
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4155
   ScaleWidth      =   5130
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   330
      Left            =   2040
      TabIndex        =   8
      Top             =   3720
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      Caption         =   "Options"
      Height          =   3015
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   4935
      Begin VB.TextBox txtLink 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   9
         Top             =   2280
         Width           =   4455
      End
      Begin RichTextLib.RichTextBox txtCaption 
         Height          =   375
         Left            =   120
         TabIndex        =   7
         Top             =   1320
         Width           =   4455
         _ExtentX        =   7858
         _ExtentY        =   661
         _Version        =   393217
         Enabled         =   -1  'True
         MultiLine       =   0   'False
         TextRTF         =   $"frmOptions.frx":0000
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin RichTextLib.RichTextBox txtTitle 
         Height          =   405
         Left            =   120
         TabIndex        =   6
         Top             =   480
         Width           =   4455
         _ExtentX        =   7858
         _ExtentY        =   714
         _Version        =   393217
         Enabled         =   -1  'True
         MultiLine       =   0   'False
         TextRTF         =   $"frmOptions.frx":008E
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Link to LOG.TXT"
         Height          =   195
         Left            =   120
         TabIndex        =   2
         Top             =   1920
         Width           =   1575
      End
      Begin VB.Line Line5 
         X1              =   4680
         X2              =   1680
         Y1              =   240
         Y2              =   240
      End
      Begin VB.Line Line4 
         X1              =   4680
         X2              =   4680
         Y1              =   2760
         Y2              =   240
      End
      Begin VB.Line Line1 
         X1              =   120
         X2              =   4680
         Y1              =   1800
         Y2              =   1800
      End
      Begin VB.Line Line2 
         X1              =   120
         X2              =   4680
         Y1              =   2760
         Y2              =   2760
      End
      Begin VB.Label Label1 
         Caption         =   "Caption of you ping.htm:"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   1080
         Width           =   1935
      End
      Begin VB.Line Line3 
         X1              =   120
         X2              =   4680
         Y1              =   960
         Y2              =   960
      End
      Begin VB.Label Label2 
         Caption         =   "Title of the ping.htm:"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   3255
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Ok"
      Height          =   330
      Left            =   3600
      TabIndex        =   0
      Top             =   3720
      Width           =   1455
   End
   Begin VB.Label Label3 
      Caption         =   "Note: It will only work if you use an HTTP server like IIS or SMALL HTTP SERVER etc. ect. ."
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   4935
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Check1_Click()
If Check1.Value = 1 Then txtLink.Enabled = True Else txtLink.Enabled = False
If Check1.Value = 1 Then txtLink.Text = "Print version" Else txtLink.Text = ""
End Sub

Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub Command1_Click()
txtTitle.SaveFile ("title.ini"), vbFlags
txtCaption.SaveFile ("caption.ini"), vbFlags
Me.Hide
End Sub

Private Sub Form_Load()
On Error Resume Next
txtTitle.LoadFile ("title.ini"), vbFlags
txtCaption.LoadFile ("caption.ini"), vbFlags
Exit Sub
End Sub


