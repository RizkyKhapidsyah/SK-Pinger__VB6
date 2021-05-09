VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "Richtx32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MsComCtl.ocx"
Begin VB.Form frmMain 
   BackColor       =   &H80000004&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Intranet     HTML-PINGER V1.0"
   ClientHeight    =   4515
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8025
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4515
   ScaleWidth      =   8025
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   735
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   7935
      Begin VB.CommandButton cmdMin 
         Caption         =   "Minimize"
         Height          =   330
         Left            =   4200
         TabIndex        =   12
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Stop pinging"
         Height          =   330
         Left            =   2520
         TabIndex        =   11
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton cmdPing 
         Caption         =   "&Ping"
         Height          =   330
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton cmdOptions 
         Caption         =   "&Options"
         Height          =   330
         Left            =   1320
         TabIndex        =   7
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton cmdAbout 
         Caption         =   "A&bout"
         Height          =   330
         Left            =   5400
         TabIndex        =   6
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton cmdExit 
         Caption         =   "E&xit"
         Height          =   330
         Left            =   6600
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Timer Timer2 
      Interval        =   50
      Left            =   3240
      Top             =   4560
   End
   Begin VB.Timer Timer1 
      Interval        =   60000
      Left            =   2760
      Top             =   4560
   End
   Begin MSComctlLib.ImageList imglMain 
      Left            =   120
      Top             =   4560
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":000C
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0326
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame frmeMain 
      BackColor       =   &H80000004&
      ForeColor       =   &H80000007&
      Height          =   3855
      Left            =   0
      TabIndex        =   3
      Top             =   600
      Width           =   7935
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   6120
         TabIndex        =   10
         Text            =   "Text1"
         Top             =   240
         Width           =   855
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   6960
         TabIndex        =   9
         Text            =   "Text2"
         Top             =   240
         Width           =   855
      End
      Begin RichTextLib.RichTextBox txtNumber 
         Height          =   330
         Left            =   2040
         TabIndex        =   1
         Top             =   240
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   582
         _Version        =   393217
         MultiLine       =   0   'False
         MaxLength       =   1
         TextRTF         =   $"frmMain.frx":0A9D
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin RichTextLib.RichTextBox txtIP 
         Height          =   330
         Left            =   120
         TabIndex        =   0
         Top             =   240
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   582
         _Version        =   393217
         Enabled         =   -1  'True
         MultiLine       =   0   'False
         TextRTF         =   $"frmMain.frx":0B1B
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin RichTextLib.RichTextBox txtOutPut 
         Height          =   3135
         Left            =   120
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   600
         Width           =   7695
         _ExtentX        =   13573
         _ExtentY        =   5530
         _Version        =   393217
         ReadOnly        =   -1  'True
         ScrollBars      =   3
         Appearance      =   0
         TextRTF         =   $"frmMain.frx":0BA3
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Const SYNCHRONIZE = &H100000
Const INFINITE = &HFFFF
Const WAIT_OBJECT_0 = 0
Const WAIT_TIMEOUT = &H102
Private Declare Function SendMessage Lib "User32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Private Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long

Private Sub cmdAbout_Click()
frmAbout.Show vbModal, Me
End Sub

Private Sub cmdExit_Click()
txtIP.SaveFile ("ip.ini")
        End
End Sub

Private Sub cmdOptions_Click()
frmOptions.Show
End Sub

Private Sub cmdPing_Click()
Dim ShellX As String
Dim lPid As Long
Dim lHnd As Long
Dim lRet As Long
Dim VarX As String
    Open ("ping.htm") For Output As #1
    Close #1
  frmMain.MousePointer = 11
  If txtIP.Text <> "" Then
    DoEvents
    ShellX = Shell("cmd.exe /c ping -n " & txtNumber.Text & " " & txtIP.Text & " > log.txt", vbHide)
    
    lPid = ShellX
    If lPid <> 0 Then
        lHnd = OpenProcess(SYNCHRONIZE, 0, lPid)
        If lHnd <> 0 Then
            lRet = WaitForSingleObject(lHnd, INFINITE)
            CloseHandle (lHnd)
        End If
            Beep
            frmMain.MousePointer = 0
            Open "log.txt" For Input As #1
            txtOutPut.Text = Input(LOF(1), 1)
            Close #1
    End If
  Else
    frmMain.MousePointer = 0
    VarX = MsgBox("You have not entered an ip address or the number of times you want to ping.", vbCritical, "Error has occured")
  End If

Open ("ping.htm") For Output As #1
Print #1, "<html>"
Print #1, "<head>"
Print #1, "<title>" & frmOptions.txtTitle.Text & "</title>"
Print #1, "</head>"
Print #1, "<body bgcolor=#000080 text=#C0C0C0 link=#FFCC66 vlink=#FFCC66 alink=#FFCC66>"
Print #1, "<p align=center><b><font size=7>" & frmOptions.txtCaption.Text & "</font></b></p>"
Print #1, "<p>"
Print #1, "<p align=center>&nbsp;"
Print #1, "<p align=center>&nbsp;"
Print #1, "<p align=center>&nbsp;"
Print #1, txtOutPut.Text
Print #1, "</P>"
Print #1, "<p align=center>&nbsp;"
Print #1, "<p align=; center; >&nbsp;</p>"
Print #1, "<p align=center>&nbsp;"
Print #1, "Last Ping: " & Text1.Text & " " & Text2.Text & " with an interval of 1 minute."
Print #1, "</p>"
Print #1, "<p align=center><a href=log.txt>" & frmOptions.txtLink & "</a></p>"
Print #1, "</body>"
Print #1, "</html>"
Close #1
Open "log.txt" For Output As #2
Print #2, txtOutPut.Text
Close #2
End Sub

Private Sub Command1_Click()
If Timer1.Enabled = True Then Timer1.Enabled = False Else Timer1.Enabled = True
If Command1.Caption = "&Stop pinging" Then Command1.Caption = "&Start pinging" Else Command1.Caption = "&Stop pinging"

End Sub

Private Sub Form_Load()
  On Error Resume Next
  frmOptions.Show
  frmOptions.Hide
  txtIP.LoadFile ("ip.ini")
    Open "log.txt" For Output As #1
  Close #1
    SendMessage cmdExit.hwnd, &HF4&, &H0&, 0&
    SendMessage cmdPing.hwnd, &HF4&, &H0&, 0&
End Sub

Private Sub SelectText(ByRef textObj As RichTextBox)
    textObj.SelStart = 0
    textObj.SelLength = Len(textObj)
End Sub

Private Sub Form_Unload(Cancel As Integer)

Timer1.Enabled = False
Timer2.Enabled = False
End Sub

Private Sub Timer1_Timer()
Dim ShellX As String
Dim lPid As Long
Dim lHnd As Long
Dim lRet As Long
Dim VarX As String
    Open ("ping.htm") For Output As #1
    Close #1
  frmMain.MousePointer = 11
  If txtIP.Text <> "" Then
    DoEvents
    ShellX = Shell("command.com /c ping -n " & txtNumber.Text & " " & txtIP.Text & " > log.txt", vbHide)
    
    lPid = ShellX
    If lPid <> 0 Then
        lHnd = OpenProcess(SYNCHRONIZE, 0, lPid)
        If lHnd <> 0 Then
            lRet = WaitForSingleObject(lHnd, INFINITE)
            CloseHandle (lHnd)
        End If
            Beep
            frmMain.MousePointer = 0
            Open "log.txt" For Input As #1
            txtOutPut.Text = Input(LOF(1), 1)
            Close #1
    End If
  Else
    frmMain.MousePointer = 0
    VarX = MsgBox("You have not entered an ip address or the number of times you want to ping.", vbCritical, "Error has occured")
  End If

Open ("ping.htm") For Output As #1
Print #1, "<html>"
Print #1, "<head>"
Print #1, "<title>" & frmOptions.txtTitle.Text & "</title>"
Print #1, "</head>"
Print #1, "<body bgcolor=#000080 text=#C0C0C0 link=#FFCC66 vlink=#FFCC66 alink=#FFCC66>"
Print #1, "<p align=center><b><font size=7>" & frmOptions.txtCaption.Text & "</font></b></p>"
Print #1, "<p>"
Print #1, "<p align=center>&nbsp;"
Print #1, "<p align=center>&nbsp;"
Print #1, "<p align=center>&nbsp;"
Print #1, txtOutPut.Text
Print #1, "</P>"
Print #1, "<p align=center>&nbsp;"
Print #1, "<p align=; center; >&nbsp;</p>"
Print #1, "<p align=center>&nbsp;"
Print #1, "Last Ping: " & Text1.Text & " " & Text2.Text & " with an interval of 1 minute."
Print #1, "</p>"
Print #1, "<p align=center><a href=log.txt>" & frmOptions.txtLink & "</a></p>"
Print #1, "</body>"
Print #1, "</html>"
Close #1
Open "log.txt" For Output As #2
Print #2, txtOutPut.Text
Close #2
End Sub

Private Sub Timer2_Timer()
Text1.Text = Time
Text2.Text = Date
End Sub
