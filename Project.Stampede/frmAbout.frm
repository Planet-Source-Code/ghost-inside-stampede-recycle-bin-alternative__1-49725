VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form About Stampede"
   ClientHeight    =   5580
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4110
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5580
   ScaleWidth      =   4110
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame frStamp 
      Caption         =   "Stampede File Deletion - Trash Alternative"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5295
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3855
      Begin VB.Label lblNote 
         Caption         =   $"frmAbout.frx":0442
         Height          =   975
         Left            =   120
         TabIndex        =   5
         Top             =   3960
         Width           =   3615
      End
      Begin VB.Label Label1 
         Caption         =   $"frmAbout.frx":051A
         Height          =   855
         Left            =   120
         TabIndex        =   4
         Top             =   3000
         Width           =   3495
      End
      Begin VB.Label lblDelete 
         Caption         =   "'Windows doesn't really delete files but I do....after stamping the hell out of  them!'"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   960
         TabIndex        =   3
         Top             =   360
         Width           =   2415
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   240
         Picture         =   "frmAbout.frx":05AE
         Top             =   360
         Width           =   480
      End
      Begin VB.Label lblHLink 
         Caption         =   "Email: John Bridle jbridle@rainydayz.com"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00634221&
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   2520
         Width           =   3495
      End
      Begin VB.Label lblTrash 
         Caption         =   $"frmAbout.frx":09F0
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   120
         TabIndex        =   1
         Top             =   1320
         Width           =   3615
      End
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const SW_SHOWNORMAL = 1
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Me.lblHLink.ForeColor = &H634221
End Sub

Private Sub Form_Terminate()
    End
End Sub

Private Sub frStamp_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Form_MouseMove Button, Shift, X, Y
End Sub
Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Form_MouseMove Button, Shift, X, Y
End Sub

Private Sub lblHLink_Click()
Dim lret        As Long
Dim sLink       As String

    sLink = "mailto:jbridle@rainydayz.com"
    lret = ShellExecute(Me.hwnd, "open", sLink, vbNull, vbNull, SW_SHOWNORMAL)
    
End Sub
Private Sub lblHLink_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Me.lblHLink.ForeColor = &HF8CDB1
    
End Sub
Private Sub lblNote_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Form_MouseMove Button, Shift, X, Y
End Sub

Private Sub lblTrash_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Form_MouseMove Button, Shift, X, Y
End Sub
