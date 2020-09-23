VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form Form2 
   Caption         =   "Folder Watch"
   ClientHeight    =   3315
   ClientLeft      =   3465
   ClientTop       =   855
   ClientWidth     =   7620
   LinkTopic       =   "Form2"
   ScaleHeight     =   3315
   ScaleWidth      =   7620
   Begin VB.Timer Timer1 
      Left            =   7680
      Top             =   240
   End
   Begin VB.Frame Frame7 
      Caption         =   " Changed folders/files log "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   3255
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7575
      Begin VB.CommandButton Command1 
         Caption         =   "E&xit"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   2760
         Width           =   1575
      End
      Begin VB.TextBox Text5 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   6480
         TabIndex        =   1
         Text            =   "0"
         Top             =   2760
         Width           =   975
      End
      Begin RichTextLib.RichTextBox rtbChangedfiles 
         Height          =   2415
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   7335
         _ExtentX        =   12938
         _ExtentY        =   4260
         _Version        =   393217
         Enabled         =   -1  'True
         ScrollBars      =   3
         Appearance      =   0
         TextRTF         =   $"log2.frx":0000
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label9 
         Caption         =   "Log Entry"
         Height          =   255
         Left            =   5760
         TabIndex        =   3
         Top             =   2760
         Width           =   735
      End
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
'exits
Unload Form2


End Sub

Private Sub Command2_Click()
'clears log file ( this form only )

rtbChangedfiles.Text = " "

counter = 0
Text1.refresh

End Sub

Private Sub Form_Load()
       
    rtbChangedfiles = Form1.rtbChangedfiles.Text
    Timer1.Interval = 30 * 1000
    Text5.Text = Form1.Text5.Text

End Sub
Private Sub Timer1_Timer()
'reload info from main roprtlog

Timer1.Interval = 30 * 1000
rtbChangedfiles = Form1.rtbChangedfiles.Text

End Sub
