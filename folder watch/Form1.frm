VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Folder Watch"
   ClientHeight    =   6000
   ClientLeft      =   3510
   ClientTop       =   2175
   ClientWidth     =   7680
   DrawWidth       =   2
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   6000
   ScaleWidth      =   7680
   Begin VB.Timer Timer1 
      Left            =   10800
      Top             =   120
   End
   Begin VB.ListBox list2 
      Height          =   255
      ItemData        =   "Form1.frx":0ECA
      Left            =   120
      List            =   "Form1.frx":0ECC
      TabIndex        =   18
      Top             =   6120
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.ListBox List1 
      Height          =   255
      ItemData        =   "Form1.frx":0ECE
      Left            =   120
      List            =   "Form1.frx":0ED0
      TabIndex        =   17
      Top             =   6000
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Frame Frame3 
      Caption         =   "Status"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   2775
      Left            =   4200
      TabIndex        =   9
      Top             =   0
      Width           =   3495
      Begin MSComctlLib.Slider Slider1 
         Height          =   255
         Left            =   1200
         TabIndex        =   21
         Top             =   240
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   450
         _Version        =   393216
         Min             =   5
         Max             =   60
         SelStart        =   5
         Value           =   5
      End
      Begin VB.Frame Frame5 
         Caption         =   " Results "
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
         Height          =   2055
         Left            =   120
         TabIndex        =   12
         Top             =   600
         Width           =   3255
         Begin VB.TextBox txtFiles 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1680
            Locked          =   -1  'True
            TabIndex        =   19
            TabStop         =   0   'False
            Text            =   "0"
            Top             =   960
            Width           =   855
         End
         Begin VB.TextBox txtStatus 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   960
            Locked          =   -1  'True
            TabIndex        =   14
            TabStop         =   0   'False
            Text            =   "Idle"
            Top             =   240
            Width           =   1575
         End
         Begin VB.TextBox txtChanges 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   960
            Locked          =   -1  'True
            TabIndex        =   13
            TabStop         =   0   'False
            Text            =   "No changes"
            Top             =   600
            Width           =   1575
         End
         Begin VB.Label Label6 
            Height          =   255
            Left            =   240
            TabIndex        =   25
            Top             =   1560
            Width           =   2655
         End
         Begin VB.Label Label1 
            Caption         =   "Total Folders/Files"
            Height          =   255
            Left            =   240
            TabIndex        =   20
            Top             =   960
            Width           =   1335
         End
         Begin VB.Label Label3 
            Caption         =   "Status"
            Height          =   255
            Left            =   240
            TabIndex        =   16
            Top             =   240
            Width           =   735
         End
         Begin VB.Label Label4 
            Caption         =   "Change"
            Height          =   255
            Left            =   240
            TabIndex        =   15
            Top             =   600
            Width           =   615
         End
      End
      Begin VB.TextBox txtTime 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   11
         TabStop         =   0   'False
         Text            =   "5 Sec"
         Top             =   240
         Width           =   975
      End
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
      Left            =   120
      TabIndex        =   7
      Top             =   2760
      Width           =   7575
      Begin VB.CommandButton Command1 
         Caption         =   "Clear &log"
         Height          =   375
         Left            =   120
         TabIndex        =   24
         Top             =   2760
         Width           =   1215
      End
      Begin VB.TextBox Text5 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   6480
         TabIndex        =   22
         Text            =   "0"
         Top             =   2760
         Width           =   975
      End
      Begin RichTextLib.RichTextBox rtbChangedfiles 
         Height          =   2415
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   7335
         _ExtentX        =   12938
         _ExtentY        =   4260
         _Version        =   393217
         Enabled         =   -1  'True
         ScrollBars      =   3
         Appearance      =   0
         TextRTF         =   $"Form1.frx":0ED2
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
         TabIndex        =   23
         Top             =   2760
         Width           =   735
      End
   End
   Begin MSComctlLib.StatusBar sbr 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   6
      Top             =   5625
      Width           =   7680
      _ExtentX        =   13547
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Bevel           =   0
            Object.Width           =   8361
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            TextSave        =   "1/6/2006"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            TextSave        =   "3:43 PM"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman Baltic"
         Size            =   9.75
         Charset         =   186
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Frame Frame2 
      Caption         =   "Commands"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   735
      Left            =   120
      TabIndex        =   3
      Top             =   0
      Width           =   3975
      Begin VB.CommandButton mini 
         Caption         =   "&Minimize"
         Height          =   375
         Left            =   1320
         TabIndex        =   10
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton cmdScan 
         Caption         =   "Start &Scan"
         Height          =   375
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton cmdExit 
         Caption         =   "E&xit"
         Height          =   375
         Left            =   2520
         TabIndex        =   4
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Navigation"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   1935
      Left            =   120
      TabIndex        =   0
      Top             =   840
      Width           =   3975
      Begin VB.DriveListBox Drive1 
         Height          =   315
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   2295
      End
      Begin VB.DirListBox strpath 
         Height          =   1215
         Left            =   120
         TabIndex        =   1
         Top             =   600
         Width           =   3735
      End
   End
   Begin VB.Menu mnuform 
      Caption         =   "Main"
      Begin VB.Menu mnushow 
         Caption         =   "&Maximize"
      End
      Begin VB.Menu mnulog 
         Caption         =   "&Log"
      End
      Begin VB.Menu mnuexit 
         Caption         =   "E&xit"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''
''''''''    Purpose of Program: To Scan FTP server(or any other folder/Drive)for
''''''''                        New or Changed Files/Folders
''''''''
''''''''
''''''''
''''''''
''''''''
''''''''
''''''''
''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim counter As Integer
Option Explicit

Private Sub Command1_Click()

rtbChangedfiles.Text = " "
Text5.Text = "0"
counter = 0
Text5.refresh

End Sub


Private Sub mini_Click()

TrayAdd hwnd, Me.Icon, "Folder Watch", MouseMove
    mnuHide_Click

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

Dim cEvent As Single

sbr.Panels(1).Text = ""
cEvent = X / Screen.TwipsPerPixelX

Select Case cEvent
    
    Case MouseMove
        Debug.Print "MouseMove"
    Case LeftUp
        Debug.Print "Left Up"
    Case LeftDown
        Debug.Print "LeftDown": PopupMenu mnuform
    Case LeftDbClick
        Debug.Print "LeftDbClick"
    Case MiddleUp
        Debug.Print "MiddleUp"
    Case MiddleDown
        Debug.Print "MiddleDown"
    Case MiddleDbClick
        Debug.Print "MiddleDbClick"
    Case RightUp
        Debug.Print "RightUp": PopupMenu mnuform
    Case RightDown
        Debug.Print "RightDown"
    Case RightDbClick
        Debug.Print "RightDbClick"

End Select

End Sub

Private Sub mnuexit_Click()
    
    cmdExit_Click

End Sub

Private Sub mnuHide_Click()

    If Not Me.WindowState = 1 Then WindowState = 1: Me.Hide
    
End Sub

Private Sub mnulog_Click()
    
    Form2.Show
    
End Sub

Private Sub mnuShow_Click()

    If Me.WindowState = 1 Then WindowState = 0: Me.Show
    TrayDelete
        
End Sub

Private Sub cmdClr_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    sbr.Panels(1).Text = "Clear Report"

End Sub
Private Sub cmdclr_click()

rtbChangedfiles = " "

End Sub

Private Sub cmdExit_Click()

TrayDelete
    
    End

End Sub

Private Sub cmdExit_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    sbr.Panels(1).Text = "Exit"

End Sub

Private Sub mini_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    sbr.Panels(1).Text = "Minimize to System Tray"

End Sub

Private Sub cmdScan_Click()
    
    cmdScan.Enabled = False
    strScanDir = strpath.List(strpath.ListIndex)
    Set strFolder = FSO.GetFolder(strScanDir)
    MapDirs (strFolder)
    txtFiles.Text = list2.ListCount
    txtStatus.Text = "Guard started"
    Text5.Text = "0"
    txtTime.Text = Slider1.Value & " Sec"
    Timer1.Interval = Slider1.Value * 1000
    TrayAdd hwnd, Me.Icon, "Folder Watch", MouseMove
    'mnuHide_Click
    Label6.Caption = "Started  " & " " & Date & "  " & Time
 
       
       
End Sub

Private Sub cmdScan_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    sbr.Panels(1).Text = "Scan Selected Directory"

End Sub

Private Sub strpath_Change()
    
    ChDir strpath.Path

End Sub

Private Sub strpath_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    sbr.Panels(1).Text = "Select Directory to Scan"

End Sub

Private Sub command1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    sbr.Panels(1).Text = "Clear Log"

End Sub

Private Sub Drive1_Change()
On Error Resume Next
    
    ChDrive Drive1.Drive
    strpath.Path = Drive1.Drive

End Sub

Private Sub Form_Load()
    
    strpath = Drive1.Drive
    Set strFolder = FSO.GetFolder(strpath)
    For Each strFile In strFolder.Files
    strDsc = ExtractExt(strFile.Name)
        
    Next


mnuform.Visible = False
    
End Sub

Private Function MapDirs(sfolder As String)

Dim pFolder As Folder
Dim fname As String

On Error Resume Next
    
    Set pFolder = FSO.GetFolder(sfolder)
    For Each strSubFolder In pFolder.SubFolders
        List1.AddItem UCase$(strSubFolder.Path)
        list2.AddItem UCase$(strSubFolder.Path)
        MapDirs (strSubFolder.Path)
        txtStatus.Text = "Updating"
        txtStatus.refresh
    Next
 
 DoEvents
    

    For Each strFile In pFolder.Files
    
        fname = sfolder + "\" + strFile.Name
        List1.AddItem LCase$(fname)
        list2.AddItem LCase$(fname)
        sbr.Panels(1).Text = "Scanning - " & strFile.Name
        txtStatus.Text = "Updating"
        txtStatus.refresh
    Next

DoEvents
    
End Function

Private Function check(sfolder As String)

Dim pFolder As Folder
Dim fname As String

On Error Resume Next
    
    Set pFolder = FSO.GetFolder(sfolder)
    
    For Each strSubFolder In pFolder.SubFolders
        List1.AddItem UCase$(strSubFolder.Path)
        check (strSubFolder.Path)
        txtStatus.Text = "Updating"
        txtStatus.refresh

    Next

DoEvents
    
    For Each strFile In pFolder.Files
    
        fname = sfolder + "\" + strFile.Name
        List1.AddItem LCase$(fname)
        sbr.Panels(1).Text = "Scanning - " & strFile.Name
        txtStatus.Text = "Updating"
        txtStatus.refresh
    
    Next
    
    
    DoEvents
    
End Function


Private Sub Timer1_Timer()

On Error Resume Next

Dim r1 As Integer
Dim r2 As Integer
Dim morethan As Integer

List1.Clear
check (strFolder)

r1 = 0
r2 = 0
 
DoEvents

j2:

While list2.ListCount > 0 And r2 <= list2.ListCount = True
txtStatus.Text = "Updating"
DoEvents

         While List1.ListCount > 0 And r1 <= List1.ListCount = True
txtStatus.Text = "Updating"
 DoEvents
    
                List1.ListIndex = r1
                list2.ListIndex = r2

                    If List1.ListCount > list2.ListCount Then
        
                            morethan = 1
        
                    ElseIf list2.ListCount > List1.ListCount Then
        
                            morethan = 2
                    Else
                    
                            morethan = 0
                
                    End If
        
        If List1.Text = list2.Text Then
            
                txtChanges.Text = " No Changes "
                
        Else

        If morethan = 1 Then
    
                list2.AddItem List1.Text, r1
                rtbChangedfiles.Text = rtbChangedfiles.Text & List1.Text & " " & Date & "  " & Time & " " & "Added!! " & vbCrLf
                txtChanges.Text = "Changes Detected!"
                counter = counter + 1
                Text5.Text = counter
                txtFiles.Text = list2.ListCount

        ElseIf morethan = 2 Then
        
                rtbChangedfiles.Text = rtbChangedfiles.Text & list2.Text & " " & Date & "  " & Time & " " & "Removed" & vbCrLf
                txtChanges.Text = "Changes Detected!"
                list2.RemoveItem (r2)
                list2.refresh
                txtFiles.Text = list2.ListCount
                counter = counter + 1
                Text5.Text = counter
        Else

            GoTo refresh
            
        
        End If
        
                r2 = r2 + 1
        
            GoTo j2
                        
                
        End If

                r1 = r1 + 1
                r2 = r2 + 1
            
    Wend
                DoEvents
Wend

refresh:


txtStatus.refresh
txtStatus.Text = "Await update"
Me.refresh
Timer1.Interval = Slider1.Value * 1000

End Sub
Private Sub StopTimer()

Timer1.Interval = 0
txtStatus.Text = "Idle"
txtStatus.refresh

End Sub
Private Sub Slider1_Change()

Dim updatetime As Long

updatetime = Slider1.Value

If Timer1.Interval <> 0 Then
    
    Timer1.Interval = (updatetime * 1000)
    txtTime = CStr(updatetime) & " Sec"

End If

End Sub
