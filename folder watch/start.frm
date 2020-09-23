VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Folder Watch V 1.0"
   ClientHeight    =   2910
   ClientLeft      =   1095
   ClientTop       =   1200
   ClientWidth     =   12330
   DrawWidth       =   2
   Icon            =   "start.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2910
   ScaleWidth      =   12330
   Begin VB.Frame Frame6 
      Height          =   615
      Left            =   120
      TabIndex        =   26
      Top             =   2280
      Width           =   6375
      Begin VB.Label Label12 
         Caption         =   "Sec"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5880
         TabIndex        =   36
         Top             =   0
         Width           =   255
      End
      Begin VB.Label Label11 
         Caption         =   "Min"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5400
         TabIndex        =   35
         Top             =   0
         Width           =   255
      End
      Begin VB.Label Label10 
         Caption         =   "Hour"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4800
         TabIndex        =   34
         Top             =   0
         Width           =   375
      End
      Begin VB.Line Line1 
         X1              =   3480
         X2              =   3480
         Y1              =   120
         Y2              =   600
      End
      Begin VB.Label Label8 
         Caption         =   ":"
         Height          =   255
         Left            =   5760
         TabIndex        =   33
         Top             =   240
         Width           =   135
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   120
         TabIndex        =   32
         Top             =   240
         Width           =   3135
      End
      Begin VB.Label Label5 
         Caption         =   ":"
         Height          =   255
         Left            =   5160
         TabIndex        =   31
         Top             =   240
         Width           =   135
      End
      Begin VB.Label lbMin 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "00"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   5400
         TabIndex        =   30
         Top             =   240
         Width           =   225
      End
      Begin VB.Label lbsec 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "00"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   5940
         TabIndex        =   29
         Top             =   240
         Width           =   225
      End
      Begin VB.Label lbHour 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "00"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   4800
         TabIndex        =   28
         Top             =   240
         Width           =   345
      End
      Begin VB.Label Label2 
         Caption         =   "Run Time"
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   3840
         TabIndex        =   27
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.Timer Timer2 
      Left            =   2160
      Top             =   5640
   End
   Begin VB.Timer Timer1 
      Left            =   1680
      Top             =   5640
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
      ForeColor       =   &H000000FF&
      Height          =   2295
      Left            =   3600
      TabIndex        =   9
      Top             =   0
      Width           =   3015
      Begin MSComctlLib.Slider Slider1 
         Height          =   2055
         Left            =   2520
         TabIndex        =   21
         Top             =   120
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   3625
         _Version        =   393216
         Orientation     =   1
         Min             =   10
         Max             =   60
         SelStart        =   10
         Value           =   10
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
         ForeColor       =   &H000000FF&
         Height          =   1575
         Left            =   120
         TabIndex        =   12
         Top             =   600
         Width           =   2415
         Begin VB.TextBox txtFiles 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1080
            Locked          =   -1  'True
            TabIndex        =   19
            TabStop         =   0   'False
            Text            =   "0"
            Top             =   960
            Width           =   1215
         End
         Begin VB.TextBox txtStatus 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   720
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
            Left            =   720
            Locked          =   -1  'True
            TabIndex        =   13
            TabStop         =   0   'False
            Text            =   "No Changes"
            Top             =   600
            Width           =   1575
         End
         Begin VB.Label Label1 
            Caption         =   "Total Folders/Files"
            Height          =   375
            Left            =   120
            TabIndex        =   20
            Top             =   960
            Width           =   975
         End
         Begin VB.Label Label3 
            Caption         =   "Status"
            Height          =   255
            Left            =   120
            TabIndex        =   16
            Top             =   240
            Width           =   735
         End
         Begin VB.Label Label4 
            Caption         =   "Change"
            Height          =   255
            Left            =   120
            TabIndex        =   15
            Top             =   600
            Width           =   615
         End
      End
      Begin VB.TextBox txtTime 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   11
         TabStop         =   0   'False
         Text            =   "10 Sec"
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label7 
         Caption         =   "Timer Scan Rate"
         Height          =   255
         Left            =   120
         TabIndex        =   25
         Top             =   240
         Width           =   1215
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
      ForeColor       =   &H000000FF&
      Height          =   2895
      Left            =   6600
      TabIndex        =   7
      Top             =   0
      Width           =   5655
      Begin VB.CommandButton Command1 
         Caption         =   "Clear &log"
         Height          =   375
         Left            =   120
         TabIndex        =   24
         Top             =   2400
         Width           =   1215
      End
      Begin VB.TextBox Text5 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   4560
         TabIndex        =   22
         Text            =   "0"
         Top             =   2520
         Width           =   975
      End
      Begin RichTextLib.RichTextBox rtbChangedfiles 
         Height          =   2055
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   5415
         _ExtentX        =   9551
         _ExtentY        =   3625
         _Version        =   393217
         ScrollBars      =   3
         Appearance      =   0
         TextRTF         =   $"start.frx":0ECA
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
         Left            =   3840
         TabIndex        =   23
         Top             =   2520
         Width           =   735
      End
   End
   Begin MSComctlLib.StatusBar sbr 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   6
      Top             =   2535
      Width           =   12330
      _ExtentX        =   21749
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Bevel           =   0
            Object.Width           =   16563
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            TextSave        =   "1/12/2006"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            TextSave        =   "9:20 AM"
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
      ForeColor       =   &H000000FF&
      Height          =   735
      Left            =   120
      TabIndex        =   3
      Top             =   0
      Width           =   3375
      Begin VB.CommandButton mini 
         Caption         =   "&Minimize"
         Height          =   375
         Left            =   1200
         TabIndex        =   10
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton cmdstart 
         Caption         =   "&Start"
         Height          =   375
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton cmdExit 
         Caption         =   "E&xit"
         Height          =   375
         Left            =   2280
         TabIndex        =   4
         Top             =   240
         Width           =   975
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
      ForeColor       =   &H000000FF&
      Height          =   1575
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   3375
      Begin VB.DriveListBox Drive1 
         Height          =   315
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   3015
      End
      Begin VB.DirListBox strpath 
         Height          =   765
         Left            =   120
         TabIndex        =   1
         Top             =   600
         Width           =   3015
      End
   End
   Begin VB.ListBox master 
      Height          =   255
      ItemData        =   "start.frx":0F4E
      Left            =   120
      List            =   "start.frx":0F50
      TabIndex        =   18
      Top             =   6120
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.ListBox checked 
      Height          =   255
      ItemData        =   "start.frx":0F52
      Left            =   120
      List            =   "start.frx":0F54
      TabIndex        =   17
      Top             =   6000
      Visible         =   0   'False
      Width           =   615
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
      Begin VB.Menu mnulast 
         Caption         =   "S&how last"
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
''''''''    FOLDER WATCH
''''''''    BY JAMES SPROULL
''''''''
''''''''
''''''''
''''''''    Purpose of Program: To Scan FTP server(or any other folder/Drive)for
''''''''                        New or Removed Files/Folders
''''''''
''''''''
''''''''    Features:
''''''''
''''''''        Scans Subdirectories
''''''''        Has Menu popups and Log files
''''''''        Auto saves log in app folder
''''''''        Keeps track of how long prg has been running
''''''''        Tested on 30k files on p4 2.2gig 1gig mem for over 24 hours
''''''''
''''''''    ORIGINAL SOURCES
''''''''
''''''''    FOLDER GUARD : http://www.Planet-Source-Code.com/vb/scripts/ShowCode.asp?txtCodeId=31350&lngWId=1
''''''''    TRAY         : http://www.Planet-Source-Code.com/vb/scripts/ShowCode.asp?txtCodeId=59003&lngWId=1
''''''''    DIR SCAN     : http://www.Planet-Source-Code.com/vb/scripts/ShowCode.asp?txtCodeId=46404&lngWId=1
''''''''    POPUP TRAY   : http://www.Planet-Source-Code.com/vb/scripts/ShowCode.asp?txtCodeId=27491&lngWId=1
''''''''
''''''''
''''''''    THANKS GUYS!
''''''''    I hope I gave credit were credit is Due!
''''''''    If Not Please let me know so I may Correct it.
''''''''
''''''''    Will take all comments/sugestions/critisism
''''''''
''''''''
''''''''    SET REFERENCES
''''''''    microsoft scripting runtime
''''''''
''''''''
''''''''
''''''''
''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim counter As Integer ' reportlog counter
Dim allertup1 As String ' menu popup
Dim fcount As Integer ' file counter
Dim folcount As Integer ' folder counter

Option Explicit

Private Sub Command1_Click()

'clears report

rtbChangedfiles.Text = " "
Text5.Text = "0"
counter = 0
Text5.refresh

End Sub

Private Sub mini_Click()

'minimizes window to system tray

TrayAdd hWnd, Me.Icon, "Folder Watch", MouseMove
    MnuHide_Click

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

' mouse move events

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

Private Sub MnuExit_Click()

'exits prg
    
    cmdExit_Click

End Sub

Private Sub MnuHide_Click()

'minimize routine

    If Not Me.WindowState = 1 Then WindowState = 1: Me.Hide
    
End Sub

Private Sub mnulog_Click()

' bring up reporlog from menu
    
    Form2.Show
    
End Sub

Private Sub MnuShow_Click()

' max's prg

    If Me.WindowState = 1 Then WindowState = 0: Me.Show
    TrayDelete
        
End Sub
Public Sub mnulast_Click()

' brings up last popup

    If allertup1 = "" Then allertup1 = " No Entries "
    
    Alert allertup1
    
End Sub
Private Sub cmdClr_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

'mouse move events
    
    sbr.Panels(1).Text = "Clear Report"

End Sub
Private Sub cmdclr_click()

'clears reportlog

rtbChangedfiles = " "

End Sub

Private Sub cmdExit_Click()

'exits

TrayDelete
    
    End

End Sub

Private Sub cmdExit_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

'mouse move
    
    sbr.Panels(1).Text = "Exit"

End Sub

Private Sub mini_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

'mouse move
    
    sbr.Panels(1).Text = "Minimize to System Tray"

End Sub

Private Sub cmdStart_Click()

' starts scanning folders/files to populate master list
    
    cmdstart.Enabled = False
    strScanDir = strpath.List(strpath.ListIndex)
    Set strFolder = FSO.GetFolder(strScanDir)
    Scandirs (strFolder)
    txtFiles.Text = master.ListCount
    txtStatus.Text = "Folder Watch Started"
    Text5.Text = "0"
    txtTime.Text = Slider1.Value & " Sec"
    Timer1.Interval = Slider1.Value * 1000
    TrayAdd hWnd, Me.Icon, "Folder Watch", MouseMove
    Label6.Caption = "Started  " & " " & Date & "  " & Time
    Timer2.Enabled = True
    Timer2.Interval = 1000
   
End Sub

Private Sub cmdstart_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

'mouse move

    sbr.Panels(1).Text = "Scan Selected Directory"

End Sub

Private Sub strpath_Change()
    
    ChDir strpath.Path

End Sub

Private Sub strpath_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

'mouse move

    sbr.Panels(1).Text = "Select Directory to Scan"

End Sub

Private Sub command1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

'mouse move

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

Private Function Scandirs(sfolder As String)

'populates master list

Dim pFolder As Folder
Dim fname As String

On Error Resume Next
    
    Set pFolder = FSO.GetFolder(sfolder)
    For Each strSubFolder In pFolder.SubFolders
        master.AddItem UCase$(strSubFolder.Path)
        Scandirs (strSubFolder.Path)
        txtStatus.Text = "Updating"
        txtStatus.refresh
        'folcount = folcount + 1
        'Text2.Text = folcount
        DoEvents
        Timer2.Interval = 1000
    Next
 
 DoEvents
    
    For Each strFile In pFolder.Files
    
        fname = sfolder + strFile.Name
        master.AddItem LCase$(fname)
        sbr.Panels(1).Text = "Scanning - " & strFile.Name
        txtStatus.Text = "Updating"
        txtStatus.refresh
        'fcount = fcount + 1
        'Text1.Text = fcount
        DoEvents
        Timer2.Interval = 1000
    Next

DoEvents

End Function

Private Function check(sfolder As String)

'populates checked list for comparision later

Dim pFolder As Folder
Dim fname As String

On Error Resume Next
    
    Set pFolder = FSO.GetFolder(sfolder)
    For Each strSubFolder In pFolder.SubFolders
        checked.AddItem UCase$(strSubFolder.Path)
        check (strSubFolder.Path)
        txtStatus.Text = "Updating"
        txtStatus.refresh

    Next

DoEvents
    
    For Each strFile In pFolder.Files
        fname = sfolder + strFile.Name
        checked.AddItem LCase$(fname)
        sbr.Panels(1).Text = "Scanning - " & strFile.Name
        txtStatus.Text = "Updating"
        txtStatus.refresh
    
    Next
     
     Timer2.Interval = 1000
     DoEvents
    
End Function

Private Sub Timer1_Timer()

' master timer loop

On Error Resume Next

Dim allertup As String

Dim r1 As Integer
Dim r2 As Integer
Dim morethan As Integer

checked.Clear
check (strFolder)

r1 = 0
r2 = 0
 
DoEvents

j2:

While master.ListCount > 0 And r2 <= master.ListCount = True

    txtStatus.Text = "Updating"
    DoEvents
    Timer2.Interval = 1000
         While checked.ListCount > 0 And r1 <= checked.ListCount = True

    txtStatus.Text = "Updating"
    DoEvents
    Timer2.Interval = 1000
                checked.ListIndex = r1
                master.ListIndex = r2

        If checked.ListCount > master.ListCount Then
        
                            morethan = 1
        
        ElseIf master.ListCount > checked.ListCount Then
        
                            morethan = 2
        Else
                    
                            morethan = 0
                
        End If
        
        If checked.Text = master.Text Then
            
                            txtChanges.Text = " No Changes "
                
        Else

        If morethan = 1 Then
         
                master.AddItem checked.Text, r1
                allertup = Date & "  " & Time & vbCrLf & checked.Text & vbCrLf & "Added"
                allertup1 = Date & "  " & Time & vbCrLf & checked.Text & vbCrLf & "Added"
                rtbChangedfiles.Text = rtbChangedfiles.Text + checked.Text & " " & Date & "  " & Time & " " & "Added!! " & vbCrLf
                saveit
                Alert allertup
                allertup = " "
                txtChanges.Text = "Changes Detected!"
                counter = counter + 1
                Text5.Text = counter
                txtFiles.Text = master.ListCount

        ElseIf morethan = 2 Then
        
                allertup = Date & "  " & Time & vbCrLf & master.Text & vbCrLf & "Removed"
                allertup1 = Date & "  " & Time & vbCrLf & master.Text & vbCrLf & "Removed"
                rtbChangedfiles.Text = rtbChangedfiles.Text + master.Text & " " & Date & "  " & Time & " " & "Removed" & vbCrLf
                saveit
                Alert allertup
                allertup = ""
                txtChanges.Text = "Changes Detected!"
                master.RemoveItem (r2)
                master.refresh
                txtFiles.Text = master.ListCount
                counter = counter + 1
                Text5.Text = counter
        Else

            GoTo refresh
            
        End If
        
            GoTo j2
                          
        End If

                r1 = r1 + 1
                r2 = r2 + 1
            
                DoEvents
    
    Wend
                
                DoEvents

Wend

refresh:
txtStatus.refresh
txtStatus.Text = "Awaiting Update"
Me.refresh
Timer1.Interval = Slider1.Value * 1000
Timer2.Interval = 1000
DoEvents

End Sub

Public Sub saveit()

' save reportlog to file reportlog.txt

Open App.Path + "\reportlog.txt" For Append As #2
 
    Print #2, rtbChangedfiles.Text

    Close #2
    
End Sub

Private Sub Slider1_Change()

'slider for adjusting scan time

Dim updatetime As Long

updatetime = Slider1.Value

If Timer1.Interval <> 0 Then
    
    Timer1.Interval = (updatetime * 1000)
    txtTime = CStr(updatetime) & " Sec"

End If

End Sub

Private Sub Alert(Text As String)

' alert box popup

    Dim AlertBox As frmAlert
    
    Set AlertBox = New frmAlert
    
    AlertBox.DisplayAlert Text, 5000

End Sub

Private Sub Timer2_Timer()

' keeps track of how long prog has been running
        
   Dim Tsec As Integer
   Dim Tmin As Integer
   Dim Thour As Integer
    
    Tsec = lbsec.Caption + 1

    If Tsec < 10 Then
        
        Tsec = "00" & Tsec
        lbsec.Caption = Tsec
        lbsec.Caption = Format(lbsec.Caption, "00")
    
    Else
        
        lbsec.Caption = Tsec
        lbsec.Caption = Format(lbsec.Caption, "00")
    
    End If

    If Tsec = 60 Then
        
        Tmin = lbMin.Caption + 1
        lbMin.Caption = Format(lbMin.Caption, "00")

        If Tmin < 10 Then
            
            Tmin = "00" & Tmin
            lbMin.Caption = Tmin
            lbMin.Caption = Format(lbMin.Caption, "00")
        
        Else
            
            lbMin.Caption = Tmin
            lbMin.Caption = Format(lbMin.Caption, "00")
        
        End If
        
        lbsec.Caption = "01"
        lbsec.Caption = Format(lbsec.Caption, "00")
    
    End If

    If Tmin = 60 Then
        
        Thour = lbHour.Caption + 1
        lbHour.Caption = Format(lbHour.Caption, "00")
        
        If Thour < 10 Then
            
            Thour = "00" & Thour
            lbHour.Caption = Thour
            lbHour.Caption = Format(lbHour.Caption, "00")
        
        Else
            
            lbHour.Caption = Thour
            lbHour.Caption = Format(lbHour.Caption, "00")
        
        End If
        
        lbMin.Caption = "01"
        lbMin.Caption = Format(lbMin.Caption, "00")
    
    End If

End Sub
