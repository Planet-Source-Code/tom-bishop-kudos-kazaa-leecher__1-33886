VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   BackColor       =   &H80000000&
   Caption         =   "KuDos | KaZaa Leecher"
   ClientHeight    =   8880
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   12030
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8880
   ScaleWidth      =   12030
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame5 
      Caption         =   "Search results"
      Height          =   3255
      Left            =   120
      TabIndex        =   29
      Top             =   5280
      Width           =   5895
      Begin VB.ListBox lstSResults 
         Height          =   2790
         Left            =   120
         TabIndex        =   30
         Top             =   240
         Width           =   5655
      End
   End
   Begin VB.ComboBox cmbShow 
      Height          =   315
      Left            =   6720
      TabIndex        =   28
      Text            =   "Combo1"
      Top             =   0
      Width           =   3135
   End
   Begin MSComDlg.CommonDialog cd1 
      Left            =   6600
      Top             =   5040
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList imgFiles 
      Left            =   7080
      Top             =   5280
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   8
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":2372
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":290C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":2D5E
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":31B0
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":5532
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":67B4
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":8F66
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":B2E8
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lvFiles 
      Height          =   8175
      Left            =   6120
      TabIndex        =   25
      Top             =   360
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   14420
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      GridLines       =   -1  'True
      _Version        =   393217
      SmallIcons      =   "imgFiles"
      ForeColor       =   -2147483640
      BackColor       =   16777215
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Filename:"
         Object.Width           =   14887
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Size (bytes):"
         Object.Width           =   2540
      EndProperty
   End
   Begin MSComctlLib.ProgressBar pb1 
      Height          =   255
      Left            =   8900
      TabIndex        =   24
      Top             =   8640
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   19
      Top             =   8625
      Width           =   12030
      _ExtentX        =   21220
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   15584
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            TextSave        =   "4/17/2002"
         EndProperty
      EndProperty
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Index           =   0
      Left            =   5160
      Top             =   2760
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList imgIcons 
      Left            =   5160
      Top             =   2400
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":B73A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lvUsers 
      Height          =   2535
      Left            =   240
      TabIndex        =   6
      Top             =   240
      Width           =   5655
      _ExtentX        =   9975
      _ExtentY        =   4471
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      OLEDragMode     =   1
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      SmallIcons      =   "imgIcons"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      OLEDragMode     =   1
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "User:"
         Object.Width           =   6068
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "IP:"
         Object.Width           =   3352
      EndProperty
   End
   Begin VB.Frame Frame1 
      Caption         =   "Current users (click to browse)"
      Height          =   4455
      Left            =   120
      TabIndex        =   7
      Top             =   0
      Width           =   5895
      Begin VB.CheckBox Check1 
         Caption         =   "Only show users that are sharing files"
         Height          =   255
         Left            =   2760
         TabIndex        =   26
         Top             =   3960
         Width           =   2990
      End
      Begin VB.Frame Frame2 
         Caption         =   "Host scan:"
         Height          =   1335
         Left            =   120
         TabIndex        =   8
         Top             =   3000
         Width           =   5655
         Begin VB.TextBox StartGroup1 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   720
            TabIndex        =   0
            Text            =   "127"
            Top             =   240
            Width           =   375
         End
         Begin VB.TextBox StartGroup2 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   1200
            TabIndex        =   1
            Text            =   "0"
            Top             =   240
            Width           =   375
         End
         Begin VB.TextBox StartGroup3 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   1680
            TabIndex        =   2
            Text            =   "0"
            Top             =   240
            Width           =   375
         End
         Begin VB.TextBox StartGroup4 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   2160
            TabIndex        =   3
            Text            =   "0"
            Top             =   240
            Width           =   375
         End
         Begin VB.TextBox EndGroup1 
            Alignment       =   2  'Center
            Enabled         =   0   'False
            Height          =   285
            Left            =   720
            TabIndex        =   11
            Text            =   "127"
            Top             =   600
            Width           =   375
         End
         Begin VB.TextBox EndGroup2 
            Alignment       =   2  'Center
            Enabled         =   0   'False
            Height          =   285
            Left            =   1200
            TabIndex        =   10
            Text            =   "0"
            Top             =   600
            Width           =   375
         End
         Begin VB.TextBox EndGroup3 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   1680
            TabIndex        =   4
            Text            =   "255"
            Top             =   600
            Width           =   375
         End
         Begin VB.TextBox EndGroup4 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   2160
            TabIndex        =   5
            Text            =   "255"
            Top             =   600
            Width           =   375
         End
         Begin VB.CommandButton cmdStartScan 
            Caption         =   "Start"
            Height          =   615
            Left            =   4800
            TabIndex        =   9
            Top             =   240
            Width           =   735
         End
         Begin VB.Timer Timer1 
            Enabled         =   0   'False
            Interval        =   1000
            Left            =   5160
            Top             =   -240
         End
         Begin MSComctlLib.Slider Slider1 
            Height          =   255
            Left            =   3240
            TabIndex        =   12
            Top             =   600
            Width           =   1425
            _ExtentX        =   2514
            _ExtentY        =   450
            _Version        =   393216
            Min             =   -1000
            Max             =   -1
            SelStart        =   -1
            TickStyle       =   3
            Value           =   -1
         End
         Begin VB.Label Label4 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Scan status"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00004000&
            Height          =   270
            Left            =   720
            TabIndex        =   18
            Top             =   960
            Width           =   1815
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Start:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   195
            Left            =   120
            TabIndex        =   17
            Top             =   240
            Width           =   480
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Stop:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   195
            Left            =   120
            TabIndex        =   16
            Top             =   600
            Width           =   465
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Found:"
            Height          =   195
            Left            =   2640
            TabIndex        =   15
            Top             =   360
            Width           =   495
         End
         Begin VB.Label lblFound 
            BackStyle       =   0  'Transparent
            Caption         =   "0"
            Height          =   255
            Left            =   3240
            TabIndex        =   14
            Top             =   375
            Width           =   1455
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Speed:"
            Height          =   195
            Left            =   2640
            TabIndex        =   13
            Top             =   600
            Width           =   510
         End
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Search"
      Height          =   615
      Left            =   120
      TabIndex        =   20
      Top             =   4560
      Width           =   5895
      Begin VB.PictureBox Picture1 
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   120
         Picture         =   "Form1.frx":C014
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   32
         Top             =   240
         Width           =   255
      End
      Begin VB.Timer tmrTimeout 
         Enabled         =   0   'False
         Interval        =   5000
         Left            =   5520
         Top             =   600
      End
      Begin VB.CommandButton cmdSearch 
         Caption         =   "Search"
         Height          =   255
         Left            =   4920
         TabIndex        =   23
         Top             =   240
         Width           =   855
      End
      Begin VB.TextBox Search 
         Height          =   285
         Left            =   1560
         TabIndex        =   22
         Top             =   240
         Width           =   3255
      End
      Begin MSWinsockLib.Winsock wscHttp 
         Left            =   4800
         Top             =   0
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackColor       =   &H80000000&
         Caption         =   "Search for:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   195
         Left            =   480
         TabIndex        =   21
         Top             =   255
         Width           =   960
      End
   End
   Begin VB.TextBox Text2 
      Height          =   3375
      Left            =   2280
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   33
      Text            =   "Form1.frx":E7B6
      Top             =   4200
      Visible         =   0   'False
      Width           =   11655
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   9960
      TabIndex        =   31
      Top             =   30
      Width           =   75
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Show:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   255
      Left            =   6120
      TabIndex        =   27
      Top             =   40
      Width           =   2055
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuPreferences 
         Caption         =   "Preferences"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuScanResults 
      Caption         =   "&Users"
      Begin VB.Menu mnuAddUser 
         Caption         =   "Add User"
      End
      Begin VB.Menu mnuDelUser 
         Caption         =   "Delete User"
      End
      Begin VB.Menu mnuClearAll 
         Caption         =   "Clear All"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuAbout 
         Caption         =   "About"
      End
   End
   Begin VB.Menu mnuCopy 
      Caption         =   "COPY"
      Visible         =   0   'False
      Begin VB.Menu mnuCopyURL 
         Caption         =   "Copy URL"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private m_strRemoteHost As String      'Who do we connect to
Private m_strFilePath As String        'Remote file path. In this case always /
Private m_strHttpResponse As String    'Response from server. Used for parsing
Private m_bResponseReceived As Boolean 'Recieved response?
Private StillExecuting As Boolean      'Set to True when socket is active

Private Sub Check1_Click()
    'Save setting to registry
    SaveSetting "kudos", "config", "ShowSharing", Check1.Value
End Sub

Public Sub cmbShow_Click()
    'If nothing is in list then exit
    If Len(currentUser) = 0 Then Exit Sub
    
    'Setup our connection variables
    m_strRemoteHost = currentUser
    m_strFilePath = "/"
    m_strHttpResponse = ""
    m_bResponseReceived = False
    socketMode = "Browse"
    
    'Update status
    StatusBar1.Panels(1).Text = "Connecting to " & m_strRemoteHost & "..."
    Label8.Caption = "Connecting to " & m_strRemoteHost & "..."
    
    lvFiles.ListItems.Clear
    
    'Connect
    With wscHttp
        .Close
        .LocalPort = 0
        .Connect m_strRemoteHost, 1214
        tmrTimeout.Enabled = True
        StillExecuting = True
    End With
End Sub

Private Sub cmdSearch_Click()
    Dim i As Long
    Dim strTemp As String
    Dim strIP As Variant
    
    On Error Resume Next
        
        'See if the list is empty
        If lvUsers.ListItems.Count = 0 Then
            MsgBox "No scan results!" & vbCrLf & "Please do a host scan and try again.", vbCritical, "Error"
            Exit Sub
        End If
        
        'Prepare search
        lstSResults.Clear
        stopSearch = False
        StillExecuting = False
        Form2.Show , Me
        Form2.pb1.Max = lvUsers.ListItems.Count
        Form2.Top = Me.Top + 1500
        Form2.Left = Me.Left + 600
        tmrTimeout.Interval = GetSetting("kudos", "config", "SearchTimeOut") & "000"
        
        'Loop through all item's in lvUsers
        For i = 1 To lvUsers.ListItems.Count
            
            'Stop search is stopSearch is True
            If stopSearch Then
                wscHttp.Close
                StatusBar1.Panels(1).Text = "Search stopped."
                Exit For
            Else
                
                'Update status
                Form2.pb1.Value = i
                Form2.StatusBar1.SimpleText = "Searching " & lvUsers.ListItems(i).Text
                
                'Setup our connection variables
                m_strRemoteHost = lvUsers.ListItems(i).ListSubItems(1).Text
                m_strFilePath = "/"
                m_strHttpResponse = ""
                m_bResponseReceived = False
                socketMode = "Search"
                
                'Update status
                StatusBar1.Panels(1).Text = "Connecting to " & m_strRemoteHost
                
                'Connect
                With wscHttp
                    .Close
                    .LocalPort = 0
                    .Connect m_strRemoteHost, 1214
                    tmrTimeout.Enabled = True
                    StillExecuting = True
                End With
                
                'StillExecuting is the variable that I set to True when I connect.
                'Once the socket closes or error's out I set it to false
                While StillExecuting
                    DoEvents
                Wend
    
            End If
        
        'Get next item in lvUsers and start over
        Next i
        
        'Hide status screen
        Form2.Hide
End Sub

Private Sub cmdStartScan_Click()
    'This is par
    Select Case cmdStartScan.Caption
        Case "Start"
        
            Set db = OpenDatabase(App.Path & "\users.mdb")
    
            With db
                Set rstInfo = .OpenRecordset("users")
            End With
        
            Timer1.Enabled = True
            cmdStartScan.Caption = "Stop"
        Case "Stop"
            Timer1.Enabled = False
            cmdStartScan.Caption = "Start"
            
            Set db = Nothing
            Set rstInfo = Nothing
            
    End Select
End Sub



Private Sub EndGroup3_GotFocus()
SelectAll EndGroup3
End Sub

Private Sub EndGroup4_GotFocus()
SelectAll EndGroup4
End Sub

Private Sub Form_Load()

On Error Resume Next

Dim listx As ListItem

imLoading = True

StartGroup1.Text = 24
StartGroup2.Text = 66
Me.Caption = Me.Caption & " - " & Winsock1(0).LocalIP

Check1.Value = GetSetting("kudos", "config", "ShowSharing")

cmbShow.AddItem "All files"
cmbShow.AddItem "Video files"
cmbShow.AddItem "Picture files"
cmbShow.AddItem "Music files"
cmbShow.Text = "All files"

strFilePath = GetSetting("kudos", "config", "DLPath")

Set db = OpenDatabase(App.Path & "\users.mdb")

With db
    Set rstInfo = .OpenRecordset("users")
End With

Set dUsers = CreateObject("Scripting.Dictionary")
dUsers.CompareMode = BinaryCompare

With rstInfo
    If .RecordCount = 0 Then
        Exit Sub
    End If

    .MoveFirst
    
    While Not .EOF
        If dUsers.Exists(.Fields("user")) = False Then
            'dUsers.Add (.Fields("user")), 1
            Set listx = lvUsers.ListItems.Add(, , , , 1)
            listx.Text = .Fields("user")
            listx.SubItems(1) = .Fields("ip")
            lblFound = lblFound + 1
            .MoveNext
        Else
            .MoveNext
        End If
    Wend
End With

Set rstInfo = Nothing
Set db = Nothing

End Sub

Private Sub Start1_Change()
EndGroup1.Text = StartGroup1.Text
End Sub

Private Sub Start2_Change()
EndGroup2.Text = StartGroup2.Text
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
If Button = 2 Then
    PopupMenu mnuFile
End If
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
If Button = 2 Then
    PopupMenu mnuFile
Else
    If RespondToTray(X) <> 0 Then
        Call ShowFormAgain(Me)
        Form5.Show
    End If
End If
End Sub

Private Sub Form_Paint()
On Error Resume Next

lvFiles.Height = Me.ScaleHeight - 700
lvFiles.Width = Me.ScaleWidth - 6200

pb1.Top = StatusBar1.Top + 65
pb1.Height = StatusBar1.Height - 100
pb1.Width = StatusBar1.Panels(2).Width - 75
pb1.Left = StatusBar1.Panels(2).Left + 25

Frame5.Height = Me.ScaleHeight - 5600
lstSResults.Height = Frame5.Height - 300

    Form2.Top = Me.Top + 1500
    Form2.Left = Me.Left + 600

End Sub

Private Sub Form_Resize()
Form_Paint
End Sub

Private Sub lstSResults_Click()

If lstSResults.ListCount = 0 Then Exit Sub

    m_strRemoteHost = lstSResults.Text
    currentUser = lstSResults.Text
    m_strFilePath = "/"
    
    m_strHttpResponse = ""
    m_bResponseReceived = False
    
    StatusBar1.Panels(1).Text = "Connecting to " & m_strRemoteHost & "..."
    Label8.Caption = "Connecting to " & m_strRemoteHost & "..."
    socketMode = "Browse"
    
    lvFiles.ListItems.Clear
    
    With wscHttp
        .Close
        .LocalPort = 0
        .Connect m_strRemoteHost, 1214
        tmrTimeout.Enabled = True
        StillExecuting = True
    End With


End Sub

Private Sub lvFiles_DblClick()
Dim listx As ListItem

On Error Resume Next

If doingDL Then

    Set listx = Form5.lvDL.ListItems.Add(, , , , 4)
        listx.Text = lvFiles.SelectedItem.Text
        listx.SubItems(1) = lvFiles.SelectedItem.ListSubItems(1).Text
        listx.SubItems(2) = "Queued"
        listx.Key = lvFiles.SelectedItem.Key
    
    Form5.Show , Me
        
Else

    Set listx = Form5.lvDL.ListItems.Add(, , , , 1)
        listx.Text = lvFiles.SelectedItem.Text
        listx.SubItems(1) = lvFiles.SelectedItem.ListSubItems(1).Text
        listx.SubItems(2) = "Downloading..."
        listx.Key = lvFiles.SelectedItem.Key
        
        currIndex = listx.Index
        
        Form5.lblRFile = listx.Key
        Form5.lblLFile = listx.Text
        Form5.lblProgress = ""
        
        Form5.pb1.Max = lvFiles.SelectedItem.ListSubItems(1).Text
        
        Form5.Show , Me
        
        With Form5.dl
            .URL = lvFiles.SelectedItem.Key
            .GetFileInformation
            .SaveLocation = strFilePath & "\" & lvFiles.SelectedItem.Text
            .DownLoad
        End With
End If
End Sub

Private Sub lvFiles_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
If Button = 2 Then
    PopupMenu mnuCopy
End If
End Sub

Private Sub lvUsers_Click()
    On Error Resume Next
    If lvUsers.ListItems.Count = 0 Then Exit Sub

    m_strRemoteHost = lvUsers.SelectedItem.ListSubItems(1).Text
    currentUser = lvUsers.SelectedItem.ListSubItems(1).Text
    m_strFilePath = "/"
    
    m_strHttpResponse = ""
    m_bResponseReceived = False
    
    StatusBar1.Panels(1).Text = "Connecting to " & m_strRemoteHost & "..."
    Label8.Caption = "Connecting to " & m_strRemoteHost & "..."
    socketMode = "Browse"
    
    lvFiles.ListItems.Clear
    
    With wscHttp
        .Close
        .LocalPort = 0
        .Connect m_strRemoteHost, 1214
        tmrTimeout.Enabled = True
        StillExecuting = True
    End With
End Sub

Private Sub lvUsers_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
If Button = 2 Then
    PopupMenu mnuScanResults
End If
End Sub

Private Sub mnuAbout_Click()
MsgBox "KuDos | KaZaa Leecher" & vbCrLf & "Written by: Thomas Bishop" & vbCrLf & "n9productions@hotmail.com", vbInformation, "About"
End Sub

Private Sub mnuAddUser_Click()
Form4.Show , Me
End Sub

Private Sub mnuClearAll_Click()

If lvUsers.ListItems.Count = 0 Then Exit Sub

If MsgBox("Are you sure?", vbQuestion + vbYesNo, "Clear all") = vbYes Then
    
    Set db = OpenDatabase(App.Path & "\users.mdb")
    Set rstInfo = db.OpenRecordset("users")
    
    With rstInfo
        .MoveFirst
        While Not .EOF
            .Delete
            .MoveNext
        Wend
        lvUsers.ListItems.Clear
        lblFound.Caption = 0
        Set rstInfo = Nothing
        Set db = Nothing
    End With
End If
End Sub

Private Sub mnuCopyURL_Click()
On Error Resume Next
Clipboard.SetText lvFiles.SelectedItem.Text
End Sub

Private Sub mnuDelUser_Click()
If lvUsers.ListItems.Count = 0 Then Exit Sub

If MsgBox("Are you sure?", vbQuestion + vbYesNo, "Delete User") = vbYes Then
    
    Set db = OpenDatabase(App.Path & "\users.mdb")
    Set rstInfo = db.OpenRecordset("users")
    
    With rstInfo
        .Move lvUsers.SelectedItem.Index - 1
        .Delete
        lvUsers.ListItems.Remove (lvUsers.SelectedItem.Index)
        lblFound.Caption = lblFound.Caption - 1
        Set rstInfo = Nothing
        Set db = Nothing
    End With
End If

End Sub

Private Sub mnuExit_Click()
End
End Sub

Private Sub mnuPreferences_Click()
Form3.Show , Me
End Sub

Private Sub Search_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    cmdSearch_Click
End If
End Sub

Private Sub StartGroup1_Change()
EndGroup1.Text = StartGroup1.Text
End Sub

Private Sub StartGroup1_GotFocus()
SelectAll StartGroup1
End Sub


Private Sub StartGroup2_Change()
EndGroup2.Text = StartGroup2.Text
End Sub

Private Sub StartGroup2_GotFocus()
SelectAll StartGroup2
End Sub

Private Sub StartGroup3_GotFocus()
SelectAll StartGroup3
End Sub

Private Sub StartGroup4_GotFocus()
SelectAll StartGroup4
End Sub

Private Sub Timer1_Timer()

On Error Resume Next

Timer1.Interval = -Slider1.Value
TimesAround = TimesAround + 1

Load Winsock1(TimesAround) 'load new winsock
If TimesAround > 50 Then Unload Winsock1(TimesAround - 50) 'unload winsock control (time out)

If Val(StartGroup4) < Val(EndGroup4) Then
    StartGroup4.Text = StartGroup4.Text + 1 'increase current ip address by one
ElseIf Val(StartGroup4) = Val(EndGroup4) Then
    StartGroup3 = StartGroup3 + 1 'increase 3rd group in ip address by one
    StartGroup4 = 0 'reset last group in ip address
End If

If Val(StartGroup3) > Val(EndGroup3) Then 'check if scan is complete
    MsgBox "Scan Complete"
    Timer1.Enabled = False
    cmdStartScan.Caption = "Start"
End If

Winsock1(TimesAround).Connect StartGroup1 & "." & StartGroup2 & "." & StartGroup3 & "." & StartGroup4, 1214 'connect to potential kazaa user
Label4.Caption = StartGroup1 & "." & StartGroup2 & "." & StartGroup3 & "." & StartGroup4 'display current ip address
End Sub

Private Sub tmrTimeout_Timer()
tmrTimeout.Enabled = False
StillExecuting = False
End Sub

Private Sub Winsock1_Connect(Index As Integer)
Winsock1(Index).SendData "PASS Admin" & vbCrLf & "NICK M{iN}M" & vbCrLf & "USER KaZaAClone " & Winsock1(Index).LocalIP & ":KaZaA"
End Sub

Private Sub Winsock1_DataArrival(Index As Integer, ByVal bytesTotal As Long)
On Error GoTo X

Dim Data As String
Dim listx As ListItem

Winsock1(Index).GetData Data, vbString
'clean up data containing username
Data = Replace(Data, "HTTP/1.0 501 Not Implemented", "", 1, Len(Data), vbTextCompare)
Data = Replace(Data, "X-Kazaa-Username: ", "", 1, Len(Data), vbTextCompare)
Data = Replace(Data, Winsock1(Index).RemoteHostIP, "", 1, Len(Data), vbTextCompare)
Data = Replace(Data, "X-Kazaa-Network: KaZaA", "", 1, Len(Data), vbTextCompare)
Data = Replace(Data, " ", "", 1, Len(Data), vbTextCompare)
Data = Replace(Data, "X-Kazaa-IP:", "", 1, Len(Data), vbTextCompare)
Data = Replace(Data, ":1214", "", 1, Len(Data), vbTextCompare)
Data = Replace(Data, vbCrLf, "", 1, Len(Data), vbTextCompare)
Data = Replace(Data, "X-Kazaa-SupernodeIP:", "", 1, Len(Data), vbTextCompare)
Data = Replace(Data, "X-Kazaa-Network:", "", 1, Len(Data), vbTextCompare)
Data = Replace(Data, "MusicCity", "", 1, Len(Data), vbTextCompare)
Data = Replace(Data, ".", "", 1, Len(Data), vbTextCompare)
Data = Replace(Data, "0", "", 1, Len(Data), vbTextCompare)
Data = Replace(Data, "1", "", 1, Len(Data), vbTextCompare)
Data = Replace(Data, "2", "", 1, Len(Data), vbTextCompare)
Data = Replace(Data, "3", "", 1, Len(Data), vbTextCompare)
Data = Replace(Data, "4", "", 1, Len(Data), vbTextCompare)
Data = Replace(Data, "5", "", 1, Len(Data), vbTextCompare)
Data = Replace(Data, "6", "", 1, Len(Data), vbTextCompare)
Data = Replace(Data, "7", "", 1, Len(Data), vbTextCompare)
Data = Replace(Data, "8", "", 1, Len(Data), vbTextCompare)
Data = Replace(Data, "9", "", 1, Len(Data), vbTextCompare)
Data = Replace(Data, Chr(10), "", 1, Len(Data), vbTextCompare)
Data = Replace(Data, " ", "", 1, Len(Data), vbTextCompare)

If Check1.Value = 1 Then
    If Right(Data, 3) = "???" Then
        Set listx = lvUsers.ListItems.Add(, , , , 1)
        listx.Text = Data
        listx.SubItems(1) = Winsock1(Index).RemoteHostIP
        
        With rstInfo
            .AddNew
            !User = Data
            !ip = Winsock1(Index).RemoteHostIP
            .Update
        End With
    End If
Else
    Set listx = lvUsers.ListItems.Add(, , , , 1)
    listx.Text = Data
    listx.SubItems(1) = Winsock1(Index).RemoteHostIP
    
    With rstInfo
        .AddNew
        !User = Data
        !ip = Winsock1(Index).RemoteHostIP
        .Update
    End With
End If

lblFound = lblFound + 1
Winsock1(Index).Close
X:
End Sub

Public Sub SelectAll(Editctr As Control)
On Error Resume Next
    With Editctr
        .SelStart = 0
        .SelLength = Len(Editctr.Text)
        .SetFocus
    End With
End Sub

Private Sub wscHttp_Close()

    On Error Resume Next

    Dim strHttpResponseHeader As String
    Dim listx As ListItem
    Dim vFiles As Variant
    Dim vTemp As Variant
    Dim i As Long

    StatusBar1.Panels(1).Text = "Done."
    Label8.Caption = "Done."

    If Not m_bResponseReceived Then
        strHttpResponseHeader = Left$(m_strHttpResponse, _
                                InStr(1, m_strHttpResponse, _
                                vbCrLf & vbCrLf) - 1)
        m_strHttpResponse = Mid(m_strHttpResponse, _
                            InStr(1, m_strHttpResponse, _
                            vbCrLf & vbCrLf) + 4)
                            
        If InStr(1, m_strHttpResponse, "404") > 0 Then
            m_bResponseReceived = True
            tmrTimeout.Enabled = False
            StillExecuting = False
            Exit Sub
        End If
        
        If socketMode = "Search" Then
                   
                If InStr(1, LCase(m_strHttpResponse), LCase(Search)) > 0 Then
                    lstSResults.AddItem wscHttp.RemoteHostIP
                End If
            
        Else

                Text2.Text = m_strHttpResponse
                Text2.Text = Text2.Text
                Text2.Text = Replace(Text2.Text, Chr(10), vbCrLf, 1, Len(Text2.Text), vbTextCompare)
                Text2.Text = Replace(Text2.Text, "<tr><td><a href=", vbCrLf, 1, Len(Text2.Text), vbTextCompare)
                Text2.Text = Replace(Text2.Text, Chr(34) & ">", "|", 1, Len(Text2.Text), vbTextCompare)
                Text2.Text = Replace(Text2.Text, Chr(34), "", 1, Len(Text2.Text), vbTextCompare)
                Text2.Text = Replace(Text2.Text, "</a><td>", "|", 1, Len(Text2.Text), vbTextCompare)
                Text2.Text = Replace(Text2.Text, "<html><body><table>", "", 1, Len(Text2.Text), vbTextCompare)
                Text2.Text = Replace(Text2.Text, "</table></body></html>", "", 1, Len(Text2.Text), vbTextCompare)
                Text2.Text = Replace(Text2.Text, vbCrLf & vbCrLf, vbCrLf, 1, Len(Text2.Text), vbTextCompare)
                If Asc(Left(Text2.Text, 1)) = 13 Then Text2.Text = Right(Text2.Text, Len(Text2.Text) - 1)
                If Asc(Left(Text2.Text, 1)) = 10 Then Text2.Text = Right(Text2.Text, Len(Text2.Text) - 1)
                If Asc(Right(Text2.Text, 1)) = 13 Then Text2.Text = Left(Text2.Text, Len(Text2.Text) - 1)
                If Asc(Right(Text2.Text, 1)) = 10 Then Text2.Text = Left(Text2.Text, Len(Text2.Text) - 1)
                If Asc(Right(Text2.Text, 1)) = 13 Then Text2.Text = Left(Text2.Text, Len(Text2.Text) - 1)
            
                vFiles = Split(Text2.Text, vbCrLf)
                
                lvFiles.ListItems.Clear
                                   
                'Display results accordingly
                
                For i = 0 To UBound(vFiles)
                vTemp = Split(vFiles(i), "|")
                
                Select Case cmbShow.Text
                
                    Case "All files"
                    
                            If Right(LCase(vTemp(1)), 3) = "mpg" Or Right(LCase(vTemp(1)), 3) = "avi" Or Right(LCase(vTemp(1)), 4) = "mpeg" Or Right(LCase(vTemp(1)), 3) = "mov" Or Right(LCase(vTemp(1)), 3) = "wma" Then
                                
                                If Not Len(vTemp(1)) = 0 Then
                                
                                    Set listx = lvFiles.ListItems.Add(, , , , 5)
                                        listx.Text = vTemp(1)
                                        listx.SubItems(1) = vTemp(2)
                                        listx.Key = "http://" & wscHttp.RemoteHostIP & ":1214" & vTemp(0)
        
                                    If Len(Search) > 0 Then
                                        If InStr(1, LCase(vTemp(1)), LCase(Search)) > 0 Then
                                            listx.Bold = True
                                            listx.ForeColor = &HFF&
                                        End If
                                    End If
                                    
                                End If
    
                                Label8.Caption = "Browsing: [" & lvUsers.SelectedItem.Text & "] " & lvFiles.ListItems.Count & " files"
    
                            ElseIf Right(LCase(vTemp(1)), 3) = "wav" Or Right(LCase(vTemp(1)), 3) = "snd" Or Right(LCase(vTemp(1)), 3) = "mp3" Then
    
                                If Not Len(vTemp(1)) = 0 Then
                                
                                    Set listx = lvFiles.ListItems.Add(, , , , 4)
                                        listx.Text = vTemp(1)
                                        listx.SubItems(1) = vTemp(2)
                                        listx.Key = "http://" & wscHttp.RemoteHostIP & ":1214" & vTemp(0)
        
                                    If Len(Search) > 0 Then
                                        If InStr(1, LCase(vTemp(1)), LCase(Search)) > 0 Then
                                            listx.Bold = True
                                            listx.ForeColor = &HFF&
                                        End If
                                    End If
                                    
                                End If
    
                                Label8.Caption = "Browsing: [" & lvUsers.SelectedItem.Text & "] " & lvFiles.ListItems.Count & " files"
    
                            ElseIf Right(LCase(vTemp(1)), 3) = "jpg" Or Right(LCase(vTemp(1)), 3) = "bmp" Or Right(LCase(vTemp(1)), 3) = "gif" Then
    
                                If Not Len(vTemp(1)) = 0 Then
                                
                                    Set listx = lvFiles.ListItems.Add(, , , , 6)
                                        listx.Text = vTemp(1)
                                        listx.SubItems(1) = vTemp(2)
                                        listx.Key = "http://" & wscHttp.RemoteHostIP & ":1214" & vTemp(0)
        
                                    If Len(Search) > 0 Then
                                        If InStr(1, LCase(vTemp(1)), LCase(Search)) > 0 Then
                                            listx.Bold = True
                                            listx.ForeColor = &HFF&
                                        End If
                                    End If
                                    
                                End If
    
                                Label8.Caption = "Browsing: [" & lvUsers.SelectedItem.Text & "] " & lvFiles.ListItems.Count & " files"
    
                            ElseIf Right(LCase(vTemp(1)), 3) = "vbs" Then
    
                                If Not Len(vTemp(1)) = 0 Then
                                
                                    Set listx = lvFiles.ListItems.Add(, , , , 8)
                                        listx.Text = vTemp(1)
                                        listx.SubItems(1) = vTemp(2)
                                        listx.Key = "http://" & wscHttp.RemoteHostIP & ":1214" & vTemp(0)
        
                                    If Len(Search) > 0 Then
                                        If InStr(1, LCase(vTemp(1)), LCase(Search)) > 0 Then
                                            listx.Bold = True
                                            listx.ForeColor = &HFF&
                                        End If
                                    End If
                                    
                                End If
    
                                Label8.Caption = "Browsing: [" & lvUsers.SelectedItem.Text & "] " & lvFiles.ListItems.Count & " files"
    
                            Else
    
                                If Not Len(vTemp(1)) = 0 Then
                                
                                    Set listx = lvFiles.ListItems.Add(, , , , 7)
                                        listx.Text = vTemp(1)
                                        listx.SubItems(1) = vTemp(2)
                                        listx.Key = "http://" & wscHttp.RemoteHostIP & ":1214" & vTemp(0)
        
                                    If Len(Search) > 0 Then
                                        If InStr(1, LCase(vTemp(1)), LCase(Search)) > 0 Then
                                            listx.Bold = True
                                            listx.ForeColor = &HFF&
                                        End If
                                    End If
                                    
                                End If
    
                                Label8.Caption = "Browsing: [" & lvUsers.SelectedItem.Text & "] " & lvFiles.ListItems.Count & " files"
    
                            End If
                        
                    Case "Video files"
                    
                        If Right(vTemp(1), 3) = "mpg" Or Right(vTemp(1), 3) = "avi" Or Right(vTemp(1), 4) = "mpeg" Or Right(vTemp(1), 3) = "mov" Or Right(vTemp(1), 3) = "wma" Then
                            
                            If Not Len(vTemp(1)) = 0 Then
                            
                                Set listx = lvFiles.ListItems.Add(, , , , 5)
                                    listx.Text = vTemp(1)
                                    listx.SubItems(1) = vTemp(2)
                                    listx.Key = "http://" & wscHttp.RemoteHostIP & ":1214" & vTemp(0)
    
                                If Len(Search) > 0 Then
                                    If InStr(1, LCase(vTemp(1)), LCase(Search)) > 0 Then
                                        listx.Bold = True
                                        listx.ForeColor = &HFF&
                                    End If
                                End If
                                
                            End If
                                
                            Label8.Caption = "Browsing: [" & lvUsers.SelectedItem.Text & "] " & lvFiles.ListItems.Count & " files"
        
                        End If
                        
                    Case "Picture files"
                        
                        If Right(vTemp(1), 3) = "jpg" Or Right(vTemp(1), 3) = "gif" Or Right(vTemp(1), 3) = "bmp" Then
                            
                            If Not Len(vTemp(1)) = 0 Then
                            
                                Set listx = lvFiles.ListItems.Add(, , , , 6)
                                    listx.Text = vTemp(1)
                                    listx.SubItems(1) = vTemp(2)
                                    listx.Key = "http://" & wscHttp.RemoteHostIP & ":1214" & vTemp(0)
    
                                If Len(Search) > 0 Then
                                    If InStr(1, LCase(vTemp(1)), LCase(Search)) > 0 Then
                                        listx.Bold = True
                                        listx.ForeColor = &HFF&
                                    End If
                                End If
                                
                            End If
                                
                            Label8.Caption = "Browsing: [" & lvUsers.SelectedItem.Text & "] " & lvFiles.ListItems.Count & " files"
        
                        End If
                        
                    Case "Music files"
                    
                        If Right(vTemp(1), 3) = "mp3" Or Right(vTemp(1), 3) = "snd" Or Right(vTemp(1), 4) = "wav" Then
                            
                            If Not Len(vTemp(1)) = 0 Then
                            
                                Set listx = lvFiles.ListItems.Add(, , , , 4)
                                    listx.Text = vTemp(1)
                                    listx.SubItems(1) = vTemp(2)
                                    listx.Key = "http://" & wscHttp.RemoteHostIP & ":1214" & vTemp(0)
    
                                If Len(Search) > 0 Then
                                    If InStr(1, LCase(vTemp(1)), LCase(Search)) > 0 Then
                                        listx.Bold = True
                                        listx.ForeColor = &HFF&
                                    End If
                                End If
                                
                            End If
                                
                            Label8.Caption = "Browsing: [" & lvUsers.SelectedItem.Text & "] " & lvFiles.ListItems.Count & " files"
        
                        End If
        
                 End Select
                 
                 Next i
                                                  
                If lvFiles.ListItems.Count = 0 Then
                    Set listx = lvFiles.ListItems.Add
                        listx.Text = "No item's to display..."
                        Label8.Caption = ""
                        StatusBar1.Panels(1).Text = ""
                End If
                        
        End If
        
        m_bResponseReceived = True
        
    End If
    
    tmrTimeout.Enabled = False
    StillExecuting = False
    
End Sub

Private Sub wscHttp_Connect()

    Dim strHttpRequest As String
    
    tmrTimeout.Enabled = False
    StatusBar1.Panels(1).Text = "Connected!"
    Label8.Caption = "Connected!"

    'MsgBox m_strFilePath

    strHttpRequest = "GET " & m_strFilePath & " HTTP/1.1" & vbCrLf
    strHttpRequest = strHttpRequest & "Host: " & m_strRemoteHost & vbCrLf
    strHttpRequest = strHttpRequest & "Connection: close" & vbCrLf
    strHttpRequest = strHttpRequest & "Accept: */*" & vbCrLf
    strHttpRequest = strHttpRequest & vbCrLf

    wscHttp.SendData strHttpRequest
End Sub

Private Sub wscHttp_DataArrival(ByVal bytesTotal As Long)
    On Error Resume Next
    Dim strData As String

    StatusBar1.Panels(1).Text = "Reading data..."
    Label8.Caption = "Reading data..."
    
    wscHttp.GetData strData
    m_strHttpResponse = m_strHttpResponse & strData
End Sub

Private Sub wscHttp_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
tmrTimeout.Enabled = False
StillExecuting = False
End Sub
