VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{D85F17FA-1A65-4C49-9E12-15A5C27E81B6}#1.0#0"; "Downloader.ocx"
Begin VB.Form Form5 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "KuDos | Download queue"
   ClientHeight    =   5190
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   7920
   Icon            =   "Form5.frx":0000
   LinkTopic       =   "Form5"
   MaxButton       =   0   'False
   ScaleHeight     =   5190
   ScaleWidth      =   7920
   Begin VB.CommandButton Command1 
      Caption         =   "&Cancel"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   1320
      Visible         =   0   'False
      Width           =   855
   End
   Begin MSComctlLib.ProgressBar pb1 
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   1320
      Width           =   7215
      _ExtentX        =   12726
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin MSComctlLib.StatusBar sb1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   4935
      Width           =   7920
      _ExtentX        =   13970
      _ExtentY        =   450
      Style           =   1
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin Downloader.DownLoad dl 
      Left            =   600
      Top             =   3240
      _ExtentX        =   688
      _ExtentY        =   688
   End
   Begin MSComctlLib.ImageList imgDL 
      Left            =   1320
      Top             =   3600
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form5.frx":2372
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form5.frx":4B24
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form5.frx":4F76
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form5.frx":53C8
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lvDL 
      Height          =   3615
      Left            =   0
      TabIndex        =   0
      Top             =   1680
      Width           =   7935
      _ExtentX        =   13996
      _ExtentY        =   6376
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      SmallIcons      =   "imgDL"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "File:"
         Object.Width           =   7832
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Size (bytes):"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Progress:"
         Object.Width           =   3529
      EndProperty
   End
   Begin VB.Label lblRate 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0kb"
      Height          =   195
      Left            =   1200
      TabIndex        =   12
      Top             =   960
      Width           =   270
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Transfer rate:"
      Height          =   195
      Left            =   120
      TabIndex        =   11
      Top             =   960
      Width           =   945
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      X1              =   0
      X2              =   7920
      Y1              =   10
      Y2              =   10
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      X1              =   0
      X2              =   7920
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Label lblProgress 
      BackStyle       =   0  'Transparent
      Caption         =   "\"
      Height          =   255
      Left            =   1200
      TabIndex        =   9
      Top             =   720
      Width           =   2775
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Progress:"
      Height          =   195
      Left            =   120
      TabIndex        =   8
      Top             =   720
      Width           =   660
   End
   Begin VB.Label lblPercent 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0"
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
      Left            =   7440
      TabIndex        =   7
      Top             =   1320
      Width           =   375
   End
   Begin VB.Label lblLFile 
      BackStyle       =   0  'Transparent
      Caption         =   "-"
      Height          =   255
      Left            =   1200
      TabIndex        =   6
      Top             =   360
      Width           =   6735
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Local file:"
      Height          =   195
      Left            =   120
      TabIndex        =   5
      Top             =   360
      Width           =   675
   End
   Begin VB.Label lblRFile 
      BackStyle       =   0  'Transparent
      Caption         =   "--"
      Height          =   255
      Left            =   1200
      TabIndex        =   4
      Top             =   120
      Width           =   6735
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Remote file:"
      Height          =   195
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   840
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuTray 
         Caption         =   "Send to tray"
      End
   End
   Begin VB.Menu mnuDownload 
      Caption         =   "Download"
      Visible         =   0   'False
      Begin VB.Menu mnuCancel 
         Caption         =   "Cancel"
      End
      Begin VB.Menu mnuDelQueue 
         Caption         =   "Remove from Queue"
      End
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

On Error Resume Next

dl.Cancel

lblRFile = "N/A"
lblLFile = "N/A"
lblPercent = "N/A"
lblProgress = "N/A"
lblRate = "N/A"
pb1.Value = 0
doingDL = False

Kill strFilePath & "\" & lvDL.SelectedItem.Text

lvDL.ListItems(currIndex).ListSubItems(2).Text = "Download stopped."
lvDL.ListItems(currIndex).SmallIcon = 3

currIndex = currIndex + 1

doingDL = False
End Sub

Private Sub dl_DLComplete()

'On Error Resume Next

lvDL.ListItems(lvDL.SelectedItem.Index).ListSubItems(2).Text = "Download complete!"
lvDL.ListItems(lvDL.SelectedItem.Index).SmallIcon = 2

currIndex = lvDL.SelectedItem.Index + 1

If currIndex > lvDL.ListItems.Count Then
    lblRFile = "N/A"
    lblLFile = "N/A"
    lblPercent = "N/A"
    lblProgress = "N/A"
    lblRate = "N/A"
    pb1.Value = 0
    doingDL = False
    Exit Sub
Else
    lvDL.ListItems(currIndex).Selected = True
    lvDL.SelectedItem.SmallIcon = 1
    lblLFile.Caption = strFilePath & "\" & lvDL.SelectedItem.Text
    lblRFile.Caption = lvDL.SelectedItem.Key
    
    pb1.Max = lvDL.SelectedItem.ListSubItems(1).Text
    
    dl.URL = lvDL.SelectedItem.Key
    dl.GetFileInformation
    dl.SaveLocation = strFilePath & "\" & lvDL.SelectedItem.Text
    dl.DownLoad
End If

End Sub

Private Sub dl_DLError(lpErrorDescription As String)

On Error Resume Next

lvDL.ListItems(currIndex).SmallIcon = 3
lvDL.ListItems(currIndex).ListSubItems(2).Text = lpErrorDescription

currIndex = currIndex + 1

If currIndex > lvDL.ListItems.Count Then
    lblRFile = "N/A"
    lblLFile = "N/A"
    lblPercent = "N/A"
    lblProgress = "N/A"
    lblRate = "N/A"
    pb1.Value = 0
    doingDL = False
    Exit Sub
Else
    lvDL.ListItems(currIndex).Selected = True
    lvDL.SelectedItem.SmallIcon = 1
    lblLFile.Caption = strFilePath & "\" & lvDL.SelectedItem.Text
    lblRFile.Caption = lvDL.SelectedItem.Key
    dl.URL = lvDL.SelectedItem.Key
    dl.GetFileInformation
    dl.SaveLocation = strFilePath & "\" & lvDL.SelectedItem.Text
    dl.DownLoad
End If

End Sub

Private Sub dl_Percent(lPercent As Long)
lblPercent = lPercent & "%"
End Sub

Private Sub dl_Rate(lpRate As String)
lblRate.Caption = lpRate
End Sub

Private Sub dl_RecievedBytes(lnumBYTES As Long)
On Error Resume Next
lblProgress.Caption = lnumBYTES & "\" & lvDL.ListItems(currIndex).ListSubItems(1).Text
pb1.Value = lnumBYTES
End Sub

Private Sub dl_StatusChange(lpStatus As String)
On Error Resume Next
lvDL.ListItems(currIndex).ListSubItems(2).Text = lpStatus
End Sub

Private Sub Form_Load()
doingDL = True
End Sub

Private Sub Form_Paint()
On Error Resume Next
lvDL.Width = Me.ScaleWidth
lvDL.Height = Me.ScaleHeight - 2000
End Sub

Private Sub Form_Resize()
Form_Paint
End Sub

Private Sub lvDL_DblClick()

currIndex = lvDL.SelectedItem.Index

lvDL.SelectedItem.SmallIcon = 1
lblLFile.Caption = strFilePath & "\" & lvDL.SelectedItem.Text
lblRFile.Caption = lvDL.SelectedItem.Key

dl.URL = lvDL.SelectedItem.Key
dl.GetFileInformation
dl.SaveLocation = strFilePath & "\" & lvDL.SelectedItem.Text
dl.DownLoad

End Sub

Private Sub lvDL_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
If Button = 2 Then
    If lvDL.SelectedItem.SmallIcon = 1 Then
        mnuCancel.Enabled = True
    Else
        mnuCancel.Enabled = False
    End If
    PopupMenu mnuDownload
End If
End Sub

Private Sub mnuCancel_Click()
Command1_Click
End Sub

Private Sub mnuDelQueue_Click()
If lvDL.SelectedItem.SmallIcon = 1 Then
    If MsgBox("This will cancel the download." & vbCrLf & "Are you sure you want to continue?", vbQuestion + vbYesNo) = vbYes Then
        Command1_Click
        lvDL.ListItems.Remove (lvDL.SelectedItem.Index)
    End If
Else
    lvDL.ListItems.Remove (lvDL.SelectedItem.Index)
End If
End Sub


Private Sub mnuTray_Click()
Me.Hide
AddToTray Form1.Icon, "KuDos", Form1
End Sub
