VERSION 5.00
Object = "{FE0065C0-1B7B-11CF-9D53-00AA003C9CB6}#1.1#0"; "COMCT232.OCX"
Begin VB.Form Form3 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Preferences"
   ClientHeight    =   2985
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4335
   Icon            =   "Form3.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2985
   ScaleWidth      =   4335
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Caption         =   "Download path:"
      Height          =   1095
      Left            =   120
      TabIndex        =   7
      Top             =   1320
      Width           =   4095
      Begin VB.TextBox txtDLPath 
         Height          =   285
         Left            =   120
         TabIndex        =   9
         Top             =   720
         Width           =   3855
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "This is where all you files downloaded by KuDos will be stored."
         Height          =   375
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   3855
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   3120
      TabIndex        =   6
      Top             =   2520
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&OK"
      Height          =   375
      Left            =   2040
      TabIndex        =   5
      Top             =   2520
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Caption         =   "Search timeout:"
      Height          =   1095
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4095
      Begin ComCtl2.UpDown UpDown1 
         Height          =   285
         Left            =   3720
         TabIndex        =   3
         Top             =   720
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   503
         _Version        =   327681
         Value           =   1
         BuddyControl    =   "txtSTimeOut"
         BuddyDispid     =   196612
         OrigLeft        =   3600
         OrigTop         =   600
         OrigRight       =   3840
         OrigBottom      =   855
         Max             =   15
         Min             =   1
         SyncBuddy       =   -1  'True
         BuddyProperty   =   0
         Enabled         =   -1  'True
      End
      Begin VB.TextBox txtSTimeOut 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   3000
         TabIndex        =   2
         Text            =   "0"
         Top             =   720
         Width           =   735
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Timeout value (seconds):"
         Height          =   255
         Left            =   1080
         TabIndex        =   4
         Top             =   720
         Width           =   1815
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "The slower your connection is the higher this value should be."
         Height          =   495
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   3855
      End
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
SaveSetting "kudos", "config", "SearchTimeOut", txtSTimeOut.Text
If Len(txtDLPath) = 0 Then
    txtDLPath = App.Path
End If
SaveSetting "kudos", "config", "DLPath", txtDLPath.Text
strFilePath = txtDLPath.Text
Unload Me
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Load()
txtSTimeOut.Text = GetSetting("kudos", "config", "SearchTimeOut")
txtDLPath.Text = GetSetting("kudos", "config", "DLPath")
End Sub

