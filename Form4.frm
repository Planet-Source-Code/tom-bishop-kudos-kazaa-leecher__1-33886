VERSION 5.00
Begin VB.Form Form4 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Add User"
   ClientHeight    =   2040
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4215
   Icon            =   "Form4.frx":0000
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2040
   ScaleWidth      =   4215
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox Check1 
      Caption         =   "Connect right away"
      Height          =   195
      Left            =   120
      TabIndex        =   6
      Top             =   1680
      Width           =   2055
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   3240
      TabIndex        =   5
      Top             =   1560
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&OK"
      Height          =   375
      Left            =   2280
      TabIndex        =   4
      Top             =   1560
      Width           =   855
   End
   Begin VB.Frame Frame2 
      Caption         =   "IP"
      Height          =   615
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   3975
      Begin VB.TextBox txtIP 
         Height          =   285
         Left            =   120
         TabIndex        =   3
         Text            =   "127.0.0.1"
         Top             =   240
         Width           =   3735
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Username"
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3975
      Begin VB.TextBox txtUsername 
         Height          =   285
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   3735
      End
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim listx As ListItem

If Len(txtUsername) = 0 Then Exit Sub
If Len(txtIP) = 0 Then Exit Sub

Set listx = Form1.lvUsers.ListItems.Add(, , , , 1)
    listx.Text = txtUsername
    listx.SubItems(1) = txtIP
    
    Set db = OpenDatabase(App.Path & "\users.mdb")
    Set rstInfo = db.OpenRecordset("users")
    
    With rstInfo
        .AddNew
        !User = txtUsername
        !ip = txtIP
        .Update
    End With
    
    Form1.lblFound = Form1.lblFound + 1
    
If Check1 Then
    currentUser = txtIP
    Form1.cmbShow_Click
End If

Unload Me
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

