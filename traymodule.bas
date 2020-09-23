Attribute VB_Name = "TrayModule"
      Type NOTIFYICONDATA
         cbSize As Long
         hWnd As Long
         uId As Long
         uFlags As Long
         uCallBackMessage As Long
         hIcon As Long
         szTip As String * 64
      End Type
      
      Global Const NIM_ADD = &H0
      Global Const NIM_MODIFY = &H1
      Global Const NIM_DELETE = &H2

      Global Const WM_MOUSEMOVE = &H200

      Global Const NIF_MESSAGE = &H1
      Global Const NIF_ICON = &H2
      Global Const NIF_TIP = &H4

      Global Const WM_LBUTTONDBLCLK = &H203   'Double-click
      Global Const WM_LBUTTONDOWN = &H201     'Button down
      Global Const WM_LBUTTONUP = &H202       'Button up

      Global Const WM_RBUTTONDBLCLK = &H206   'Double-click
      Global Const WM_RBUTTONDOWN = &H204     'Button down
      Global Const WM_RBUTTONUP = &H205       'Button up

      Declare Function Shell_NotifyIcon Lib "shell32" _
         Alias "Shell_NotifyIconA" _
         (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean

      Global nid As NOTIFYICONDATA

Sub AddToTray(TrayIcon, TrayText As String, TrayForm As Form)
         'Set the individual values of the NOTIFYICONDATA data type.
         nid.cbSize = Len(nid)
         nid.hWnd = TrayForm.hWnd
         nid.uId = vbNull
         nid.uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
         nid.uCallBackMessage = WM_MOUSEMOVE
         nid.hIcon = TrayIcon 'You can replace form1.icon with loadpicture=("icon's file name")
         nid.szTip = TrayText & vbNullChar

         'Call the Shell_NotifyIcon function to add the icon to the taskbar
         'status area.
         Shell_NotifyIcon NIM_ADD, nid
         TrayForm.Hide
End Sub
Sub ModifyTray(TrayIcon, TrayText As String, TrayForm As Form)
         'Set the individual values of the NOTIFYICONDATA data type.
         nid.cbSize = Len(nid)
         nid.hWnd = TrayForm.hWnd
         nid.uId = vbNull
         nid.uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
         nid.uCallBackMessage = WM_MOUSEMOVE
         nid.hIcon = TrayIcon 'You can replace form1.icon with loadpicture=("icon's file name")
         nid.szTip = TrayText & vbNullChar

         'Call the Shell_NotifyIcon function to modify the icon to the taskbar
         'status area.
         Shell_NotifyIcon NIM_MODIFY, nid
End Sub

Sub RemoveFromTray()
Shell_NotifyIcon NIM_DELETE, nid
End Sub

Function RespondToTray(X As Single)
          'Call this sub from the mousemove event on a form
          'Event occurs when the mouse pointer is within the rectangular
          'boundaries of the icon in the taskbar status area.
          RespondToTray = 0
          Dim Msg As Long
          Dim sFilter As String
          If Form1.ScaleMode <> 3 Then Msg = X / Screen.TwipsPerPixelX Else: Msg = X
          Select Case Msg
             Case WM_LBUTTONDOWN
             Case WM_LBUTTONUP
             Case WM_LBUTTONDBLCLK 'Left button double-clicked
             RespondToTray = 1
             Case WM_RBUTTONDOWN 'Right button pressed
             RespondToTray = 2
             Case WM_RBUTTONUP
             Case WM_RBUTTONDBLCLK
          End Select
End Function


      Sub ShowFormAgain(TrayForm As Form)
      Call RemoveFromTray
      TrayForm.Show
      End Sub


