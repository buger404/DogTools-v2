Attribute VB_Name = "uwpChecker"
Public uwpSwitch As Long, uwpRet As Long
Public Function uwpChild(ByVal Hwnd As Long, ByVal lParam As Long) As Boolean
    If uwpRet <> 0 Then GoTo last

    Dim class As String * 255
    GetClassNameA Hwnd, class, 255
    
    If UnSpace(class) = "Windows.UI.Core.CoreWindow" Then uwpRet = Hwnd
    
last:
    uwpChild = True
End Function
Public Function uwpFind(ByVal Hwnd As Long) As Long

    uwpRet = 0
    EnumChildWindows Hwnd, AddressOf uwpChild, 0&
    uwpFind = uwpRet
    
End Function
