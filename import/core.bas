Attribute VB_Name = "core"
'================================================================================================
' Automating Edge IE Mode Based On Internet Explorer Server Class Object
'------------------------------------------------------------------------------------------------
' Author(s)   :
'       Unknown Author(s)
' Contributors:
'       Long Vh (long.hoang.vu@hsbc.com.sg)
' Last Update :
'       01/06/22 - Long Vh: Tidied up originals + commentations + added TestEdgeIeMode procedure
' Descriptions:
'       The codes were enhanced for both VBA7 (64-bit) and others (32-bit) by Long Vh.
'       To test, run the procedure in the "demo" module.
'       Please give due credit to the author(s) before distributing the codes.
' References Required:
'       Nil.
'================================================================================================

#If VBA7 Then

    Declare PtrSafe Sub Sleep Lib "kernel32.dll" (ByVal time As Long)
    Declare PtrSafe Function SetForegroundWindow Lib "user32.dll" (ByVal hwnd As LongPtr) As Integer
    Declare PtrSafe Function EnumWindows Lib "user32" (ByVal lpEnumFunc As LongPtr, ByVal lParam As LongPtr) As LongPtr
    Declare PtrSafe Function EnumChildWindows Lib "user32" (ByVal hWndParent As LongPtr, ByVal lpEnumFunc As LongPtr, ByVal lParam As LongPtr) As LongPtr
    Declare PtrSafe Function GetClassName Lib "user32.dll" Alias "GetClassNameA" (ByVal hwnd As LongPtr, ByVal Buffer As String, ByVal bufferLength As Integer) As Long
    Declare PtrSafe Function GetWindowText Lib "user32.dll" Alias "GetWindowTextA" (ByVal hwnd As LongPtr, ByVal Buffer As String, ByVal bufferLength As Integer) As Long
    Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
    Declare PtrSafe Function sendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As LongPtr, ByVal wMsg As Long, ByVal wParam As LongPtr, lParam As Any) As Long
    Declare PtrSafe Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As LongPtr, ByVal hWnd2 As LongPtr, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
    Declare PtrSafe Function RegisterWindowMessageW Lib "user32" (ByVal lpString As LongPtr) As Long
    Declare PtrSafe Function SendMessageTimeoutW Lib "user32" (ByVal hwnd As LongPtr, ByVal msg As Long, ByVal wParam As LongPtr, ByRef lParam As LongPtr, ByVal fuFlags As Long, ByVal uTimeout As Long, ByRef lpdwResult As Long) As LongPtr
    Declare PtrSafe Function ObjectFromLresult Lib "oleacc" (ByVal lResult As Long, ByRef riid As Currency, ByVal wParam As LongPtr, ppvObject As Any) As Long

    Private resultHwnd As LongPtr

#Else

    Declare Sub Sleep Lib "kernel32.dll" (ByVal time As Long)
    Declare Function SetForegroundWindow Lib "user32.dll" (ByVal hwnd As Long) As Integer
    Declare Function EnumWindows Lib "user32" (ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long
    Declare Function EnumChildWindows Lib "user32" (ByVal hWndParent As Long, ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long
    Declare Function GetClassName Lib "user32.dll" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal buffer As String, ByVal bufferLength As Integer) As Long
    Declare Function GetWindowText Lib "user32.dll" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal buffer As String, ByVal bufferLength As Integer) As Long
    Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
    Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
    Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
    Declare Function RegisterWindowMessageW Lib "user32" (ByVal lpString As Long) As Long
    Declare Function SendMessageTimeoutW Lib "user32" (ByVal hwnd As Long, ByVal msg As Long, ByVal wParam As Long, ByRef lParam As Long, ByVal fuFlags As Long, ByVal uTimeout As Long, ByRef lpdwResult As Long) As Long
    Declare Function ObjectFromLresult Lib "oleacc" (ByVal lResult As Long, ByRef riid As Currency, ByVal wParam As Long, ppvObject As Any) As Long

    Private resultHwnd As Long

#End If

Private targetTitle As String   'Global var for functions to check the user-defined title of the target window
Private resultTitle As String   'Global var for functions to check the current window title during enumerations
Private searchTimeStart, searchTimeOut


Function GetEdgeIeDOM(ByVal WinTitle As String) As Object
'----------------------------------------------------------------------------------
'Note : This function uses GetHTMLDocument function to obtain the HTMLDoc Object
'       based on the window handle obtained from SearchForIEWindow function using
'       the user-provided windows title.
'----------------------------------------------------------------------------------
    
   'Enable for IntelliSense
   'Requires MS HTML Object Library Reference
    'Dim TargetWebDoc As MSHTML.HTMLDocument
    
    resultHwnd = 0
    resultTitle = ""
    targetTitle = WinTitle  'Assign title to global variable
    
    Call SearchForIEWindow
    
    If resultHwnd <> 0 Then
        Set TargetWebDoc = GetHtmlDocument(resultHwnd)
        Set GetEdgeIeDOM = TargetWebDoc
    Else
        Set GetEdgeIeDOM = Nothing
    End If
    
    resultHwnd = 0
    resultTitle = ""
    targetTitle = ""
    
End Function
 

Function SearchForIEWindow()
'----------------------------------------------------------------------------------
'Note : This method will continuously look for the target window or timeout
'       Default timeout is 3 seconds (timeoutseconds)
'----------------------------------------------------------------------------------
        
    searchTimeStart = Timer
    timeoutseconds = 3          'in seconds
    
    While resultHwnd = 0
        If Timer - searchTimeStart > timeoutseconds Then Exit Function
        EnumWindows AddressOf EnumWindowsProc, 0&
        DoEvents
    Wend
    
End Function


Function EnumWindowsProc(ByVal hwnd As Long, ByVal lParam As Long) As Boolean
'----------------------------------------------------------------------------------
'Note : This function tries to find the Chrome_WidgetWin Class object
'       This object shall include a handle to the Internet Explorer_Server class
'----------------------------------------------------------------------------------
 
    EnumWindowsProc = True
 
    Dim sTitle As String
    Dim sClassName As String
    
    sTitle = Space(1024)
    sClassName = Space(1024)
    
    GetWindowText hwnd, sTitle, 1024
    GetClassName hwnd, sClassName, 1024
    
    sTitle = Replace(Trim(sTitle), Chr(0), "")
    sClassName = Replace(Trim(sClassName), Chr(0), "")
    
    If InStr(sClassName, "Chrome_WidgetWin") Then
        If InStr(sTitle, targetTitle) Then
            resultTitle = sTitle
            EnumWindowsProc = False
        End If
            
        If FindWindowEx(hwnd, 0&, vbNullString, vbNullString) > 0 Then
            EnumChildWindows hwnd, AddressOf EnumChildWindowsProc, 0&
        End If
        
        If resultHwnd <> 0 Then EnumWindowsProc = False    'Set to False to stop enum
    End If
    
End Function


Function EnumChildWindowsProc(ByVal hwnd As Long, ByVal lParam As Long) As Boolean
'----------------------------------------------------------------------------------
'Note : This function tries to find the Internet Explorer_Server Class object
'       This object is required for DOM automation of IE object.
'----------------------------------------------------------------------------------
 
    EnumChildWindowsProc = True
 
    Dim sTitle2 As String
    Dim sClassName2 As String
    
    sTitle2 = Space(1024)
    sClassName2 = Space(1024)
    
    GetWindowText hwnd, sTitle2, 1024
    GetClassName hwnd, sClassName2, 1024
    
    sTitle2 = Replace(Trim(sTitle2), Chr(0), "")
    sClassName2 = Replace(Trim(sClassName2), Chr(0), "")
    
    If Len(resultTitle) = 0 And InStr(sTitle2, targetTitle) Then resultTitle = sTitle2
    
    If sClassName2 = "Internet Explorer_Server" Then
        If Len(resultTitle) > 0 Then
            resultHwnd = hwnd
            EnumChildWindowsProc = False
        End If
    End If
    
End Function


Function GetHtmlDocument(ByVal hwnd_InternetExplorer_Server As LongPtr, Optional ByVal uTimeout As Long = 1000, Optional ByVal documentVersion As Integer = 1) As Object
'----------------------------------------------------------------------------------
'Note : Obtains IE Dom object using winAPI
'----------------------------------------------------------------------------------

    Dim lngMsg As Long
    Dim interfaceID(1) As Currency
    Set GetHtmlDocument = Nothing
 
    lngMsg = RegisterWindowMessageW(StrPtr("WM_HTML_GETOBJECT"))
    
    If lngMsg <> 0 Then
        Dim lpdwResult As Long
        If SendMessageTimeoutW(hwnd_InternetExplorer_Server, lngMsg, 0, 0, 2, uTimeout, lpdwResult) <> 0 Then   '2 = Abort If Hung
            Dim hResult As Long
            hResult = ObjectFromLresult(lpdwResult, interfaceID(0), 0, GetHtmlDocument)
            If hResult <> 0 Then err.Raise hResult, "GetHtmlDocument", "Unable to get the DOM object of Internet Explorer_Server"
        End If
    End If
    
End Function
