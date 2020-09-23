Attribute VB_Name = "mPropertyPage"
Option Explicit


Private Const MAX_PATH = 260

Private Declare Function ILCreateFromPathW Lib "shell32.dll" (ByVal pwszPath As Long) As Long
Private Declare Sub ILFree Lib "shell32.dll" (ByVal Pidl As Long)

Private Const CSIDL_DRIVES As Long = &H11
Private Declare Function SHGetSpecialFolderLocation Lib "shell32.dll" (ByVal hwndOwner As Long, ByVal nFolder As Long, Pidl As Long) As Long



Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
Private Declare Function GetWindowRect Lib "user32.dll" (ByVal hWnd As Long, lpRect As RECT) As Long


Private Const SWP_NOZORDER = &H4
Private Declare Function SetWindowPos Lib "user32.dll" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Private Declare Function SetWindowText Lib "user32" Alias "SetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String) As Long
Private Declare Function SendMessageAsString Lib "user32.dll" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As String) As Long

Private Type BROWSEINFO
    hwndOwner As Long
    pIDLRoot As Long
    pszDisplayName As String 'return display name of item selected
    lpszTitle As String
    ulFlags As Long
    lpfnCallback As Long
    lParam As Long
    iImage As Long
End Type

Private Const WM_USER As Long = &H400
Private Const BFFM_SETSELECTIONA As Long = (WM_USER + 102)
Private Const BFFM_INITIALIZED As Long = 1
Private Const BIF_RETURNONLYFSDIRS = &H1       'Only returns file system directories
Private Declare Function SHBrowseForFolder Lib "shell32.dll" Alias "SHBrowseForFolderA" (lpBrowseInfo As BROWSEINFO) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32.dll" Alias "SHGetPathFromIDListA" (ByVal Pidl As Long, ByVal pszPath As String) As Long


Private Declare Function PathRemoveBackslash Lib "shlwapi.dll" Alias "PathRemoveBackslashA" (ByVal sPath As String) As Long
Private Declare Function PathIsDirectory Lib "shlwapi.dll" Alias "PathIsDirectoryA" (ByVal pszPath As String) As Long


Private Const LMEM_FIXED = &H0
Private Const LMEM_ZEROINIT = &H40
Private Const LPTR = (LMEM_FIXED Or LMEM_ZEROINIT)
Private Declare Function lstrcpyA Lib "kernel32" (lpString1 As Any, lpString2 As Any) As Long
Private Declare Function lstrlenA Lib "kernel32" (lpString As Any) As Long
Private Declare Function LstrCat Lib "kernel32" Alias "lstrcatA" (ByVal lpString1 As String, ByVal lpString2 As String) As Long
Private Declare Sub CoTaskMemFree Lib "ole32.dll" (ByVal pv As Long)

Private m_Caption As String
Private m_StartingFolder As String

Public Function BrowseFolderForPropertyPage(ByVal hWnd As Long, _
                                            ByVal Caption As String, _
                                            ByVal Title As String, _
                                            ByVal RootFolder As String, _
                                            ByVal StartingFolder As String) As String


    Dim i As Long
    Dim BI As BROWSEINFO
    Dim NewPathPidl As Long
    Dim RootPidl As Long
    Dim StartingPidl As Long
    Dim RetPath As String
    

    m_Caption = Caption

    With BI
        .hwndOwner = hWnd

        If PathIsDirectory(RootFolder) <> 0 Then
            RootPidl = ILCreateFromPathW(StrPtr(RootFolder))
            .pIDLRoot = RootPidl
        Else 'obtain the pidl to the special folder 'drives'
            Call SHGetSpecialFolderLocation(0&, CSIDL_DRIVES, RootPidl)
            .pIDLRoot = RootPidl
        End If
        
        .lpszTitle = Title

        .lpfnCallback = Val(AddressOf BrowseFolderForPropertyPageCallbackProc)

        If PathIsDirectory(StartingFolder) <> 0 Then
            StartingPidl = ILCreateFromPathW(StrPtr(StartingFolder))
            .lParam = StartingPidl
            m_StartingFolder = pRemoveBackslash(StartingFolder) & vbNullChar
        Else
            m_StartingFolder = ""
        End If

    End With

    NewPathPidl = SHBrowseForFolder(BI)

    If NewPathPidl Then
        RetPath = Space$(512)
        If SHGetPathFromIDList(NewPathPidl, RetPath) Then
            i = InStr(RetPath, vbNullChar) - 1
            If i > 0 Then RetPath = Left$(RetPath, i)

            BrowseFolderForPropertyPage = RetPath
            'free the NewPathPidl from SHBrowseForFolder call
            Call CoTaskMemFree(NewPathPidl)
        End If
    End If


    ' Free our allocated string pointer
    If (StartingPidl <> 0 And StartingPidl <> RootPidl) Then Call ILFree(StartingPidl)
    If RootPidl <> 0 Then Call ILFree(RootPidl)

End Function
Private Function BrowseFolderForPropertyPageCallbackProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal lParam As Long, ByVal lpData As Long) As Long
    '-----------------------------------------------------------------------------------------------------------------------------
    'This function is used by the Browse Folder dialog to call back for instructions
    '-----------------------------------------------------------------------------------------------------------------------------
    Dim BrowserRect As RECT
    Dim BrowserNewRect As RECT


    Select Case uMsg
        Case BFFM_INITIALIZED
            If Len(m_Caption) > 0 Then
                Call SetWindowText(hWnd, m_Caption)
            End If
            GetWindowRect hWnd, BrowserRect
            BrowserNewRect.Left = ((Screen.Width / Screen.TwipsPerPixelX) - (BrowserRect.Right - BrowserRect.Left)) \ 2
            BrowserNewRect.Top = ((Screen.Height / Screen.TwipsPerPixelY) - (BrowserRect.Bottom - BrowserRect.Top)) \ 2
            BrowserNewRect.Right = BrowserRect.Right - BrowserRect.Left
            BrowserNewRect.Bottom = BrowserRect.Bottom - BrowserRect.Top

            SetWindowPos hWnd, 0, BrowserNewRect.Left, BrowserNewRect.Top, BrowserNewRect.Right, BrowserNewRect.Bottom, SWP_NOZORDER
            
            If Len(m_StartingFolder) > 0 Then
                Call SendMessageAsString(hWnd, BFFM_SETSELECTIONA, ByVal 1, ByVal m_StartingFolder)
            End If

        Case Else:

    End Select

End Function




Private Function pRemoveBackslash(Path As String) As String
    Dim NewPath As String
    NewPath = Left(Trim(Path) & String(MAX_PATH, 0), MAX_PATH)
    Call PathRemoveBackslash(NewPath)
    pRemoveBackslash = Left$(NewPath, InStr(NewPath, vbNullChar) - 1)
End Function


