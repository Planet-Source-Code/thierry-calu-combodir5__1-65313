VERSION 5.00
Begin VB.UserControl ucComboDir5 
   BackColor       =   &H8000000E&
   ClientHeight    =   315
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3630
   PropertyPages   =   "ucComboDir5.ctx":0000
   ScaleHeight     =   21
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   242
   ToolboxBitmap   =   "ucComboDir5.ctx":002F
End
Attribute VB_Name = "ucComboDir5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'---------------------------------------------------
'Thousand thanks to:
'Paul Caton, vbAccelerator, Matthew Curland, Franseco Balena etc...
'and a special thank to Matthew Hood (DWD Multi-Column ComboBox - ActiveX Control)
'http://www.Planet-Source-Code.com/vb/scripts/ShowCode.asp?txtCodeId=26216&lngWId=1
'---------------------------------------------------

'---------------------------------------------------
'Public Enumerations
'---------------------------------------------------
Public Enum enmXPTheme
    [As Application] = 0
    [Disbladed Theme] = 1
    [Enabled Theme] = 2
End Enum
'---------------------------------------------------

'---------------------------------------------------
' member variable for properties
'---------------------------------------------------
Private m_ResolveSharedFolder As Boolean: Private Const m_ResolveSharedFolder_ByDef As Boolean = False
Private m_ResolveNetworkDrive As Boolean: Private Const m_ResolveNetworkDrive_ByDef As Boolean = False
Private m_ResolveCSIDL As Boolean: Private Const m_ResolveCSIDL_ByDef As Boolean = True
Private m_RootFolder As String
Private m_RootNodeHighLight As Boolean: Private Const m_RootNodeHighLight_ByDef As Boolean = False
Private m_RootNodeBold As Boolean: Private Const m_RootNodeBold_ByDef As Boolean = False
Private m_StartingFolder As String
Private m_StartingNodeHighLight As Boolean: Private Const m_StartingNodeHighLight_ByDef As Boolean = False
Private m_StartingNodeBold As Boolean: Private Const m_StartingNodeBold_ByDef As Boolean = False
Private m_XPTheme As enmXPTheme: Private Const m_XPTheme_ByDef As Boolean = [As Application]
Private m_Enabled As Boolean
Private m_CurrentFolder As String
'---------------------------------------------------
'TODO
Private m_DropDownHeight As Long
Private m_DropDownWidth As Long
'---------------------------------------------------
Private Type typInvariantFolder
    LocalName As String
    AliasName As String
End Type
Private m_InvariantFolders() As typInvariantFolder
Private m_NbrOfInvariantFolders As Integer
'---------------------------------------------------


Private m_hWnd As Long
Private m_OwnerFormHwnd As Long
Private m_hFont As Long
Private m_BrowserDlgHwnd As Long
Private m_SysTreeView32Hwnd As Long
Private m_OkButtonHwnd As Long

Private m_ScrollBoxWidth As Long
Private m_UseXpTheme As Boolean
Private m_ButtonDown As Boolean
Private m_IsSubClassing As Boolean

Private m_IconRect As RECT
Private m_ButtonRect As RECT
Private m_UserControlRect As RECT
Private m_TextRect As RECT


'======================================================================================================
Private Const MAX_PATH = 260

Private Declare Function PathRemoveBackslash Lib "shlwapi.dll" Alias "PathRemoveBackslashA" (ByVal sPath As String) As Long
Private Declare Function PathIsDirectory Lib "shlwapi.dll" Alias "PathIsDirectoryA" (ByVal pszPath As String) As Long

'Browser controls
Private Const IDOK = &H1
Private Const IDCANCEL = &H2
Private Const IDTREEVIEW = &H3741
Private Declare Function GetDlgItem Lib "user32.dll" (ByVal hDlg As Long, ByVal nIDDlgItem As Long) As Long


'To dertermine the real parent
Private Const GA_ROOT As Long = 2
Private Declare Function GetAncestor Lib "user32.dll" (ByVal hWnd As Long, ByVal gaFlags As Long) As Long

'To dertermine size of the ComboDir button
Private Const SM_CXHTHUMB As Long = 10
Private Declare Function GetSystemMetrics Lib "user32.dll" (ByVal nIndex As Long) As Long

'used to position windows
Private Type POINTAPI
    X As Long
    Y As Long
End Type

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Declare Function GetClientRect Lib "user32.dll" (ByVal hWnd As Long, lpRect As RECT) As Long
Private Declare Function SetRect Lib "user32.dll" (lpRect As RECT, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function FillRect Lib "user32.dll" (ByVal hdc As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function OffsetRect Lib "user32.dll" (lpRect As RECT, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function GetWindowRect Lib "user32.dll" (ByVal hWnd As Long, lpRect As RECT) As Long

Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long



Private Const SWP_NOZORDER = &H4
Private Declare Function SetWindowPos Lib "user32.dll" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

'UXTheme...
Private Declare Function SetWindowTheme Lib "UXTheme.dll" (ByVal hWnd As Long, ByVal pszSubAppName As Long, ByVal pszSubIdList As Long) As Long
Private Declare Function OpenThemeData Lib "UXTheme.dll" (ByVal hWnd As Long, ByVal pszClassList As Long) As Long
Private Declare Function CloseThemeData Lib "UXTheme.dll" (ByVal hTheme As Long) As Long
Private Declare Function DrawThemeBackground Lib "UXTheme.dll" (ByVal hTheme As Long, ByVal hdc As Long, ByVal iPartId As Long, ByVal iStateId As Long, ByRef pRect As RECT, ByRef pClipRect As Any) As Long



'To dertermine if XP Theme enabled / possible
Private Type OSVersionInfo
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
    szCSDVersion As String * 128    ' Maintenance string for PSS usage
End Type

Private Type DLLVersionInfo
    cbSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
End Type
Private Const VER_PLATFORM_WIN32_NT As Long = 2
Private Declare Function GetVersionEx Lib "kernel32.dll" Alias "GetVersionExA" (ByRef lpVersionInfo As OSVersionInfo) As Long
Private Declare Function DllGetVersion Lib "comctl32.dll" (ByRef PDVI As DLLVersionInfo) As Long
Private Declare Function IsThemeActive Lib "UXTheme.dll" () As Long
Private Declare Function IsAppThemed Lib "UXTheme.dll" () As Long

'Drawing
Private Const BF_RECT = &HF    '=>BF_LEFT Or BF_TOP Or BF_RIGHT Or BF_BOTTOM
Private Const BDR_SUNKEN = &HA
Private Declare Function DrawEdge Lib "user32.dll" (ByVal hdc As Long, qrc As RECT, ByVal edge As Long, ByVal grfFlags As Long) As Long

Private Const DFC_SCROLL As Long = 3
Private Const DFCS_SCROLLDOWN As Long = &H1
Private Const DFCS_PUSHED As Long = &H200
Private Const DFCS_FLAT As Long = &H4000
Private Const DFCS_INACTIVE As Long = &H100
Private Declare Function DrawFrameControl Lib "user32.dll" (ByVal hdc As Long, lpRect As RECT, ByVal un1 As Long, ByVal un2 As Long) As Long

Private Const COLOR_HIGHLIGHT = 13
Private Const COLOR_HIGHLIGHTTEXT = 14
Private Const COLOR_MENU = 4
Private Const COLOR_MENUTEXT = 7
Private Const COLOR_GRAYTEXT = 17
Private Declare Function GetSysColor Lib "user32.dll" (ByVal nIndex As Long) As Long
Private Declare Function GetSysColorBrush Lib "user32.dll" (ByVal nIndex As Long) As Long
Private Declare Function SetTextColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Private Declare Function SetBkColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long

Private Const OPAQUE = 2
Private Const TRANSPARENT = 1
Private Declare Function SetBkMode Lib "gdi32" (ByVal hdc As Long, ByVal nBkMode As Long) As Long

Private Const DT_BOTTOM As Long = &H8
Private Const DT_LEFT = &H0
Private Declare Function DrawText Lib "user32.dll" Alias "DrawTextA" (ByVal hdc As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long

Private Const DI_MASK = &H1
Private Const DI_IMAGE = &H2
Private Const DI_NORMAL = DI_MASK Or DI_IMAGE
Private Declare Function DrawIconEx Lib "user32.dll" (ByVal hdc As Long, ByVal xLeft As Long, ByVal yTop As Long, ByVal hIcon As Long, ByVal cxWidth As Long, ByVal cyWidth As Long, ByVal istepIfAniCur As Long, ByVal hbrFlickerFreeDraw As Long, ByVal diFlags As Long) As Long
Private Declare Function DestroyIcon Lib "user32.dll" (ByVal hIcon As Long) As Long


Private Type BROWSEINFO
    hwndOwner As Long
    pIDLRoot As Long
    pszDisplayName As String   'return display name of item selected
    lpszTitle As Long
    ulFlags As Long
    lpfnCallback As Long
    lParam As Long
    iImage As Long
End Type

Private Const BIF_RETURNONLYFSDIRS = &H1       'Only returns file system directories
Private Const BIF_RETURNFSANCESTORS = &H8  'Only returns file system ancestors.
Private Const BIF_BROWSEFORCOMPUTER As Long = &H1000

'======================================================================================================
'To get the correct icons and displayname of a path
'======================================================================================================
Private Type SHFILEINFO
    hIcon As Long
    iIcon As Long
    dwAttributes As Long
    szDisplayName As String * 260
    szTypeName As String * 80
End Type

Private Const SHGFI_ICON = &H100                         '  get icon
Private Const SHGFI_SMALLICON = &H1                      '  get small icon
Private Const SHGFI_DISPLAYNAME = &H200                  '  get display name
Private Const SHGFI_SELECTED As Long = &H10000
Private Const SHGFI_PIDL = &H8&

Private Declare Function ILCreateFromPathW Lib "shell32.dll" (ByVal pwszPath As Long) As Long
Private Declare Sub ILFree Lib "shell32.dll" (ByVal Pidl As Long)
Private Declare Function SHGetFileInfo Lib "shell32.dll" Alias "SHGetFileInfoA" (ByVal pszPath As String, ByVal dwFileAttributes As Long, psfi As SHFILEINFO, ByVal cbFileInfo As Long, ByVal uFlags As Long) As Long
Private Declare Function SHGetPidlInfo Lib "shell32.dll" Alias "SHGetFileInfoA" (ByVal ppidl As Long, ByVal dwFileAttributes As Long, ByRef psfi As SHFILEINFO, ByVal cbFileInfo As Long, ByVal uFlags As Long) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32.dll" Alias "SHGetPathFromIDListA" (ByVal Pidl As Long, ByVal pszPath As String) As Long
Private Declare Function SHBrowseForFolder Lib "shell32.dll" Alias "SHBrowseForFolderA" (lpBrowseInfo As BROWSEINFO) As Long
Private Declare Function SHGetSpecialFolderLocation Lib "shell32.dll" (ByVal hwndOwner As Long, ByVal nFolder As Long, Pidl As Long) As Long



Private Declare Function GetDesktopWindow Lib "user32.dll" () As Long
Private Declare Function FindWindowEx Lib "user32.dll" Alias "FindWindowExA" (ByVal HwndParent As Long, ByVal hWndChildAfter As Long, ByVal lpszClassName As String, ByVal lpszCaption As String) As Long
Private Declare Function IsWindow Lib "user32.dll" (ByVal hWnd As Long) As Long
Private Declare Function DestroyWindow Lib "user32.dll" (ByVal hWnd As Long) As Long
Private Declare Function SendMessageAsLong Lib "user32.dll" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Long) As Long
Private Declare Function SendMessageAsAny Lib "user32.dll" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function SendMessageAsString Lib "user32.dll" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As String) As Long
Private Declare Function PostMessage Lib "user32.dll" Alias "PostMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function IsWindowEnabled Lib "user32.dll" (ByVal hWnd As Long) As Long

Private Declare Function GetFocus Lib "user32.dll" () As Long

Private Declare Function lstrlenA Lib "kernel32" (ByVal lpString As Any) As Long
Private Declare Function lstrlenW Lib "kernel32" (ByVal lpString As Long) As Long
Private Declare Function lstrcpyA Lib "kernel32" (ByVal lpString1 As Any, ByVal lpString2 As Any) As Long

Private Declare Function LstrCat Lib "kernel32" Alias "lstrcatA" (ByVal lpString1 As String, ByVal lpString2 As String) As Long
Private Declare Function LocalAlloc Lib "kernel32" (ByVal uFlags As Long, ByVal uBytes As Long) As Long
Private Declare Function LocalLock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function LocalUnlock Lib "kernel32" (ByVal hMem As Long) As Long

Private Declare Sub CoTaskMemFree Lib "ole32.dll" (ByVal pv As Long)

Private Const SW_HIDE               As Long = 0
Private Declare Function ShowWindow Lib "user32.dll" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long

'======================================================================================================
'To remove border of the original BrowseForFolderDlg and SysTreeView32
Private Const WS_CAPTION            As Long = &HC00000
Private Const WS_THICKFRAME         As Long = &H40000
Private Const WS_EX_CLIENTEDGE      As Long = &H200
Private Const WS_EX_WINDOWEDGE      As Long = &H100
Private Const WS_EX_STATICEDGE      As Long = &H20000

Private Const GWL_STYLE = (-16)
Private Const GWL_EXSTYLE = (-20)
Private Declare Function SetWindowLong Lib "user32.dll" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function GetWindowLong Lib "user32.dll" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long



Private Const WM_USER               As Long = &H400
Private Const WM_SETFONT            As Long = &H30
Private Const WM_LBUTTONDBLCLK      As Long = &H203
Private Const WM_ACTIVATE           As Long = &H6
Private Const WA_INACTIVE           As Long = 0
Private Const WM_KEYDOWN            As Long = &H100
Private Const WM_NCACTIVATE         As Long = &H86
Private Const WM_COMMAND            As Long = &H111&
Private Const WM_CLOSE              As Long = &H10&

'message from browser
Private Const BFFM_INITIALIZED      As Long = 1
Private Const BFFM_SELCHANGED       As Long = 2
Private Const BFFM_VALIDATEFAILEDA  As Long = 3     '// lParam:szPath ret:1(cont),0(EndDialog)
Private Const BFFM_VALIDATEFAILEDW  As Long = 4    '// lParam:wzPath ret:1(cont),0(EndDialog)
'// messages to browser
Private Const BFFM_SETSTATUSTEXT    As Long = (WM_USER + 100)
Private Const BFFM_ENABLEOK         As Long = (WM_USER + 101)
Private Const BFFM_SETSELECTIONA    As Long = (WM_USER + 102)
Private Const BFFM_SETSELECTIONW    As Long = (WM_USER + 103)
Private Const BFFM_SETSTATUSTEXTA   As Long = (WM_USER + 100)
Private Const BFFM_SETSTATUSTEXTW   As Long = (WM_USER + 104)

Private Const DRIVE_REMOVABLE       As Long = 2
Private Const DRIVE_FIXED           As Long = 3
Private Const DRIVE_REMOTE          As Long = 4
Private Const DRIVE_CDROM           As Long = 5
Private Const DRIVE_RAMDISK         As Long = 6
Private Declare Function GetDriveType Lib "kernel32" Alias "GetDriveTypeA" (ByVal nDrive As String) As Long


'======================================================================================================
'use for Font property
'======================================================================================================
Private Type LOGFONT
    lfHeight         As Long
    lfWidth          As Long
    lfEscapement     As Long
    lfOrientation    As Long
    lfWeight         As Long
    lfItalic         As Byte
    lfUnderline      As Byte
    lfStrikeOut      As Byte
    lfCharSet        As Byte
    lfOutPrecision   As Byte
    lfClipPrecision  As Byte
    lfQuality        As Byte
    lfPitchAndFamily As Byte
    lfFaceName(32)   As Byte
End Type
Private Const LOGPIXELSY        As Long = 90
Private Const FW_NORMAL         As Long = 400
Private Const FW_BOLD           As Long = 700

Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, ByVal nIndex As Long) As Long
Private Declare Function CreateFontIndirect Lib "gdi32" Alias "CreateFontIndirectA" (lpLogFont As LOGFONT) As Long
Private Declare Function MulDiv Lib "kernel32" (ByVal nNumber As Long, ByVal nNumerator As Long, ByVal nDenominator As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long


'======================================================================================================
'use for Hightlight or/and Bold nodes
'======================================================================================================
Private Type TV_ITEM
    mask As Long
    hItem As Long
    state As Long
    stateMask As Long
    pszText As String
    cchTextMax As Long
    iImage As Long
    iSelectedImage As Long
    cChildren As Long
    lParam As Long
End Type

Private Const TV_FIRST              As Long = &H1100
Private Const TVM_GETNEXTITEM       As Long = (TV_FIRST + 10)
Private Const TVM_GETITEM           As Long = (TV_FIRST + 12)
Private Const TVM_SETITEM           As Long = (TV_FIRST + 13)
Private Const TVM_SETTEXTCOLOR      As Long = (TV_FIRST + 30)

Private Const TVGN_CARET            As Long = &H9
Private Const TVGN_ROOT             As Long = &H0

Private Const TVIF_STATE            As Long = &H8
Private Const TVIF_HANDLE           As Long = &H10
Private Const TVIS_BOLD             As Long = &H10
Private Const TVIS_DROPHILITED      As Long = &H8

Private Declare Function OleTranslateColor Lib "olepro32" (ByVal OLE_COLOR As Long, ByVal hPalette As Long, pccolorref As Long) As Long

'======================================================================================================
'use for Resolve CSIDL paths
'To deal with more invariant CSIDL change the routine: pvGetResolvedFolders
'Limited set of invariant CSIDL-Folders
'Goto http://vbnet.mvps.org/index.html?code/browse/csidlversions.htm
'======================================================================================================
Private Const CSIDL_PERSONAL            As Long = &H5
Private Const CSIDL_MYMUSIC             As Long = &HD
Private Const CSIDL_MYPICTURES          As Long = &H27
Private Const CSIDL_MYVIDEO             As Long = &HE
Private Const CSIDL_DESKTOPDIRECTORY    As Long = &H10
Private Const CSIDL_DRIVES              As Long = &H11
Private Const CSIDL_TEMPLATES           As Long = &H15
Private Const CSIDL_APPDATA             As Long = &H1A
Private Const CSIDL_COMMON_TEMPLATES    As Long = &H2D
Private Const CSIDL_COMMON_DOCUMENTS    As Long = &H2E
Private Const CSIDL_COMMON_MUSIC        As Long = &H35
Private Const CSIDL_COMMON_PICTURES     As Long = &H36
Private Const CSIDL_COMMON_VIDEO        As Long = &H37
'---------------------------------------------------

'======================================================================================================
'use for Resolve UNC paths
'http://vbnet.mvps.org/index.html?code/network/index.html
'======================================================================================================
Private Const MAX_PREFERRED_LENGTH  As Long = -1
Private Const RESOURCETYPE_ANY      As Long = &H0
Private Const RESOURCE_CONNECTED    As Long = &H1
Private Const ERROR_MORE_DATA       As Long = &HEA
Private Const NO_ERROR              As Long = &H0
Private Const STYPE_DISKTREE        As Long = 0
Private Const ERROR_ACCESS_DENIED   As Long = 5&

Private Type NETRESOURCE
    dwScope As Long
    dwType As Long
    dwDisplayType As Long
    dwUsage As Long
    lpLocalName As Long
    lpRemoteName As Long
    lpComment As Long
    lpProvider As Long
End Type

Private Declare Function WNetGetConnection Lib "mpr.dll" Alias "WNetGetConnectionA" (ByVal lpszLocalName As String, ByVal lpszRemoteName As String, cbRemoteName As Long) As Long
Private Declare Function WNetOpenEnum Lib "mpr.dll" Alias "WNetOpenEnumA" (ByVal dwScope As Long, ByVal dwType As Long, ByVal dwUsage As Long, lpNetResource As Any, lphEnum As Long) As Long
Private Declare Function WNetEnumResource Lib "mpr.dll" Alias "WNetEnumResourceA" (ByVal hEnum As Long, lpcCount As Long, lpBuffer As Any, lpBufferSize As Long) As Long
Private Declare Function WNetCloseEnum Lib "mpr.dll" (ByVal hEnum As Long) As Long

Private Declare Function NetShareEnum Lib "netapi32.dll" (ByVal ServerName As Long, ByVal Level As Long, bufptr As Long, ByVal prefmaxlen As Long, EntriesRead As Long, TotalEntries As Long, resume_handle As Long) As Long
Private Declare Function NetApiBufferFree Lib "netapi32.dll" (ByVal Buffer As Long) As Long

'======================================================================================================
'Subclass APIs
'======================================================================================================
Private Enum eMsgWhen
    MSG_AFTER = 1                                                                   'Message calls back after the original (previous) WndProc
    MSG_BEFORE = 2                                                                  'Message calls back before the original (previous) WndProc
    MSG_BEFORE_AND_AFTER = MSG_AFTER Or MSG_BEFORE                                  'Message calls back before and after the original (previous) WndProc
End Enum

Private Const ALL_MESSAGES  As Long = -1                         'All messages added or deleted
Private Const GMEM_FIXED    As Long = 0                        'Fixed memory GlobalAlloc flag
Private Const GWL_WNDPROC   As Long = -4                        'Get/SetWindow offset to the WndProc procedure address
Private Const PATCH_04      As Long = 88                     'Table B (before) address patch offset
Private Const PATCH_05      As Long = 93                     'Table B (before) entry count patch offset
Private Const PATCH_08      As Long = 132                    'Table A (after) address patch offset
Private Const PATCH_09      As Long = 137                    'Table A (after) entry count patch offset

Private Type tSubData                                                               'Subclass data type
    hWnd As Long        'Handle of the window being subclassed
    nAddrSub As Long            'The address of our new WndProc (allocated memory).
    nAddrOrig As Long             'The address of the pre-existing WndProc
    nMsgCntA As Long            'Msg after table entry count
    nMsgCntB As Long            'Msg before table entry count
    aMsgTblA() As Long              'Msg after table array
    aMsgTblB() As Long              'Msg Before table array
End Type

Private sc_aSubData() As tSubData                     'Subclass data array


Private Declare Sub RtlMoveMemory Lib "kernel32" (pDest As Any, pSource As Any, ByVal dwLength As Long)
Private Declare Function GetModuleHandleA Lib "kernel32" (ByVal lpModuleName As String) As Long
Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function GlobalFree Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function SetWindowLongA Lib "user32.dll" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

'======================================================================================================
'Subclass handler - MUST be the first Public routine in this file. That includes public properties also
'======================================================================================================
Public Sub zzSubclass_Proc(ByVal bBefore As Boolean, _
                           ByRef bHandled As Boolean, _
                           ByRef lReturn As Long, _
                           ByRef lng_hWnd As Long, _
                           ByRef uMsg As Long, _
                           ByRef lParam As Long, _
                           ByRef wParam As Long)
Attribute zzSubclass_Proc.VB_MemberFlags = "40"
    'Parameters:
    'bBefore  - Indicates whether the the message is being processed before or after the default handler - only really needed if a message is set to callback both before & after.
    'bHandled - Set this variable to True in a 'before' callback to prevent the message being subsequently processed by the default handler... and if set, an 'after' callback
    'lReturn  - Set this variable as per your intentions and requirements, see the MSDN documentation for each individual message value.
    'hWnd     - The window handle
    'uMsg     - The message number
    'lParam   - Message related data
    'wParam   - Message related data
    'Notes:
    'If you really know what you're doing, it's possible to change the values of the
    'hWnd, uMsg, lParam and wParam parameters in a 'before' callback so that different
    'values get passed to the default handler.. and optionaly, the 'after' callback
    Dim RC As RECT
    Dim lLoW As Long
    Dim lstrDisplayName As String
    Dim SHFI As SHFILEINFO
    Dim lForeColor As Long
    Dim CurPos As POINTAPI

    Select Case uMsg
        Case BFFM_INITIALIZED
            m_BrowserDlgHwnd = lng_hWnd

            'the form container will be un-activate soon...
            Call zSubclass_Start(m_OwnerFormHwnd)
            Call zSubclass_AddMsg(m_OwnerFormHwnd, WM_ACTIVATE, MSG_AFTER)

            GetWindowRect UserControl.hWnd, RC
            'BrowseForFolder Module has Initialized, so set the Starting Path
            SetWindowLong lng_hWnd, GWL_STYLE, GetWindowLong(lng_hWnd, GWL_STYLE) And Not WS_CAPTION
            SetWindowLong lng_hWnd, GWL_EXSTYLE, 0
            SetWindowPos lng_hWnd, 0, RC.Left, RC.Bottom, RC.Right - RC.Left, 322, SWP_NOZORDER

            m_OkButtonHwnd = GetDlgItem(lng_hWnd, IDOK)
            Call ShowWindow(m_OkButtonHwnd, vbHide)
            Call ShowWindow(GetDlgItem(lng_hWnd, IDCANCEL), vbHide)


            ' Set a little border around the TreeView:
            Call GetClientRect(lng_hWnd, RC)
            m_SysTreeView32Hwnd = GetDlgItem(lng_hWnd, &H3741)
            SetWindowLong m_SysTreeView32Hwnd, GWL_STYLE, GetWindowLong(m_SysTreeView32Hwnd, GWL_STYLE) Or WS_THICKFRAME
            SetWindowLong m_SysTreeView32Hwnd, GWL_EXSTYLE, WS_EX_STATICEDGE
            SetWindowPos m_SysTreeView32Hwnd, 0, 0, 0, RC.Right, RC.Bottom, SWP_NOZORDER
            
            ' Set the colour in the TreeView:
            OleTranslateColor UserControl.ForeColor, 0&, lForeColor
            Call SendMessageAsLong(m_SysTreeView32Hwnd, TVM_SETTEXTCOLOR, 0&, ByVal lForeColor)
            
            ' Set the font of the TreeView:
            If pvStdFontToFontIndirect Then
                Call SendMessageAsLong(m_SysTreeView32Hwnd, WM_SETFONT, m_hFont, 0)
            End If


            If Not m_UseXpTheme Then
                Call SetWindowTheme(m_SysTreeView32Hwnd, StrPtr(" "), StrPtr(" "))
            End If


            Call zSubclass_Start(m_SysTreeView32Hwnd)
            Call zSubclass_AddMsg(m_SysTreeView32Hwnd, WM_LBUTTONDBLCLK, MSG_AFTER)
            Call zSubclass_Start(m_BrowserDlgHwnd)
            Call zSubclass_AddMsg(m_BrowserDlgHwnd, WM_ACTIVATE, MSG_AFTER)
            Call zSubclass_AddMsg(m_hWnd, BFFM_VALIDATEFAILEDA, MSG_AFTER)
            Call zSubclass_AddMsg(m_hWnd, BFFM_SELCHANGED, MSG_AFTER)


            If PathIsDirectory(m_StartingFolder) Then
                Call SendMessageAsString(lng_hWnd, BFFM_SETSELECTIONA, ByVal 1, ByVal m_StartingFolder)
            End If
            Call pvTVSetState(m_SysTreeView32Hwnd)


        Case BFFM_SELCHANGED
            'sPath = Space$(MAX_PATH) 'a buffer for SHGetPathFromIDList
            'If SHGetPathFromIDList(lParam, sPath) Then
            '    If PathIsDirectory(sPath) Then
            '        If (GetDriveType(sPath) Or DRIVE_FIXED) Then
            '            If SHGetPidlInfo(ByVal lParam, 0&, SHFI, Len(SHFI), SHGFI_ICON Or SHGFI_SMALLICON Or SHGFI_DISPLAYNAME Or SHGFI_PIDL) <> 0 Then
            '                lstrDisplayName = Left(SHFI.szDisplayName, InStr(SHFI.szDisplayName, vbNullChar) - 1)
            '                Call yPaint(True, False, True, lstrDisplayName, SHFI.hIcon)
            '                If SHFI.hIcon <> 0 Then DestroyIcon SHFI.hIcon
            '                'Free memory allocated for PIDL
            '                'CoTaskMemFree llngPidl
            '                bFlag = 1
            '            End If
            '        End If
            '        'Else
            '    End If
            'End If
            'Call SendMessageAsLong(lng_hWnd, BFFM_ENABLEOK, 0, ByVal bFlag)
        
            If SHGetPidlInfo(ByVal lParam, 0&, SHFI, Len(SHFI), SHGFI_ICON Or SHGFI_SMALLICON Or SHGFI_DISPLAYNAME Or SHGFI_PIDL) <> 0 Then
                lstrDisplayName = Left(SHFI.szDisplayName, InStr(SHFI.szDisplayName, vbNullChar) - 1)
                Call yPaint(True, False, True, lstrDisplayName, SHFI.hIcon)
                If SHFI.hIcon <> 0 Then DestroyIcon SHFI.hIcon
            End If
        
        
        
        Case BFFM_VALIDATEFAILEDA
            ' Invalid path selected...
            Call SendMessageAsLong(lng_hWnd, BFFM_ENABLEOK, 0, ByVal 0&)

        Case WM_LBUTTONDBLCLK
            If IsWindowEnabled(m_OkButtonHwnd) <> 0 Then
                'Debug.Print "WM_LBUTTONDBLCLK"
                GoTo CloseBrowseForFolderDlg
            End If
            'lReturn = 0

        Case WM_ACTIVATE
            'Get Lo values of the lParam
            lLoW = lParam And &HFFFF&
            If (lLoW = WA_INACTIVE) Then
                If lng_hWnd = m_OwnerFormHwnd Then
                    Call SendMessageAsLong(m_OwnerFormHwnd, WM_NCACTIVATE, 1, 0)
                    Call zSubclass_DelMsg(m_OwnerFormHwnd, WM_ACTIVATE)
                    lReturn = 1
                ElseIf lng_hWnd = m_BrowserDlgHwnd Then
                    'To bad:
                    'If their is a Browser's msgbox ... we don't want to close the list
                    'but how to know why we loose the focus ???
                    'This is my uggly way to do this:
                    Call GetCursorPos(CurPos)
                    If Not WindowFromPoint(CurPos.X, CurPos.Y) = m_BrowserDlgHwnd Then
                        GoTo CloseBrowseForFolderDlg
                    End If
                End If

            End If
    End Select

    Exit Sub

CloseBrowseForFolderDlg:
    On Error Resume Next
    Call zSubclass_DelMsg(m_hWnd, BFFM_VALIDATEFAILEDA)
    Call zSubclass_DelMsg(m_hWnd, BFFM_SELCHANGED)
    Call zSubclass_DelMsg(m_BrowserDlgHwnd, WM_ACTIVATE)
    Call zSubclass_DelMsg(m_SysTreeView32Hwnd, WM_LBUTTONDBLCLK)
    If IsWindow(m_BrowserDlgHwnd) <> 0 Then
        'Close the dialog
        If IsWindowEnabled(m_OkButtonHwnd) <> 0 Then
            PostMessage m_BrowserDlgHwnd, WM_COMMAND, IDOK, 1&
        Else
            PostMessage m_BrowserDlgHwnd, WM_COMMAND, IDCANCEL, 1&
        End If
    End If
    lReturn = 0
End Sub

'Stop all subclassing
Private Sub zSubclass_StopAll()
    Dim i As Long

    i = UBound(sc_aSubData())                                                       'Get the upper bound of the subclass data array
    Do While i >= 0                                                                 'Iterate through each element
        With sc_aSubData(i)
            If .hWnd <> 0 Then                                                      'If not previously zSubclass_Stop'd
                Call zSubclass_Stop(.hWnd)                                           'zSubclass_Stop
            End If
        End With
        i = i - 1                                                                   'Next element
    Loop
End Sub

'Stop subclassing the passed window handle
Private Sub zSubclass_Stop(ByVal lng_hWnd As Long)
    'Parameters:
    'lng_hWnd  - The handle of the window to stop being subclassed
    With sc_aSubData(zIdx(lng_hWnd))
        Call SetWindowLongA(.hWnd, GWL_WNDPROC, .nAddrOrig)                         'Restore the original WndProc
        Call zPatchVal(.nAddrSub, PATCH_05, 0)                                      'Patch the Table B entry count to ensure no further 'before' callbacks
        Call zPatchVal(.nAddrSub, PATCH_09, 0)                                      'Patch the Table A entry count to ensure no further 'after' callbacks
        Call GlobalFree(.nAddrSub)                                                  'Release the machine code memory
        .hWnd = 0                                                                   'Mark the sc_aSubData element as available for re-use
        .nMsgCntB = 0                                                               'Clear the before table
        .nMsgCntA = 0                                                               'Clear the after table
        Erase .aMsgTblB                                                             'Erase the before table
        Erase .aMsgTblA                                                             'Erase the after table
    End With
End Sub

'Start subclassing the passed window handle
Private Function zSubclass_Start(ByVal lng_hWnd As Long) As Long
    'Parameters:
    'lng_hWnd  - The handle of the window to be subclassed
    'Returns;
    'The sc_aSubData() index
    Const CODE_LEN As Long = 204                          'Length of the machine code in bytes
    Const FUNC_CWP As String = "CallWindowProcA"          'We use CallWindowProc to call the original WndProc
    Const FUNC_EBM As String = "EbMode"                   'VBA's EbMode function allows the machine code thunk to know if the IDE has stopped or is on a breakpoint
    Const FUNC_SWL As String = "SetWindowLongA"           'SetWindowLongA allows the cSubclasser machine code thunk to unsubclass the subclasser itself if it detects via the EbMode function that the IDE has stopped
    Const MOD_USER As String = "user32.dll"                   'Location of the SetWindowLongA & CallWindowProc functions
    Const MOD_VBA5 As String = "vba5"                     'Location of the EbMode function if running VB5
    Const MOD_VBA6 As String = "vba6"                     'Location of the EbMode function if running VB6
    Const PATCH_01 As Long = 18                           'Code buffer offset to the location of the relative address to EbMode
    Const PATCH_02 As Long = 68                           'Address of the previous WndProc
    Const PATCH_03 As Long = 78                           'Relative address of SetWindowsLong
    Const PATCH_06 As Long = 116                          'Address of the previous WndProc
    Const PATCH_07 As Long = 121                          'Relative address of CallWindowProc
    Const PATCH_0A As Long = 186                          'Address of the owner object
    Static aBuf(1 To CODE_LEN) As Byte                                            'Static code buffer byte array
    Static pCWP As Long                             'Address of the CallWindowsProc
    Static pEbMode As Long                                'Address of the EbMode IDE break/stop/running function
    Static pSWL As Long                             'Address of the SetWindowsLong function
    Dim i As Long                       'Loop index
    Dim j As Long                       'Loop index
    Dim nSubIdx As Long                             'Subclass data index
    Dim sHex As String                        'Hex code string

    'If it's the first time through here..
    If aBuf(1) = 0 Then

        'The hex pair machine code representation.
        sHex = "5589E583C4F85731C08945FC8945F8EB0EE80000000083F802742185C07424E830000000837DF800750AE838000000E84D00" & _
               "00005F8B45FCC9C21000E826000000EBF168000000006AFCFF7508E800000000EBE031D24ABF00000000B900000000E82D00" & _
               "0000C3FF7514FF7510FF750CFF75086800000000E8000000008945FCC331D2BF00000000B900000000E801000000C3E33209" & _
               "C978078B450CF2AF75278D4514508D4510508D450C508D4508508D45FC508D45F85052B800000000508B00FF90A4070000C3"

        'Convert the string from hex pairs to bytes and store in the static machine code buffer
        i = 1
        Do While j < CODE_LEN
            j = j + 1
            aBuf(j) = Val("&H" & Mid$(sHex, i, 2))                                  'Convert a pair of hex characters to an eight-bit value and store in the static code buffer array
            i = i + 2
        Loop                                                                        'Next pair of hex characters

        'Get API function addresses
        If zSubclass_InIDE Then                                                      'If we're running in the VB IDE
            aBuf(16) = &H90                                                         'Patch the code buffer to enable the IDE state code
            aBuf(17) = &H90                                                         'Patch the code buffer to enable the IDE state code
            pEbMode = zAddrFunc(MOD_VBA6, FUNC_EBM)                                 'Get the address of EbMode in vba6.dll
            If pEbMode = 0 Then                                                     'Found?
                pEbMode = zAddrFunc(MOD_VBA5, FUNC_EBM)                             'VB5 perhaps
            End If
        End If

        pCWP = zAddrFunc(MOD_USER, FUNC_CWP)                                        'Get the address of the CallWindowsProc function
        pSWL = zAddrFunc(MOD_USER, FUNC_SWL)                                        'Get the address of the SetWindowLongA function
        ReDim sc_aSubData(0 To 0) As tSubData                                       'Create the first sc_aSubData element
    Else
        nSubIdx = zIdx(lng_hWnd, True)
        If nSubIdx = -1 Then                                                        'If an sc_aSubData element isn't being re-cycled
            nSubIdx = UBound(sc_aSubData()) + 1                                     'Calculate the next element
            ReDim Preserve sc_aSubData(0 To nSubIdx) As tSubData                    'Create a new sc_aSubData element
        End If
        zSubclass_Start = nSubIdx
    End If

    With sc_aSubData(nSubIdx)
        .hWnd = lng_hWnd                                                            'Store the hWnd
        .nAddrSub = GlobalAlloc(GMEM_FIXED, CODE_LEN)                               'Allocate memory for the machine code WndProc
        .nAddrOrig = SetWindowLongA(.hWnd, GWL_WNDPROC, .nAddrSub)                  'Set our WndProc in place
        Call RtlMoveMemory(ByVal .nAddrSub, aBuf(1), CODE_LEN)                      'Copy the machine code from the static byte array to the code array in sc_aSubData
        Call zPatchRel(.nAddrSub, PATCH_01, pEbMode)                                'Patch the relative address to the VBA EbMode api function, whether we need to not.. hardly worth testing
        Call zPatchVal(.nAddrSub, PATCH_02, .nAddrOrig)                             'Original WndProc address for CallWindowProc, call the original WndProc
        Call zPatchRel(.nAddrSub, PATCH_03, pSWL)                                   'Patch the relative address of the SetWindowLongA api function
        Call zPatchVal(.nAddrSub, PATCH_06, .nAddrOrig)                             'Original WndProc address for SetWindowLongA, unsubclass on IDE stop
        Call zPatchRel(.nAddrSub, PATCH_07, pCWP)                                   'Patch the relative address of the CallWindowProc api function
        Call zPatchVal(.nAddrSub, PATCH_0A, ObjPtr(Me))                             'Patch the address of this object instance into the static machine code buffer
    End With
End Function

'Return whether we're running in the IDE.
Private Function zSubclass_InIDE() As Boolean
    Debug.Assert zSetTrue(zSubclass_InIDE)
End Function

'Delete a message from the table of those that will invoke a callback.
Private Sub zSubclass_DelMsg(ByVal lng_hWnd As Long, ByVal uMsg As Long, Optional ByVal When As eMsgWhen = MSG_AFTER)
    'Parameters:
    'lng_hWnd  - The handle of the window for which the uMsg is to be removed from the callback table
    'uMsg      - The message number that will be removed from the callback table. NB Can also be ALL_MESSAGES, ie all messages will callback
    'When      - Whether the msg is to be removed from the before, after or both callback tables
    With sc_aSubData(zIdx(lng_hWnd))
        If When And eMsgWhen.MSG_BEFORE Then
            Call zDelMsg(uMsg, .aMsgTblB, .nMsgCntB, eMsgWhen.MSG_BEFORE, .nAddrSub)
        End If
        If When And eMsgWhen.MSG_AFTER Then
            Call zDelMsg(uMsg, .aMsgTblA, .nMsgCntA, eMsgWhen.MSG_AFTER, .nAddrSub)
        End If
    End With
End Sub

'======================================================================================================
'Subclass code - The programmer may call any of the following Subclass_??? routines

'Add a message to the table of those that will invoke a callback. You should Subclass_Subclass first and then add the messages
Private Sub zSubclass_AddMsg(ByVal lng_hWnd As Long, ByVal uMsg As Long, Optional ByVal When As eMsgWhen = MSG_AFTER)
    'Parameters:
    'lng_hWnd  - The handle of the window for which the uMsg is to be added to the callback table
    'uMsg      - The message number that will invoke a callback. NB Can also be ALL_MESSAGES, ie all messages will callback
    'When      - Whether the msg is to callback before, after or both with respect to the the default (previous) handler
    With sc_aSubData(zIdx(lng_hWnd))
        If When And eMsgWhen.MSG_BEFORE Then
            Call zAddMsg(uMsg, .aMsgTblB, .nMsgCntB, eMsgWhen.MSG_BEFORE, .nAddrSub)
        End If
        If When And eMsgWhen.MSG_AFTER Then
            Call zAddMsg(uMsg, .aMsgTblA, .nMsgCntA, eMsgWhen.MSG_AFTER, .nAddrSub)
        End If
    End With
End Sub

'Worker function for zSubclass_InIDE
Private Function zSetTrue(ByRef bValue As Boolean) As Boolean
    zSetTrue = True
    bValue = True
End Function
'======================================================================================================
'   End SubClass Sections
'======================================================================================================

'Patch the machine code buffer at the indicated offset with the passed value
Private Sub zPatchVal(ByVal nAddr As Long, ByVal nOffset As Long, ByVal nValue As Long)
    Call RtlMoveMemory(ByVal nAddr + nOffset, nValue, 4)
End Sub

'Patch the machine code buffer at the indicated offset with the relative address to the target address.
Private Sub zPatchRel(ByVal nAddr As Long, ByVal nOffset As Long, ByVal nTargetAddr As Long)
    Call RtlMoveMemory(ByVal nAddr + nOffset, nTargetAddr - nAddr - nOffset - 4, 4)
End Sub

'Get the sc_aSubData() array index of the passed hWnd
Private Function zIdx(ByVal lng_hWnd As Long, Optional ByVal bAdd As Boolean = False) As Long
    'Get the upper bound of sc_aSubData() - If you get an error here, you're probably sc_AddMsg-ing before zSubclass_Start
    zIdx = UBound(sc_aSubData)
    Do While zIdx >= 0                                                              'Iterate through the existing sc_aSubData() elements
        With sc_aSubData(zIdx)
            If .hWnd = lng_hWnd Then                                                'If the hWnd of this element is the one we're looking for
                If Not bAdd Then                                                    'If we're searching not adding
                    Exit Function                                                   'Found
                End If
            ElseIf .hWnd = 0 Then                                                   'If this an element marked for reuse.
                If bAdd Then                                                        'If we're adding
                    Exit Function                                                   'Re-use it
                End If
            End If
        End With
        zIdx = zIdx - 1                                                                 'Decrement the index
    Loop

    If Not bAdd Then
        Debug.Assert False                                                          'hWnd not found, programmer error
    End If

    'If we exit here, we're returning -1, no freed elements were found
End Function

'Worker sub for sc_DelMsg
Private Sub zDelMsg(ByVal uMsg As Long, ByRef aMsgTbl() As Long, ByRef nMsgCnt As Long, ByVal When As eMsgWhen, ByVal nAddr As Long)
    Dim nEntry As Long

    If uMsg = ALL_MESSAGES Then                                                     'If deleting all messages
        nMsgCnt = 0                                                                 'Message count is now zero
        If When = eMsgWhen.MSG_BEFORE Then                                          'If before
            nEntry = PATCH_05                                                       'Patch the before table message count location
        Else                                                                        'Else after
            nEntry = PATCH_09                                                       'Patch the after table message count location
        End If
        Call zPatchVal(nAddr, nEntry, 0)                                            'Patch the table message count to zero
    Else                                                                            'Else deleteting a specific message
        Do While nEntry < nMsgCnt                                                   'For each table entry
            nEntry = nEntry + 1
            If aMsgTbl(nEntry) = uMsg Then                                          'If this entry is the message we wish to delete
                aMsgTbl(nEntry) = 0                                                 'Mark the table slot as available
                Exit Do                                                             'Bail
            End If
        Loop                                                                        'Next entry
    End If
End Sub

'Return the memory address of the passed function in the passed dll
Private Function zAddrFunc(ByVal sDLL As String, ByVal sProc As String) As Long
    zAddrFunc = GetProcAddress(GetModuleHandleA(sDLL), sProc)
    Debug.Assert zAddrFunc                                                          'You may wish to comment out this line if you're using vb5 else the EbMode GetProcAddress will stop here everytime because we look for vba6.dll first
End Function

'======================================================================================================
'These z??? routines are exclusively called by the Subclass_??? routines.

'Worker sub for sc_AddMsg
Private Sub zAddMsg(ByVal uMsg As Long, ByRef aMsgTbl() As Long, ByRef nMsgCnt As Long, ByVal When As eMsgWhen, ByVal nAddr As Long)
    Dim nEntry As Long                                                            'Message table entry index
    Dim nOff1 As Long                                                           'Machine code buffer offset 1
    Dim nOff2 As Long                                                           'Machine code buffer offset 2

    If uMsg = ALL_MESSAGES Then                                                     'If all messages
        nMsgCnt = ALL_MESSAGES                                                      'Indicates that all messages will callback
    Else                                                                            'Else a specific message number
        Do While nEntry < nMsgCnt                                                   'For each existing entry. NB will skip if nMsgCnt = 0
            nEntry = nEntry + 1

            If aMsgTbl(nEntry) = 0 Then                                             'This msg table slot is a deleted entry
                aMsgTbl(nEntry) = uMsg                                              'Re-use this entry
                Exit Sub                                                            'Bail
            ElseIf aMsgTbl(nEntry) = uMsg Then                                      'The msg is already in the table!
                Exit Sub                                                            'Bail
            End If
        Loop                                                                        'Next entry
        nMsgCnt = nMsgCnt + 1                                                       'New slot required, bump the table entry count
        ReDim Preserve aMsgTbl(1 To nMsgCnt) As Long                                'Bump the size of the table.
        aMsgTbl(nMsgCnt) = uMsg                                                     'Store the message number in the table
    End If

    If When = eMsgWhen.MSG_BEFORE Then                                              'If before
        nOff1 = PATCH_04                                                            'Offset to the Before table
        nOff2 = PATCH_05                                                            'Offset to the Before table entry count
    Else                                                                            'Else after
        nOff1 = PATCH_08                                                            'Offset to the After table
        nOff2 = PATCH_09                                                            'Offset to the After table entry count
    End If

    If uMsg <> ALL_MESSAGES Then
        Call zPatchVal(nAddr, nOff1, VarPtr(aMsgTbl(1)))                            'Address of the msg table, has to be re-patched because Redim Preserve will move it in memory.
    End If
    Call zPatchVal(nAddr, nOff2, nMsgCnt)                                           'Patch the appropriate table entry count
End Sub

Private Sub yShowFolderBrowse()

    If Not UserControl.Ambient.UserMode Then Exit Sub

    Dim BI As BROWSEINFO
    Dim RetPath As String
    Dim i As Integer
    Dim RootPidl As Long
    Dim StartingPidl As Long
    Dim NewPathPidl As Long


    'Set the properties of the folder dialog
    With BI
        'Debug.Print "m_RootFolder="; m_RootFolder
        'Debug.Print "m_StartingFolder="; m_StartingFolder

        If PathIsDirectory(m_RootFolder) <> 0 Then
            RootPidl = ILCreateFromPathW(StrPtr(m_RootFolder))
            .pIDLRoot = RootPidl
        Else    'obtain the pidl to the special folder 'drives'
            Call SHGetSpecialFolderLocation(0&, CSIDL_DRIVES, RootPidl)
            .pIDLRoot = RootPidl
        End If

        .hwndOwner = GetDesktopWindow

        ' What style?
        .ulFlags = BIF_RETURNONLYFSDIRS Or BIF_RETURNFSANCESTORS Or BIF_BROWSEFORCOMPUTER

        ' Get address of function.
        .lpfnCallback = sc_aSubData(0).nAddrSub

        If PathIsDirectory(m_StartingFolder) <> 0 Then
            StartingPidl = ILCreateFromPathW(StrPtr(m_StartingFolder))
            .lParam = StartingPidl
        Else
            StartingPidl = RootPidl
            .lParam = StartingPidl
        End If

    End With
    '   Show the Browse For Folder dialog
    NewPathPidl = SHBrowseForFolder(BI)


    RetPath = Space$(512)
    If SHGetPathFromIDList(ByVal NewPathPidl, ByVal RetPath) <> 0 Then
        'Trim off the null chars ending the path
        'and display the returned folder
        i = InStr(RetPath, vbNullChar)
        If i > 0 Then RetPath = Left$(RetPath, i - 1)
        m_StartingFolder = RetPath
        'Free memory allocated for PIDL
        CoTaskMemFree NewPathPidl
    End If

    ' Free our allocated string pointer
    If (StartingPidl <> 0 And StartingPidl <> RootPidl) Then Call ILFree(StartingPidl)
    If RootPidl <> 0 Then Call ILFree(RootPidl)


    If IsWindow(m_BrowserDlgHwnd) <> 0 Then
        DestroyWindow m_BrowserDlgHwnd
    End If

    m_SysTreeView32Hwnd = 0
    m_BrowserDlgHwnd = 0
    m_ButtonDown = False
    UserControl.Refresh

End Sub

Private Sub yPaint(ByVal ButtonDown As Boolean, _
                   ByVal IsSelected As Boolean, _
                   ByVal IsEnabled As Boolean, _
                   ByVal DisplayName As String, _
                   ByVal IconHandle As Long)


    Dim llngHdc As Long
    Dim llngState As Long
    Dim hTheme As Long
    Dim lForeColor As Long

    UserControl.Cls

    llngHdc = UserControl.hdc



    'Draw Border
    If m_UseXpTheme Then
        hTheme = OpenThemeData(0&, StrPtr("Edit"))
        DrawThemeBackground hTheme, llngHdc, 1, 1, m_UserControlRect, ByVal &O0
        CloseThemeData hTheme
    Else
        DrawEdge ByVal llngHdc, m_UserControlRect, BDR_SUNKEN, BF_RECT
    End If


    'Draw Button
    If m_UseXpTheme Then
        hTheme = OpenThemeData(0&, StrPtr("ComboBox"))
        If IsEnabled Then
            If ButtonDown Then
                llngState = 3
            Else
                llngState = 1
            End If
        Else
            llngState = 4
        End If
        Call DrawThemeBackground(hTheme, llngHdc, 1, llngState, m_ButtonRect, ByVal &O0)
        Call CloseThemeData(hTheme)
    Else
        If IsEnabled Then
            If ButtonDown Then
                llngState = DFCS_SCROLLDOWN Or DFCS_PUSHED Or DFCS_FLAT
            Else
                llngState = DFCS_SCROLLDOWN
            End If
        Else
            llngState = DFCS_SCROLLDOWN Or DFCS_INACTIVE
        End If

        Call DrawFrameControl(ByVal llngHdc, m_ButtonRect, DFC_SCROLL, llngState)
    End If




    If Len(DisplayName) > 1 Then
        lForeColor = UserControl.ForeColor
        m_TextRect.Right = UserControl.TextWidth(DisplayName) + m_TextRect.Left + 2
        
        '************************************
        'Set background and forecolor appropriately
        If IsSelected And IsEnabled Then
            FillRect llngHdc, m_TextRect, GetSysColorBrush(COLOR_HIGHLIGHT)     ' paint blue background for selection
            SetTextColor llngHdc, GetSysColor(COLOR_HIGHLIGHTTEXT)     ' write in a color that will be readable through blue background
        ElseIf Not IsEnabled Then
            FillRect llngHdc, m_TextRect, GetSysColorBrush(COLOR_MENU)     ' paint gray background
            SetTextColor llngHdc, vbWhite     ' text white for disabled (we'll write this text again as gray, offset by one pixel to look 'disabled')
        Else
            FillRect llngHdc, m_TextRect, GetSysColorBrush(COLOR_MENU)     ' paint gray background
            SetTextColor llngHdc, lForeColor 'GetSysColor(COLOR_MENUTEXT)     ' normal text color
        End If


        '*****************************************
        'Do the caption
        'OffsetRect m_TextRect, 26, 0
        SetBkMode llngHdc, TRANSPARENT     ' write text transparent
        'SetTextColor llngHdc, GetSysColor(COLOR_MENUTEXT)
        'If IsSelected Then SetTextColor llngHdc, GetSysColor(COLOR_HIGHLIGHTTEXT)
        'If Not IsEnabled Then SetTextColor llngHdc, vbWhite
        DrawText llngHdc, DisplayName, Len(DisplayName), m_TextRect, DT_LEFT Or DT_BOTTOM
        If Not IsEnabled Then
            'Do it again with gray and offset by one pixel
            SetTextColor llngHdc, GetSysColor(COLOR_GRAYTEXT)
            OffsetRect m_TextRect, -1, -1
            DrawText llngHdc, DisplayName, Len(DisplayName), m_TextRect, DT_LEFT Or DT_BOTTOM
            OffsetRect m_TextRect, 1, 1
        End If
    End If

    '*****************************************
    'Draw the icon
    If IconHandle <> 0 Then
        DrawIconEx llngHdc, m_IconRect.Left, m_IconRect.Top, IconHandle, m_IconRect.Right - m_IconRect.Left, m_IconRect.Bottom - m_IconRect.Top, 0, 0, DI_NORMAL
    End If

End Sub

Public Property Get XpTheme() As enmXPTheme

    XpTheme = m_XPTheme
End Property

Public Property Let XpTheme(Value As enmXPTheme)
    Dim lblnUseXpTheme As Boolean

    If Not Value = m_XPTheme Then
        m_XPTheme = Value
        PropertyChanged "XPTheme"
    End If

    If m_XPTheme = [As Application] Then
        lblnUseXpTheme = pvVisualStylesEnabled
    ElseIf m_XPTheme = [Disbladed Theme] Then
        lblnUseXpTheme = False
    Else    'If m_XPTheme = [Enabled Theme] Then
        lblnUseXpTheme = pvVisualStylesPossible
    End If

    If m_UseXpTheme <> lblnUseXpTheme Then
        m_UseXpTheme = lblnUseXpTheme
        Call pvGetRects
        If Not UserControl.Ambient.UserMode Then
            UserControl.Refresh
        End If
    End If


End Property

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    With PropBag
        Call .WriteProperty("Enabled", UserControl.Enabled, True)
        Call .WriteProperty("RootFolder", LocalPathToInvariantPath(m_RootFolder), vbNullString)
        Call .WriteProperty("StartingFolder", LocalPathToInvariantPath(m_StartingFolder), vbNullString)
        Call .WriteProperty("FontName", UserControl.Font.Name, "MS Sans Serif")
        Call .WriteProperty("FontCharset", UserControl.Font.Charset)
        Call .WriteProperty("ForeColor", UserControl.ForeColor, vbWindowText)
        Call .WriteProperty("RootNodeBold", m_RootNodeBold, m_RootNodeBold_ByDef)
        Call .WriteProperty("RootNodeHighLight", m_RootNodeHighLight, m_RootNodeHighLight_ByDef)
        Call .WriteProperty("StartingNodeBold", m_StartingNodeBold, m_StartingNodeBold_ByDef)
        Call .WriteProperty("StartingNodeHighLight", m_StartingNodeHighLight, m_StartingNodeHighLight_ByDef)
        Call .WriteProperty("ResolveCSIDL", m_ResolveCSIDL, m_ResolveCSIDL_ByDef)
        Call .WriteProperty("ResolveSharedFolder", m_ResolveSharedFolder, m_ResolveSharedFolder_ByDef)
        Call .WriteProperty("ResolveNetworkDrive", m_ResolveNetworkDrive, m_ResolveNetworkDrive_ByDef)
        Call .WriteProperty("XPTheme", m_XPTheme, m_XPTheme_ByDef)
    End With

End Sub

Private Sub UserControl_Terminate()
    On Error GoTo ToBad

    If m_IsSubClassing Then
        'Stop all subclassing
        Call zSubclass_StopAll
        '   Set our Flag that were done....
        m_IsSubClassing = False
    End If

ToBad:
        Call pvDestroyFont

End Sub

Private Sub UserControl_Show()
    Call pvGetRects
End Sub

Private Sub UserControl_Resize()
    On Error GoTo UserControl_FINISH
    Dim llnHeight As Long

    If m_ScrollBoxWidth = 0 Then m_ScrollBoxWidth = GetSystemMetrics(ByVal SM_CXHTHUMB)

    If Not UserControl.Ambient.UserMode Then
        llnHeight = UserControl.ScaleY(m_ScrollBoxWidth + 4, vbPixels, UserControl.Parent.ScaleMode)

        If UserControl.Height <> llnHeight Then
            Debug.Print llnHeight
            UserControl.Height = llnHeight
            Exit Sub
        End If
    End If
    Call pvGetRects

    Exit Sub
UserControl_FINISH:
    Err.Clear
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)


    With PropBag

        UserControl.Enabled = .ReadProperty("Enabled", True)
        UserControl.FontName = .ReadProperty("FontName", UserControl.Ambient.Font.Name)
        UserControl.Font.Charset = .ReadProperty("FontCharset", UserControl.Ambient.Font.Charset)
        UserControl.ForeColor = .ReadProperty("ForeColor", vbWindowText)

        m_RootNodeBold = .ReadProperty("RootNodeBold", m_RootNodeBold_ByDef)
        m_RootNodeHighLight = .ReadProperty("RootNodeHighLight", m_RootNodeHighLight_ByDef)
        m_StartingNodeBold = .ReadProperty("StartingNodeBold", m_StartingNodeBold_ByDef)
        m_StartingNodeHighLight = .ReadProperty("StartingNodeHighLight", m_StartingNodeHighLight_ByDef)



        m_ResolveCSIDL = .ReadProperty("ResolveCSIDL", m_ResolveCSIDL_ByDef)
        m_ResolveSharedFolder = .ReadProperty("ResolveSharedFolder", m_ResolveSharedFolder_ByDef)
        m_ResolveNetworkDrive = .ReadProperty("ResolveNetworkDrive", m_ResolveNetworkDrive_ByDef)

        Call pvGetResolvedFolders

        m_RootFolder = .ReadProperty("RootFolder", vbNullString)
        If Len(m_RootFolder) > 0 Then
            m_RootFolder = InvariantPathToLocalPath(m_RootFolder)
        End If
        m_StartingFolder = .ReadProperty("StartingFolder", vbNullString)
        If Len(m_StartingFolder) > 0 Then
            m_StartingFolder = InvariantPathToLocalPath(m_StartingFolder)
        End If
        m_XPTheme = .ReadProperty("XPTheme", m_XPTheme_ByDef)
        If m_XPTheme = [As Application] Then
            m_UseXpTheme = pvVisualStylesEnabled
        ElseIf m_XPTheme = [Disbladed Theme] Then
            m_UseXpTheme = False
        Else    'If m_XPTheme = [Enabled Theme] Then
            m_UseXpTheme = pvVisualStylesPossible
        End If
    End With



    If (Ambient.UserMode) Then
        'If we're not in design mode

        If Not (UserControl.Parent Is Nothing) Then
            If TypeOf UserControl.Parent Is VB.MDIForm Then
                m_OwnerFormHwnd = UserControl.Extender.Parent.hWnd
            ElseIf TypeOf UserControl.Parent Is VB.Form Then
                If UserControl.Parent.MDIChild Then
                    m_OwnerFormHwnd = GetAncestor(UserControl.Parent.hWnd, GA_ROOT)
                Else
                    m_OwnerFormHwnd = UserControl.Parent.hWnd
                End If
            ElseIf TypeOf UserControl.Parent Is VB.PropertyPage Then
                m_OwnerFormHwnd = FindWindowEx(0, 0, "IDEOwner", vbNullString)
            End If
        End If

        m_hWnd = UserControl.hWnd
        '   Start Subclassing using our Handle
        Call zSubclass_Start(m_hWnd)
        '   Subclass the BrowseForFolder Message
        Call zSubclass_AddMsg(m_hWnd, BFFM_INITIALIZED, MSG_AFTER)

        '   Store our Flag that we are Now Subclassing
        m_IsSubClassing = True
    End If

End Sub

Private Sub UserControl_Paint()
    Dim lblnIsSelected As Boolean
    Dim lblnIsEnabled As Boolean
    Dim lstrDisplayName As String
    Dim llngIconHandle As Long
    Dim llngFlags As Long
    Dim SHFI As SHFILEINFO


    'If UserControl.Ambient.UserMode Then
    lblnIsSelected = (GetFocus = UserControl.hWnd)
    lblnIsEnabled = UserControl.Enabled

    'm_ButtonDown = (IsWindow(m_BrowserDlgHwnd) = 0)
    If PathIsDirectory(m_StartingFolder) <> 0 Then
        If lblnIsSelected Then
            llngFlags = SHGFI_ICON Or SHGFI_SMALLICON Or SHGFI_SELECTED Or SHGFI_DISPLAYNAME
        Else
            llngFlags = SHGFI_ICON Or SHGFI_SMALLICON Or SHGFI_DISPLAYNAME
        End If

        Call SHGetFileInfo(m_StartingFolder, 0&, SHFI, Len(SHFI), llngFlags)

        lstrDisplayName = Left(SHFI.szDisplayName, InStr(SHFI.szDisplayName, vbNullChar) - 1)
        llngIconHandle = SHFI.hIcon

    End If

    Call yPaint(m_ButtonDown, lblnIsSelected, lblnIsEnabled, lstrDisplayName, llngIconHandle)

    If llngIconHandle <> 0 Then Call DestroyIcon(llngIconHandle)
    'End If
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If IsWindow(m_BrowserDlgHwnd) = 0 Then
        m_ButtonDown = True
        yShowFolderBrowse
    End If
End Sub

Private Sub UserControl_LostFocus()
    UserControl.Refresh
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyPageDown, vbKeyDown, vbKeyF4, vbKeyReturn
            If IsWindow(m_BrowserDlgHwnd) = 0 Then
                m_ButtonDown = True
                yShowFolderBrowse
            End If
    End Select
End Sub

Private Sub UserControl_InitProperties()
    Set Font = UserControl.Ambient.Font
    m_Enabled = True
End Sub

Private Sub UserControl_Initialize()

    m_hWnd = UserControl.hWnd

    m_RootFolder = ""
    m_RootNodeBold = m_RootNodeBold_ByDef
    m_RootNodeHighLight = m_RootNodeHighLight_ByDef
    
    m_StartingFolder = ""
    m_StartingNodeBold = m_StartingNodeBold_ByDef
    m_StartingNodeHighLight = m_StartingNodeHighLight_ByDef
    
    m_ResolveCSIDL = m_ResolveCSIDL_ByDef
    m_ResolveSharedFolder = m_ResolveSharedFolder_ByDef
    m_ResolveNetworkDrive = m_ResolveNetworkDrive_ByDef
    

End Sub

Private Sub UserControl_GotFocus()
    UserControl.Refresh
End Sub

Private Sub UserControl_ExitFocus()
    'Debug.Print "UserControl_ExitFocus"
End Sub

Private Sub UserControl_EnterFocus()
    'Debug.Print "UserControl_EnterFocus"
End Sub

' The StartingNodeHighLight property

Public Property Get StartingNodeHighLight() As Boolean
    StartingNodeHighLight = m_StartingNodeHighLight
End Property

Public Property Let StartingNodeHighLight(ByVal Value As Boolean)
    m_StartingNodeHighLight = Value
    PropertyChanged "StartingNodeHighLight"
End Property

Public Property Get StartingNodeBold() As Boolean
    StartingNodeBold = m_StartingNodeBold
End Property

Public Property Let StartingNodeBold(ByVal Value As Boolean)
    m_StartingNodeBold = Value
    PropertyChanged "StartingNodeBold"
End Property

Public Property Get StartingFolder() As String
Attribute StartingFolder.VB_ProcData.VB_Invoke_Property = "pgStartingProperty"
    StartingFolder = m_StartingFolder
End Property

Public Property Let StartingFolder(ByVal Value As String)

    If PathIsDirectory(Value) Then
        Value = pvRemoveBackslash(Value)
        If InStr(UCase$(Value), UCase$(m_RootFolder)) = 0 Then
            Value = ""
        End If
    Else
        Value = ""
    End If

    If m_StartingFolder <> Value Then
        m_StartingFolder = Value
        PropertyChanged "StartingFolder"
    End If

End Property

' The RootNodeHighLight property

Public Property Get RootNodeHighLight() As Boolean
    RootNodeHighLight = m_RootNodeHighLight
End Property

Public Property Let RootNodeHighLight(ByVal Value As Boolean)
    m_RootNodeHighLight = Value
    PropertyChanged "RootNodeHighLight"
End Property

Public Property Get RootNodeBold() As Boolean
    RootNodeBold = m_RootNodeBold
End Property

Public Property Let RootNodeBold(ByVal Value As Boolean)
    m_RootNodeBold = Value
    PropertyChanged "RootNodeBold"
End Property

Public Property Get RootFolder() As String
Attribute RootFolder.VB_ProcData.VB_Invoke_Property = "pgRootProperty"
    RootFolder = m_RootFolder
End Property

Public Property Let RootFolder(ByVal Value As String)

    If PathIsDirectory(Value) <> 0 Then
        Value = pvRemoveBackslash(Value)
    Else
        Value = ""
    End If

    If m_RootFolder <> Value Then
        m_RootFolder = Value
        PropertyChanged "RootFolder"
        Me.StartingFolder = m_StartingFolder
    End If


End Property

' To make local shared folder invariant, only available at design time .
Public Property Get ResolveSharedFolder() As Boolean
Attribute ResolveSharedFolder.VB_Description = "To make local shared folder invariant, only available at design time ."
    ResolveSharedFolder = m_ResolveSharedFolder
End Property

Public Property Let ResolveSharedFolder(ByVal Value As Boolean)
    If UserControl.Ambient.UserMode = False Then
        If m_ResolveSharedFolder <> Value Then
            m_ResolveSharedFolder = Value
            Call pvGetResolvedFolders(True)
            PropertyChanged "ResolveSharedFolder"
        End If
    End If
End Property

' To make Network Drive invariant
Public Property Get ResolveNetworkDrive() As Boolean
Attribute ResolveNetworkDrive.VB_Description = "To make Network Drive invariant"
    ResolveNetworkDrive = m_ResolveNetworkDrive
End Property

Public Property Let ResolveNetworkDrive(ByVal Value As Boolean)
    If UserControl.Ambient.UserMode = False Then
        If m_ResolveNetworkDrive <> Value Then
            m_ResolveNetworkDrive = Value
            Call pvGetResolvedFolders(True)
            PropertyChanged "ResolveNetworkDrive"
        End If
    End If
End Property

' To make a limited set of CSIDL-Folders invariant.

Public Property Get ResolveCSIDL() As Boolean
Attribute ResolveCSIDL.VB_Description = "To make a limited set of CSIDL-Folders invariant."
    ResolveCSIDL = m_ResolveCSIDL
End Property

Public Property Let ResolveCSIDL(ByVal Value As Boolean)
    If UserControl.Ambient.UserMode = False Then
        If m_ResolveCSIDL <> Value Then
            m_ResolveCSIDL = Value
            Call pvGetResolvedFolders(True)
            PropertyChanged "ResolveCSIDL"
        End If
    End If
End Property

Public Sub Refresh()
    UserControl.Refresh
End Sub

Private Property Get pvVisualStylesPossible() As Boolean
    On Error GoTo VisualStylesIsImpossible
    Dim OS As OSVersionInfo

    OS.dwOSVersionInfoSize = Len(OS)
    Call GetVersionEx(OS)

    If ((OS.dwPlatformId = VER_PLATFORM_WIN32_NT) And ( _
        ((OS.dwMajorVersion = 5) And (OS.dwMinorVersion >= 1)) Or (OS.dwMajorVersion > 5))) Then
        pvVisualStylesPossible = (CBool(IsThemeActive()) And CBool(IsAppThemed()))
    End If
    Exit Property
VisualStylesIsImpossible:
End Property

Private Property Get pvVisualStylesEnabled() As Boolean
    Dim OS As OSVersionInfo
    Dim Version As DLLVersionInfo

    OS.dwOSVersionInfoSize = Len(OS)
    Call GetVersionEx(OS)

    If ((OS.dwPlatformId = VER_PLATFORM_WIN32_NT) And ( _
        ((OS.dwMajorVersion = 5) And (OS.dwMinorVersion >= 1)) Or (OS.dwMajorVersion > 5))) Then
        Version.cbSize = Len(Version)

        If (DllGetVersion(Version) = 0) Then _
           pvVisualStylesEnabled = (Version.dwMajorVersion > 5) And _
           CBool(IsThemeActive()) And CBool(IsAppThemed())
    End If

End Property

Private Sub pvTVSetState(hwndTV As Long)
    Dim lState As Long
    Dim TVI As TV_ITEM
    Dim hitemTV As Long

    'If the item is selected, use TVGN_CARET.
    'To highlight the first item in the root, use TVGN_ROOT
    'To hilight the first visible, use TVGN_FIRSTVISIBLE
    'To hilight the selected item, use TVGN_CARET
    If PathIsDirectory(m_RootFolder) <> 0 And (m_RootNodeBold Or m_RootNodeHighLight) Then
        If m_RootNodeBold And m_RootNodeHighLight Then
            lState = TVIS_DROPHILITED Or TVIS_BOLD
        ElseIf m_RootNodeBold Then
            lState = TVIS_BOLD
        Else
            lState = TVIS_DROPHILITED
        End If

        hitemTV = SendMessageAsLong(hwndTV, TVM_GETNEXTITEM, TVGN_ROOT, ByVal 0&)
        'if a valid handle get and set the item's state attributes
        If hitemTV > 0 Then
            With TVI
                .hItem = hitemTV
                .mask = TVIF_STATE
                .stateMask = lState
                Call SendMessageAsAny(hwndTV, TVM_GETITEM, 0&, TVI)
                .state = lState
            End With
            Call SendMessageAsAny(hwndTV, TVM_SETITEM, 0&, TVI)
        End If
    End If
    
    If PathIsDirectory(m_StartingFolder) <> 0 And (m_StartingNodeBold Or m_StartingNodeHighLight) Then
        If m_StartingNodeBold And m_StartingNodeHighLight Then
            lState = TVIS_DROPHILITED Or TVIS_BOLD
        ElseIf m_StartingNodeBold Then
            lState = TVIS_BOLD
        Else
            lState = TVIS_DROPHILITED
        End If

        hitemTV = SendMessageAsLong(hwndTV, TVM_GETNEXTITEM, TVGN_CARET, ByVal 0&)
        'if a valid handle get and set the item's state attributes
        If hitemTV > 0 Then
            With TVI
                .hItem = hitemTV
                .mask = TVIF_STATE
                .stateMask = lState
                Call SendMessageAsAny(hwndTV, TVM_GETITEM, 0&, TVI)
                .state = lState
            End With
            Call SendMessageAsAny(hwndTV, TVM_SETITEM, 0&, TVI)
        End If
    End If
    

End Sub

Private Function pvStrFromPtrW(ByVal dwData As Long) As String

    Dim tmp() As Byte
    Dim tmplen As Long

    If dwData <> 0 Then
        tmplen = lstrlenW(dwData) * 2
        If tmplen <> 0 Then
            ReDim tmp(0 To (tmplen - 1)) As Byte
            RtlMoveMemory tmp(0), ByVal dwData, tmplen
            pvStrFromPtrW = tmp
        End If

    End If

End Function

Private Function pvStrFromPtrA(ByVal lpszA As Long) As String
    Dim i As Long

    pvStrFromPtrA = String$(lstrlenA(ByVal lpszA), 0)
    Call lstrcpyA(ByVal pvStrFromPtrA, ByVal lpszA)
    i = InStr(pvStrFromPtrA, Chr$(0))
    If i > 0 Then
        pvStrFromPtrA = Left$(pvStrFromPtrA, InStr(pvStrFromPtrA, Chr$(0)) - 1)
    End If

End Function

Private Function pvStdFontToFontIndirect() As Boolean
    On Error GoTo pvStdFontToFontIndirect_ERR
    Dim uLogFont As LOGFONT
    Dim lChar As Long

    If (m_hFont) Then
        If (DeleteObject(m_hFont)) Then
            m_hFont = 0
        End If
    End If

     With uLogFont
         
         For lChar = 1 To Len(UserControl.Font.Name)
             .lfFaceName(lChar - 1) = CByte(Asc(Mid$(UserControl.Font.Name, lChar, 1)))
         Next lChar
         .lfHeight = -MulDiv(UserControl.Font.Size, GetDeviceCaps(UserControl.hdc, LOGPIXELSY), 72)
         .lfItalic = UserControl.Font.Italic
         .lfWeight = IIf(UserControl.Font.Bold, FW_BOLD, FW_NORMAL)
         .lfUnderline = UserControl.Font.Underline
         .lfStrikeOut = UserControl.Font.Strikethrough
         .lfCharSet = UserControl.Font.Charset
    End With
    
    m_hFont = CreateFontIndirect(uLogFont)
    pvStdFontToFontIndirect = (m_hFont <> 0)
    Exit Function

pvStdFontToFontIndirect_ERR:
    If (m_hFont) Then
        If (DeleteObject(m_hFont)) Then
            m_hFont = 0
        End If
    End If
End Function

Private Function pvRemoveBackslash(ByVal Path As String) As String
    Dim NewPath As String
    NewPath = Left$(Trim$(Path) & String(MAX_PATH, 0), MAX_PATH)
    Call PathRemoveBackslash(NewPath)
    pvRemoveBackslash = Left$(NewPath, InStr(NewPath, vbNullChar) - 1)
End Function

Private Function pvLongFromPtr(ByVal lpDWord As Long) As Long
    Call RtlMoveMemory(pvLongFromPtr, ByVal lpDWord, 4)
End Function

Private Function pvGetResolvedFolders(Optional ByVal Reset As Boolean = False) As Boolean

    If Reset Then
        m_NbrOfInvariantFolders = 0
        Erase m_InvariantFolders
    End If

    If m_NbrOfInvariantFolders = 0 Then

        If m_ResolveCSIDL Then
            Call pvGetCsIdlPath(CSIDL_DESKTOPDIRECTORY, "CSIDL_DESKTOPDIRECTORY")
            Call pvGetCsIdlPath(CSIDL_PERSONAL, "CSIDL_PERSONAL")
            Call pvGetCsIdlPath(CSIDL_TEMPLATES, "CSIDL_TEMPLATES")
            Call pvGetCsIdlPath(CSIDL_MYMUSIC, "CSIDL_MYMUSIC")
            Call pvGetCsIdlPath(CSIDL_MYPICTURES, "CSIDL_MYPICTURES")
            Call pvGetCsIdlPath(CSIDL_MYVIDEO, "CSIDL_MYVIDEO")
            Call pvGetCsIdlPath(CSIDL_APPDATA, "CSIDL_APPDATA")
            Call pvGetCsIdlPath(CSIDL_COMMON_TEMPLATES, "CSIDL_COMMON_TEMPLATES")
            Call pvGetCsIdlPath(CSIDL_COMMON_DOCUMENTS, "CSIDL_COMMON_DOCUMENTS")
            Call pvGetCsIdlPath(CSIDL_COMMON_MUSIC, "CSIDL_COMMON_MUSIC")
            Call pvGetCsIdlPath(CSIDL_COMMON_PICTURES, "CSIDL_COMMON_PICTURES")
            Call pvGetCsIdlPath(CSIDL_COMMON_VIDEO, "CSIDL_COMMON_VIDEO")
        End If

        If m_ResolveSharedFolder Then
            Call pvGetLocalSharedFolder
        End If

        If m_ResolveNetworkDrive Then
            Call pvGetNetworkDrive
        End If

    End If
End Function

Private Sub pvGetRects()

    Call GetClientRect(UserControl.hWnd, m_UserControlRect)

    If m_UseXpTheme Then
        Call SetRect(m_IconRect, 2, m_UserControlRect.Bottom - 2 - 16, 18, m_UserControlRect.Bottom - 2)
        Call SetRect(m_TextRect, m_IconRect.Right + 2, 2, 200, m_UserControlRect.Bottom - 2)
        Call SetRect(m_ButtonRect, m_UserControlRect.Right - m_ScrollBoxWidth - 1, 1, m_UserControlRect.Right - 1, m_UserControlRect.Bottom - 1)
    Else
        Call SetRect(m_IconRect, 2, m_UserControlRect.Bottom - 2 - 16, 18, m_UserControlRect.Bottom - 2)
        Call SetRect(m_TextRect, m_IconRect.Right + 2, 3, 200, m_UserControlRect.Bottom - 3)
        Call SetRect(m_ButtonRect, m_UserControlRect.Right - m_ScrollBoxWidth - 2, 2, m_UserControlRect.Right - 2, m_UserControlRect.Bottom - 2)
    End If

End Sub

Private Function pvGetNetworkDrive() As Boolean

    Dim hEnum As Long
    Dim dwBuffSize As Long
    Dim nStructSize As Long
    Dim dwEntries As Long
    Dim cnt As Long
    Dim success As Long
    Dim netres() As NETRESOURCE


    'obtain an enumeration handle that can be used in
    'a subsequent call to WNetEnumResource
    success = WNetOpenEnum(RESOURCE_CONNECTED, _
                           RESOURCETYPE_ANY, _
                           0&, _
                           ByVal 0&, _
                           hEnum)

    'if no error and a handle obtained..
    If success = NO_ERROR And hEnum <> 0 Then
        'set number of dwEntries and redim a NETRESOURCE array
        'to hold the data returned
        dwEntries = 1024
        ReDim netres(0 To dwEntries - 1) As NETRESOURCE

        'calculate the size of the buffer
        'being passed
        nStructSize = LenB(netres(0))
        dwBuffSize = 1024& * nStructSize

        'and call WNetEnumResource
        success = WNetEnumResource(hEnum, _
                                   dwEntries, _
                                   netres(0), _
                                   dwBuffSize)

        If success = 0 Then

            'loop through the returned data
            For cnt = 0 To dwEntries - 1

                'If the returned NETRESOURCE members are valid
                If netres(cnt).lpLocalName <> 0 And _
                   netres(cnt).lpRemoteName <> 0 Then

                    'Get the local name  (drive letter) returned
                    m_NbrOfInvariantFolders = m_NbrOfInvariantFolders + 1
                    ReDim Preserve m_InvariantFolders(0 To m_NbrOfInvariantFolders) As typInvariantFolder
                    m_InvariantFolders(m_NbrOfInvariantFolders).LocalName = pvStrFromPtrA(netres(cnt).lpLocalName)
                    m_InvariantFolders(m_NbrOfInvariantFolders).AliasName = pvStrFromPtrA(netres(cnt).lpRemoteName)

                End If

            Next cnt
        End If  'If success = 0 (WNetEnumResource)
    End If  'If success = 0 (WNetOpenEnum)

    'clean up
    Call WNetCloseEnum(hEnum)

End Function

'
Private Function pvGetLocalSharedFolder() As Boolean

    Dim Level As Long
    Dim lpBuffer As Long
    Dim EntriesRead As Long
    Dim Offset As Long
    Dim nRet As Long
    Dim i As Long
    Dim ShareType As Long
    Dim sComputerName As String


    sComputerName = "\\" & Environ$("COMPUTERNAME") & "\"


    ' convert ServerName to null pointer => asking all available shares on the local machine.
    Level = 2    ' try level 2 first
    nRet = NetShareEnum(StrPtr(vbNullChar), Level, lpBuffer, MAX_PREFERRED_LENGTH, EntriesRead, 0&, 0&)

    If nRet = ERROR_ACCESS_DENIED Then
        ' bummer -- need admin privs for level 2, drop to level 1
        Level = 1
        nRet = NetShareEnum(StrPtr(vbNullChar), Level, lpBuffer, MAX_PREFERRED_LENGTH, EntriesRead, 0&, 0&)
    End If

    If nRet = NO_ERROR Then
        ' make sure there are shares to decipher
        If EntriesRead > 0 Then
            ' loop through API buffer, extracting each element
            For i = 0 To EntriesRead - 1
                ShareType = pvLongFromPtr(lpBuffer + Offset + 4)
                If ShareType = STYPE_DISKTREE Then
                    'Get the local name  (drive letter) returned
                    m_NbrOfInvariantFolders = m_NbrOfInvariantFolders + 1
                    ReDim Preserve m_InvariantFolders(0 To m_NbrOfInvariantFolders) As typInvariantFolder
                    m_InvariantFolders(m_NbrOfInvariantFolders).LocalName = pvStrFromPtrW(pvLongFromPtr(lpBuffer + Offset + 24))
                    m_InvariantFolders(m_NbrOfInvariantFolders).AliasName = sComputerName & pvStrFromPtrW(pvLongFromPtr(lpBuffer + Offset))
                End If
                If Level = 2 Then
                    Offset = Offset + 32    'Len(SHARE_INFO_2)
                Else
                    Offset = Offset + 12  ' Len(SHARE_INFO_1)
                End If
            Next i
        End If
    End If

    ' clean up
    If lpBuffer Then Call NetApiBufferFree(lpBuffer)

End Function

Private Function pvGetCsIdlPath(ByVal CSIDL As Long, ByVal AliasName As String) As Boolean

    Dim sPath As String
    Dim Pidl As Long

    'fill the idl structure with the specified folder item
    If SHGetSpecialFolderLocation(0&, CSIDL, Pidl) = 0 Then
        'NB 0 => Ok
        'if the pidl is returned, initialize  and get the path from the id list
        sPath = Space$(MAX_PATH)
        If SHGetPathFromIDList(ByVal Pidl, ByVal sPath) Then
            'return the path
            m_NbrOfInvariantFolders = m_NbrOfInvariantFolders + 1
            ReDim Preserve m_InvariantFolders(0 To m_NbrOfInvariantFolders) As typInvariantFolder
            m_InvariantFolders(m_NbrOfInvariantFolders).LocalName = Left(sPath, InStr(sPath, Chr$(0)) - 1)
            m_InvariantFolders(m_NbrOfInvariantFolders).AliasName = "[" & AliasName & "]"    '"[&H" & Hex$(CSIDL) & "]"
            pvGetCsIdlPath = True
        End If
        'free the pidl
        Call CoTaskMemFree(Pidl)
    End If

End Function

Private Sub pvDestroyFont()

    If (m_hFont) Then
        If (DeleteObject(m_hFont)) Then
            m_hFont = 0
        End If
    End If
End Sub

Public Function LocalPathToInvariantPath(ByVal LocalPath As String) As String

    'On Error GoTo LocalPathToInvariantPath_ERR
    Dim i As Long
    Dim llngMaxLen As Long
    Dim lstrAliasName As String

    LocalPath = Trim$(LocalPath)

    For i = 1 To m_NbrOfInvariantFolders
        If InStr(1, LocalPath, m_InvariantFolders(i).LocalName, vbTextCompare) = 1 Then
            If Len(m_InvariantFolders(i).LocalName) > llngMaxLen Then
                llngMaxLen = Len(m_InvariantFolders(i).LocalName)
                lstrAliasName = m_InvariantFolders(i).AliasName
            End If
        End If
    Next i

    If llngMaxLen > 0 Then
        LocalPathToInvariantPath = lstrAliasName & Right$(LocalPath, Len(LocalPath) - llngMaxLen)
    Else
        LocalPathToInvariantPath = LocalPath
    End If

End Function

Public Function InvariantPathToLocalPath(ByVal InvariantPath As String) As String

    Dim i As Long
    Dim llngMaxLen As Long
    Dim lstrLocalName As String

    InvariantPath = Trim$(InvariantPath)
    For i = 1 To m_NbrOfInvariantFolders
        If InStr(1, InvariantPath, m_InvariantFolders(i).AliasName, vbTextCompare) = 1 Then
            If Len(m_InvariantFolders(i).AliasName) > llngMaxLen Then
                llngMaxLen = Len(m_InvariantFolders(i).AliasName)
                lstrLocalName = m_InvariantFolders(i).LocalName
                If Left$(m_InvariantFolders(i).AliasName, 1) = "[" Then
                    'It's an invariant CSIDL Alias (example: [CSIDL_PERSONAL])
                    'so we don't need to search more
                    Exit For
                End If
            End If
        End If
    Next i

    If llngMaxLen > 0 Then
        InvariantPathToLocalPath = lstrLocalName & Right$(InvariantPath, Len(InvariantPath) - llngMaxLen)
    Else
        InvariantPathToLocalPath = InvariantPath
    End If

End Function

Public Property Get ForeColor() As OLE_COLOR
    ForeColor = UserControl.ForeColor
End Property

Public Property Let ForeColor(ByVal Value As OLE_COLOR)
    Dim lColor As Long
    Call OleTranslateColor(Value, 0, lColor)
    If UserControl.ForeColor <> lColor Then
        UserControl.ForeColor = lColor
        PropertyChanged "ForeColor"
    End If
End Property

Public Property Get Font() As StdFont
    Set Font = UserControl.Font
End Property

Public Property Set Font(newFont As StdFont)
    UserControl.Font.Name = newFont.Name
    UserControl.Font.Charset = newFont.Charset
End Property

' Determines whether the object can respond to user-generated events
Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Determines whether the object can respond to user-generated events"
Attribute Enabled.VB_UserMemId = -514
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal Value As Boolean)
    UserControl.Enabled = Value
    UserControl_Paint
End Property

Public Property Get CurrentFolder() As String
Attribute CurrentFolder.VB_MemberFlags = "400"
    CurrentFolder = m_CurrentFolder
End Property
Public Property Get DropDownHeight() As Long
    DropDownHeight = m_DropDownHeight
End Property
Public Property Let DropDownHeight(ByVal Value As Long)
    m_DropDownHeight = Value
    Call PropertyChanged("DropDownHeight")
End Property

Public Property Get DropDownWidth() As Long
    DropDownWidth = m_DropDownWidth
End Property
Public Property Let DropDownWidth(ByVal Value As Long)

    If (Value < -1) Then
        Err.Raise 380
        Exit Property
    Else

    'm_DropDownWidthAutoSize = (Value = UserControl.Width)

    m_DropDownWidth = Value

    Call PropertyChanged("DrowDownWidth")
    End If
End Property

