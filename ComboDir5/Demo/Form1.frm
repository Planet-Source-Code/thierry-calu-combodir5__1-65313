VERSION 5.00
Object = "{0B349903-6F43-41A7-B8BE-B506C90374B1}#1.0#0"; "ComboDir5.ocx"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   2595
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7740
   LinkTopic       =   "Form1"
   ScaleHeight     =   2595
   ScaleWidth      =   7740
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame fraTestInvariantPaths 
      Caption         =   "Root && Starting paths"
      Height          =   1935
      Left            =   240
      TabIndex        =   3
      Top             =   360
      Width           =   2775
      Begin VB.PictureBox Picture1 
         BorderStyle     =   0  'None
         Height          =   1455
         Left            =   120
         ScaleHeight     =   1455
         ScaleWidth      =   2295
         TabIndex        =   4
         Top             =   240
         Width           =   2295
         Begin VB.OptionButton optPaths 
            Caption         =   "App.Path"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   8
            Top             =   120
            Width           =   2055
         End
         Begin VB.OptionButton optPaths 
            Caption         =   "My Documents"
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   7
            Top             =   480
            Width           =   2055
         End
         Begin VB.OptionButton optPaths 
            Caption         =   "My Music"
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   6
            Top             =   840
            Width           =   2055
         End
         Begin VB.OptionButton optPaths 
            Caption         =   "UserControl property"
            Height          =   255
            Index           =   3
            Left            =   120
            TabIndex        =   5
            Top             =   1200
            Value           =   -1  'True
            Width           =   2055
         End
      End
   End
   Begin ComboDir5.ucComboDir5 ucComboDir51 
      Height          =   315
      Left            =   3480
      TabIndex        =   2
      Top             =   1680
      Width           =   3975
      _extentx        =   6376
      _extenty        =   556
      fontcharset     =   0
      forecolor       =   -2147483630
      startingnodebold=   -1
      rootfolder      =   "[CSIDL_PERSONAL]"
   End
   Begin VB.TextBox txtStartingPath 
      Height          =   285
      Left            =   3480
      TabIndex        =   1
      Top             =   1080
      Width           =   3975
   End
   Begin VB.TextBox txtRootPath 
      Height          =   285
      Left            =   3480
      TabIndex        =   0
      Top             =   360
      Width           =   3975
   End
   Begin VB.Label Label2 
      Caption         =   "Starting path:"
      Height          =   255
      Left            =   3480
      TabIndex        =   10
      Top             =   840
      Width           =   2295
   End
   Begin VB.Label Label1 
      Caption         =   "Root path:"
      Height          =   255
      Left            =   3480
      TabIndex        =   9
      Top             =   120
      Width           =   2295
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function InitCommonControls Lib "COMCTL32" () As Long

Private Const MAX_PATH = 260
Private Const S_FALSE = 1
Private Const S_OK = 0

Private Const CSIDL_PERSONAL As Long = &H5    'My Documents
Private Const CSIDL_MYMUSIC As Long = &HD    '"My Music" folder

Private Declare Function SHGetSpecialFolderLocation Lib "shell32.dll" (ByVal hwndOwner As Long, ByVal nFolder As Long, Pidl As Long) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32.dll" Alias "SHGetPathFromIDListA" (ByVal Pidl As Long, ByVal pszPath As String) As Long

Private Declare Sub CoTaskMemFree Lib "ole32.dll" (ByVal pv As Long)


Private Function pGetCsIdlPath(ByVal CSIDL As Long) As String

    Dim sPath As String
    Dim Pidl As Long
    Dim i As Long

    'fill the idl structure with the specified folder item
    If SHGetSpecialFolderLocation(0&, CSIDL, Pidl) = S_OK Then
        'if the pidl is returned, initialize  and get the path from the id list
        sPath = Space$(MAX_PATH)
        If SHGetPathFromIDList(ByVal Pidl, ByVal sPath) Then
            'return the path
            i = InStr(sPath, vbNullChar)
            If i Then
                sPath = Left$(sPath, i - 1)
            End If

            pGetCsIdlPath = sPath
        End If
        'free the pidl
        Call CoTaskMemFree(Pidl)
    End If

End Function



Private Sub Form_Load()
    'optPaths(0).Value = True
End Sub

Private Sub optPaths_Click(Index As Integer)
    Dim lstrAppPath As String


    Select Case Index
        Case 0
            'With traditionnal path every thinks work
            lstrAppPath = App.Path
            txtRootPath.Text = Left$(lstrAppPath, InStrRev(lstrAppPath, "\"))
            txtStartingPath.Text = lstrAppPath
        Case 1
            'here olso only the Root path work !!! why ???????
            lstrAppPath = pGetCsIdlPath(CSIDL_PERSONAL)
            txtRootPath.Text = lstrAppPath
            txtStartingPath.Text = pGetSubPath(lstrAppPath & "\")

        Case 2
            'here olso only the Root path work !!! why ???????
            lstrAppPath = pGetCsIdlPath(CSIDL_MYMUSIC)
            txtRootPath.Text = lstrAppPath
            txtStartingPath.Text = pGetSubPath(lstrAppPath & "\")
        Case 3
            txtRootPath.Text = ""
            txtStartingPath.Text = ""

    End Select

    ucComboDir51.RootFolder = txtRootPath.Text
    ucComboDir51.StartingFolder = txtStartingPath.Text
    ucComboDir51.Refresh

End Sub
Private Function pGetSubPath(Path As String) As String
    Dim lstrSubPath As String
    lstrSubPath = Dir(Path, vbDirectory)
    Do While lstrSubPath <> ""   ' Commence la boucle.
        ' Ignore le dossier courant et le dossier
        ' contenant le dossier courant.
        If lstrSubPath <> "." And lstrSubPath <> ".." Then
            ' Utilise une comparaison au niveau du bit pour
            ' vérifier que MyName est un dossier.
            If (GetAttr(Path & lstrSubPath) And vbDirectory) = vbDirectory Then
                pGetSubPath = Path & lstrSubPath
                Exit Function
            End If
        End If
        lstrSubPath = Dir   ' Extrait l'entrée suivante.
    Loop

End Function

