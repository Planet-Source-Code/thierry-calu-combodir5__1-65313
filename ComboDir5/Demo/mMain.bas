Attribute VB_Name = "mMain"
Option Explicit

Public Type tagInitCommonControlsEx
   lngSize As Long
   lngICC As Long
End Type
Public Declare Function InitCommonControlsEx Lib "comctl32.dll" (iccex As tagInitCommonControlsEx) As Boolean
Public Const ICC_USEREX_CLASSES = &H200

Public Sub Main()

    'Goto http://www.vbaccelerator.com/home/VB/Code/Libraries/XP%5FVisual%5FStyles/Using_XP_Visual_Styles_in_VB/article.asp
   ' we need to call InitCommonControls before we can use XP visual styles.  Here I'm using
   ' InitCommonControlsEx, which is the extended version provided in v4.72 upwards
   'NB you need v6.00 or higher to get XP styles...
   On Error Resume Next
   ' this will fail if Comctl not available
   '  - unlikely now though!
   Dim iccex As tagInitCommonControlsEx
   With iccex
       .lngSize = LenB(iccex)
       .lngICC = ICC_USEREX_CLASSES
   End With
   InitCommonControlsEx iccex
   
   ' now start the application
   On Error GoTo 0
   Form1.Show
   
End Sub


