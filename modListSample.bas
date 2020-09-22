Attribute VB_Name = "modListSample"
Option Explicit

'API TYPE DECLARATIONS

Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type

'API FUNCTION DECLARATIONS

'Note: first two have been modified from the API Viewer version
'(all have been reformatted for readability)
Declare Function SendMessageByString Lib "user32" Alias "SendMessageA" _
    (ByVal hwnd As Long, _
    ByVal wMsg As Long, _
    ByVal wParam As Long, _
    ByVal lParam As String) As Long

Declare Function SendMessageByNum Lib "user32" Alias "SendMessageA" _
    (ByVal hwnd As Long, _
    ByVal wMsg As Long, _
    ByVal wParam As Long, _
    ByVal lParam As Long) As Long

Declare Function GetSystemMetrics Lib "user32" _
    (ByVal nIndex As Long) As Long

Declare Function GetClientRect Lib "user32" _
    (ByVal hwnd As Long, _
    lpRect As RECT) As Long

'API CONSTANT DECLARATIONS

Public Const LB_FINDSTRING = &H18F
Public Const LB_FINDSTRINGEXACT = &H1A2
Public Const LB_GETITEMHEIGHT = &H1A1
Public Const SM_CYVSCROLL = 20


Public Function ChangeInList(lst As ListBox, ByVal strOld As String, ByVal strNew As String) As Boolean
Dim lngPos As Long
    'Look for old value in list
    lngPos = FindExactInList(lst, strOld)
    If lngPos >= 0 Then
        'Found!  Change it...
        lst.List(lngPos) = strNew
        ChangeInList = True
    End If 'if lngPos=-1 (not found), function makes no change and returns False
End Function

Public Function FindExactInList(lst As ListBox, ByVal strItem As String) As Long
Dim hList As Long
    'Search list from beginning for WHOLE string
    hList = lst.hwnd
    FindExactInList = SendMessageByString(hList, LB_FINDSTRINGEXACT, -1, strItem)
End Function

Public Function FindBeginsWithInList(lst As ListBox, ByVal strItem As String) As Long
Dim hList As Long
    'Search list from beginning (the -1) for anything beginning with string
    hList = lst.hwnd
    FindBeginsWithInList = SendMessageByString(hList, LB_FINDSTRING, -1, strItem)
End Function

Public Function FindNextBeginsInList(lst As ListBox, ByVal strItem As String, lngLast As Long) As Long
Dim hList As Long
    'Search list starting at next item from last one found
    hList = lst.hwnd
    FindNextBeginsInList = SendMessageByString(hList, LB_FINDSTRING, lngLast + 1, strItem)
End Function

