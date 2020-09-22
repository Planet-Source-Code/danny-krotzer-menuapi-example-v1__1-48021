Attribute VB_Name = "mMenuAPI"
'__________________________________________________________________
'
' Author:   Danny K (danny@xi0n.com)
' Site:     http://www.xi0n.com
' Module:   ...for MenuAPI example
'__________________________________________________________________

Option Explicit


' API declarations
'__________________________________________________________________

Public Declare Function IsWindow Lib "user32" ( _
        ByVal hwnd As Long) As Long

Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" ( _
        ByVal hwnd As Long, _
        ByVal nIndex As Long, _
        ByVal dwNewLong As Long) As Long
        
Public Declare Function EnumWindows Lib "user32" ( _
        ByVal lpEnumFunc As Long, _
        ByVal lParam As Long) As Long

Public Declare Function GetMenuString Lib "user32" Alias "GetMenuStringA" ( _
        ByVal hMenu As Long, _
        ByVal wIDItem As Long, _
        ByVal lpString As String, _
        ByVal nMaxCount As Long, _
        ByVal wFlag As Long) As Long
        
Public Declare Function GetMenuItemID Lib "user32" ( _
        ByVal hMenu As Long, _
        ByVal nPos As Long) As Long
        
Public Declare Function GetMenuItemCount Lib "user32" ( _
        ByVal hMenu As Long) As Long
        
Public Declare Function GetSubMenu Lib "user32" ( _
        ByVal hMenu As Long, _
        ByVal nPos As Long) As Long
        
Public Declare Function GetMenu Lib "user32" ( _
        ByVal hwnd As Long) As Long

Public Declare Function AppendMenu Lib "user32" Alias "AppendMenuA" ( _
        ByVal hMenu As Long, _
        ByVal wFlags As Long, _
        ByVal wIDNewItem As Long, _
        ByVal lpNewItem As Any) As Long

Public Declare Function DeleteMenu Lib "user32" ( _
        ByVal hMenu As Long, _
        ByVal nPosition As Long, _
        ByVal wFlags As Long) As Long

Public Declare Function CreatePopupMenu Lib "user32" () As Long

Public Declare Function DrawMenuBar Lib "user32" ( _
        ByVal hwnd As Long) As Long

Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" ( _
        ByVal lpPrevWndFunc As Long, _
        ByVal hwnd As Long, _
        ByVal Msg As Long, _
        ByVal wParam As Long, _
        ByVal lParam As Long) As Long

Public Declare Function PostMessage Lib "user32" Alias "PostMessageA" ( _
        ByVal hwnd As Long, _
        ByVal wMsg As Long, _
        ByVal wParam As Long, _
        ByVal lParam As Long) As Long

Public Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" ( _
        ByVal hwnd As Long) As Long

Public Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" ( _
        ByVal hwnd As Long, _
        ByVal lpString As String, _
        ByVal cch As Long) As Long
        
Public Declare Function IsWindowVisible Lib "user32" ( _
        ByVal hwnd As Long) As Long



' Constants
'__________________________________________________________________

Public Const WM_COMMAND = &H111
Public Const GWL_WNDPROC = (-4)
Public Const MF_BYPOSITION = &H400&
Public Const MF_BYCOMMAND = &H0&
Public Const MF_POPUP = &H10&
Public Const MF_STRING = &H0&
Public Const MF_SEPARATOR = &H800&


' Misc Variables
'__________________________________________________________________

Public lWindowHwnd As Long 'Targeted Window Handle
Public plOldProc As Long   'Original WinProc Address (for subclassing)

'__________________________________________________________________
'
' Subclassing - This function processes our forms messages
'__________________________________________________________________

Public Function WindowProc(ByVal hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    
    Select Case uMsg
    
    Case WM_COMMAND
                
        'make sure its not a CommandButton from our form
        
        If wParam = 1 Then
            WindowProc = CallWindowProc(plOldProc, hwnd, uMsg, wParam, lParam)
        End If
        Debug.Print
        
        'wParam holds the menu item ID
        Call PostMessage(lWindowHwnd, WM_COMMAND, wParam, 0&)
        
        
    Case Else
        
        ' otherwise let the original procedure handle it.
        WindowProc = CallWindowProc(plOldProc, hwnd, uMsg, wParam, lParam)
        
    End Select
    
End Function

'__________________________________________________________________
'
' Processes all returned window handles from EnumWindows()
'__________________________________________________________________

Public Function EnumWindowsProc(ByVal hwnd As Long, ByVal lParam As Long) As Boolean
    
 'caption length
 Dim lLength As Long
 lLength = GetWindowTextLength(hwnd)
    
 Dim lReturn As Long
 lReturn = IsWindowVisible(hwnd)
 
 If lLength > 0 And lReturn <> 0 Then
 
    'get caption
    Dim strCaption As String
    strCaption = String$(lLength, vbNullChar)
    GetWindowText hwnd, strCaption, lLength + 1
 
    'add to list
    With frmMain.lstWindows
        .AddItem strCaption
        .ItemData(.NewIndex) = hwnd
    End With
 
 End If
 
 'all good
 EnumWindowsProc = True
 
End Function
