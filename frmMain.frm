VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Run Menu API"
   ClientHeight    =   3780
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   8160
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3780
   ScaleWidth      =   8160
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame fraCaption 
      Caption         =   "Target Window Caption"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3135
      Left            =   240
      TabIndex        =   1
      Top             =   240
      Width           =   7695
      Begin VB.CommandButton cmdRefresh 
         Caption         =   "Refresh List"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   4
         Top             =   2520
         Width           =   3495
      End
      Begin VB.ListBox lstWindows 
         Height          =   2010
         Left            =   240
         TabIndex        =   3
         Top             =   360
         Width           =   7215
      End
      Begin VB.CommandButton cmdGrab 
         Caption         =   "Get Menu"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4080
         TabIndex        =   2
         Top             =   2520
         Width           =   3375
      End
   End
   Begin VB.CommandButton cmdRun 
      Caption         =   "Run Selected"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   6480
      Width           =   2535
   End
   Begin VB.Menu mnuMain 
      Caption         =   "Menu"
      Index           =   0
      Visible         =   0   'False
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'__________________________________________________________________
'
' Name:     MenuAPI Example
'
' Author:   Danny K. (danny@xi0n.com)
'           Violent (admin@sniip3r.com)
'
' Site:     http://www.xi0n.com
'           http://www.sniip3r.com
'
' Purpose:  How to use API to access & clone another apps menu, as well as
'           subclass your form to catch messages (for custom menu items)
'
' Notes:    This example tries to keep things relatively simple...
'           I don't have it check menu-item states (checked,enabled,etc)
'
'           Also I dont have it loop through all possible nested SubMenus
'           It will only load the main menu items and their submenus,
'           so for nested SubMenus, you'll have to add another loop.
'
'           And sometimes you'll find some menu-items don't work, such as
'           Notepad's Cut/Copy/Paste. That's usually because the app is
'           using built in system ID's.
'
'           But that just means a little research ...;)
'__________________________________________________________________
'
     

'___________________________________________________________________________
'
' Form - Subclass Window Messages to catch Menu Clicks
'___________________________________________________________________________

Private Sub Form_Load()
 
 'populate list with all open windows
 Call ListWindows
 
 'subclass - replace current WindProc address
 plOldProc = SetWindowLong(frmMain.hwnd, GWL_WNDPROC, AddressOf WindowProc)
 
End Sub

Private Sub Form_Unload(Cancel As Integer)

 'restore previous window procedure
 Call SetWindowLong(frmMain.hwnd, GWL_WNDPROC, plOldProc)

End Sub

'___________________________________________________________________________
'
' Command Button - Get Window Handle from Caption
'___________________________________________________________________________

Private Sub cmdRefresh_Click()

 'refresh windows list
 Call ListWindows

End Sub

'___________________________________________________________________________
'
' Command Button - Get Window Handle from Caption
'___________________________________________________________________________

Private Sub cmdGrab_Click()
 
 'clear any menu items we might have loaded
 Call ClearMenu
 
 'try to find the window
 lWindowHwnd = lstWindows.ItemData(lstWindows.ListIndex)

 'if window found
 If IsWindow(lWindowHwnd) <> 0 Then
    'try to acquire menu
    GetMenuInfo lWindowHwnd
 Else
    MsgBox "No window with that caption was found. Try refreshing the list...", vbExclamation
 End If

End Sub

'___________________________________________________________________________
'
' Sub - Grab menu items and list in a ListBox w/MenuID stored in ItemData
'___________________________________________________________________________

Private Sub GetMenuInfo(lHwnd As Long)

'then get menu handle
Dim lMenuHwnd As Long
lMenuHwnd = GetMenu(lHwnd)

'no menu found
If lMenuHwnd <= 0 Then
    MsgBox "Selected window doesnt have a valid menu.", vbInformation
    Exit Sub
End If

'then find menu count
Dim lMenuCount As Long
lMenuCount = GetMenuItemCount(lMenuHwnd)

'get menu items
Dim lMenuID As Long
Dim strMenuItem As String
Dim strSubMenu As String
Dim lSubHwnd As Long
Dim lSubMenuID As Long

'for each menu item
Dim i As Integer
For i = 0 To lMenuCount - 1
    
    'get menu item text
    strMenuItem = String$(100, vbNullChar)
    Call GetMenuString(lMenuHwnd, i, strMenuItem, 100, MF_BYPOSITION)
    
    'get the ID
    lMenuID = GetMenuItemID(lMenuHwnd, i)
    
    'if ID is -1, then its a submenu
    If lMenuID = -1 Then
        AddMenuItem strMenuItem, lMenuID, True
    Else
        AddMenuItem strMenuItem, lMenuID, False
        GoTo NextItem
    End If
    
    'deal with the menu item's submenu
    lSubHwnd = GetSubMenu(lMenuHwnd, i)
    
    'submenu Count
    Dim lSubCount As Long
    lSubCount = GetMenuItemCount(lSubHwnd)

    'for each submenu item
    Dim X As Long
    For X = 0 To lSubCount - 1
        
        lSubMenuID = GetMenuItemID(lSubHwnd, X)
                
        'if its a Separator
        If lSubMenuID = 0 Then
            
            AddSubMenuItem "-", 0, i
        
        'if its a submenu popup
        ElseIf lSubMenuID = -1 Then
            
            'get menu item text
            strSubMenu = String$(100, vbNullChar)
            Call GetMenuString(lSubHwnd, X, strSubMenu, 100, MF_BYPOSITION)
            
            AddSubMenuItem strSubMenu, 0, i
        
        'its a regular submenu item
        Else
            
            'get menu item text
            strSubMenu = String$(100, vbNullChar)
            Call GetMenuString(lSubHwnd, X, strSubMenu, 100, MF_BYPOSITION)
            
            'some separators still have an ID so check caption
            If Left$(strSubMenu, 1) = vbNullChar Then
                AddSubMenuItem "-", 0, i
            Else
                AddSubMenuItem strSubMenu, lSubMenuID, i
            End If
            
        End If
        
    Next 'subitems loop

NextItem:

Next 'menu items loop

End Sub


'___________________________________________________________________________
'
' Sub - Add Item to Forms Menu
'___________________________________________________________________________

Private Sub AddMenuItem(ByVal MenuName As String, ByVal MenuID As String, PopUp As Boolean)
 
 'get our menu handle
 Dim hMenu As Long
 hMenu = GetMenu(Me.hwnd)
 
 'if its a popup
 If PopUp Then
 
    'create a popup for it to use if needed
    Dim hSubMenu As Long
    hSubMenu = CreatePopupMenu()
 
    'now add new item with its popup
    AppendMenu hMenu, MF_STRING Or MF_POPUP, hSubMenu, MenuName
 
 Else
    
    'add new item
    AppendMenu hMenu, MF_STRING, MenuID, MenuName
 
 End If
 
 'redraw the updated menu
 DrawMenuBar Me.hwnd

End Sub

'___________________________________________________________________________
'
' Sub - Add SubMenu to a Menu Item
'___________________________________________________________________________

Private Sub AddSubMenuItem(ByVal SubMenuName As String, ByVal SubMenuID As String, MenuIndex As Integer)

 'get our menu handle
 Dim hMenu As Long
 hMenu = GetMenu(Me.hwnd)

 'get its submenu handle
 Dim hSubMenu As Long
 hSubMenu = GetSubMenu(hMenu, MenuIndex)

 'insert Separator
 If SubMenuName = "-" Then
    
    AppendMenu hSubMenu, MF_SEPARATOR, 0, vbNullString
 
 'insert another Popup
 ElseIf SubMenuID = 0 Then
    
    Dim hPopup As Long
    hPopup = CreatePopupMenu()
    
    AppendMenu hSubMenu, MF_STRING Or MF_POPUP, hPopup, SubMenuName
 
 'insert regular submenu item
 Else
    
    AppendMenu hSubMenu, MF_STRING Or MF_POPUP, SubMenuID, SubMenuName
 
 End If

 'redraw updated menu
 DrawMenuBar Me.hwnd

End Sub

'___________________________________________________________________________
'
' Sub - Remove all created MenuItems
'___________________________________________________________________________

Private Sub ClearMenu()

Dim hMenu As Long
hMenu = GetMenu(Me.hwnd)

Dim lCount As Long
lCount = GetMenuItemCount(hMenu)

Dim i As Integer
For i = 1 To lCount
    Call DeleteMenu(hMenu, 0, MF_BYPOSITION)
Next

'redraw updated menu
 DrawMenuBar Me.hwnd

End Sub

'___________________________________________________________________________
'
' Sub - Load all current windows into List
'___________________________________________________________________________

Private Sub ListWindows()

 'clear list
 lstWindows.Clear

 'ask for all open window handles
 Call EnumWindows(AddressOf EnumWindowsProc, 0&)
 
End Sub





