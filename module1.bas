Attribute VB_Name = "module1"
'**************************************
'**************************************
 'here is a tip
'-----
'You can add two bitmaps to a menu entry
'One for the checked and one for the unchecked
'state.
'**************************************
'**************************************



'change anything you want it's all free
'greetz would be cool!
'any guestions or comments email me jimidotcom@yahoo.com
Declare Function GetMenu Lib "user32" _
(ByVal hwnd As Long) As Long

Declare Function GetSubMenu Lib "user32" _
(ByVal hMenu As Long, ByVal nPos As Long) As Long

Declare Function GetMenuItemID Lib "user32" _
(ByVal hMenu As Long, ByVal nPos As Long) As Long

Declare Function SetMenuItemBitmaps Lib "user32" _
(ByVal hMenu As Long, ByVal nPosition As Long, _
ByVal wFlags As Long, ByVal hBitmapUnchecked As Long, _
ByVal hBitmapChecked As Long) As Long

Declare Function GetMenuItemCount Lib "user32" _
(ByVal hMenu As Long) As Long

Declare Function GetMenuItemInfo Lib "user32" _
Alias "GetMenuItemInfoA" (ByVal hMenu As Long, _
ByVal un As Long, ByVal b As Boolean, _
lpMenuItemInfo As MENUITEMINFO) As Boolean


Public Const MF_BITMAP = &H4&

Type MENUITEMINFO
    cbSize As Long
    fMask As Long
    fType As Long
    fState As Long
    wID As Long
    hSubMenu As Long
    hbmpChecked As Long
    hbmpUnchecked As Long
    dwItemData As Long
    dwTypeData As String
    cch As Long
End Type '


Public Const MIIM_ID = &H2
Public Const MIIM_TYPE = &H10
Public Const MFT_STRING = &H0&

Public Sub pictopop1()

'Get the menuhandle of your app
'MENU 1 is what i named the main menu.
'the nameing is not important yet
hMenu& = GetMenu(form1.hwnd)

'Get the handle of the first submenu MENU 2
hSubMenu& = GetSubMenu(hMenu&, 0)

' change the number id here,
'this is where you edit to move the
'image from one menu to another
'(hSubMenu&, 0) for first item in submenu,
'(hSubMenu&, 1) for second item in submenu etc....
hID& = GetMenuItemID(hSubMenu&, 0)

'Add the bitmap
'so far it's just bitmaps but i know you can do icons i just haven't
'figured that out yet
SetMenuItemBitmaps hMenu&, hID&, MF_BITMAP, _
form1.Picture1.Picture, _
form1.Picture1.Picture
'The stuff below is just for the demo
form1.Text1 = "0"
form1.Text2 = "mnufile1"
form1.Text3 = "MENU2"
form1.Text1.SetFocus
End Sub
Public Sub pictopop2()

hMenu& = GetMenu(form1.hwnd)


hSubMenu& = GetSubMenu(hMenu&, 0)


hID& = GetMenuItemID(hSubMenu&, 1)


SetMenuItemBitmaps hMenu&, hID&, MF_BITMAP, _
form1.Picture2.Picture, _
form1.Picture2.Picture
'The stuff below is just for the demo
form1.Text1 = "1"
form1.Text2 = "mnufile2"
form1.Text3 = "MENU3"
form1.Text1.SetFocus
End Sub
Public Sub pictopop3()


hMenu& = GetMenu(form1.hwnd)


hSubMenu& = GetSubMenu(hMenu&, 0)


hID& = GetMenuItemID(hSubMenu&, 2)


SetMenuItemBitmaps hMenu&, hID&, MF_BITMAP, _
form1.Picture3.Picture, _
form1.Picture3.Picture
'The stuff below is just for the demo
form1.Text1 = "2"
form1.Text2 = "mnufile3"
form1.Text3 = "MENU4"
form1.Text1.SetFocus
End Sub
Public Sub pictopop4()


hMenu& = GetMenu(form1.hwnd)

hSubMenu& = GetSubMenu(hMenu&, 0)


hID& = GetMenuItemID(hSubMenu&, 3)


SetMenuItemBitmaps hMenu&, hID&, MF_BITMAP, _
form1.Picture4.Picture, _
form1.Picture4.Picture
'The stuff below is just for the demo
form1.Text1 = "3"
form1.Text2 = "mnufile4"
form1.Text3 = "MENU5"
form1.Text1.SetFocus
End Sub
