Attribute VB_Name = "VBKeyboard"
Public Declare Sub InstallHook Lib "VBKeyboardHook.Dll" (ByVal hWnd As Long)
Public Declare Sub RemoveHook Lib "VBKeyboardHook.Dll" ()
