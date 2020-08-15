Attribute VB_Name = "Module1"
Declare Sub GetKeyboardState Lib "user32" (lpKeyState As Any)
Declare Sub SetKeyboardState Lib "user32" (lpKeyState As Any)
Public Const VK_CAPITAL = &H14
Public Const VK_NUMLOCK = &H90

