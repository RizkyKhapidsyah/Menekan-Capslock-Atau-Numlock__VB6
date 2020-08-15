VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Menekan Capslock atau Numlock"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5430
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   5430
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "NumLock"
      Height          =   375
      Left            =   2760
      TabIndex        =   1
      Top             =   2160
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "CapsLock"
      Height          =   375
      Left            =   1080
      TabIndex        =   0
      Top             =   2160
      Width           =   1455
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()      'Menekan tombol
                                  'CapsLock.
ReDim KeyboardBuffer(256) As Byte
  GetKeyboardState KeyboardBuffer(0)
  'Contoh ini menunjukkan bagaimana Anda dapat menekan
  'tombol CapsLock.
  If KeyboardBuffer(VK_CAPITAL) And 1 Then
     KeyboardBuffer(VK_CAPITAL) = 0
  Else
     KeyboardBuffer(VK_CAPITAL) = 1
  End If
  SetKeyboardState KeyboardBuffer(0)
End Sub

Private Sub Command2_Click()
'Menekan tombol NumLock.
ReDim KeyboardBuffer(256) As Byte
  GetKeyboardState KeyboardBuffer(0)
  'Contoh ini menunjukkan bagaimana Anda dapat menekan
  'tombol NumLock.
  If KeyboardBuffer(VK_NUMLOCK) And 1 Then
     KeyboardBuffer(VK_NUMLOCK) = 0
  Else
     KeyboardBuffer(VK_NUMLOCK) = 1
  End If
  SetKeyboardState KeyboardBuffer(0)
End Sub

