VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   Caption         =   "0"
   ClientHeight    =   8040
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8955
   DrawWidth       =   5
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   536
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   597
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text2 
      Height          =   255
      Left            =   840
      TabIndex        =   4
      Text            =   "0"
      Top             =   7680
      Width           =   2715
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   315
      Left            =   5280
      TabIndex        =   3
      Top             =   7680
      Width           =   1635
   End
   Begin VB.PictureBox picBack 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   3825
      Left            =   120
      Picture         =   "Form1.frx":000C
      ScaleHeight     =   3825
      ScaleWidth      =   4500
      TabIndex        =   0
      Top             =   3660
      Visible         =   0   'False
      Width           =   4500
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "ËÎÌå"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   7575
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   1
      Text            =   "Form1.frx":61CF
      Top             =   0
      Width           =   8955
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Exit"
      Height          =   360
      Left            =   7680
      TabIndex        =   2
      Top             =   7620
      Width           =   1200
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Command2_Click()
BitBlt Form1.hdc, 0, 0, TxtWid, TxtHei, hMemDC, 0, 0, vbSrcCopy
'Text1.Text = ""
End Sub

Private Sub Form_Load()
Me.ScaleMode = vbPixels
PrepPic picBack, Text1

Call subclass(Text1.hwnd)
Call subclass(Me.hwnd)
End Sub

Private Sub Form_Resize()
'Text1.Move 0, 0, Me.ScaleWidth, Me.ScaleHeight
End Sub

Private Sub Form_Unload(Cancel As Integer)
DeleteDC hMemDC
DeleteDC hBuffDC
'DeleteDC hBackDC

DeleteObject hBitmap
DeleteObject hBuffBitMap
'DeleteObject hBackBitMap

Call EndSubclass(Text1.hwnd)
Call EndSubclass(Me.hwnd)
End Sub

Private Sub subclass(lWnd As Long)
Dim ret As Long
 
  '¼ÇÂ¼Window ProcedureµÄµØÖ·
  preWinProc = GetWindowLong(lWnd, GWL_WNDPROC)
  '¿ªÊ¼½ØÈ¡ÏûÏ¢,²¢½«ÏûÏ¢½»¸øwndproc¹ý³Ì´¦Àí.
  ret = SetWindowLong(lWnd, GWL_WNDPROC, AddressOf wndproc)
End Sub
Private Sub EndSubclass(lWnd As Long)
Dim ret As Long
  'È¡ÏûÏûÏ¢½ØÈ¡£¬½áÊø×Ó·ÖÀà¹ý³Ì.
  ret = SetWindowLong(lWnd, GWL_WNDPROC, preWinProc)
End Sub
