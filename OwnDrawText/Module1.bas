Attribute VB_Name = "Module1"
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Const GWL_WNDPROC = (-4)
Public Const WM_PAINT = &HF
Public Const WM_LBUTTONDOWN = &H201
Public Const WM_COMMAND = &H111
Public Const EN_HSCROLL = &H601
Public Const EN_VSCROLL = &H602
Public Const WM_CTLCOLOREDIT = &H133
Public Const WM_KEYDOWN = &H100

Public Declare Function CreateCaret Lib "user32" (ByVal hwnd As Long, ByVal hBitmap As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Public Declare Function DestroyCaret Lib "user32" () As Long
Public Declare Function ShowCaret Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function HideCaret Lib "user32" (ByVal hwnd As Long) As Long

Public Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function SetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal crlColoror As Long) As Long
Public Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
Public Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Public Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Public Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Public Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Public Declare Function Rectangle Lib "gdi32" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Public Declare Function CreatePatternBrush Lib "gdi32" (ByVal hBitmap As Long) As Long
Public Declare Function CreateSolidBrush Lib "gdi32" (ByVal crlColoror As Long) As Long
Public Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crlColoror As Long) As Long
Public Const PS_NULL = 5

Public Declare Function SetBkMode Lib "gdi32" (ByVal hdc As Long, ByVal nBkMode As Long) As Long
Public Const TRANSPARENT = 1

Public hMemDC As Long, hTxtDC As Long
Public hBuffDC As Long ', hBackDC As Long
Public hBitmap As Long, hBuffBitMap As Long ', hBackBitMap As Long
Public TxtWid As Long, TxtHei As Long

'全局变量,存放控件标志性数据
Public preWinProc As Long
  
'本函数就是用来接收子分类时截取的消息的
Public Function wndproc(ByVal hwnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Dim wParamHi As Long

  '截取下来的消息存放在msg参数中.
  Select Case msg
    Case WM_CTLCOLOREDIT
      If lParam = Form1.Text1.hwnd Then
        BitBlt hTxtDC, 0, 0, TxtWid, TxtHei, hMemDC, 0, 0, vbMergePaint
        wndproc = CallWindowProc(preWinProc, hwnd, msg, wParam, lParam)
        Form1.Caption = hTxtDC
      End If
        wndproc = CallWindowProc(preWinProc, hwnd, msg, wParam, lParam)

'    Case WM_COMMAND
'      If hwnd = Form1.hwnd Then
'        wParamHi = ((wParam And &H7FFF0000) \ &H10000) ' Or &H8000&
'        If (wParamHi = EN_HSCROLL) Or (wParamHi = EN_VSCROLL) Then
'          BitBlt hTxtDC, 0, 0, txtWid, txtHei, hMemDC, 0, 0, vbMergePaint
'           Form1.Caption = Hex(wParamHi)
'       End If
'      End If
'      wndproc = CallWindowProc(preWinProc, hwnd, msg, wParam, lParam)
    
    Case WM_PAINT, WM_LBUTTONDOWN, WM_KEYDOWN
      '检测到消息,这里就可以加入我们的处理代码
      '需要注意,如果这儿不加入任何代码,则相当于吃掉了这条消息.
      wndproc = CallWindowProc(preWinProc, hwnd, msg, wParam, lParam)
      If hwnd = Form1.Text1.hwnd Then DrawTextPic (msg)
    
    Case Else
    '如果我们不是我们需要处理的消息,则将之送回原来的程序.
      wndproc = CallWindowProc(preWinProc, hwnd, msg, wParam, lParam)
      'DrawTextPic (msg)
  End Select
End Function

Public Sub Main()
  Form1.Show
End Sub


Public Sub PrepPic(picBG As PictureBox, txtDest As TextBox)
Dim i As Long, j As Long
Dim lColor As Long
Dim hPBrush As Long, hPen As Long
Dim hTmpBitMap As Long
Dim hTmpDC As Long
Dim PicWid As Long, PicHei As Long
Dim CX As Long, CY As Long

TxtWid = txtDest.Width
TxtHei = txtDest.Height
PicWid = picBG.Width * 2
PicHei = picBG.Height * 2

hMemDC = CreateCompatibleDC(0)
hBitmap = CreateCompatibleBitmap(picBG.hdc, TxtWid, TxtHei)
SelectObject hMemDC, hBitmap

hBuffDC = CreateCompatibleDC(0)
hBuffBitMap = CreateCompatibleBitmap(picBG.hdc, TxtWid, TxtHei)
SelectObject hBuffDC, hBuffBitMap

'hBackDC = CreateCompatibleDC(0)
'hBackBitMap = CreateCompatibleBitmap(picBG.hdc, txtWid, txtHei)
'SelectObject hBackDC, hBackBitMap

'建立临时画面
hTmpDC = CreateCompatibleDC(0)
hTmpBitMap = CreateCompatibleBitmap(picBG.hdc, PicWid, PicHei)
SelectObject hTmpDC, hTmpBitMap
'在临时画面上画上白色
hPBrush = CreateSolidBrush(vbWhite)
SelectObject hTmpDC, hPBrush
hPen = CreatePen(PS_NULL, 0, 0)
SelectObject hTmpDC, hPen
Rectangle hTmpDC, 0, 0, TxtWid, TxtHei
DeleteObject hPen
DeleteObject hPBrush

'打散图像
For i = 0 To picBG.Width - 1
   For j = 0 To picBG.Height - 1
      lColor = GetPixel(picBG.hdc, i, j)
      SetPixel hTmpDC, i * 2, j * 2, lColor
   Next
Next

'平铺到 MemDC
hPBrush = CreatePatternBrush(hTmpBitMap)
r = SelectObject(hMemDC, hPBrush)
Rectangle hMemDC, 0, 0, TxtWid, TxtHei
DeleteObject hPBrush

'While TxtHei > CY + PicHei
'  While TxtWid > CX + PicWid
'    BitBlt hMemDC, CX, CY, PicWid, PicHei, hTmpDC, 0, 0, vbSrcCopy
'    CX = CX + PicWid
'  Wend
'  CX = 0
'  CY = CY + PicHei
'Wend

hTxtDC = GetDC(txtDest.hwnd)

DeleteDC hTmpDC
DeleteObject hTmpBitMap

End Sub

Public Sub DrawTextPic(msg)

BitBlt hBuffDC, 0, 0, TxtWid, TxtHei, hTxtDC, 0, 0, vbSrcCopy
BitBlt hBuffDC, 0, 0, TxtWid, TxtHei, hMemDC, 0, 0, vbSrcAnd
BitBlt hTxtDC, 0, 0, TxtWid, TxtHei, hBuffDC, 0, 0, vbSrcCopy

'Form1.Caption = Form1.Caption + 1
Debug.Print Hex(msg)
End Sub
