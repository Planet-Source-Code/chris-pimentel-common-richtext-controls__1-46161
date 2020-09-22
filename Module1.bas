Attribute VB_Name = "Module1"
Option Explicit

Private Type Rect
   Left As Long
   Top As Long
   Right As Long
   Bottom As Long
End Type

Private Type CharRange
  cpMin As Long
  cpMax As Long
End Type

Private Type FormatRange
  hdc As Long
  hdcTarget As Long
  rc As Rect
  rcPage As Rect
  chrg As CharRange
End Type

Private Const WM_USER As Long = &H400
Private Const EM_FORMATRANGE As Long = WM_USER + 57
Private Const EM_SETTARGETDEVICE As Long = WM_USER + 72
Private Const PHYSICALOFFSETX As Long = 112
Private Const PHYSICALOFFSETY As Long = 113

Private Declare Function GetDeviceCaps Lib "gdi32" ( _
   ByVal hdc As Long, ByVal nIndex As Long) As Long
Private Declare Function SendMessage Lib "USER32" Alias "SendMessageA" _
   (ByVal hWnd As Long, ByVal msg As Long, ByVal wp As Long, _
   lp As Any) As Long
Private Declare Function CreateDC Lib "gdi32" Alias "CreateDCA" _
   (ByVal lpDriverName As String, ByVal lpDeviceName As String, _
   ByVal lpOutput As Long, ByVal lpInitData As Long) As Long

Public Sub WYSIWYG_RTF(RTF As RichTextBox, LeftMarginWidth As Long, RightMarginWidth As Long, TopMarginWidth As Long, BottomMarginWidth As Long, PrintableWidth As Long, PrintableHeight As Long)
   Dim LeftOffset As Long
   Dim LeftMargin As Long
   Dim RightMargin As Long
   Dim TopOffset As Long
   Dim TopMargin As Long
   Dim BottomMargin As Long
   Dim PrinterhDC As Long
   Dim r As Long

  
   Printer.Print Space(1)
   Printer.ScaleMode = vbTwips
   
  
   LeftOffset = GetDeviceCaps(Printer.hdc, PHYSICALOFFSETX)
   LeftOffset = Printer.ScaleX(LeftOffset, vbPixels, vbTwips)
   
  
   LeftMargin = LeftMarginWidth - LeftOffset
   RightMargin = (Printer.Width - RightMarginWidth) - LeftOffset
   
   
   PrintableWidth = RightMargin - LeftMargin
   
   
   TopOffset = GetDeviceCaps(Printer.hdc, PHYSICALOFFSETY)
   TopOffset = Printer.ScaleX(TopOffset, vbPixels, vbTwips)
   
   
   TopMargin = TopMarginWidth - TopOffset
   BottomMargin = (Printer.Height - BottomMarginWidth) - TopOffset
   

   PrintableHeight = BottomMargin - TopMargin
    
   

   PrinterhDC = CreateDC(Printer.DriverName, Printer.DeviceName, 0, 0)


   r = SendMessage(RTF.hWnd, EM_SETTARGETDEVICE, PrinterhDC, _
      ByVal PrintableWidth)

   
   Printer.KillDoc
End Sub

Public Sub PrintRTF(RTF As RichTextBox, LeftMarginWidth As Long, _
   TopMarginHeight, RightMarginWidth, BottomMarginHeight)
   Dim LeftOffset As Long, TopOffset As Long
   Dim LeftMargin As Long, TopMargin As Long
   Dim RightMargin As Long, BottomMargin As Long
   Dim fr As FormatRange
   Dim rcDrawTo As Rect
   Dim rcPage As Rect
   Dim TextLength As Long
   Dim NextCharPosition As Long
   Dim r As Long

   
   Printer.Print Space(1)
   Printer.ScaleMode = vbTwips

   
   LeftOffset = Printer.ScaleX(GetDeviceCaps(Printer.hdc, _
      PHYSICALOFFSETX), vbPixels, vbTwips)
   TopOffset = Printer.ScaleY(GetDeviceCaps(Printer.hdc, _
      PHYSICALOFFSETY), vbPixels, vbTwips)

   
   LeftMargin = LeftMarginWidth - LeftOffset
   TopMargin = TopMarginHeight - TopOffset
   RightMargin = (Printer.Width - RightMarginWidth) - LeftOffset
   BottomMargin = (Printer.Height - BottomMarginHeight) - TopOffset

  
   rcPage.Left = 0
   rcPage.Top = 0
   rcPage.Right = Printer.ScaleWidth
   rcPage.Bottom = Printer.ScaleHeight

   
   rcDrawTo.Left = LeftMargin
   rcDrawTo.Top = TopMargin
   rcDrawTo.Right = RightMargin
   rcDrawTo.Bottom = BottomMargin

   
   fr.hdc = Printer.hdc
   fr.hdcTarget = Printer.hdc
   fr.rc = rcDrawTo
   fr.rcPage = rcPage
   fr.chrg.cpMin = 0
   fr.chrg.cpMax = -1

   
   TextLength = Len(RTF.Text)

   
   Do
      
      NextCharPosition = SendMessage(RTF.hWnd, EM_FORMATRANGE, True, fr)
      If NextCharPosition >= TextLength Then Exit Do
      fr.chrg.cpMin = NextCharPosition
      Printer.NewPage
      Printer.Print Space(1)
      fr.hdc = Printer.hdc
      fr.hdcTarget = Printer.hdc
   Loop


   Printer.EndDoc

  
   r = SendMessage(RTF.hWnd, EM_FORMATRANGE, False, ByVal CLng(0))
End Sub




