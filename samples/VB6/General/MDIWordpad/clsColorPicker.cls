VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsColorPicker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

  Implements ISubclassedWindow


  Private Const INVALIDCOLORCELLINDEX = -1
  Private Const MORECOLORSCELLINDEX = -2
  Private Const DEFAULTCOLORCELLINDEX = -3
  Private Const MAXNUMBEROFCOLORCELLS = 100

  Private Const CLR_DEFAULT = &HFF000000

  Private Type COLORTABLEENTRY
    rgbColor As Long
    clrName As String
  End Type

  Private Type MONITORINFO
    cbSize As Long
    rcMonitor As RECTANGLE
    rcWork As RECTANGLE
    dwFlags As Long
  End Type

  Private Type PAINTSTRUCT
    hDC As Long
    fErase As Long
    rcPaint As RECTANGLE
    fRestore As Long
    fIncUpdate As Long
    rgbReserved(0 To 31) As Byte
  End Type

  Private Type POINT
    x As Long
    y As Long
  End Type

  Private Type MSGDETAILS
    hWnd As Long
    message As Long
    wParam As Long
    lParam As Long
    time As Long
    pt As POINT
  End Type

  Private Type TCHOOSECOLOR
    lStructSize As Long
    hwndOwner As Long
    hInstance As Long
    rgbResult As Long
    lpCustColors As Long
    Flags As Long
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As Long
  End Type

  Private Type Size
    cx As Long
    cy As Long
  End Type

  Private Type WNDCLASSEX
    cbSize As Long
    style As Long
    lpfnWndProc As Long
    cbClsExtra As Long
    cbWndExtra As Long
    hInstance As Long
    hIcon As Long
    hCursor As Long
    hbrBackground As Long
    lpszMenuName As Long
    lpszClassName As Long
    hIconSm As Long
  End Type

  Private Type PICKERWINDOWPROPERTIES
    hoveredColorCell As Long
    selectedColorCell As Long
    hoveredColor As Long
    numberOfPredefinedColors As Long
    numberOfColumns As Long
    numberOfRows As Long
    hPalette As Long
    cancelled As Boolean
    moreColorsTextBoundingRectangle As RECTANGLE
    defaultColorTextBoundingRectangle As RECTANGLE
    colorGridBoundingRectangle As RECTANGLE
    colorCellSize As Size
    BackColor As Long
    highlightBorderColor As Long
    hoverBackColor As Long
    hoverTextColor As Long
    selectionTextColor As Long
    selectionBackColor As Long
    has3DBorder As Boolean
  End Type


  Private m_predefinedColors() As COLORTABLEENTRY
  Private m_customColors(0 To 15) As Long
  Private m_innerTextCellBorderSize As Size
  Private m_outerTextCellBorderSize As Size
  Private m_innerColorCellBorderSize As Size
  Private m_outerColorCellBorderSize As Size
  Private m_colorCellContentSize As Size

  Private m_bIgnoreFirstLeftMouseButtonUp As Boolean
  Private m_bTrackSelection As Boolean
  Private m_defaultColor As Long
  Private m_hWndDropDown As Long
  Private m_hWndOwner As Long
  Private m_hWndParent As Long
  Private m_pickerProperties As PICKERWINDOWPROPERTIES
  Private m_selectedColor As Long
  Private m_toolTipControl As clsToolTip
  Private m_defaultColorText As String
  Private m_moreColorsText As String


  Private Declare Function BeginPaint Lib "user32.dll" (ByVal hWnd As Long, lpPaint As PAINTSTRUCT) As Long
  Private Declare Function ChooseColor Lib "comdlg32.dll" Alias "ChooseColorW" (lpcc As TCHOOSECOLOR) As Long
  Private Declare Function CreatePalette Lib "gdi32.dll" (ByVal lplgpl As Long) As Long
  Private Declare Function CreatePen Lib "gdi32.dll" (ByVal fnPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
  Private Declare Function CreateWindowEx Lib "user32.dll" Alias "CreateWindowExW" (ByVal dwExStyle As Long, ByVal lpClassName As Long, ByVal lpWindowName As Long, ByVal dwStyle As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, ByVal lpParam As Long) As Long
  Private Declare Function DeleteObject Lib "gdi32.dll" (ByVal hObject As Long) As Long
  Private Declare Function DestroyWindow Lib "user32.dll" (ByVal hWnd As Long) As Long
  Private Declare Function DispatchMessage Lib "user32.dll" Alias "DispatchMessageW" (lpMsg As MSGDETAILS) As Long
  Private Declare Function DrawEdge Lib "user32.dll" (ByVal hDC As Long, qrc As RECTANGLE, ByVal edge As Long, ByVal grfFlags As Long) As Long
  Private Declare Function DrawRectangle Lib "gdi32.dll" Alias "Rectangle" (ByVal hDC As Long, ByVal nLeftRect As Long, ByVal nTopRect As Long, ByVal nRightRect As Long, ByVal nBottomRect As Long) As Long
  Private Declare Function DrawText Lib "user32.dll" Alias "DrawTextW" (ByVal hDC As Long, ByVal lpchText As Long, ByVal nCount As Long, lpRect As RECTANGLE, ByVal uFormat As Long) As Long
  Private Declare Function EndPaint Lib "user32.dll" (ByVal hWnd As Long, lpPaint As PAINTSTRUCT) As Long
  Private Declare Function ExtTextOut Lib "gdi32.dll" Alias "ExtTextOutW" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal fuOptions As Long, lprc As RECTANGLE, ByVal lpString As Long, ByVal cbCount As Long, ByVal lpDx As Long) As Long
  Private Declare Function GetCapture Lib "user32.dll" () As Long
  Private Declare Function GetClientRect Lib "user32.dll" (ByVal hWnd As Long, lpRect As RECTANGLE) As Long
  Private Declare Function GetDC Lib "user32.dll" (ByVal hWnd As Long) As Long
  Private Declare Function GetDeviceCaps Lib "gdi32.dll" (ByVal hDC As Long, ByVal nIndex As Long) As Long
  Private Declare Function GetMessage Lib "user32.dll" Alias "GetMessageW" (lpMsg As MSGDETAILS, ByVal hWnd As Long, ByVal wMsgFilterMin As Long, ByVal wMsgFilterMax As Long) As Long
  Private Declare Function GetModuleHandle Lib "kernel32.dll" Alias "GetModuleHandleW" (ByVal lpModuleName As Long) As Long
  Private Declare Function GetMonitorInfo Lib "user32.dll" Alias "GetMonitorInfoW" (ByVal hMonitor As Long, lpmi As MONITORINFO) As Long
  Private Declare Function GetProcAddress Lib "kernel32.dll" (ByVal hModule As Long, ByVal lpProcName As String) As Long
  Private Declare Function GetProcessHeap Lib "kernel32.dll" () As Long
  Private Declare Function GetSysColor Lib "user32.dll" (ByVal nIndex As Long) As Long
  Private Declare Function GetSystemMetrics Lib "user32.dll" (ByVal nIndex As Long) As Long
  Private Declare Function GetTextExtentPoint32 Lib "gdi32.dll" Alias "GetTextExtentPoint32W" (ByVal hDC As Long, ByVal lpString As Long, ByVal c As Long, lpSize As Size) As Long
  Private Declare Function GetWindowRect Lib "user32.dll" (ByVal hWnd As Long, lpRect As RECTANGLE) As Long
  Private Declare Function HeapAlloc Lib "kernel32.dll" (ByVal hHeap As Long, ByVal dwFlags As Long, ByVal dwBytes As Long) As Long
  Private Declare Function HeapFree Lib "kernel32.dll" (ByVal hHeap As Long, ByVal dwFlags As Long, ByVal lpMem As Long) As Long
  Private Declare Function InflateRect Lib "user32.dll" (lprc As RECTANGLE, ByVal dx As Long, ByVal dy As Long) As Long
  Private Declare Function InvalidateRect Lib "user32.dll" (ByVal hWnd As Long, ByVal lpRect As Long, ByVal bErase As Long) As Long
  Private Declare Function LoadImage Lib "user32.dll" Alias "LoadImageW" (ByVal hinst As Long, ByVal lpszName As Long, ByVal uType As Long, ByVal cxDesired As Long, ByVal cyDesired As Long, ByVal fuLoad As Long) As Long
  Private Declare Function MonitorFromWindow Lib "user32.dll" (ByVal hWnd As Long, ByVal dwFlags As Long) As Long
  Private Declare Function MoveWindow Lib "user32.dll" (ByVal hWnd As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
  Private Declare Function OffsetRect Lib "user32.dll" (lprc As RECTANGLE, ByVal dx As Long, ByVal dy As Long) As Long
  Private Declare Function PeekMessage Lib "user32.dll" Alias "PeekMessageW" (lpMsg As MSGDETAILS, ByVal hWnd As Long, ByVal wMsgFilterMin As Long, ByVal wMsgFilterMax As Long, ByVal wRemoveMsg As Long) As Long
  Private Declare Sub PostQuitMessage Lib "user32.dll" (ByVal nExitCode As Long)
  Private Declare Function PtInRect Lib "user32.dll" (lprc As RECTANGLE, ByVal ptx As Long, ByVal pty As Long) As Long
  Private Declare Function RealizePalette Lib "gdi32.dll" (ByVal hDC As Long) As Long
  Private Declare Function RegisterClassEx Lib "user32.dll" Alias "RegisterClassExW" (lpwcx As WNDCLASSEX) As Long
  Private Declare Function ReleaseCapture Lib "user32.dll" () As Long
  Private Declare Function ReleaseDC Lib "user32.dll" (ByVal hWnd As Long, ByVal hDC As Long) As Long
  Private Declare Function SelectObject Lib "gdi32.dll" (ByVal hDC As Long, ByVal hObject As Long) As Long
  Private Declare Function SelectPalette Lib "gdi32.dll" (ByVal hDC As Long, ByVal hPal As Long, ByVal bForceBackground As Long) As Long
  Private Declare Function SendMessageAsLong Lib "user32.dll" Alias "SendMessageW" (ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
  Private Declare Function SetBkColor Lib "gdi32.dll" (ByVal hDC As Long, ByVal crColor As Long) As Long
  Private Declare Function SetBkMode Lib "gdi32.dll" (ByVal hDC As Long, ByVal iBkMode As Long) As Long
  Private Declare Function SetCapture Lib "user32.dll" (ByVal hWnd As Long) As Long
  Private Declare Function SetTextColor Lib "gdi32.dll" (ByVal hDC As Long, ByVal crColor As Long) As Long
  Private Declare Function ShowWindow Lib "user32.dll" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
  Private Declare Function SystemParametersInfo Lib "user32.dll" Alias "SystemParametersInfoW" (ByVal uiAction As Long, ByVal uiParam As Long, ByVal pvParam As Long, ByVal fWinIni As Long) As Long
  Private Declare Function UnregisterClass Lib "user32.dll" Alias "UnregisterClassW" (ByVal lpClassName As Long, ByVal hInstance As Long) As Long


  Public Event HotTrackingColor(ByVal hotColor As Long, ByVal colorIsValid As Boolean)


Private Function ISubclassedWindow_HandleMessage(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long, ByVal eSubclassID As EnumSubclassID, bCallDefProc As Boolean) As Long
  Dim lRet As Long

  On Error GoTo StdHandler_Error
  Select Case eSubclassID
    Case EnumSubclassID.escidColorPicker
      lRet = HandleMessage_DropDown(hWnd, uMsg, wParam, lParam, bCallDefProc)
    Case Else
      Debug.Print "clsColorPicker.ISubclassedWindow_HandleMessage: Unknown Subclassing ID " & CStr(eSubclassID)
  End Select

StdHandler_Ende:
  ISubclassedWindow_HandleMessage = lRet
  Exit Function

StdHandler_Error:
  Debug.Print "Error in clsColorPicker.ISubclassedWindow_HandleMessage (SubclassID=" & CStr(eSubclassID) & ": ", Err.number, Err.Description
  Resume StdHandler_Ende
End Function

Private Function HandleMessage_DropDown(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long, bCallDefProc As Boolean) As Long
  Const WM_PAINT = &HF
  Const WM_PALETTECHANGED = &H311
  Const WM_QUERYNEWPALETTE = &H30F
  Dim lRet As Long

  On Error GoTo StdHandler_Error
  Select Case uMsg
    Case WM_PAINT
      lRet = OnPickerPaint(uMsg, wParam, lParam)
      bCallDefProc = False
    Case WM_PALETTECHANGED
      lRet = DefSubclassProc(hWnd, uMsg, wParam, lParam)
      OnPickerPaletteChanged uMsg, wParam, lParam
      bCallDefProc = False
    Case WM_QUERYNEWPALETTE
      lRet = OnPickerQueryNewPalette(uMsg, wParam, lParam)
      bCallDefProc = False
  End Select

StdHandler_Ende:
  HandleMessage_DropDown = lRet
  Exit Function

StdHandler_Error:
  Debug.Print "Error in clsColorPicker.HandleMessage_DropDown: ", Err.number, Err.Description
  Resume StdHandler_Ende
End Function


Private Sub Class_Initialize()
  Const ALPHACHANNEL = 48
  Const COLOR_MENU = 4
  Const COLOR_MENUTEXT = 7
  Const COLOR_HIGHLIGHT = 13
  Const COLOR_HIGHLIGHTTEXT = 14
  Const COLOR_BTNFACE = 15
  Const COLOR_MENUHILIGHT = 29
  Const SPI_GETFLATMENU = &H1022
  Dim i As Long
  Dim clr As Long

  m_defaultColorText = "Automatic"
  m_moreColorsText = "More Colors..."

  ReDim m_predefinedColors(0 To 39) As COLORTABLEENTRY
  m_predefinedColors(0).rgbColor = RGB(&H0, &H0, &H0)
  m_predefinedColors(0).clrName = "Black"
  m_predefinedColors(1).rgbColor = RGB(&HA5, &H2A, &H0)
  m_predefinedColors(1).clrName = "Brown"
  m_predefinedColors(2).rgbColor = RGB(&H0, &H40, &H40)
  m_predefinedColors(2).clrName = "Dark Olive Green"
  m_predefinedColors(3).rgbColor = RGB(&H0, &H55, &H0)
  m_predefinedColors(3).clrName = "Dark Green"
  m_predefinedColors(4).rgbColor = RGB(&H0, &H0, &H5E)
  m_predefinedColors(4).clrName = "Dark Teal"
  m_predefinedColors(5).rgbColor = RGB(&H0, &H0, &H8B)
  m_predefinedColors(5).clrName = "Dark Blue"
  m_predefinedColors(6).rgbColor = RGB(&H4B, &H0, &H82)
  m_predefinedColors(6).clrName = "Indigo"
  m_predefinedColors(7).rgbColor = RGB(&H28, &H28, &H28)
  m_predefinedColors(7).clrName = "Dark Grey"

  m_predefinedColors(8).rgbColor = RGB(&H8B, &H0, &H0)
  m_predefinedColors(8).clrName = "Dark Red"
  m_predefinedColors(9).rgbColor = RGB(&HFF, &H68, &H20)
  m_predefinedColors(9).clrName = "Orange"
  m_predefinedColors(10).rgbColor = RGB(&H8B, &H8B, &H0)
  m_predefinedColors(10).clrName = "Dark Yellow"
  m_predefinedColors(11).rgbColor = RGB(&H0, &H93, &H0)
  m_predefinedColors(11).clrName = "Green"
  m_predefinedColors(12).rgbColor = RGB(&H38, &H8E, &H8E)
  m_predefinedColors(12).clrName = "Teal"
  m_predefinedColors(13).rgbColor = RGB(&H0, &H0, &HFF)
  m_predefinedColors(13).clrName = "Blue"
  m_predefinedColors(14).rgbColor = RGB(&H7B, &H7B, &HC0)
  m_predefinedColors(14).clrName = "Blue-Grey"
  m_predefinedColors(15).rgbColor = RGB(&H66, &H66, &H66)
  m_predefinedColors(15).clrName = "Grey - 40"

  m_predefinedColors(16).rgbColor = RGB(&HFF, &H0, &H0)
  m_predefinedColors(16).clrName = "Red"
  m_predefinedColors(17).rgbColor = RGB(&HFF, &HAD, &H5B)
  m_predefinedColors(17).clrName = "Light Orange"
  m_predefinedColors(18).rgbColor = RGB(&H32, &HCD, &H32)
  m_predefinedColors(18).clrName = "Lime"
  m_predefinedColors(19).rgbColor = RGB(&H3C, &HB3, &H71)
  m_predefinedColors(19).clrName = "Sea Green"
  m_predefinedColors(20).rgbColor = RGB(&H7F, &HFF, &HD4)
  m_predefinedColors(20).clrName = "Aqua"
  m_predefinedColors(21).rgbColor = RGB(&H7D, &H9E, &HC0)
  m_predefinedColors(21).clrName = "Light Blue"
  m_predefinedColors(22).rgbColor = RGB(&H80, &H0, &H80)
  m_predefinedColors(22).clrName = "Violet"
  m_predefinedColors(23).rgbColor = RGB(&H7F, &H7F, &H7F)
  m_predefinedColors(23).clrName = "Grey - 50"

  m_predefinedColors(24).rgbColor = RGB(&HFF, &HC0, &HCB)
  m_predefinedColors(24).clrName = "Pink"
  m_predefinedColors(25).rgbColor = RGB(&HFF, &HD7, &H0)
  m_predefinedColors(25).clrName = "Gold"
  m_predefinedColors(26).rgbColor = RGB(&HFF, &HFF, &H0)
  m_predefinedColors(26).clrName = "Yellow"
  m_predefinedColors(27).rgbColor = RGB(&H0, &HFF, &H0)
  m_predefinedColors(27).clrName = "Bright Green"
  m_predefinedColors(28).rgbColor = RGB(&H40, &HE0, &HD0)
  m_predefinedColors(28).clrName = "Torquoise"
  m_predefinedColors(29).rgbColor = RGB(&HC0, &HFF, &HFF)
  m_predefinedColors(29).clrName = "Skyblue"
  m_predefinedColors(30).rgbColor = RGB(&H48, &H0, &H48)
  m_predefinedColors(30).clrName = "Plum"
  m_predefinedColors(31).rgbColor = RGB(&HC0, &HC0, &HC0)
  m_predefinedColors(31).clrName = "Light Grey"

  m_predefinedColors(32).rgbColor = RGB(&HFF, &HE4, &HE1)
  m_predefinedColors(32).clrName = "Rose"
  m_predefinedColors(33).rgbColor = RGB(&HD2, &HB4, &H8C)
  m_predefinedColors(33).clrName = "Tan"
  m_predefinedColors(34).rgbColor = RGB(&HFF, &HFF, &HE9)
  m_predefinedColors(34).clrName = "Light Yellow"
  m_predefinedColors(35).rgbColor = RGB(&H98, &HFB, &H98)
  m_predefinedColors(35).clrName = "Pale Green"
  m_predefinedColors(36).rgbColor = RGB(&HAF, &HEE, &HEE)
  m_predefinedColors(36).clrName = "Pale Torquoise"
  m_predefinedColors(37).rgbColor = RGB(&H68, &H83, &H8B)
  m_predefinedColors(37).clrName = "Pale Blue"
  m_predefinedColors(38).rgbColor = RGB(&HE6, &HE6, &HFA)
  m_predefinedColors(38).clrName = "Lavender"
  m_predefinedColors(39).rgbColor = RGB(&HFF, &HFF, &HFF)
  m_predefinedColors(39).clrName = "White"

  For i = LBound(m_customColors) To UBound(m_customColors)
    m_customColors(i) = RGB(255, 255, 255)
  Next i

  m_innerTextCellBorderSize.cx = 3
  m_innerTextCellBorderSize.cy = 3
  m_outerTextCellBorderSize.cx = 2
  m_outerTextCellBorderSize.cy = 2
  m_innerColorCellBorderSize.cx = 2
  m_innerColorCellBorderSize.cy = 2
  m_outerColorCellBorderSize.cx = 0
  m_outerColorCellBorderSize.cy = 0
  m_colorCellContentSize.cx = 14
  m_colorCellContentSize.cy = 14

  m_defaultColor = CLR_DEFAULT
  m_selectedColor = CLR_DEFAULT

  clr = GetSysColor(COLOR_MENU)
  If (Abs(GetRed(clr) - GetRed(vbWhite)) <= 10) And (Abs(GetGreen(clr) - GetGreen(vbWhite)) <= 10) And (Abs(GetBlue(clr) - GetBlue(vbWhite)) <= 10) Then
    clr = GetSysColor(COLOR_BTNFACE)
  End If
  m_pickerProperties.BackColor = clr
  m_pickerProperties.highlightBorderColor = GetSysColor(COLOR_HIGHLIGHT)
  m_pickerProperties.hoverBackColor = GetSysColor(COLOR_MENUHILIGHT)
  SystemParametersInfo SPI_GETFLATMENU, 0, VarPtr(m_pickerProperties.has3DBorder), 0
  m_pickerProperties.hoverTextColor = GetSysColor(COLOR_HIGHLIGHTTEXT)
  m_pickerProperties.selectionTextColor = GetSysColor(COLOR_MENUTEXT)
  m_pickerProperties.selectionBackColor = RGB(GetRed(GetRed(&HF0F0F0) * (255 - ALPHACHANNEL) + GetRed(&HFF9933) * ALPHACHANNEL), GetRed(GetGreen(&HF0F0F0) * (255 - ALPHACHANNEL) + GetGreen(&HFF9933) * ALPHACHANNEL), GetRed(GetBlue(&HF0F0F0) * (255 - ALPHACHANNEL) + GetBlue(&HFF9933) * ALPHACHANNEL))
End Sub

Private Sub Class_Terminate()
  If m_pickerProperties.hPalette Then
    DeleteObject m_pickerProperties.hPalette
    m_pickerProperties.hPalette = 0
  End If
End Sub


Public Property Get DefaultColor() As Long
  DefaultColor = m_defaultColor
End Property

Public Property Let DefaultColor(ByVal newVal As Long)
  m_defaultColor = newVal
End Property

Public Property Get DefaultColorText() As String
  DefaultColorText = m_defaultColorText
End Property

Public Property Let DefaultColorText(ByVal newVal As String)
  m_defaultColorText = newVal
End Property

Public Property Get hWndDropDown() As Long
  hWndDropDown = m_hWndDropDown
End Property

Public Property Get hwndOwner() As Long
  hwndOwner = m_hWndOwner
End Property

Public Property Let hwndOwner(ByVal newVal As Long)
  m_hWndOwner = newVal
End Property

Public Property Get hWndParent() As Long
  hWndParent = m_hWndParent
End Property

Public Property Let hWndParent(ByVal newVal As Long)
  m_hWndParent = newVal
End Property

Public Property Get IgnoreFirstLeftMouseButtonUp() As Boolean
  IgnoreFirstLeftMouseButtonUp = m_bIgnoreFirstLeftMouseButtonUp
End Property

Public Property Let IgnoreFirstLeftMouseButtonUp(ByVal newVal As Boolean)
  m_bIgnoreFirstLeftMouseButtonUp = newVal
End Property

Public Property Get MoreColorsText() As String
  MoreColorsText = m_moreColorsText
End Property

Public Property Let MoreColorsText(ByVal newVal As String)
  m_moreColorsText = newVal
End Property

Public Property Get SelectedColor() As Long
  SelectedColor = m_selectedColor
End Property

Public Property Let SelectedColor(ByVal newVal As Long)
  m_selectedColor = newVal
End Property

Public Property Get TrackSelection() As Boolean
  TrackSelection = m_bTrackSelection
End Property

Public Property Let TrackSelection(ByVal newVal As Boolean)
  m_bTrackSelection = newVal
End Property

Public Sub SetPredefinedColors(colorsArray As Variant, namesArray As Variant)
  Dim i As Long

  Erase m_predefinedColors
  ReDim m_predefinedColors(0 To UBound(colorsArray) - LBound(colorsArray)) As COLORTABLEENTRY
  For i = LBound(colorsArray) To UBound(colorsArray)
    m_predefinedColors(i - LBound(colorsArray)).rgbColor = colorsArray(i)
    m_predefinedColors(i - LBound(colorsArray)).clrName = namesArray(i)
  Next i
End Sub


Private Sub ChangePickerSelection(ByVal cellIndex As Long)
  Dim Color As Long
  Dim colorIsValid As Boolean
  Dim hDCClient As Long
  Dim previousSelection As Long

  hDCClient = GetDC(m_hWndDropDown)

  If cellIndex > m_pickerProperties.numberOfPredefinedColors Then
    cellIndex = MORECOLORSCELLINDEX
  End If

  ' if the current selection is valid, redraw the old selection without it being selected
  If ((m_pickerProperties.hoveredColorCell >= 0) And (m_pickerProperties.hoveredColorCell < m_pickerProperties.numberOfPredefinedColors)) Or (m_pickerProperties.hoveredColorCell = MORECOLORSCELLINDEX) Or (m_pickerProperties.hoveredColorCell = DEFAULTCOLORCELLINDEX) Then
    previousSelection = m_pickerProperties.hoveredColorCell
    m_pickerProperties.hoveredColorCell = INVALIDCOLORCELLINDEX
    DrawPickerCell hDCClient, previousSelection
  End If

  ' set the current selection as row/col and draw (it will be drawn selected)
  m_pickerProperties.hoveredColorCell = cellIndex
  DrawPickerCell hDCClient, m_pickerProperties.hoveredColorCell

  ' store the current color
  colorIsValid = True
  Color = RGB(0, 0, 0)
  If m_pickerProperties.hoveredColorCell = MORECOLORSCELLINDEX Then
    Color = m_defaultColor
  ElseIf m_pickerProperties.hoveredColorCell = DEFAULTCOLORCELLINDEX Then
    m_pickerProperties.hoveredColor = CLR_DEFAULT
    Color = m_pickerProperties.hoveredColor
  ElseIf m_pickerProperties.hoveredColorCell = INVALIDCOLORCELLINDEX Then
    Color = RGB(0, 0, 0)
    colorIsValid = False
  ElseIf m_pickerProperties.hoveredColorCell >= 0 And m_pickerProperties.hoveredColorCell <= UBound(m_predefinedColors) Then
    m_pickerProperties.hoveredColor = m_predefinedColors(m_pickerProperties.hoveredColorCell).rgbColor
    Color = m_pickerProperties.hoveredColor
  End If

  ' send notification
  If m_bTrackSelection Then
    RaiseEvent HotTrackingColor(Color, colorIsValid)
  End If
  ReleaseDC m_hWndDropDown, hDCClient
End Sub

Private Sub DrawPickerCell(ByVal hTargetDC As Long, ByVal cellIndex As Long)
  Const COLOR_BTNSHADOW = 16
  Const COLOR_3DSHADOW = COLOR_BTNSHADOW
  Const DT_CENTER = &H1
  Const DT_VCENTER = &H4
  Const DT_SINGLELINE = &H20
  Const ETO_OPAQUE = &H2
  Const RASTERCAPS = 38
  Const RC_PALETTE = &H100
  Const TRANSPARENT = 1
  Const WM_GETFONT = &H31
  Dim backgroundColor As Long
  Dim cellRectangle As RECTANGLE
  Dim cellText As String
  Dim Color As Long
  Dim hPreviousFont As Long
  Dim hPreviousPalette As Long
  Dim innerBorderSize As Size
  Dim isSelected As Boolean
  Dim outerBorderSize As Size
  Dim previousBackColor As Long
  Dim previousBkMode As Long
  Dim previousTextColor As Long
  Dim textColor As Long

  ' get the drawing rect
  If Not GetPickerCellRectangle(cellIndex, cellRectangle) Then
    Exit Sub
  End If

  ' get the text and colors
  If cellIndex = MORECOLORSCELLINDEX Then
    cellText = m_moreColorsText
    outerBorderSize = m_outerTextCellBorderSize
    innerBorderSize = m_innerTextCellBorderSize
  ElseIf cellIndex = DEFAULTCOLORCELLINDEX Then
    cellText = m_defaultColorText
    outerBorderSize = m_outerTextCellBorderSize
    innerBorderSize = m_innerTextCellBorderSize
  Else
    Color = m_predefinedColors(cellIndex).rgbColor
    outerBorderSize = m_outerColorCellBorderSize
    innerBorderSize = m_innerColorCellBorderSize
  End If

  ' based on the selectons, get our colors
  If m_pickerProperties.hoveredColorCell = cellIndex Then
    isSelected = True
    backgroundColor = m_pickerProperties.hoverBackColor
    textColor = m_pickerProperties.hoverTextColor
  ElseIf m_pickerProperties.selectedColorCell = cellIndex Then
    isSelected = True
    backgroundColor = m_pickerProperties.selectionBackColor
    textColor = m_pickerProperties.selectionTextColor
  Else
    isSelected = False
    backgroundColor = m_pickerProperties.BackColor
    textColor = m_pickerProperties.selectionTextColor
  End If

  ' select and realize the palette
  If Len(cellText) = 0 Then
    If (m_pickerProperties.hPalette <> 0) And ((GetDeviceCaps(hTargetDC, RASTERCAPS) And RC_PALETTE) <> 0) Then
      hPreviousPalette = SelectPalette(hTargetDC, m_pickerProperties.hPalette, 0)
      RealizePalette hTargetDC
    End If
  End If

  previousBackColor = SetBkColor(hTargetDC, RGB(0, 0, 0))
  If isSelected Then
    ' draw the background
    If (outerBorderSize.cx > 0) Or (outerBorderSize.cy > 0) Then
      SetBkColor hTargetDC, m_pickerProperties.BackColor
      ExtTextOut hTargetDC, 0, 0, ETO_OPAQUE, cellRectangle, 0, 0, 0
      InflateRect cellRectangle, -outerBorderSize.cx, -outerBorderSize.cy
    End If

    ' draw the selection rectangle
    SetBkColor hTargetDC, m_pickerProperties.highlightBorderColor
    ExtTextOut hTargetDC, 0, 0, ETO_OPAQUE, cellRectangle, 0, 0, 0
    InflateRect cellRectangle, -1, -1

    ' draw the inner coloring
    SetBkColor hTargetDC, backgroundColor
    ExtTextOut hTargetDC, 0, 0, ETO_OPAQUE, cellRectangle, 0, 0, 0
    InflateRect cellRectangle, -(innerBorderSize.cx - 1), -(innerBorderSize.cy - 1)
  Else
    ' draw the background
    SetBkColor hTargetDC, m_pickerProperties.BackColor
    ExtTextOut hTargetDC, 0, 0, ETO_OPAQUE, cellRectangle, 0, 0, 0
    InflateRect cellRectangle, -(outerBorderSize.cx + innerBorderSize.cx), -(outerBorderSize.cy + innerBorderSize.cy)
  End If

  ' draw the text or color
  If Len(cellText) > 0 Then
    If m_hWndOwner Then
      hPreviousFont = SelectObject(hTargetDC, SendMessageAsLong(m_hWndOwner, WM_GETFONT, 0, 0))
    Else
      hPreviousFont = SelectObject(hTargetDC, SendMessageAsLong(m_hWndParent, WM_GETFONT, 0, 0))
    End If
    previousTextColor = SetTextColor(hTargetDC, textColor)
    previousBkMode = SetBkMode(hTargetDC, TRANSPARENT)
    DrawText hTargetDC, StrPtr(cellText), -1, cellRectangle, DT_CENTER Or DT_VCENTER Or DT_SINGLELINE
    SetBkMode hTargetDC, previousBkMode
    SetTextColor hTargetDC, previousTextColor
    SelectObject hTargetDC, hPreviousFont
  Else
    SetBkColor hTargetDC, GetSysColor(COLOR_3DSHADOW)
    ExtTextOut hTargetDC, 0, 0, ETO_OPAQUE, cellRectangle, 0, 0, 0
    InflateRect cellRectangle, -1, -1
    If cellIndex >= 0 And cellIndex <= UBound(m_predefinedColors) Then
      SetBkColor hTargetDC, m_predefinedColors(cellIndex).rgbColor
    End If
    ExtTextOut hTargetDC, 0, 0, ETO_OPAQUE, cellRectangle, 0, 0, 0
  End If
  SetBkColor hTargetDC, previousBackColor

  ' restore the palette
  If (hPreviousPalette <> 0) And ((GetDeviceCaps(hTargetDC, RASTERCAPS) And RC_PALETTE) <> 0) Then
    SelectPalette hTargetDC, hPreviousPalette, 0
  End If
End Sub

Private Sub EndPickerSelection(ByVal cancelled As Boolean)
  ReleaseCapture
  m_pickerProperties.cancelled = cancelled
End Sub

Private Function FindPickerCellFromColor(ByVal Color As Long) As Long
  Dim i As Long

  If (Color = CLR_DEFAULT) And HasDefaultColorText() Then
    FindPickerCellFromColor = DEFAULTCOLORCELLINDEX
    Exit Function
  End If

  For i = 0 To m_pickerProperties.numberOfPredefinedColors - 1
    If m_predefinedColors(i).rgbColor = Color Then
      FindPickerCellFromColor = i
      Exit Function
    End If
  Next i

  If HasMoreColorsText() Then
    FindPickerCellFromColor = MORECOLORSCELLINDEX
  Else
    FindPickerCellFromColor = INVALIDCOLORCELLINDEX
  End If
End Function

Private Function GetPickerCellRectangle(ByVal cellIndex As Long, pRectangle As RECTANGLE) As Boolean
  If cellIndex = MORECOLORSCELLINDEX Then
    pRectangle = m_pickerProperties.moreColorsTextBoundingRectangle
    GetPickerCellRectangle = True
  ElseIf cellIndex = DEFAULTCOLORCELLINDEX Then
    pRectangle = m_pickerProperties.defaultColorTextBoundingRectangle
    GetPickerCellRectangle = True
  ElseIf (cellIndex < 0) Or (cellIndex >= m_pickerProperties.numberOfPredefinedColors) Then
    GetPickerCellRectangle = False
  Else
    pRectangle.Left = (cellIndex Mod m_pickerProperties.numberOfColumns) * m_pickerProperties.colorCellSize.cx + m_pickerProperties.colorGridBoundingRectangle.Left
    pRectangle.Top = (cellIndex \ m_pickerProperties.numberOfColumns) * m_pickerProperties.colorCellSize.cy + m_pickerProperties.colorGridBoundingRectangle.Top
    pRectangle.Right = pRectangle.Left + m_pickerProperties.colorCellSize.cx
    pRectangle.Bottom = pRectangle.Top + m_pickerProperties.colorCellSize.cy
    GetPickerCellRectangle = True
  End If
End Function

Private Function GetBlue(ByVal Color As Long) As Long
  Dim ret As Byte

  CopyMemory VarPtr(ret), VarPtr(Color) + 2 * LenB(ret), LenB(ret)
  GetBlue = ret
End Function

Private Function GetGreen(ByVal Color As Long) As Long
  Dim ret As Byte

  CopyMemory VarPtr(ret), VarPtr(Color) + LenB(ret), LenB(ret)
  GetGreen = ret
End Function

Private Function GetRed(ByVal Color As Long) As Long
  Dim ret As Byte

  CopyMemory VarPtr(ret), VarPtr(Color), LenB(ret)
  GetRed = ret
End Function

Private Function HasMoreColorsText() As Boolean
  HasMoreColorsText = Len(m_moreColorsText) > 0
End Function

Private Function HasDefaultColorText() As Boolean
  HasDefaultColorText = Len(m_defaultColorText) > 0
End Function

Private Function HiWord(ByVal Value As Long) As Integer
  Dim ret As Integer

  CopyMemory VarPtr(ret), VarPtr(Value) + LenB(ret), LenB(ret)

  HiWord = ret
End Function

Private Function LoWord(ByVal Value As Long) As Integer
  Dim ret As Integer

  CopyMemory VarPtr(ret), VarPtr(Value), LenB(ret)

  LoWord = ret
End Function

Private Function OnPickerKeyDown(ByVal message As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
  Const VK_RETURN = &HD
  Const VK_ESCAPE = &H1B
  Const VK_SPACE = &H20
  Const VK_LEFT = &H25
  Const VK_UP = &H26
  Const VK_RIGHT = &H27
  Const VK_DOWN = &H28
  Dim newSelection As Long
  Dim offset As Long

  ' get the offset for movement
  Select Case wParam
    Case VK_DOWN
      offset = m_pickerProperties.numberOfColumns
    Case VK_UP
      offset = -m_pickerProperties.numberOfColumns
    Case VK_RIGHT
      offset = 1
    Case VK_LEFT
      offset = -1
    Case VK_ESCAPE
      m_pickerProperties.hoveredColor = m_selectedColor
      EndPickerSelection True
    Case VK_RETURN, VK_SPACE
      If m_pickerProperties.hoveredColorCell = INVALIDCOLORCELLINDEX Then
        m_pickerProperties.hoveredColor = m_selectedColor
      End If
      EndPickerSelection m_pickerProperties.hoveredColorCell = INVALIDCOLORCELLINDEX
  End Select

  If offset Then
    ' compute a new position based on our current position
    newSelection = 0
    If m_pickerProperties.hoveredColorCell = INVALIDCOLORCELLINDEX Then
      newSelection = IIf(offset > 0, DEFAULTCOLORCELLINDEX, MORECOLORSCELLINDEX)
    ElseIf m_pickerProperties.hoveredColorCell = DEFAULTCOLORCELLINDEX Then
      newSelection = IIf(offset > 0, 0, MORECOLORSCELLINDEX)
    ElseIf m_pickerProperties.hoveredColorCell = MORECOLORSCELLINDEX Then
      newSelection = IIf(offset > 0, DEFAULTCOLORCELLINDEX, m_pickerProperties.numberOfPredefinedColors - 1)
    Else
      newSelection = m_pickerProperties.hoveredColorCell + offset
      If newSelection < 0 Then
        newSelection = DEFAULTCOLORCELLINDEX
      ElseIf newSelection >= m_pickerProperties.numberOfPredefinedColors Then
        newSelection = MORECOLORSCELLINDEX
      End If
    End If

    ' ensure we don't select cells that we don't have
    Do While True
      If (newSelection = DEFAULTCOLORCELLINDEX) And Not HasDefaultColorText Then
        newSelection = IIf(offset > 0, 0, MORECOLORSCELLINDEX)
      ElseIf (newSelection = MORECOLORSCELLINDEX) And Not HasMoreColorsText Then
        newSelection = IIf(offset > 0, DEFAULTCOLORCELLINDEX, m_pickerProperties.numberOfPredefinedColors - 1)
      Else
        Exit Do
      End If
    Loop

    ' apply the new selection
    ChangePickerSelection newSelection
  End If
End Function

Private Function OnPickerLButtonUp(ByVal message As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
  Dim mousePosition As POINT
  Dim newSelection As Long

  ' do a hit test
  mousePosition.x = LoWord(lParam)
  mousePosition.y = HiWord(lParam)
  newSelection = PickerHitTest(mousePosition)

  If newSelection <> m_pickerProperties.hoveredColorCell Then
    ChangePickerSelection newSelection
  End If
  EndPickerSelection newSelection = INVALIDCOLORCELLINDEX
End Function

Private Function OnPickerMouseMove(ByVal message As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
  Dim mousePosition As POINT
  Dim newSelection As Long

  ' do a hit test
  mousePosition.x = LoWord(lParam)
  mousePosition.y = HiWord(lParam)
  newSelection = PickerHitTest(mousePosition)

  If newSelection <> m_pickerProperties.hoveredColorCell Then
    ChangePickerSelection newSelection
  End If
End Function

Private Function OnPickerPaint(ByVal message As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
  Const BF_LEFT = &H1
  Const BF_TOP = &H2
  Const BF_RIGHT = &H4
  Const BF_BOTTOM = &H8
  Const BF_RECT = BF_LEFT Or BF_TOP Or BF_RIGHT Or BF_BOTTOM
  Const BDR_RAISEDOUTER = &H1
  Const BDR_RAISEDINNER = &H4
  Const EDGE_RAISED = BDR_RAISEDOUTER Or BDR_RAISEDINNER
  Const COLOR_GRAYTEXT = 17
  Const PS_SOLID = 0
  Dim clientRectangle As RECTANGLE
  Dim hPen As Long
  Dim hPreviousPen As Long
  Dim hTargetDC As Long
  Dim i As Long
  Dim paintData As PAINTSTRUCT

  hTargetDC = BeginPaint(m_hWndDropDown, paintData)

  ' Draw a raised window edge. The extended window style WS_EX_WINDOWEDGE is supposed to do this, but for some reason isn't.

  GetClientRect m_hWndDropDown, clientRectangle
  If m_pickerProperties.has3DBorder Then
    hPen = CreatePen(PS_SOLID, 0, GetSysColor(COLOR_GRAYTEXT))
    hPreviousPen = SelectObject(hTargetDC, hPen)
    DrawRectangle hTargetDC, clientRectangle.Left, clientRectangle.Top, clientRectangle.Right, clientRectangle.Bottom
    SelectObject hTargetDC, hPreviousPen
  Else
    DrawEdge hTargetDC, clientRectangle, EDGE_RAISED, BF_RECT
  End If

  ' draw the "Automatic" text
  If HasDefaultColorText Then
    DrawPickerCell hTargetDC, DEFAULTCOLORCELLINDEX
  End If

  ' draw the color cells
  For i = 0 To m_pickerProperties.numberOfPredefinedColors - 1
    DrawPickerCell hTargetDC, i
  Next i

  ' draw the "More Colors" text
  If HasMoreColorsText Then
    DrawPickerCell hTargetDC, MORECOLORSCELLINDEX
  End If

  EndPaint m_hWndDropDown, paintData
End Function

Private Function OnPickerPaletteChanged(ByVal message As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
  If wParam <> m_hWndOwner Then
    If m_hWndOwner Then
      InvalidateRect m_hWndOwner, 0, 1
    End If
    If hWndDropDown Then
      InvalidateRect hWndDropDown, 0, 1
    End If
  End If
End Function

Private Function OnPickerQueryNewPalette(ByVal message As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
  If m_hWndOwner Then
    InvalidateRect m_hWndOwner, 0, 1
  End If
  If hWndDropDown Then
    InvalidateRect hWndDropDown, 0, 1
  End If
  OnPickerQueryNewPalette = 1
End Function

Private Function PickerHitTest(pt As POINT) As Long
  Dim cellIndex As Long
  Dim col As Long
  Dim row As Long

  If PtInRect(m_pickerProperties.moreColorsTextBoundingRectangle, pt.x, pt.y) Then
    PickerHitTest = MORECOLORSCELLINDEX
  ElseIf PtInRect(m_pickerProperties.defaultColorTextBoundingRectangle, pt.x, pt.y) Then
    PickerHitTest = DEFAULTCOLORCELLINDEX
  ElseIf PtInRect(m_pickerProperties.colorGridBoundingRectangle, pt.x, pt.y) = 0 Then
    PickerHitTest = INVALIDCOLORCELLINDEX
  Else
    row = (pt.y - m_pickerProperties.colorGridBoundingRectangle.Top) \ m_pickerProperties.colorCellSize.cy
    col = (pt.x - m_pickerProperties.colorGridBoundingRectangle.Left) \ m_pickerProperties.colorCellSize.cx
    If (row < 0) Or (row >= m_pickerProperties.numberOfRows) Or (col < 0) Or (col >= m_pickerProperties.numberOfColumns) Then
      PickerHitTest = INVALIDCOLORCELLINDEX
      Exit Function
    End If
    cellIndex = row * m_pickerProperties.numberOfColumns + col
    If cellIndex >= m_pickerProperties.numberOfPredefinedColors Then
      PickerHitTest = INVALIDCOLORCELLINDEX
      Exit Function
    End If
    PickerHitTest = cellIndex
  End If
End Function

Public Function Popup(buttonRectangle As RECTANGLE) As Boolean
  Const CC_RGBINIT = &H1
  Const CC_FULLOPEN = &H2
  Const CC_ANYCOLOR = &H100
  Const COLOR_MENU = 4
  Const CS_VREDRAW = &H1
  Const CS_HREDRAW = &H2
  Const CS_CLASSDC = &H40
  Const CS_SAVEBITS = &H800
  Const CS_DROPSHADOW = &H20000
  Const IMAGE_CURSOR = 2
  Const LR_DEFAULTSIZE = &H40
  Const LR_SHARED = &H8000
  Const OCR_NORMAL = 32512&
  Const PM_NOREMOVE = &H0
  Const SPI_GETDROPSHADOW = &H1024
  Const SW_SHOWNA = 8
  Const WS_POPUP = &H80000000
  Const WM_PAINT = &HF
  Const WM_KEYDOWN = &H100
  Const WM_KEYUP = &H101
  Const WM_MOUSEFIRST = &H200
  Const WM_MOUSELAST = &H20E
  Const WM_MOUSEMOVE = &H200
  Const WM_LBUTTONUP = &H202
  Const WM_RBUTTONDOWN = &H204
'  Const WM_USER = &H400
'  Const TTM_RELAYEVENT = (WM_USER + 7)
  Dim cancelled As Boolean
  Dim colorDialogData As TCHOOSECOLOR
  Dim dropShadow As Long
  Dim hMod As Long
  Dim isFirstLButtonUp As Boolean
  Dim Msg As MSGDETAILS
  Dim windowClass As WNDCLASSEX
  Dim wndClassAtom As Long

  cancelled = True
  If m_hWndParent = 0 Then
    Exit Function
  End If

  ' initialize some sizing parameters
  SetNumberOfPredefinedColors UBound(m_predefinedColors) - LBound(m_predefinedColors) + 1

  ' initialize our state
  m_pickerProperties.hoveredColorCell = INVALIDCOLORCELLINDEX
  m_pickerProperties.selectedColorCell = INVALIDCOLORCELLINDEX
  m_pickerProperties.hoveredColor = m_selectedColor

  ' register the window class used for the picker
  windowClass.cbSize = LenB(windowClass)
  windowClass.style = CS_CLASSDC Or CS_SAVEBITS Or CS_HREDRAW Or CS_VREDRAW
  hMod = GetModuleHandle(StrPtr("user32.dll"))
  If hMod Then
    windowClass.lpfnWndProc = GetProcAddress(hMod, "DefWindowProcW")
  End If
  windowClass.hInstance = App.hInstance
  windowClass.hCursor = LoadImage(0, OCR_NORMAL, IMAGE_CURSOR, 0, 0, LR_DEFAULTSIZE Or LR_SHARED)
  windowClass.hbrBackground = COLOR_MENU + 1
  windowClass.lpszClassName = StrPtr("ColorPicker")
  SystemParametersInfo SPI_GETDROPSHADOW, 0, VarPtr(dropShadow), 0
  If dropShadow Then
    windowClass.style = windowClass.style Or CS_DROPSHADOW
  End If
  wndClassAtom = RegisterClassEx(windowClass)

  ' create the window
  m_hWndDropDown = CreateWindowEx(0, wndClassAtom, StrPtr(""), WS_POPUP, buttonRectangle.Left, buttonRectangle.Bottom, 100, 100, m_hWndParent, 0, App.hInstance, 0)

  If m_hWndDropDown Then
    If Not SubclassWindow(m_hWndDropDown, Me, EnumSubclassID.escidColorPicker) Then
      Debug.Print "Subclassing failed!"
    End If
    ' set the window size
    SetPickerWindowSize

    ' create the tooltips
    SetupPickerToolTips

    ' find which cell (if any) corresponds to the initial color
    m_pickerProperties.selectedColorCell = FindPickerCellFromColor(m_selectedColor)

    ' make visible
    ShowWindow m_hWndDropDown, SW_SHOWNA

    ' purge the message queue of paints
    While PeekMessage(Msg, 0, WM_PAINT, WM_PAINT, PM_NOREMOVE)
      If GetMessage(Msg, 0, WM_PAINT, WM_PAINT) = 0 Then
        cancelled = True
        GoTo ret
      End If
      DispatchMessage Msg
    Wend

    ' set capture to the window which received this message
    SetCapture m_hWndDropDown

    ' get messages until capture lost or cancelled/accepted
    isFirstLButtonUp = True
    Do While m_hWndDropDown = GetCapture()
      If GetMessage(Msg, 0, 0, 0) = 0 Then
        PostQuitMessage Msg.wParam
        Exit Do
      End If

      If Msg.message >= WM_MOUSEFIRST And Msg.message <= WM_MOUSELAST Then
        'SendMessageAsLong m_toolTipControl.hWnd, TTM_RELAYEVENT, 0, VarPtr(Msg)
        m_toolTipControl.RelayMouseEvent Msg.hWnd, Msg.message, Msg.wParam, Msg.lParam
      End If
      Select Case Msg.message
        Case WM_LBUTTONUP
          If isFirstLButtonUp And m_bIgnoreFirstLeftMouseButtonUp Then
            isFirstLButtonUp = False
          Else
            OnPickerLButtonUp Msg.message, Msg.wParam, Msg.lParam
          End If
        Case WM_MOUSEMOVE
          OnPickerMouseMove Msg.message, Msg.wParam, Msg.lParam
        Case WM_KEYUP
          '
        Case WM_KEYDOWN
          OnPickerKeyDown Msg.message, Msg.wParam, Msg.lParam
        Case WM_RBUTTONDOWN
          ReleaseCapture
          m_pickerProperties.cancelled = True
        Case Else
          DispatchMessage Msg
      End Select
    Loop
    ReleaseCapture
    cancelled = m_pickerProperties.cancelled

    ' destroy the window
    m_toolTipControl.Destroy
    Set m_toolTipControl = Nothing
    If Not UnSubclassWindow(m_hWndDropDown, EnumSubclassID.escidColorPicker) Then
      Debug.Print "UnSubclassing failed!"
    End If
    DestroyWindow m_hWndDropDown

    ' if needed, show the color selection dialog
    If Not cancelled Then
      If m_pickerProperties.hoveredColorCell = MORECOLORSCELLINDEX Then
        colorDialogData.lStructSize = LenB(colorDialogData)
        colorDialogData.lpCustColors = VarPtr(m_customColors(0))
        colorDialogData.hwndOwner = m_hWndOwner
        colorDialogData.Flags = CC_FULLOPEN Or CC_ANYCOLOR Or CC_RGBINIT
        colorDialogData.rgbResult = m_selectedColor
        If ChooseColor(colorDialogData) Then
          m_selectedColor = colorDialogData.rgbResult
        Else
          cancelled = True
        End If
      Else
        m_selectedColor = m_pickerProperties.hoveredColor
      End If
    End If
  End If

ret:
  ' unregister our class
  UnregisterClass wndClassAtom, App.hInstance
  Popup = Not cancelled
End Function

Private Sub SetNumberOfPredefinedColors(ByVal number As Long)
  Dim dummy As Integer
  Dim i As Long
  Dim p As Long
  Dim pLogPalette As Long

  If number > MAXNUMBEROFCOLORCELLS Then number = MAXNUMBEROFCOLORCELLS
  m_pickerProperties.numberOfPredefinedColors = number
  ' retrieve the number of columns and rows
  m_pickerProperties.numberOfColumns = 8
  m_pickerProperties.numberOfRows = m_pickerProperties.numberOfPredefinedColors \ m_pickerProperties.numberOfColumns
  If (m_pickerProperties.numberOfPredefinedColors Mod m_pickerProperties.numberOfColumns) <> 0 Then
    m_pickerProperties.numberOfRows = m_pickerProperties.numberOfRows + 1
  End If

  ' create the palette
  If m_pickerProperties.hPalette Then
    DeleteObject m_pickerProperties.hPalette
    m_pickerProperties.hPalette = 0
  End If

  pLogPalette = HeapAlloc(GetProcessHeap(), 0, 8 + MAXNUMBEROFCOLORCELLS * 4)
  If pLogPalette Then
    dummy = &H300
    CopyMemory pLogPalette, VarPtr(dummy), LenB(dummy)
    dummy = m_pickerProperties.numberOfPredefinedColors
    CopyMemory pLogPalette + LenB(dummy), VarPtr(dummy), LenB(dummy)
    p = pLogPalette + 2 * LenB(dummy)

    For i = 0 To m_pickerProperties.numberOfPredefinedColors - 1
      CopyMemory p, VarPtr(m_predefinedColors(i).rgbColor), 4
      p = p + 4
    Next i
    m_pickerProperties.hPalette = CreatePalette(pLogPalette)
    HeapFree GetProcessHeap(), 0, pLogPalette
    pLogPalette = 0
  End If
End Sub

Private Sub SetPickerWindowSize()
  Const MONITOR_DEFAULTTONEAREST = &H2
  Const SM_CXSCREEN = 0
  Const SM_CYSCREEN = 1
  Const SM_CXEDGE = 45
  Const SM_CYEDGE = 46
  Const WM_GETFONT = &H31
  Dim hClientDC As Long
  Dim hPreviousFont As Long
  Dim minimumWidth As Long
  Dim monInfo As MONITORINFO
  Dim parentWindowRectangle As RECTANGLE
  Dim rc As RECTANGLE
  Dim screenRectangle As RECTANGLE
  Dim sz As Size
  Dim TextSize As Size
  Dim totalWidth As Long
  Dim windowRectangle As RECTANGLE

  ' if we are showing a "More Colors" or "Automatic" text area, get the font and text size
  If HasMoreColorsText() Or HasDefaultColorText() Then
    hClientDC = GetDC(m_hWndDropDown)
    hPreviousFont = SelectObject(hClientDC, SendMessageAsLong(m_hWndOwner, WM_GETFONT, 0, 0))

    ' get the size of the "More Colors" text
    If HasMoreColorsText() Then
      GetTextExtentPoint32 hClientDC, StrPtr(m_moreColorsText), Len(m_moreColorsText), TextSize
    End If

    ' get the size of the "Automatic" text
    If HasDefaultColorText() Then
      GetTextExtentPoint32 hClientDC, StrPtr(m_defaultColorText), Len(m_defaultColorText), sz
      TextSize.cx = IIf(TextSize.cx > sz.cx, TextSize.cx, sz.cx)
      TextSize.cy = IIf(TextSize.cy > sz.cy, TextSize.cy, sz.cy)
    End If
    SelectObject hClientDC, hPreviousFont

    ' commpute the final size
    TextSize.cx = TextSize.cx + 2 * (m_outerTextCellBorderSize.cx + m_innerTextCellBorderSize.cx)
    TextSize.cy = TextSize.cy + 2 * (m_outerTextCellBorderSize.cy + m_innerTextCellBorderSize.cy)
    ReleaseDC m_hWndDropDown, hClientDC
  End If

  ' initiailize our box size
  m_pickerProperties.colorCellSize.cx = m_colorCellContentSize.cx + (m_innerColorCellBorderSize.cx + m_outerColorCellBorderSize.cx) * 2
  m_pickerProperties.colorCellSize.cy = m_colorCellContentSize.cy + (m_innerColorCellBorderSize.cy + m_outerColorCellBorderSize.cy) * 2

  ' compute the minimum width
  totalWidth = m_pickerProperties.numberOfColumns * m_pickerProperties.colorCellSize.cx
  minimumWidth = IIf(totalWidth > TextSize.cx, totalWidth, TextSize.cx)

  ' create the rectangle for the "Automatic" text
  m_pickerProperties.defaultColorTextBoundingRectangle.Left = 0
  m_pickerProperties.defaultColorTextBoundingRectangle.Top = 0
  m_pickerProperties.defaultColorTextBoundingRectangle.Right = m_pickerProperties.defaultColorTextBoundingRectangle.Left + minimumWidth
  m_pickerProperties.defaultColorTextBoundingRectangle.Bottom = m_pickerProperties.defaultColorTextBoundingRectangle.Top + IIf(HasDefaultColorText, TextSize.cy, 0)

  ' initialize the color box rectangle
  m_pickerProperties.colorGridBoundingRectangle.Left = (minimumWidth - totalWidth) \ 2
  m_pickerProperties.colorGridBoundingRectangle.Top = m_pickerProperties.defaultColorTextBoundingRectangle.Bottom
  m_pickerProperties.colorGridBoundingRectangle.Right = m_pickerProperties.colorGridBoundingRectangle.Left + totalWidth
  m_pickerProperties.colorGridBoundingRectangle.Bottom = m_pickerProperties.colorGridBoundingRectangle.Top + m_pickerProperties.numberOfRows * m_pickerProperties.colorCellSize.cy

  ' create the rectangle for the "More Colors" text
  m_pickerProperties.moreColorsTextBoundingRectangle.Left = 0
  m_pickerProperties.moreColorsTextBoundingRectangle.Top = m_pickerProperties.colorGridBoundingRectangle.Bottom
  m_pickerProperties.moreColorsTextBoundingRectangle.Right = m_pickerProperties.moreColorsTextBoundingRectangle.Left + minimumWidth
  m_pickerProperties.moreColorsTextBoundingRectangle.Bottom = m_pickerProperties.moreColorsTextBoundingRectangle.Top + IIf(HasMoreColorsText, TextSize.cy, 0)

  ' get the current window position, and set the new size
  windowRectangle.Left = m_pickerProperties.defaultColorTextBoundingRectangle.Left
  windowRectangle.Top = m_pickerProperties.defaultColorTextBoundingRectangle.Top
  windowRectangle.Right = m_pickerProperties.moreColorsTextBoundingRectangle.Right
  windowRectangle.Bottom = m_pickerProperties.moreColorsTextBoundingRectangle.Bottom
  GetWindowRect m_hWndDropDown, rc
  OffsetRect windowRectangle, rc.Left, rc.Top

  ' adjust the rects for the border
  windowRectangle.Right = windowRectangle.Right + 2 * GetSystemMetrics(SM_CXEDGE)
  windowRectangle.Bottom = windowRectangle.Bottom + 2 * GetSystemMetrics(SM_CYEDGE)
  OffsetRect m_pickerProperties.defaultColorTextBoundingRectangle, GetSystemMetrics(SM_CXEDGE), GetSystemMetrics(SM_CYEDGE)
  OffsetRect m_pickerProperties.colorGridBoundingRectangle, GetSystemMetrics(SM_CXEDGE), GetSystemMetrics(SM_CYEDGE)
  OffsetRect m_pickerProperties.moreColorsTextBoundingRectangle, GetSystemMetrics(SM_CXEDGE), GetSystemMetrics(SM_CYEDGE)

  ' get the screen rectangle
  screenRectangle.Right = screenRectangle.Left + GetSystemMetrics(SM_CXSCREEN)
  screenRectangle.Bottom = screenRectangle.Top + GetSystemMetrics(SM_CYSCREEN)
  monInfo.cbSize = LenB(monInfo)
  GetMonitorInfo MonitorFromWindow(m_hWndOwner, MONITOR_DEFAULTTONEAREST), monInfo
  screenRectangle = monInfo.rcWork

  ' need to check it'll fit on screen: Too far right?
  If windowRectangle.Right > screenRectangle.Right Then
    OffsetRect windowRectangle, screenRectangle.Right - windowRectangle.Right, 0
  End If
  ' Too far left?
  If windowRectangle.Left < screenRectangle.Left Then
    OffsetRect windowRectangle, screenRectangle.Left - windowRectangle.Left, 0
  End If

  ' Bottom falling out of screen?  If so, move the whole popup above the parent window.
  If windowRectangle.Bottom > screenRectangle.Bottom Then
    GetWindowRect m_hWndOwner, parentWindowRectangle
    OffsetRect windowRectangle, 0, -((parentWindowRectangle.Bottom - parentWindowRectangle.Top) + (windowRectangle.Bottom - windowRectangle.Top))
  End If

  ' set the window size and position
  MoveWindow m_hWndDropDown, windowRectangle.Left, windowRectangle.Top, windowRectangle.Right - windowRectangle.Left, windowRectangle.Bottom - windowRectangle.Top, 1
End Sub

Private Sub SetupPickerToolTips()
  Dim cellRectangle As RECTANGLE
  Dim i As Long

  Set m_toolTipControl = New clsToolTip
  m_toolTipControl.Create m_hWndDropDown
  m_toolTipControl.Activate

  ' add a tool for each cell
  For i = 0 To m_pickerProperties.numberOfPredefinedColors - 1
    If GetPickerCellRectangle(i, cellRectangle) Then
      m_toolTipControl.AddTool 1, m_predefinedColors(i).clrName, True, False, False, True, cellRectangle.Left, cellRectangle.Top, cellRectangle.Right, cellRectangle.Bottom
    End If
  Next i
End Sub
