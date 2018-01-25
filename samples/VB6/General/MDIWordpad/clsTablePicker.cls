VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTablePicker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

  Implements ISubclassedWindow


  Private Const INVALIDCELLINDEX = -1
  Private Const CAPTIONCELLINDEX = -2
  Private Const SIZECELLINDEX = -3

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
    hoveredCell As Long
    numberOfColumns As Long
    numberOfRows As Long
    cancelled As Boolean
    captionBoundingRectangle As RECTANGLE
    gridBoundingRectangle As RECTANGLE
    sizeBoundingRectangle As RECTANGLE
    cellSize As Size
    BackColor As Long
    has3DBorder As Boolean
  End Type


  Private m_cellContentSize As Size

  Private m_bIgnoreFirstLeftMouseButtonUp As Boolean
  Private m_caption As String
  Private m_hWndDropDown As Long
  Private m_hWndOwner As Long
  Private m_hWndParent As Long
  Private m_pickerProperties As PICKERWINDOWPROPERTIES
  Private m_selectedColumnCount As Long
  Private m_selectedRowCount As Long


  Private Declare Function BeginPaint Lib "user32.dll" (ByVal hWnd As Long, lpPaint As PAINTSTRUCT) As Long
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
  Private Declare Function RegisterClassEx Lib "user32.dll" Alias "RegisterClassExW" (lpwcx As WNDCLASSEX) As Long
  Private Declare Function ReleaseCapture Lib "user32.dll" () As Long
  Private Declare Function ReleaseDC Lib "user32.dll" (ByVal hWnd As Long, ByVal hDC As Long) As Long
  Private Declare Function SelectObject Lib "gdi32.dll" (ByVal hDC As Long, ByVal hObject As Long) As Long
  Private Declare Function SendMessageAsLong Lib "user32.dll" Alias "SendMessageW" (ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
  Private Declare Function SetBkColor Lib "gdi32.dll" (ByVal hDC As Long, ByVal crColor As Long) As Long
  Private Declare Function SetBkMode Lib "gdi32.dll" (ByVal hDC As Long, ByVal iBkMode As Long) As Long
  Private Declare Function SetCapture Lib "user32.dll" (ByVal hWnd As Long) As Long
  Private Declare Function SetTextColor Lib "gdi32.dll" (ByVal hDC As Long, ByVal crColor As Long) As Long
  Private Declare Function ShowWindow Lib "user32.dll" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
  Private Declare Function SystemParametersInfo Lib "user32.dll" Alias "SystemParametersInfoW" (ByVal uiAction As Long, ByVal uiParam As Long, ByVal pvParam As Long, ByVal fWinIni As Long) As Long
  Private Declare Function UnregisterClass Lib "user32.dll" Alias "UnregisterClassW" (ByVal lpClassName As Long, ByVal hInstance As Long) As Long


Private Function ISubclassedWindow_HandleMessage(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long, ByVal eSubclassID As EnumSubclassID, bCallDefProc As Boolean) As Long
  Dim lRet As Long

  On Error GoTo StdHandler_Error
  Select Case eSubclassID
    Case EnumSubclassID.escidTablePicker
      lRet = HandleMessage_DropDown(hWnd, uMsg, wParam, lParam, bCallDefProc)
    Case Else
      Debug.Print "clsTablePicker.ISubclassedWindow_HandleMessage: Unknown Subclassing ID " & CStr(eSubclassID)
  End Select

StdHandler_Ende:
  ISubclassedWindow_HandleMessage = lRet
  Exit Function

StdHandler_Error:
  Debug.Print "Error in clsTablePicker.ISubclassedWindow_HandleMessage (SubclassID=" & CStr(eSubclassID) & ": ", Err.number, Err.Description
  Resume StdHandler_Ende
End Function

Private Function HandleMessage_DropDown(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long, bCallDefProc As Boolean) As Long
  Const WM_PAINT = &HF
  Dim lRet As Long

  On Error GoTo StdHandler_Error
  Select Case uMsg
    Case WM_PAINT
      lRet = OnPickerPaint(uMsg, wParam, lParam)
      bCallDefProc = False
  End Select

StdHandler_Ende:
  HandleMessage_DropDown = lRet
  Exit Function

StdHandler_Error:
  Debug.Print "Error in clsTablePicker.HandleMessage_DropDown: ", Err.number, Err.Description
  Resume StdHandler_Ende
End Function


Private Sub Class_Initialize()
  Const COLOR_MENU = 4
  Const COLOR_BTNFACE = 15
  Const SPI_GETFLATMENU = &H1022
  Dim clr As Long

  m_caption = "Insert Table"

  m_cellContentSize.cx = 18
  m_cellContentSize.cy = 18

  clr = GetSysColor(COLOR_MENU)
  If (Abs(GetRed(clr) - GetRed(vbWhite)) <= 10) And (Abs(GetGreen(clr) - GetGreen(vbWhite)) <= 10) And (Abs(GetBlue(clr) - GetBlue(vbWhite)) <= 10) Then
    clr = GetSysColor(COLOR_BTNFACE)
  End If
  m_pickerProperties.BackColor = clr
  SystemParametersInfo SPI_GETFLATMENU, 0, VarPtr(m_pickerProperties.has3DBorder), 0

  m_pickerProperties.numberOfColumns = 10
  m_pickerProperties.numberOfRows = 15
End Sub


Public Property Get Caption() As String
  Caption = m_caption
End Property

Public Property Let Caption(ByVal newVal As String)
  m_caption = newVal
End Property

Public Property Get HasCaption() As Boolean
  HasCaption = Len(m_caption) > 0
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

Public Property Get SelectedColumnCount() As Long
  SelectedColumnCount = m_selectedColumnCount
End Property

Public Property Get SelectedRowCount()
  SelectedRowCount = m_selectedRowCount
End Property


Private Sub ChangePickerSelection(ByVal cellIndex As Long)
  Dim hDCClient As Long
  Dim i As Long

  hDCClient = GetDC(m_hWndDropDown)

  ' redraw all cells
  m_pickerProperties.hoveredCell = cellIndex
  For i = 0 To m_pickerProperties.numberOfRows * m_pickerProperties.numberOfColumns - 1
    DrawPickerCell hDCClient, i
  Next i

  ' draw the size cell
  DrawPickerCell hDCClient, SIZECELLINDEX

  ReleaseDC m_hWndDropDown, hDCClient
End Sub

Private Sub DrawPickerCell(ByVal hTargetDC As Long, ByVal cellIndex As Long)
  Const COLOR_MENUHILIGHT = 29
  Const COLOR_MENUTEXT = 7
  Const COLOR_WINDOW = 5
  Const DT_LEFT = &H0
  Const DT_VCENTER = &H4
  Const DT_SINGLELINE = &H20
  Const ETO_OPAQUE = &H2
  Const SM_CXEDGE = 45
  Const TRANSPARENT = 1
  Const WM_GETFONT = &H31
  Dim backgroundColor As Long
  Dim cellRectangle As RECTANGLE
  Dim cellText As String
  Dim hPreviousFont As Long
  Dim isSelected As Boolean
  Dim previousBkMode As Long
  Dim previousTextColor As Long
  Dim textColor As Long
  Dim columnsToDisplay As Long
  Dim rowsToDisplay As Long

  ' get the drawing rect
  If Not GetPickerCellRectangle(cellIndex, cellRectangle) Then
    Exit Sub
  End If

  If cellIndex = CAPTIONCELLINDEX Then
    cellText = m_caption
    textColor = GetSysColor(COLOR_MENUTEXT)
  ElseIf cellIndex = SIZECELLINDEX Then
    If m_pickerProperties.hoveredCell < 0 Then
      columnsToDisplay = 0
      rowsToDisplay = 0
    Else
      columnsToDisplay = (m_pickerProperties.hoveredCell Mod m_pickerProperties.numberOfColumns) + 1
      rowsToDisplay = (m_pickerProperties.hoveredCell \ m_pickerProperties.numberOfColumns) + 1
    End If
    cellText = CStr(columnsToDisplay) & " x " & CStr(rowsToDisplay)
    textColor = GetSysColor(COLOR_MENUTEXT)
  End If

  If Len(cellText) > 0 Then
    If m_hWndOwner Then
      hPreviousFont = SelectObject(hTargetDC, SendMessageAsLong(m_hWndOwner, WM_GETFONT, 0, 0))
    Else
      hPreviousFont = SelectObject(hTargetDC, SendMessageAsLong(m_hWndParent, WM_GETFONT, 0, 0))
    End If
    previousTextColor = SetTextColor(hTargetDC, textColor)
    previousBkMode = SetBkMode(hTargetDC, TRANSPARENT)
    SetBkColor hTargetDC, m_pickerProperties.BackColor
    ExtTextOut hTargetDC, 0, 0, ETO_OPAQUE, cellRectangle, 0, 0, 0
    InflateRect cellRectangle, -GetSystemMetrics(SM_CXEDGE), 0
    DrawText hTargetDC, StrPtr(cellText), -1, cellRectangle, DT_LEFT Or DT_VCENTER Or DT_SINGLELINE
    SetBkMode hTargetDC, previousBkMode
    SetTextColor hTargetDC, previousTextColor
    SelectObject hTargetDC, hPreviousFont
  Else
    ' based on the selectons, get our colors
    If (m_pickerProperties.hoveredCell Mod m_pickerProperties.numberOfColumns) >= (cellIndex Mod m_pickerProperties.numberOfColumns) Then
      If (m_pickerProperties.hoveredCell \ m_pickerProperties.numberOfColumns) >= (cellIndex \ m_pickerProperties.numberOfColumns) Then
        isSelected = True
      End If
    End If
    backgroundColor = GetSysColor(IIf(isSelected, COLOR_MENUHILIGHT, COLOR_WINDOW))

    ' draw the background
    SetBkColor hTargetDC, backgroundColor
    ExtTextOut hTargetDC, 0, 0, ETO_OPAQUE, cellRectangle, 0, 0, 0
  End If
End Sub

Private Sub EndPickerSelection(ByVal cancelled As Boolean)
  ReleaseCapture
  m_pickerProperties.cancelled = cancelled
End Sub

Private Function GetPickerCellRectangle(ByVal cellIndex As Long, pRectangle As RECTANGLE) As Boolean
  If cellIndex = CAPTIONCELLINDEX Then
    pRectangle = m_pickerProperties.captionBoundingRectangle
    GetPickerCellRectangle = True
  ElseIf cellIndex = SIZECELLINDEX Then
    pRectangle = m_pickerProperties.sizeBoundingRectangle
    GetPickerCellRectangle = True
  ElseIf (cellIndex < 0) Or (cellIndex >= m_pickerProperties.numberOfRows * m_pickerProperties.numberOfColumns) Then
    GetPickerCellRectangle = False
  Else
    pRectangle.Left = (cellIndex Mod m_pickerProperties.numberOfColumns) * m_pickerProperties.cellSize.cx + m_pickerProperties.gridBoundingRectangle.Left
    pRectangle.Top = (cellIndex \ m_pickerProperties.numberOfColumns) * m_pickerProperties.cellSize.cy + m_pickerProperties.gridBoundingRectangle.Top
    pRectangle.Right = pRectangle.Left + m_pickerProperties.cellSize.cx - 2
    pRectangle.Bottom = pRectangle.Top + m_pickerProperties.cellSize.cy - 2
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
      EndPickerSelection True
    Case VK_RETURN, VK_SPACE
      EndPickerSelection m_pickerProperties.hoveredCell < 0
  End Select

  If offset Then
    ' compute a new position based on our current position
    newSelection = 0
    If m_pickerProperties.hoveredCell = INVALIDCELLINDEX Then
      newSelection = 0
    Else
      newSelection = m_pickerProperties.hoveredCell + offset
      If newSelection < 0 Then
        newSelection = 0
      ElseIf newSelection >= m_pickerProperties.numberOfRows * m_pickerProperties.numberOfColumns Then
        newSelection = m_pickerProperties.numberOfRows * m_pickerProperties.numberOfColumns - 1
      End If
    End If

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

  If newSelection <> m_pickerProperties.hoveredCell Then
    ChangePickerSelection newSelection
  End If
  EndPickerSelection newSelection < 0
End Function

Private Function OnPickerMouseMove(ByVal message As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
  Dim mousePosition As POINT
  Dim newSelection As Long

  ' do a hit test
  mousePosition.x = LoWord(lParam)
  mousePosition.y = HiWord(lParam)
  newSelection = PickerHitTest(mousePosition)

  If newSelection <> m_pickerProperties.hoveredCell Then
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
  Const ETO_OPAQUE = &H2
  Const PS_SOLID = 0
  Dim clientRectangle As RECTANGLE
  Dim hPen As Long
  Dim hPreviousPen As Long
  Dim hTargetDC As Long
  Dim i As Long
  Dim paintData As PAINTSTRUCT
  Dim previousBackColor As Long

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

  previousBackColor = SetBkColor(hTargetDC, RGB(0, 0, 0))

  ' draw the background
  InflateRect clientRectangle, -1, -1
  SetBkColor hTargetDC, m_pickerProperties.BackColor
  ExtTextOut hTargetDC, 0, 0, ETO_OPAQUE, clientRectangle, 0, 0, 0

  ' draw the caption
  If HasCaption Then
    DrawPickerCell hTargetDC, CAPTIONCELLINDEX
  End If

  ' draw the cells
  DrawPickerCell hTargetDC, CAPTIONCELLINDEX
  For i = 0 To m_pickerProperties.numberOfRows * m_pickerProperties.numberOfColumns - 1
    DrawPickerCell hTargetDC, i
  Next i

  ' draw the size cell
  DrawPickerCell hTargetDC, SIZECELLINDEX

  SetBkColor hTargetDC, previousBackColor

  EndPaint m_hWndDropDown, paintData
End Function

Private Function PickerHitTest(pt As POINT) As Long
  Dim cellIndex As Long
  Dim col As Long
  Dim row As Long

  If PtInRect(m_pickerProperties.captionBoundingRectangle, pt.x, pt.y) Then
    PickerHitTest = CAPTIONCELLINDEX
  ElseIf PtInRect(m_pickerProperties.sizeBoundingRectangle, pt.x, pt.y) Then
    PickerHitTest = SIZECELLINDEX
  ElseIf PtInRect(m_pickerProperties.gridBoundingRectangle, pt.x, pt.y) = 0 Then
    PickerHitTest = INVALIDCELLINDEX
  Else
    row = (pt.y - m_pickerProperties.gridBoundingRectangle.Top) \ m_pickerProperties.cellSize.cy
    col = (pt.x - m_pickerProperties.gridBoundingRectangle.Left) \ m_pickerProperties.cellSize.cx
    If (row < 0) Or (row >= m_pickerProperties.numberOfRows) Or (col < 0) Or (col >= m_pickerProperties.numberOfColumns) Then
      PickerHitTest = INVALIDCELLINDEX
      Exit Function
    End If
    cellIndex = row * m_pickerProperties.numberOfColumns + col
    PickerHitTest = cellIndex
  End If
End Function

Public Function Popup(buttonRectangle As RECTANGLE) As Boolean
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
  Const WM_MOUSEMOVE = &H200
  Const WM_LBUTTONUP = &H202
  Const WM_RBUTTONDOWN = &H204
  Dim cancelled As Boolean
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

  ' initialize our state
  m_pickerProperties.hoveredCell = INVALIDCELLINDEX

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
  windowClass.lpszClassName = StrPtr("TablePicker")
  SystemParametersInfo SPI_GETDROPSHADOW, 0, VarPtr(dropShadow), 0
  If dropShadow Then
    windowClass.style = windowClass.style Or CS_DROPSHADOW
  End If
  wndClassAtom = RegisterClassEx(windowClass)

  ' create the window
  m_hWndDropDown = CreateWindowEx(0, wndClassAtom, StrPtr(""), WS_POPUP, buttonRectangle.Left, buttonRectangle.Bottom, 100, 100, m_hWndParent, 0, App.hInstance, 0)

  If m_hWndDropDown Then
    If Not SubclassWindow(m_hWndDropDown, Me, EnumSubclassID.escidTablePicker) Then
      Debug.Print "Subclassing failed!"
    End If
    ' set the window size
    SetPickerWindowSize

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
    If Not UnSubclassWindow(m_hWndDropDown, EnumSubclassID.escidTablePicker) Then
      Debug.Print "UnSubclassing failed!"
    End If
    DestroyWindow m_hWndDropDown

    ' if needed, show the color selection dialog
    If Not cancelled Then
      m_selectedColumnCount = (m_pickerProperties.hoveredCell Mod m_pickerProperties.numberOfColumns) + 1
      m_selectedRowCount = (m_pickerProperties.hoveredCell \ m_pickerProperties.numberOfColumns) + 1
    End If
  End If

ret:
  ' unregister our class
  UnregisterClass wndClassAtom, App.hInstance
  Popup = Not cancelled
End Function

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

  hClientDC = GetDC(m_hWndDropDown)
  hPreviousFont = SelectObject(hClientDC, SendMessageAsLong(m_hWndOwner, WM_GETFONT, 0, 0))

  If HasCaption Then
    GetTextExtentPoint32 hClientDC, StrPtr(m_caption), Len(m_caption), TextSize
  End If
  GetTextExtentPoint32 hClientDC, StrPtr("15 x 10"), Len("15 x 10"), sz
  TextSize.cx = IIf(TextSize.cx > sz.cx, TextSize.cx, sz.cx)
  TextSize.cy = IIf(TextSize.cy > sz.cy, TextSize.cy, sz.cy)
  SelectObject hClientDC, hPreviousFont

  ' commpute the final size
  TextSize.cx = TextSize.cx + 2 * 4 + 2 * GetSystemMetrics(SM_CXEDGE)
  TextSize.cy = TextSize.cy + 2 * 4
  ReleaseDC m_hWndDropDown, hClientDC

  ' initiailize our box size
  m_pickerProperties.cellSize = m_cellContentSize

  ' compute the minimum width
  totalWidth = m_pickerProperties.numberOfColumns * m_pickerProperties.cellSize.cx
  minimumWidth = 2 * GetSystemMetrics(SM_CXEDGE) + IIf(totalWidth > TextSize.cx, totalWidth, TextSize.cx)

  ' create the rectangle for the caption
  m_pickerProperties.captionBoundingRectangle.Left = (minimumWidth - totalWidth) \ 2
  m_pickerProperties.captionBoundingRectangle.Top = GetSystemMetrics(SM_CYEDGE)
  m_pickerProperties.captionBoundingRectangle.Right = m_pickerProperties.captionBoundingRectangle.Left + totalWidth
  m_pickerProperties.captionBoundingRectangle.Bottom = m_pickerProperties.captionBoundingRectangle.Top + IIf(HasCaption, TextSize.cy, 0)

  ' initialize the grid rectangle
  m_pickerProperties.gridBoundingRectangle.Left = (minimumWidth - totalWidth) \ 2
  m_pickerProperties.gridBoundingRectangle.Top = m_pickerProperties.captionBoundingRectangle.Bottom
  m_pickerProperties.gridBoundingRectangle.Right = m_pickerProperties.gridBoundingRectangle.Left + totalWidth
  m_pickerProperties.gridBoundingRectangle.Bottom = m_pickerProperties.gridBoundingRectangle.Top + m_pickerProperties.numberOfRows * m_pickerProperties.cellSize.cy + GetSystemMetrics(SM_CYEDGE)

  ' create the rectangle for the current size
  m_pickerProperties.sizeBoundingRectangle.Left = (minimumWidth - totalWidth) \ 2
  m_pickerProperties.sizeBoundingRectangle.Top = m_pickerProperties.gridBoundingRectangle.Bottom
  m_pickerProperties.sizeBoundingRectangle.Right = m_pickerProperties.sizeBoundingRectangle.Left + totalWidth
  m_pickerProperties.sizeBoundingRectangle.Bottom = m_pickerProperties.sizeBoundingRectangle.Top + TextSize.cy + GetSystemMetrics(SM_CYEDGE)

  ' get the current window position, and set the new size
  windowRectangle.Left = 0
  windowRectangle.Top = 0
  windowRectangle.Right = m_pickerProperties.gridBoundingRectangle.Right
  windowRectangle.Bottom = m_pickerProperties.sizeBoundingRectangle.Bottom
  GetWindowRect m_hWndDropDown, rc
  OffsetRect windowRectangle, rc.Left, rc.Top

  ' adjust the rects for the border
  windowRectangle.Right = windowRectangle.Right + 2 * GetSystemMetrics(SM_CXEDGE)
  windowRectangle.Bottom = windowRectangle.Bottom + 2 * GetSystemMetrics(SM_CYEDGE)
  OffsetRect m_pickerProperties.gridBoundingRectangle, GetSystemMetrics(SM_CXEDGE), GetSystemMetrics(SM_CYEDGE)

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
