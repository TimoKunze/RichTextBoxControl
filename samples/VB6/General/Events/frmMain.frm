VERSION 5.00
Object = "{CD7AC637-45AE-4969-81B8-A56464ECD6C4}#1.0#0"; "RichTextBoxCtlU.ocx"
Begin VB.Form frmMain 
   Caption         =   "RichTextBox 1.0 - Events Sample"
   ClientHeight    =   6495
   ClientLeft      =   60
   ClientTop       =   750
   ClientWidth     =   10335
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   433
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   689
   StartUpPosition =   2  'Bildschirmmitte
   Begin RichTextBoxCtlLibUCtl.RichTextBox RichTxtBoxU 
      Height          =   6255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5175
      _cx             =   9128
      _cy             =   11033
      AcceptTabKey    =   -1  'True
      AllowMathZoneInsertion=   0   'False
      AllowTableInsertion=   -1  'True
      AlwaysShowScrollBars=   0   'False
      AlwaysShowSelection=   0   'False
      Appearance      =   1
      AutoDetectURLs  =   3
      AutoScrolling   =   3
      AutoSelectWordOnTrackSelection=   0   'False
      BackColor       =   -2147483643
      BeepOnMaxText   =   0   'False
      BorderStyle     =   0
      CharacterConversion=   0
      DefaultMathZoneHAlignment=   1
      DefaultTabWidth =   36
      DenoteEmptyMathArguments=   0
      DetectDragDrop  =   0   'False
      DisabledEvents  =   0
      DisableIMEOperations=   0   'False
      DisplayHyperlinkTooltips=   -1  'True
      DisplayZeroWidthTableCellBorders=   -1  'True
      DraftMode       =   0   'False
      DragScrollTimeBase=   -1
      DropWordsOnWordBoundariesOnly=   -1  'True
      EmulateSimpleTextBox=   0   'False
      Enabled         =   -1  'True
      ExtendFontBackColorToWholeLine=   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FormattingRectangleHeight=   0
      FormattingRectangleLeft=   0
      FormattingRectangleTop=   0
      FormattingRectangleWidth=   0
      GrowNAryOperators=   -1  'True
      HoverTime       =   -1
      HyphenationWordWrapZoneWidth=   180
      IMEConversionMode=   0
      IMEMode         =   -1
      IntegralLimitsLocation=   0
      LeftMargin      =   -1
      LetClientHandleAllIMEOperations=   0   'False
      LinkMousePointer=   0
      LogicalCaret    =   0   'False
      MathLineBreakBehavior=   0
      MaxTextLength   =   -1
      MaxUndoQueueSize=   100
      Modified        =   0   'False
      MousePointer    =   0
      MultiLine       =   -1  'True
      MultiSelect     =   0   'False
      NAryLimitsLocation=   1
      OLEDragImageStyle=   0
      ProcessContextMenuKeys=   -1  'True
      RawSubScriptAndSuperScriptOperators=   -1  'True
      ReadOnly        =   0   'False
      RegisterForOLEDragDrop=   1
      RichEditVersion =   2
      RightMargin     =   -1
      RightToLeft     =   0
      RightToLeftMathZones=   -1  'True
      ScrollBars      =   3
      ScrollToTopOnKillFocus=   0   'False
      ShowSelectionBar=   0   'False
      SmartSpacingOnDrop=   -1  'True
      SupportOLEDragImages=   -1  'True
      TextFlow        =   0
      UseBkAcetateColorForTextSelection=   -1  'True
      UseBuiltInSpellChecking=   0   'False
      UseCustomFormattingRectangle=   0   'False
      UseSmallerFontForNestedFractions=   0   'False
      UseWindowsThemes=   -1  'True
      ZoomRatio       =   0
      AutoDetectedURLSchemes=   "frmMain.frx":0000
      Text            =   "frmMain.frx":0020
   End
   Begin VB.TextBox txtLog 
      Height          =   5655
      Left            =   5400
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Beides
      TabIndex        =   1
      Top             =   120
      Width           =   4815
   End
   Begin VB.CheckBox chkLog 
      Caption         =   "&Log"
      Height          =   255
      Left            =   5520
      TabIndex        =   2
      Top             =   6000
      Width           =   975
   End
   Begin VB.CommandButton cmdAbout 
      Caption         =   "&About..."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6630
      TabIndex        =   3
      Top             =   5880
      Width           =   2415
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

  Private Const CF_OEMTEXT = 7
  Private Const CF_UNICODETEXT = 13

  Private Const MAX_PATH = 260


  Private bLog As Boolean


  Private Declare Sub InitCommonControls Lib "comctl32.dll" ()


Private Sub chkLog_Click()
  bLog = (chkLog.Value = CheckBoxConstants.vbChecked)
End Sub

Private Sub cmdAbout_Click()
  RichTxtBoxU.About
End Sub

Private Sub Form_Click()
  '
End Sub

Private Sub Form_Initialize()
  InitCommonControls
End Sub

Private Sub Form_Load()
  chkLog.Value = CheckBoxConstants.vbChecked
End Sub

Private Sub Form_Resize()
  Dim cx As Long

  If Me.WindowState <> vbMinimized Then
    cx = 0.45 * Me.ScaleWidth
    txtLog.Move Me.ScaleWidth - cx, 0, cx, Me.ScaleHeight - cmdAbout.Height - 10
    cmdAbout.Move txtLog.Left + (cx / 2) - cmdAbout.Width / 2, txtLog.Top + txtLog.Height + 5
    chkLog.Move txtLog.Left, cmdAbout.Top + 5
    RichTxtBoxU.Move 0, 0, txtLog.Left - 5, Me.ScaleHeight
  End If
End Sub

Private Sub RichTxtBoxU_BeginDrag(ByVal TextRange As RichTextBoxCtlLibUCtl.IRichTextRange, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As RichTextBoxCtlLibUCtl.HitTestConstants)
  AddLogEntry "RichTxtBoxU_BeginDrag: TextRange=(" & TextRange.RangeStart & "-" & TextRange.RangeEnd & "), button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)

  RichTxtBoxU.OLEDrag , , , TextRange
End Sub

Private Sub RichTxtBoxU_BeginRDrag(ByVal TextRange As RichTextBoxCtlLibUCtl.IRichTextRange, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As RichTextBoxCtlLibUCtl.HitTestConstants)
  AddLogEntry "RichTxtBoxU_BeginRDrag: TextRange=(" & TextRange.RangeStart & "-" & TextRange.RangeEnd & "), button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)
End Sub

Private Sub RichTxtBoxU_Click(ByVal TextRange As RichTextBoxCtlLibUCtl.IRichTextRange, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As RichTextBoxCtlLibUCtl.HitTestConstants)
  AddLogEntry "RichTxtBoxU_Click: TextRange=(" & TextRange.RangeStart & "-" & TextRange.RangeEnd & "), button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)
End Sub

Private Sub RichTxtBoxU_ContextMenu(ByVal menuType As RichTextBoxCtlLibUCtl.MenuTypeConstants, ByVal TextRange As RichTextBoxCtlLibUCtl.IRichTextRange, ByVal SelectionType As RichTextBoxCtlLibUCtl.SelectionTypeConstants, ByVal OLEObject As RichTextBoxCtlLibUCtl.IRichOLEObject, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As RichTextBoxCtlLibUCtl.HitTestConstants, ByVal isRightMouseDrop As Boolean, ByVal draggedData As RichTextBoxCtlLibUCtl.IOLEDataObject, hMenuToDisplay As stdole.OLE_HANDLE)
  Dim files() As String
  Dim str As String

  str = "RichTxtBoxU_ContextMenu: menuType=0x" & Hex(menuType) & ", TextRange=(" & TextRange.RangeStart & "-" & TextRange.RangeEnd & "), selectionType=0x" & Hex(SelectionType) & ", OLEObject="
  If OLEObject Is Nothing Then
    str = str & "Nothing"
  Else
    str = str & OLEObject.ProgID
  End If
  str = str & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails) & ", isRightMouseDrop=" & isRightMouseDrop & ", draggedData="
  If draggedData Is Nothing Then
    str = str & "Nothing"
  Else
    On Error Resume Next
    files = draggedData.GetData(vbCFFiles)
    If Err Then
      str = str & "0 files"
    Else
      str = str & (UBound(files) - LBound(files) + 1) & " files"
    End If
  End If
  str = str & ", hMenuToDisplay=0x" & Hex(hMenuToDisplay)

  AddLogEntry str
End Sub

Private Sub RichTxtBoxU_DblClick(ByVal TextRange As RichTextBoxCtlLibUCtl.IRichTextRange, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As RichTextBoxCtlLibUCtl.HitTestConstants)
  AddLogEntry "RichTxtBoxU_DblClick: TextRange=(" & TextRange.RangeStart & "-" & TextRange.RangeEnd & "), button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)
End Sub

Private Sub RichTxtBoxU_DestroyedControlWindow(ByVal hWnd As Long)
  AddLogEntry "RichTxtBoxU_DestroyedControlWindow: hWnd=0x" & Hex(hWnd)
End Sub

Private Sub RichTxtBoxU_DragDrop(Source As Control, x As Single, y As Single)
  AddLogEntry "RichTxtBoxU_DragDrop"
End Sub

Private Sub RichTxtBoxU_DragDropDone()
  AddLogEntry "RichTxtBoxU_DragDropDone"
End Sub

Private Sub RichTxtBoxU_DragOver(Source As Control, x As Single, y As Single, State As Integer)
  AddLogEntry "RichTxtBoxU_DragOver"
End Sub

Private Sub RichTxtBoxU_GetDataObject(ByVal TextRange As RichTextBoxCtlLibUCtl.IRichTextRange, ByVal operationType As RichTextBoxCtlLibUCtl.ClipboardOperationTypeConstants, pDataObject As Long, useCustomDataObject As Boolean)
  AddLogEntry "RichTxtBoxU_GetDataObject: TextRange=(" & TextRange.RangeStart & "-" & TextRange.RangeEnd & "), operationType=" & operationType & ", pDataObject=0x" & Hex(pDataObject) & ", useCustomDataObject=" & useCustomDataObject
End Sub

Private Sub RichTxtBoxU_GetDragDropEffect(ByVal getSourceEffects As Boolean, ByVal button As Integer, ByVal shift As Integer, effect As RichTextBoxCtlLibUCtl.OLEDropEffectConstants, skipDefaultProcessing As Boolean)
  AddLogEntry "RichTxtBoxU_GetDragDropEffect: getSourceEffects=" & getSourceEffects & ", button=" & button & ", shift=" & shift & ", effect=" & effect & ", skipDefaultProcessing=" & skipDefaultProcessing
End Sub

Private Sub RichTxtBoxU_GotFocus()
  AddLogEntry "RichTxtBoxU_GotFocus"
End Sub

Private Sub RichTxtBoxU_InsertingOLEObject(ByVal TextRange As RichTextBoxCtlLibUCtl.IRichTextRange, ClassID As String, ByVal pStorage As Long, cancelInsertion As Boolean)
  AddLogEntry "RichTxtBoxU_InsertingOLEObject: TextRange=(" & TextRange.RangeStart & "-" & TextRange.RangeEnd & ") ClassID=" & ClassID & ", pStorage=0x" & Hex(pStorage) & ", cancelInsertion=" & cancelInsertion
End Sub

Private Sub RichTxtBoxU_KeyDown(keyCode As Integer, ByVal shift As Integer)
  AddLogEntry "RichTxtBoxU_KeyDown: keyCode=" & keyCode & ", shift=" & shift
End Sub

Private Sub RichTxtBoxU_KeyPress(keyAscii As Integer)
  AddLogEntry "RichTxtBoxU_KeyPress: keyAscii=" & keyAscii
End Sub

Private Sub RichTxtBoxU_KeyUp(keyCode As Integer, ByVal shift As Integer)
  AddLogEntry "RichTxtBoxU_KeyUp: keyCode=" & keyCode & ", shift=" & shift
End Sub

Private Sub RichTxtBoxU_LostFocus()
  AddLogEntry "RichTxtBoxU_LostFocus"
End Sub

Private Sub RichTxtBoxU_MClick(ByVal TextRange As RichTextBoxCtlLibUCtl.IRichTextRange, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As RichTextBoxCtlLibUCtl.HitTestConstants)
  AddLogEntry "RichTxtBoxU_MClick: TextRange=(" & TextRange.RangeStart & "-" & TextRange.RangeEnd & "), button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)
End Sub

Private Sub RichTxtBoxU_MDblClick(ByVal TextRange As RichTextBoxCtlLibUCtl.IRichTextRange, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As RichTextBoxCtlLibUCtl.HitTestConstants)
  AddLogEntry "RichTxtBoxU_MDblClick: TextRange=(" & TextRange.RangeStart & "-" & TextRange.RangeEnd & "), button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)
End Sub

Private Sub RichTxtBoxU_MouseDown(ByVal TextRange As RichTextBoxCtlLibUCtl.IRichTextRange, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As RichTextBoxCtlLibUCtl.HitTestConstants)
  AddLogEntry "RichTxtBoxU_MouseDown: TextRange=(" & TextRange.RangeStart & "-" & TextRange.RangeEnd & "), button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)
End Sub

Private Sub RichTxtBoxU_MouseEnter(ByVal TextRange As RichTextBoxCtlLibUCtl.IRichTextRange, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As RichTextBoxCtlLibUCtl.HitTestConstants)
  AddLogEntry "RichTxtBoxU_MouseEnter: TextRange=(" & TextRange.RangeStart & "-" & TextRange.RangeEnd & "), button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)
End Sub

Private Sub RichTxtBoxU_MouseHover(ByVal TextRange As RichTextBoxCtlLibUCtl.IRichTextRange, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As RichTextBoxCtlLibUCtl.HitTestConstants)
  AddLogEntry "RichTxtBoxU_MouseHover: TextRange=(" & TextRange.RangeStart & "-" & TextRange.RangeEnd & "), button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)
End Sub

Private Sub RichTxtBoxU_MouseLeave(ByVal TextRange As RichTextBoxCtlLibUCtl.IRichTextRange, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As RichTextBoxCtlLibUCtl.HitTestConstants)
  AddLogEntry "RichTxtBoxU_MouseLeave: TextRange=(" & TextRange.RangeStart & "-" & TextRange.RangeEnd & "), button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)
End Sub

Private Sub RichTxtBoxU_MouseMove(ByVal TextRange As RichTextBoxCtlLibUCtl.IRichTextRange, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As RichTextBoxCtlLibUCtl.HitTestConstants)
  'AddLogEntry "RichTxtBoxU_MouseMove: TextRange=(" & TextRange.RangeStart & "-" & TextRange.RangeEnd & "), button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)
End Sub

Private Sub RichTxtBoxU_MouseUp(ByVal TextRange As RichTextBoxCtlLibUCtl.IRichTextRange, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As RichTextBoxCtlLibUCtl.HitTestConstants)
  AddLogEntry "RichTxtBoxU_MouseUp: TextRange=(" & TextRange.RangeStart & "-" & TextRange.RangeEnd & "), button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)
End Sub

Private Sub RichTxtBoxU_MouseWheel(ByVal TextRange As RichTextBoxCtlLibUCtl.IRichTextRange, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As RichTextBoxCtlLibUCtl.HitTestConstants, ByVal scrollAxis As RichTextBoxCtlLibUCtl.ScrollAxisConstants, ByVal wheelDelta As Integer)
  AddLogEntry "RichTxtBoxU_MouseWheel: TextRange=(" & TextRange.RangeStart & "-" & TextRange.RangeEnd & "), button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails) & ", scrollAxis=" & scrollAxis & ", wheelDelta=" & wheelDelta
End Sub

Private Sub RichTxtBoxU_OLECompleteDrag(ByVal data As RichTextBoxCtlLibUCtl.IOLEDataObject, ByVal performedEffect As RichTextBoxCtlLibUCtl.OLEDropEffectConstants)
  Dim files() As String
  Dim str As String

  str = "RichTxtBoxU_OLECompleteDrag: data="
  If data Is Nothing Then
    str = str & "Nothing"
  Else
    On Error Resume Next
    files = data.GetData(vbCFFiles)
    If Err Then
      str = str & "0 files"
    Else
      str = str & (UBound(files) - LBound(files) + 1) & " files"
    End If
  End If
  str = str & ", performedEffect=" & performedEffect

  AddLogEntry str

  If data.GetFormat(vbCFFiles) Then
    files = data.GetData(vbCFFiles)
    str = Join(files, vbNewLine)
    MsgBox "Dragged files:" & vbNewLine & str
  End If
End Sub

Private Sub RichTxtBoxU_OLEDragDrop(ByVal data As RichTextBoxCtlLibUCtl.IOLEDataObject, effect As RichTextBoxCtlLibUCtl.OLEDropEffectConstants, ByVal dropTarget As RichTextBoxCtlLibUCtl.IRichTextRange, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As RichTextBoxCtlLibUCtl.HitTestConstants)
  Dim files() As String
  Dim str As String

  str = "RichTxtBoxU_OLEDragDrop: data="
  If data Is Nothing Then
    str = str & "Nothing"
  Else
    On Error Resume Next
    files = data.GetData(vbCFFiles)
    If Err Then
      str = str & "0 files"
    Else
      str = str & (UBound(files) - LBound(files) + 1) & " files"
    End If
  End If
  str = str & ", effect=" & effect & ", dropTarget=(" & dropTarget.RangeStart & "-" & dropTarget.RangeEnd & "), button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)

  AddLogEntry str

  If data.GetFormat(vbCFFiles) Then
    files = data.GetData(vbCFFiles)
    str = Join(files, vbNewLine)
    RichTxtBoxU.FinishOLEDragDrop
    MsgBox "Dropped files:" & vbNewLine & str
  End If
End Sub

Private Sub RichTxtBoxU_OLEDragEnter(ByVal data As RichTextBoxCtlLibUCtl.IOLEDataObject, effect As RichTextBoxCtlLibUCtl.OLEDropEffectConstants, ByVal dropTarget As RichTextBoxCtlLibUCtl.IRichTextRange, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As RichTextBoxCtlLibUCtl.HitTestConstants, autoHScrollVelocity As Long, autoVScrollVelocity As Long)
  Dim files() As String
  Dim str As String

  str = "RichTxtBoxU_OLEDragEnter: data="
  If data Is Nothing Then
    str = str & "Nothing"
  Else
    On Error Resume Next
    files = data.GetData(vbCFFiles)
    If Err Then
      str = str & "0 files"
    Else
      str = str & (UBound(files) - LBound(files) + 1) & " files"
    End If
  End If
  str = str & ", effect=" & effect & ", dropTarget=(" & dropTarget.RangeStart & "-" & dropTarget.RangeEnd & "), button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails) & ", autoHScrollVelocity=" & autoHScrollVelocity & ", autoVScrollVelocity=" & autoVScrollVelocity

  AddLogEntry str
End Sub

Private Sub RichTxtBoxU_OLEDragEnterPotentialTarget(ByVal hWndPotentialTarget As Long)
  AddLogEntry "RichTxtBoxU_OLEDragEnterPotentialTarget: hWndPotentialTarget=0x" & Hex(hWndPotentialTarget)
End Sub

Private Sub RichTxtBoxU_OLEDragLeave(ByVal data As RichTextBoxCtlLibUCtl.IOLEDataObject, ByVal dropTarget As RichTextBoxCtlLibUCtl.IRichTextRange, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As RichTextBoxCtlLibUCtl.HitTestConstants)
  Dim files() As String
  Dim str As String

  str = "RichTxtBoxU_OLEDragLeave: data="
  If data Is Nothing Then
    str = str & "Nothing"
  Else
    On Error Resume Next
    files = data.GetData(vbCFFiles)
    If Err Then
      str = str & "0 files"
    Else
      str = str & (UBound(files) - LBound(files) + 1) & " files"
    End If
  End If
  str = str & ", dropTarget=(" & dropTarget.RangeStart & "-" & dropTarget.RangeEnd & "), button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)

  AddLogEntry str
End Sub

Private Sub RichTxtBoxU_OLEDragLeavePotentialTarget()
  AddLogEntry "RichTxtBoxU_OLEDragLeavePotentialTarget"
End Sub

Private Sub RichTxtBoxU_OLEDragMouseMove(ByVal data As RichTextBoxCtlLibUCtl.IOLEDataObject, effect As RichTextBoxCtlLibUCtl.OLEDropEffectConstants, ByVal dropTarget As RichTextBoxCtlLibUCtl.IRichTextRange, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As RichTextBoxCtlLibUCtl.HitTestConstants, autoHScrollVelocity As Long, autoVScrollVelocity As Long)
  Dim files() As String
  Dim str As String

  str = "RichTxtBoxU_OLEDragMouseMove: data="
  If data Is Nothing Then
    str = str & "Nothing"
  Else
    On Error Resume Next
    files = data.GetData(vbCFFiles)
    If Err Then
      str = str & "0 files"
    Else
      str = str & (UBound(files) - LBound(files) + 1) & " files"
    End If
  End If
  str = str & ", effect=" & effect & ", dropTarget=(" & dropTarget.RangeStart & "-" & dropTarget.RangeEnd & "), button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails) & ", autoHScrollVelocity=" & autoHScrollVelocity & ", autoVScrollVelocity=" & autoVScrollVelocity

  AddLogEntry str
End Sub

Private Sub RichTxtBoxU_OLEGiveFeedback(ByVal effect As RichTextBoxCtlLibUCtl.OLEDropEffectConstants, useDefaultCursors As Boolean)
  AddLogEntry "RichTxtBoxU_OLEGiveFeedback: effect=" & effect & ", useDefaultCursors=" & useDefaultCursors
End Sub

Private Sub RichTxtBoxU_OLEQueryContinueDrag(ByVal pressedEscape As Boolean, ByVal button As Integer, ByVal shift As Integer, actionToContinueWith As RichTextBoxCtlLibUCtl.OLEActionToContinueWithConstants)
  AddLogEntry "RichTxtBoxU_OLEQueryContinueDrag: pressedEscape=" & pressedEscape & ", button=" & button & ", shift=" & shift & ", actionToContinueWith=0x" & Hex(actionToContinueWith)
End Sub

Private Sub RichTxtBoxU_OLEReceivedNewData(ByVal data As RichTextBoxCtlLibUCtl.IOLEDataObject, ByVal formatID As Long, ByVal Index As Long, ByVal dataOrViewAspect As Long)
  Dim files() As String
  Dim str As String

  str = "RichTxtBoxU_OLEReceivedNewData: data="
  If data Is Nothing Then
    str = str & "Nothing"
  Else
    On Error Resume Next
    files = data.GetData(vbCFFiles)
    If Err Then
      str = str & "0 files"
    Else
      str = str & (UBound(files) - LBound(files) + 1) & " files"
    End If
  End If
  str = str & ", formatID=" & formatID & ", Index=" & Index & ", dataOrViewAspect=" & dataOrViewAspect

  AddLogEntry str
End Sub

Private Sub RichTxtBoxU_OLESetData(ByVal data As RichTextBoxCtlLibUCtl.IOLEDataObject, ByVal formatID As Long, ByVal Index As Long, ByVal dataOrViewAspect As Long)
  Dim files() As String
  Dim str As String

  str = "RichTxtBoxU_OLESetData: data="
  If data Is Nothing Then
    str = str & "Nothing"
  Else
    On Error Resume Next
    files = data.GetData(vbCFFiles)
    If Err Then
      str = str & "0 files"
    Else
      str = str & (UBound(files) - LBound(files) + 1) & " files"
    End If
  End If
  str = str & ", formatID=" & formatID & ", Index=" & Index & ", dataOrViewAspect=" & dataOrViewAspect

  AddLogEntry str
End Sub

Private Sub RichTxtBoxU_OLEStartDrag(ByVal data As RichTextBoxCtlLibUCtl.IOLEDataObject)
  Dim files() As String
  Dim str As String

  str = "RichTxtBoxU_OLEStartDrag: data="
  If data Is Nothing Then
    str = str & "Nothing"
  Else
    On Error Resume Next
    files = data.GetData(vbCFFiles)
    If Err Then
      str = str & "0 files"
    Else
      str = str & (UBound(files) - LBound(files) + 1) & " files"
    End If
  End If

  AddLogEntry str

  data.SetData vbCFText
  data.SetData CF_OEMTEXT
  data.SetData CF_UNICODETEXT
End Sub

Private Sub RichTxtBoxU_QueryAcceptData(ByVal data As RichTextBoxCtlLibUCtl.IOLEDataObject, formatID As Long, ByVal operationType As RichTextBoxCtlLibUCtl.ClipboardOperationTypeConstants, ByVal performOperation As Boolean, ByVal hMetafilePicture As stdole.OLE_HANDLE, acceptData As RichTextBoxCtlLibUCtl.QueryAcceptDataConstants)
  Dim files() As String
  Dim str As String

  str = "RichTxtBoxU_QueryAcceptData: data="
  If data Is Nothing Then
    str = str & "Nothing"
  Else
    On Error Resume Next
    files = data.GetData(vbCFFiles)
    If Err Then
      str = str & "0 files"
    Else
      str = str & (UBound(files) - LBound(files) + 1) & " files"
    End If
  End If
  str = str & ", formatID=" & formatID & ", operationType=" & operationType & ", performOperation=" & performOperation & ", hMetafilePicture=0x" & Hex(hMetafilePicture) & ", acceptData=" & acceptData

  AddLogEntry str
End Sub

Private Sub RichTxtBoxU_RClick(ByVal TextRange As RichTextBoxCtlLibUCtl.IRichTextRange, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As RichTextBoxCtlLibUCtl.HitTestConstants)
  AddLogEntry "RichTxtBoxU_RClick: TextRange=(" & TextRange.RangeStart & "-" & TextRange.RangeEnd & "), button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)
End Sub

Private Sub RichTxtBoxU_RDblClick(ByVal TextRange As RichTextBoxCtlLibUCtl.IRichTextRange, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As RichTextBoxCtlLibUCtl.HitTestConstants)
  AddLogEntry "RichTxtBoxU_RDblClick: TextRange=(" & TextRange.RangeStart & "-" & TextRange.RangeEnd & "), button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)
End Sub

Private Sub RichTxtBoxU_RecreatedControlWindow(ByVal hWnd As Long)
  AddLogEntry "RichTxtBoxU_RecreatedControlWindow: hWnd=0x" & Hex(hWnd)
End Sub

Private Sub RichTxtBoxU_RemovingOLEObject(ByVal OLEObject As RichTextBoxCtlLibUCtl.IRichOLEObject, reserved As Boolean)
  Dim str As String

  str = "RichTxtBoxU_RemovingOLEObject: OLEObject="
  If OLEObject Is Nothing Then
    str = str & "Nothing"
  Else
    str = str & OLEObject.ProgId
  End If
  str = str & ", reserved=" & reserved

  AddLogEntry str
End Sub

Private Sub RichTxtBoxU_ResizedControlWindow()
  AddLogEntry "RichTxtBoxU_ResizedControlWindow"
End Sub

Private Sub RichTxtBoxU_Scrolling(ByVal axis As RichTextBoxCtlLibUCtl.ScrollAxisConstants)
  AddLogEntry "RichTxtBoxU_Scrolling: axis=0x" & Hex(axis)
End Sub

Private Sub RichTxtBoxU_SelectionChanged(ByVal SelectedTextRange As RichTextBoxCtlLibUCtl.IRichTextRange, ByVal SelectionType As RichTextBoxCtlLibUCtl.SelectionTypeConstants)
  AddLogEntry "RichTxtBoxU_SelectionChanged: SelectedTextRange=(" & SelectedTextRange.RangeStart & "-" & SelectedTextRange.RangeEnd & "), selectionType=0x" & Hex(SelectionType)
End Sub

Private Sub RichTxtBoxU_ShouldResizeControlWindow(suggestedControlSize As RichTextBoxCtlLibUCtl.RECTANGLE)
  AddLogEntry "RichTxtBoxU_ShouldResizeControlWindow: suggestedControlSize=(" & suggestedControlSize.Left & "," & suggestedControlSize.Top & ")-(" & suggestedControlSize.Right & "," & suggestedControlSize.Bottom & ")"
End Sub

Private Sub RichTxtBoxU_TextChanged()
  AddLogEntry "RichTxtBoxU_TextChanged"
End Sub

Private Sub RichTxtBoxU_TruncatedText()
  AddLogEntry "RichTxtBoxU_TruncatedText"
End Sub

Private Sub RichTxtBoxU_Validate(Cancel As Boolean)
  AddLogEntry "RichTxtBoxU_Validate"
End Sub

Private Sub RichTxtBoxU_XClick(ByVal TextRange As RichTextBoxCtlLibUCtl.IRichTextRange, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As RichTextBoxCtlLibUCtl.HitTestConstants)
  AddLogEntry "RichTxtBoxU_XClick: TextRange=(" & TextRange.RangeStart & "-" & TextRange.RangeEnd & "), button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)
End Sub

Private Sub RichTxtBoxU_XDblClick(ByVal TextRange As RichTextBoxCtlLibUCtl.IRichTextRange, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As RichTextBoxCtlLibUCtl.HitTestConstants)
  AddLogEntry "RichTxtBoxU_XDblClick: TextRange=(" & TextRange.RangeStart & "-" & TextRange.RangeEnd & "), button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)
End Sub


Private Sub AddLogEntry(ByVal txt As String)
  Dim pos As Long
  Static cLines As Long
  Static oldtxt As String

  If bLog Then
    cLines = cLines + 1
    If cLines > 50 Then
      ' delete the first line
      pos = InStr(oldtxt, vbNewLine)
      oldtxt = Mid(oldtxt, pos + Len(vbNewLine))
      cLines = cLines - 1
    End If
    oldtxt = oldtxt & (txt & vbNewLine)

    txtLog.Text = oldtxt
    txtLog.SelStart = Len(oldtxt)
  End If
End Sub
