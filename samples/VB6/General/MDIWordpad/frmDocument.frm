VERSION 5.00
Object = "{CD7AC637-45AE-4969-81B8-A56464ECD6C4}#1.0#0"; "RichTextBoxCtlU.ocx"
Begin VB.Form frmDocument 
   Caption         =   "Document 1"
   ClientHeight    =   3645
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   6585
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   243
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   439
   Begin RichTextBoxCtlLibUCtl.RichTextBox RTB 
      Height          =   1455
      Left            =   1440
      TabIndex        =   0
      Top             =   960
      Width           =   3135
      _cx             =   5530
      _cy             =   2566
      AcceptTabKey    =   -1  'True
      AdjustLineHeightForEastAsianLanguages=   -1  'True
      AllowInputThroughTSF=   -1  'True
      AllowMathZoneInsertion=   -1  'True
      AllowObjectInsertionThroughTSF=   -1  'True
      AllowTableInsertion=   -1  'True
      AllowTSFProofingTips=   -1  'True
      AllowTSFSmartTags=   -1  'True
      AlwaysShowScrollBars=   0   'False
      AlwaysShowSelection=   -1  'True
      Appearance      =   0
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
      DisabledEvents  =   2051
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
         Name            =   "Calibri"
         Size            =   11.25
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
      MultiSelect     =   -1  'True
      NAryLimitsLocation=   1
      NoInputSequenceCheck=   0   'False
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
      TextOrientation =   0
      TSFModeBias     =   0
      UseBkAcetateColorForTextSelection=   -1  'True
      UseBuiltInSpellChecking=   -1  'True
      UseCustomFormattingRectangle=   0   'False
      UseSmallerFontForNestedFractions=   0   'False
      UseTextServicesFramework=   -1  'True
      UseWindowsThemes=   -1  'True
      ZoomRatio       =   0
      AutoDetectedURLSchemes=   "frmDocument.frx":0000
      Text            =   "frmDocument.frx":0020
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   300
      Left            =   1440
      Top             =   2520
   End
End
Attribute VB_Name = "frmDocument"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

  Implements IDocument


  Private Type POINT
    x As Long
    y As Long
  End Type

  Private Type MENUITEMINFO
    cbSize As Long
    fMask As Long
    fType As Long
    fState As Long
    wID As Long
    hSubMenu As Long
    hbmpChecked As Long
    hbmpUnchecked As Long
    dwItemData As Long
    dwTypeData As Long
    cch As Long
    hBitmap As Long
  End Type


  Private activeCellBackColor As Long
  Private activeTextBackColor As Long
  Private activeTextForeColor As Long
  Private documentFileName As String
  Private tempFont As TextFont
  Private windowID As Long


  Public mdiFrame As IMDIFrame


  Private Declare Function ClientToScreen Lib "user32.dll" (ByVal hWnd As Long, lpPoint As POINT) As Long
  Private Declare Function CreatePopupMenu Lib "user32.dll" () As Long
  Private Declare Function DestroyMenu Lib "user32.dll" (ByVal hMenu As Long) As Long
  Private Declare Function InsertMenuItem Lib "user32.dll" Alias "InsertMenuItemW" (ByVal hMenu As Long, ByVal uItem As Long, ByVal fByPosition As Long, lpmii As MENUITEMINFO) As Long
  Private Declare Sub PathStripPath Lib "shlwapi.dll" Alias "PathStripPathW" (ByVal pszPath As Long)
  Private Declare Function TrackPopupMenuEx Lib "user32.dll" (ByVal hMenu As Long, ByVal fuFlags As Long, ByVal x As Long, ByVal y As Long, ByVal hWnd As Long, ByVal lptpm As Long) As Long


Private Property Get IDocument_CellBackColor() As Long
  IDocument_CellBackColor = activeCellBackColor
End Property

Private Property Let IDocument_CellBackColor(ByVal RHS As Long)
  activeCellBackColor = RHS
  IDocument_ExecuteCommand CMDID_TABLE_CELLBACKCOLOR
End Property

Private Property Get IDocument_DocumentTextRange() As RichTextBoxCtlLibUCtl.IRichTextRange
  Set IDocument_DocumentTextRange = RTB.TextRange
End Property

Private Sub IDocument_EndFontChangePreview(ByVal apply As Boolean)
  'Call RTB.Undo(-9999994)
  'If apply And Not tempFont Is Nothing Then
  '  RTB.SelectedTextRange.Font.CopySettings tempFont
  'End If
  Set tempFont = Nothing
End Sub

Private Sub IDocument_ExecuteCommand(ByVal commandID As Long, Optional ByVal params As Variant)
  Dim tbl As Table
  Dim tblRow As TableRow
  Dim tblCell As TableCell
  Dim i As Long
  Dim bIsFirstHorizontal As Boolean
  Dim bIsFirstVertical As Boolean
  Dim itemsToDelete As Variant
  Dim para As TextParagraph
  Dim commonDlg As clsCommonDialog
  Dim fileName As String
  Dim obj As OLEObject
  Dim str As String

  Select Case commandID
    Case CMDID_HELP_ABOUT
      RTB.About
    Case CMDID_FILE_PRINT
      MsgBox "Command not implemented"
      Me.SetFocus
    Case CMDID_EDIT_UNDO
      RTB.Undo
    Case CMDID_EDIT_REDO
      RTB.Redo
    Case CMDID_EDIT_CUT
      If RTB.SelectedTextRange.Cut Then
        If Not (mdiFrame Is Nothing) Then
          mdiFrame.EnableDisableCommand CMDID_EDIT_CUT, False
          mdiFrame.EnableDisableCommand CMDID_EDIT_COPY, False
          mdiFrame.EnableDisableCommand CMDID_EDIT_PASTE, RTB.SelectedTextRange.CanPaste
        End If
      End If
    Case CMDID_EDIT_COPY
      If RTB.SelectedTextRange.Copy Then
        If Not (mdiFrame Is Nothing) Then
          mdiFrame.EnableDisableCommand CMDID_EDIT_PASTE, RTB.SelectedTextRange.CanPaste
        End If
      End If
    Case CMDID_EDIT_PASTE
      Call RTB.SelectedTextRange.Paste
    Case CMDID_EDIT_FIND
      Call frmFind.ShowForm(mdiFrame, Me)
    Case CMDID_EDIT_SELECTALL
      RTB.TextRange.Select
    Case CMDID_EDIT_SPELLCHECKING
      RTB.UseBuiltInSpellChecking = Not RTB.UseBuiltInSpellChecking
      
    Case CMDID_FORMAT_ALIGNLEFT
      With RTB.SelectedTextRange.Paragraph
        .HAlignment = halLeft
      End With
    Case CMDID_FORMAT_ALIGNCENTER
      With RTB.SelectedTextRange.Paragraph
        .HAlignment = halCenter
      End With
    Case CMDID_FORMAT_ALIGNRIGHT
      With RTB.SelectedTextRange.Paragraph
        .HAlignment = halRight
      End With
    Case CMDID_FORMAT_ALIGNJUSTIFY
      With RTB.SelectedTextRange.Paragraph
        .HAlignment = halJustify
      End With
      
    Case CMDID_FORMAT_BULLETLIST
      With RTB.SelectedTextRange.Paragraph
        Set para = .Clone
        If (para.ListNumberingStyle And lnsTypeMask) = lnsBullet Then
          ' turn off list mode
          para.ListLevelIndex = 0
          para.ListNumberingStyle = lnsNoList
          para.LeftIndent = 0
        ElseIf para.ListLevelIndex = 0 Then
          ' activate list mode
          para.ListLevelIndex = 1
          para.ListNumberingTabStop = 15
          para.ListNumberingStyle = lnsBullet
          para.LeftIndent = 15
        Else
          ' just switch list numbering type
          para.ListNumberingTabStop = 15
          para.ListNumberingStyle = lnsBullet
          mdiFrame.CheckCommand CMDID_FORMAT_NUMBEREDLIST, False
        End If
        Call .CopySettings(para)
        mdiFrame.EnableDisableCommand CMDID_FORMAT_INCREASEINDENT, .ListLevelIndex > 0
        mdiFrame.EnableDisableCommand CMDID_FORMAT_DECREASEINDENT, .ListLevelIndex > 0
      End With
    Case CMDID_FORMAT_NUMBEREDLIST
      With RTB.SelectedTextRange.Paragraph
        Set para = .Clone
        If (para.ListNumberingStyle And lnsTypeMask) = lnsArabicNumbers Then
          ' turn off list mode
          para.ListLevelIndex = 0
          para.ListNumberingStyle = lnsNoList
          para.LeftIndent = 0
        ElseIf para.ListLevelIndex = 0 Then
          ' activate list mode
          para.ListLevelIndex = 1
          para.ListNumberingStart = 1
          para.ListNumberingTabStop = 15
          para.ListNumberingStyle = lnsArabicNumbers Or lnsAddFollowedByPeriod
          para.LeftIndent = 15
        Else
          ' just switch list numbering type
          para.ListNumberingStart = 1
          para.ListNumberingTabStop = 15
          para.ListNumberingStyle = lnsArabicNumbers Or lnsAddFollowedByPeriod
          mdiFrame.CheckCommand CMDID_FORMAT_BULLETLIST, False
        End If
        ' Note that for ListNumberingHAlignment<>halLeft a LeftIndent has to be specified.
        Call .CopySettings(para)
        mdiFrame.EnableDisableCommand CMDID_FORMAT_INCREASEINDENT, .ListLevelIndex > 0
        mdiFrame.EnableDisableCommand CMDID_FORMAT_DECREASEINDENT, .ListLevelIndex > 0
      End With
      
    Case CMDID_FORMAT_INCREASEINDENT
      With RTB.SelectedTextRange.Paragraph
        Set para = .Clone
        If para.ListLevelIndex > 0 Then
          para.ListLevelIndex = para.ListLevelIndex + 1
          para.LeftIndent = para.LeftIndent + 15
        End If
        Call .CopySettings(para)
        Call RTB.SelectedTextRange.Select     ' Ensure that the cursor is displayed at the correct position
        mdiFrame.EnableDisableCommand CMDID_FORMAT_INCREASEINDENT, (.ListLevelIndex > 0) And CaretIsAtLineStartOrWholeLineIsSelected
        mdiFrame.EnableDisableCommand CMDID_FORMAT_DECREASEINDENT, (.ListLevelIndex > 0) And CaretIsAtLineStartOrWholeLineIsSelected
      End With
    Case CMDID_FORMAT_DECREASEINDENT
      With RTB.SelectedTextRange.Paragraph
        Set para = .Clone
        If para.ListLevelIndex > 1 Then
          para.ListLevelIndex = para.ListLevelIndex - 1
          para.LeftIndent = para.LeftIndent - 15
        End If
        Call .CopySettings(para)
        Call RTB.SelectedTextRange.Select     ' Ensure that the cursor is displayed at the correct position
        mdiFrame.EnableDisableCommand CMDID_FORMAT_INCREASEINDENT, (.ListLevelIndex > 0) And CaretIsAtLineStartOrWholeLineIsSelected
        mdiFrame.EnableDisableCommand CMDID_FORMAT_DECREASEINDENT, (.ListLevelIndex > 0) And CaretIsAtLineStartOrWholeLineIsSelected
      End With
    Case CMDID_FORMAT_CONVERTTOHYPERLINK
      str = RTB.SelectedTextRange.URL
      If Left$(str, 1) = """" Then str = Mid$(str, 2)
      If Right$(str, 1) = """" Then str = Left$(str, Len(str) - 1)
      str = InputBox("Please insert the URL to link to.", "Convert to hyperlink", str)
      Me.SetFocus
      If StrPtr(str) = 0 Then
        ' cancelled
      Else
        RTB.SelectedTextRange.URL = str
      End If
      
    Case CMDID_FONT_BOLD
      With RTB.SelectedTextRange.Font
        .Bold = IIf(.Bold = bpvTrue, bpvFalse, bpvTrue)
      End With
    Case CMDID_FONT_ITALIC
      With RTB.SelectedTextRange.Font
        .Italic = IIf(.Italic = bpvTrue, bpvFalse, bpvTrue)
      End With
    Case CMDID_FONT_UNDERLINE
      With RTB.SelectedTextRange.Font
        .UnderlineType = IIf(.UnderlineType <= utNone, utSingle, utNone)
      End With
    Case CMDID_FONT_STRIKETHROUGH
      With RTB.SelectedTextRange.Font
        .Strikethrough = IIf(.Strikethrough = bpvTrue, bpvFalse, bpvTrue)
      End With
    Case CMDID_FONT_SUBSCRIPT
      With RTB.SelectedTextRange.Font
        .Subscript = IIf(.Subscript = bpvTrue, bpvFalse, bpvTrue)
      End With
    Case CMDID_FONT_SUPERSCRIPT
      With RTB.SelectedTextRange.Font
        .Superscript = IIf(.Superscript = bpvTrue, bpvFalse, bpvTrue)
      End With
    Case CMDID_FONT_TEXTBACKCOLOR
      With RTB.SelectedTextRange.Font
        .BackColor = activeTextBackColor
      End With
    Case CMDID_FONT_TEXTFORECOLOR
      With RTB.SelectedTextRange.Font
        .ForeColor = activeTextForeColor
      End With
      
    Case CMDID_OBJECT_INSERTIMAGE
      Set commonDlg = New clsCommonDialog
      With commonDlg
        Call .AddFilter("Images", "*.png;*.jpg;*.jpeg;*.bmp;*.gif;*.ico;*.emf;*.wmf")
        Call .AddFilter("All Files", "*.*")
        .ActiveFilter = 0
        .hWndParent = Me.hWnd
        .OpenFlags = OFN_ENABLESIZING Or OFN_FILEMUSTEXIST Or OFN_LONGNAMES Or OFN_PATHMUSTEXIST
        .Caption = "Insert Image"
        If Not .ShowOpen(fileName) Then
          fileName = ""
        End If
      End With
      If fileName <> "" Then
        On Error Resume Next
        Set obj = RTB.EmbeddedObjects.AddImage(fileName)
        If Not obj Is Nothing Then
          obj.TextRange.Select
        Else
          MsgBox "The file could not be inserted as an image."
          Me.SetFocus
        End If
      End If
      
    Case CMDID_OBJECT_SETBKIMAGE
      Set commonDlg = New clsCommonDialog
      With commonDlg
        Call .AddFilter("Images", "*.png;*.jpg;*.jpeg;*.bmp;*.gif;*.ico;*.emf;*.wmf")
        Call .AddFilter("All Files", "*.*")
        .ActiveFilter = 0
        .hWndParent = Me.hWnd
        .OpenFlags = OFN_ENABLESIZING Or OFN_FILEMUSTEXIST Or OFN_LONGNAMES Or OFN_PATHMUSTEXIST
        .Caption = "Choose Background Image"
        If Not .ShowOpen(fileName) Then
          fileName = ""
        End If
      End With
      If fileName <> "" Then
        On Error Resume Next
        Set obj = RTB.EmbeddedObjects.AddImage(fileName, , oofUseAsBackground)
        If Not obj Is Nothing Then
          '
        Else
          MsgBox "The file could not be set as background image."
          Me.SetFocus
        End If
      End If
      
    Case CMDID_OBJECT_INSERTOBJECT
      str = InputBox("Please enter the ProgID or CLSID of the object to insert.", "Insert OLE Object", "WordPad.Document.1")
      Me.SetFocus
      If str <> "" Then
        On Error Resume Next
        Set obj = RTB.EmbeddedObjects.Add(str)
        If Not obj Is Nothing Then
          obj.TextRange.Select
          obj.ExecuteVerb 0
        Else
          MsgBox "The object could not be inserted."
          Me.SetFocus
        End If
      End If
      
    Case CMDID_OBJECT_INSERTOBJECTFROMFILE
      Set commonDlg = New clsCommonDialog
      With commonDlg
        Call .AddFilter("All Files", "*.*")
        .ActiveFilter = 0
        .hWndParent = Me.hWnd
        .OpenFlags = OFN_ENABLESIZING Or OFN_FILEMUSTEXIST Or OFN_LONGNAMES Or OFN_PATHMUSTEXIST
        .Caption = "Insert OLE Object"
        If Not .ShowOpen(fileName) Then
          fileName = ""
        End If
      End With
      If fileName <> "" Then
        On Error Resume Next
        Set obj = RTB.EmbeddedObjects.Add(, fileName)
        If Not obj Is Nothing Then
          obj.TextRange.Select
        Else
          MsgBox "The file could not be inserted as object."
          Me.SetFocus
        End If
      End If
      
    Case CMDID_TABLE_INSERTTABLE
      With RTB.SelectedTextRange
        On Error Resume Next
        Set tbl = .ReplaceWithTable(params(0), params(1))
        If Not (tbl Is Nothing) Then
          tbl.Rows(0).Cells(0).TextRange.Select
        End If
      End With
    Case CMDID_TABLE_DELETETABLE
      If RTB.SelectedTextRange.IsWithinTable(tbl) Then
        tbl.TextRange.Delete
      End If
    Case CMDID_TABLE_INSERTROW
      If RTB.SelectedTextRange.IsWithinTable(tbl, tblRow) Then
        If Not tblRow Is Nothing Then
          tbl.Rows.Add(tblRow.Index + 1).Cells(0).TextRange.Select
        Else
          tbl.Rows.Add().Cells(0).TextRange.Select
        End If
      End If
    Case CMDID_TABLE_DELETEROW
      If RTB.SelectedTextRange.IsWithinTable(tbl, tblRow) Then
        If Not tblRow Is Nothing Then
          ' only one row is selected
          tbl.Rows.Remove tblRow.Index
        Else
          ' multiple rows are selected
          For Each tblRow In tbl.Rows
            For Each tblCell In tblRow.Cells
              If tblCell.TextRange.RangeStart >= RTB.SelectedTextRange.RangeStart Then
                If tblCell.TextRange.RangeStart <= RTB.SelectedTextRange.RangeEnd Then
                  ' this cell is selected
                  If IsEmpty(itemsToDelete) Then
                    ReDim itemsToDelete(1 To 1) As Long
                  Else
                    ReDim Preserve itemsToDelete(LBound(itemsToDelete) To UBound(itemsToDelete) + 1)
                  End If
                  itemsToDelete(UBound(itemsToDelete)) = tblRow.Index
                  Exit For
                End If
              End If
            Next tblCell
          Next tblRow
          For i = UBound(itemsToDelete) To LBound(itemsToDelete) Step -1
            tbl.Rows.Remove itemsToDelete(i)
          Next i
        End If
      End If
    Case CMDID_TABLE_INSERTCOLUMN
      If RTB.SelectedTextRange.IsWithinTable(tbl, , tblCell) Then
        If Not tblCell Is Nothing Then
          i = tblCell.Index
          For Each tblRow In tbl.Rows
            tblRow.Cells.Add 600, i + 1
          Next tblRow
        End If
      End If
    Case CMDID_TABLE_DELETECOLUMN
      If RTB.SelectedTextRange.IsWithinTable(tbl, , tblCell) Then
        If Not tblCell Is Nothing Then
          ' only one cell is selected
          i = tblCell.Index
          For Each tblRow In tbl.Rows
            tblRow.Cells.Remove i
          Next tblRow
        Else
          ' multiple cells (and maybe rows) are selected
          For Each tblRow In tbl.Rows
            If IsEmpty(itemsToDelete) Then
              For Each tblCell In tblRow.Cells
                If tblCell.TextRange.RangeStart >= RTB.SelectedTextRange.RangeStart Then
                  If tblCell.TextRange.RangeStart <= RTB.SelectedTextRange.RangeEnd Then
                    ' this cell is selected
                    If IsEmpty(itemsToDelete) Then
                      ReDim itemsToDelete(1 To 1) As Long
                    Else
                      ReDim Preserve itemsToDelete(LBound(itemsToDelete) To UBound(itemsToDelete) + 1)
                    End If
                    itemsToDelete(UBound(itemsToDelete)) = tblCell.Index
                  End If
                End If
              Next tblCell
            End If
            For i = UBound(itemsToDelete) To LBound(itemsToDelete) Step -1
              If tblRow.Cells.Count > 1 Then
                tblRow.Cells.Remove itemsToDelete(i)
              Else
                tbl.Rows.Remove tblRow.Index
              End If
            Next i
          Next tblRow
        End If
      End If
    Case CMDID_TABLE_MERGECELLS
      If RTB.SelectedTextRange.IsWithinTable(tbl, tblRow, tblCell) Then
        If tblRow Is Nothing Then
          ' multiple rows are selected
          bIsFirstVertical = True
          For Each tblRow In tbl.Rows
            bIsFirstHorizontal = True
            For Each tblCell In tblRow.Cells
              If tblCell.TextRange.RangeStart >= RTB.SelectedTextRange.RangeStart Then
                If tblCell.TextRange.RangeStart <= RTB.SelectedTextRange.RangeEnd Then
                  ' this cell is selected
                  tblCell.CellMergeFlags = IIf(bIsFirstVertical, cmTopCellInVerticallyMergedSet, cmContinueVerticallyMergedSet) Or _
                                           IIf(bIsFirstHorizontal, cmStartCellInHorizontallyMergedSet, cmContinueHorizontallyMergedSet)
                  bIsFirstHorizontal = False
                  bIsFirstVertical = False
                End If
              End If
            Next tblCell
          Next tblRow
        ElseIf tblCell Is Nothing Then
          ' multiple columns are selected
          bIsFirstHorizontal = True
          For Each tblCell In tblRow.Cells
            If tblCell.TextRange.RangeStart >= RTB.SelectedTextRange.RangeStart Then
              If tblCell.TextRange.RangeStart <= RTB.SelectedTextRange.RangeEnd Then
                ' this cell is selected
                tblCell.CellMergeFlags = IIf(bIsFirstHorizontal, cmStartCellInHorizontallyMergedSet, cmContinueHorizontallyMergedSet)
                bIsFirstHorizontal = False
              End If
            End If
          Next tblCell
        End If
      End If
    Case CMDID_TABLE_ALIGNTOP, CMDID_TABLE_ALIGNCENTER, CMDID_TABLE_ALIGNBOTTOM, CMDID_TABLE_CELLBACKCOLOR
      If RTB.SelectedTextRange.IsWithinTable(tbl, tblRow, tblCell) Then
        If tblRow Is Nothing Then
          ' multiple rows are selected
          For Each tblRow In tbl.Rows
            For Each tblCell In tblRow.Cells
              If tblCell.TextRange.RangeStart >= RTB.SelectedTextRange.RangeStart Then
                If tblCell.TextRange.RangeStart <= RTB.SelectedTextRange.RangeEnd Then
                  ' this cell is selected
                  Select Case commandID
                    Case CMDID_TABLE_ALIGNTOP
                      tblCell.VAlignment = valTop
                    Case CMDID_TABLE_ALIGNCENTER
                      tblCell.VAlignment = valCenter
                    Case CMDID_TABLE_ALIGNBOTTOM
                      tblCell.VAlignment = valBottom
                    Case CMDID_TABLE_CELLBACKCOLOR
                      tblCell.BackColor1 = activeCellBackColor
                  End Select
                End If
              End If
            Next tblCell
          Next tblRow
        ElseIf tblCell Is Nothing Then
          ' multiple columns are selected
          For Each tblCell In tblRow.Cells
            If tblCell.TextRange.RangeStart >= RTB.SelectedTextRange.RangeStart Then
              If tblCell.TextRange.RangeStart <= RTB.SelectedTextRange.RangeEnd Then
                ' this cell is selected
                Select Case commandID
                  Case CMDID_TABLE_ALIGNTOP
                    tblCell.VAlignment = valTop
                  Case CMDID_TABLE_ALIGNCENTER
                    tblCell.VAlignment = valCenter
                  Case CMDID_TABLE_ALIGNBOTTOM
                    tblCell.VAlignment = valBottom
                  Case CMDID_TABLE_CELLBACKCOLOR
                    tblCell.BackColor1 = activeCellBackColor
                End Select
              End If
            End If
          Next tblCell
        Else
          ' only one cell is selected
          Select Case commandID
            Case CMDID_TABLE_ALIGNTOP
              tblCell.VAlignment = valTop
            Case CMDID_TABLE_ALIGNCENTER
              tblCell.VAlignment = valCenter
            Case CMDID_TABLE_ALIGNBOTTOM
              tblCell.VAlignment = valBottom
            Case CMDID_TABLE_CELLBACKCOLOR
              tblCell.BackColor1 = activeCellBackColor
          End Select
        End If
      End If

    Case CMDID_MATH_TOGGLEMATHZONE
      With RTB.SelectedTextRange
        .Font.IsMathZone = bpvToggle
        ' Make professional display the default
        ' NOTE: We could also call BuildUpMath directly, without setting IsMathZone.
        .BuildUpMath bumMathAutoCorrect Or bumMathCollapseSel
      End With
    Case CMDID_MATH_BUILDUP
      RTB.SelectedTextRange.BuildUpMath bumMathAutoCorrect Or bumMathCollapseSel
    Case CMDID_MATH_BUILDDOWN
      RTB.SelectedTextRange.BuildDownMath bumMathAlphabetics
      ' also possible: RTB.SelectedTextRange.BuildUpMath bumMathBuildDown Or bumMathAlphabetics
    Case CMDID_MATH_INSERTROOTSIGN
      With RTB.SelectedTextRange
        .Text = ChrW(&H221A&) & "(&)"
        .BuildUpMath bumMathAutoCorrect Or bumMathCollapseSel
      End With
    Case CMDID_MATH_INSERTSUMSIGN
      With RTB.SelectedTextRange
        .Text = ChrW(&H2211&) & "_i^j" & ChrW(&H2592&) & "x"
        .BuildUpMath bumMathAutoCorrect Or bumMathCollapseSel
      End With
    Case CMDID_MATH_INSERTPRODUCTSIGN
      With RTB.SelectedTextRange
        .Text = ChrW(&H220F&) & "_i^j" & ChrW(&H2592&) & "x"
        .BuildUpMath bumMathAutoCorrect Or bumMathCollapseSel
      End With
    Case CMDID_MATH_INSERTINTSIGN
      With RTB.SelectedTextRange
        .Text = ChrW(&H222B&) & "_i^j"
        .BuildUpMath bumMathAutoCorrect Or bumMathCollapseSel
      End With
    Case CMDID_MATH_INSERTLIMES
      With RTB.SelectedTextRange
        .Text = "lim" & ChrW(&H252C&) & "(n" & ChrW(&H2192&) & ChrW(&H221E&) & ")" & ChrW(&H2061&) & "()"
        .BuildUpMath bumMathAutoCorrect Or bumMathCollapseSel
      End With
    Case CMDID_MATH_INSERTMATRIX
      With RTB.SelectedTextRange
        .Text = "(" & ChrW(&H25A0&) & "8(1&0@0&1))"
        .BuildUpMath bumMathAutoCorrect Or bumMathCollapseSel
      End With
  End Select
End Sub

Private Property Get IDocument_FileName() As String
  IDocument_FileName = documentFileName
End Property

Private Function IDocument_GetCommandToolTip(ByVal commandID As Long) As String
  Dim ret As String

  Select Case commandID
    Case CMDID_EDIT_UNDO
      If RTB.CanUndo Then
        Select Case RTB.NextUndoActionType
          Case uatAutoTable
            ret = "Undo: Automatic table insertion"
          Case uatCut
            ret = "Undo: Cut"
          Case uatDelete
            ret = "Undo: Delete"
          Case uatDragDrop
            ret = "Undo: Drag&Drop"
          Case uatPaste
            ret = "Undo: Paste"
          Case uatTyping
            ret = "Undo: Typing"
          Case uatUnknown
            ret = "Undo"
        End Select
      Else
        ret = "Undo not possible"
      End If
    Case CMDID_EDIT_REDO
      If RTB.CanRedo Then
        Select Case RTB.NextRedoActionType
          Case uatAutoTable
            ret = "Redo: Automatic table insertion"
          Case uatCut
            ret = "Redo: Cut"
          Case uatDelete
            ret = "Redo: Delete"
          Case uatDragDrop
            ret = "Redo: Drag&Drop"
          Case uatPaste
            ret = "Redo: Paste"
          Case uatTyping
            ret = "Redo: Typing"
          Case uatUnknown
            ret = "Redo"
        End Select
      Else
        ret = "Redo not possible"
      End If
  End Select

  IDocument_GetCommandToolTip = ret
End Function

Private Function IDocument_OpenFile(ByVal File As String) As Boolean
  Dim str As String

  If RTB.LoadFromFile(File) Then
    documentFileName = File
    str = documentFileName
    Call PathStripPath(StrPtr(str))
    Me.Caption = str
    RTB.Modified = False
    IDocument_OpenFile = True
  End If
End Function

Private Function IDocument_SaveFile(ByVal File As String) As Boolean
  Dim str As String

  If RTB.SaveToFile(File) Then
    documentFileName = File
    str = documentFileName
    Call PathStripPath(StrPtr(str))
    Me.Caption = str
    RTB.Modified = False
    IDocument_SaveFile = True
  End If
End Function

Private Property Get IDocument_SelectedTextRange() As TextRange
  Set IDocument_SelectedTextRange = RTB.SelectedTextRange
End Property

Private Property Get IDocument_SelectionCellBackColor() As Long
  Dim lRet As Long
  Dim tbl As Table
  Dim tblRow As TableRow
  Dim tblCell As TableCell

  lRet = -1
  If RTB.SelectedTextRange.IsWithinTable(tbl, tblRow, tblCell) Then
    If tblRow Is Nothing Then
      ' multiple rows are selected
      For Each tblRow In tbl.Rows
        For Each tblCell In tblRow.Cells
          If tblCell.TextRange.RangeStart >= RTB.SelectedTextRange.RangeStart Then
            If tblCell.TextRange.RangeStart <= RTB.SelectedTextRange.RangeEnd Then
              ' this cell is selected
              If (lRet <> -1) And (lRet <> tblCell.BackColor1) Then
                lRet = -1
                Exit For
              Else
                lRet = tblCell.BackColor1
              End If
            End If
          End If
        Next tblCell
      Next tblRow
    ElseIf tblCell Is Nothing Then
      ' multiple columns are selected
      For Each tblCell In tblRow.Cells
        If tblCell.TextRange.RangeStart >= RTB.SelectedTextRange.RangeStart Then
          If tblCell.TextRange.RangeStart <= RTB.SelectedTextRange.RangeEnd Then
            ' this cell is selected
            If (lRet <> -1) And (lRet <> tblCell.BackColor1) Then
              lRet = -1
              Exit For
            Else
              lRet = tblCell.BackColor1
            End If
          End If
        End If
      Next tblCell
    Else
      ' only one cell is selected
      lRet = tblCell.BackColor1
    End If
  End If
  IDocument_SelectionCellBackColor = lRet
End Property

Private Sub IDocument_SetFocus()
  RTB.SetFocus
End Sub

Private Sub IDocument_SetFontName(ByVal fontName As String)
  If tempFont Is Nothing Then
    RTB.SelectedTextRange.Font.Name = fontName
  Else
    tempFont.Name = fontName
  End If
End Sub

Private Sub IDocument_SetFontSize(ByVal fontSize As Single)
  If tempFont Is Nothing Then
    RTB.SelectedTextRange.Font.Size = fontSize
  Else
    tempFont.Size = fontSize
  End If
End Sub

Private Sub IDocument_StartFontChangePreview()
  Set tempFont = RTB.SelectedTextRange.Font
  'Call RTB.Undo(-9999995)
End Sub

Private Property Get IDocument_SelectionTextBackColor() As Long
  If tempFont Is Nothing Then
    IDocument_SelectionTextBackColor = RTB.SelectedTextRange.Font.BackColor
  Else
    IDocument_SelectionTextBackColor = tempFont.BackColor
  End If
End Property

Private Property Get IDocument_TextBackColor() As Long
  IDocument_TextBackColor = activeTextBackColor
End Property

Private Property Let IDocument_TextBackColor(ByVal RHS As Long)
  If tempFont Is Nothing Then
    activeTextBackColor = RHS
    RTB.SelectedTextRange.Font.BackColor = RHS
  Else
    tempFont.BackColor = RHS
  End If
End Property

Private Property Get IDocument_SelectionTextForeColor() As Long
  If tempFont Is Nothing Then
    IDocument_SelectionTextForeColor = RTB.SelectedTextRange.Font.ForeColor
  Else
    IDocument_SelectionTextForeColor = tempFont.ForeColor
  End If
End Property

Private Property Get IDocument_TextForeColor() As Long
  IDocument_TextForeColor = activeTextForeColor
End Property

Private Property Let IDocument_TextForeColor(ByVal RHS As Long)
  If tempFont Is Nothing Then
    activeTextForeColor = RHS
    RTB.SelectedTextRange.Font.ForeColor = RHS
  Else
    tempFont.ForeColor = RHS
  End If
End Property

Private Sub Form_Activate()
  mdiFrame.ActivatedDocument Me
  ' might as well be ">= reav75", but it does not seem to work with Office 2013 (reav80) on Windows 10,
  ' while it works with the native (reav75) control of Windows 10
  If RTB.RichEditAPIVersion = reav75 Then
    mdiFrame.EnableDisableCommand CMDID_EDIT_SPELLCHECKING, True
    mdiFrame.CheckCommand CMDID_EDIT_SPELLCHECKING, RTB.UseBuiltInSpellChecking
  Else
    mdiFrame.EnableDisableCommand CMDID_EDIT_SPELLCHECKING, False
    mdiFrame.CheckCommand CMDID_EDIT_SPELLCHECKING, False
  End If
End Sub

Private Sub Form_GotFocus()
  RTB.SetFocus
End Sub

Private Sub Form_Load()
  activeCellBackColor = -1
  activeTextBackColor = -1
  activeTextForeColor = -1
  If Not (mdiFrame Is Nothing) Then
    mdiFrame.OpenedDocument Me, windowID
  End If
  Me.Caption = "Document " & CStr(windowID)
  Clipboard.Clear
  RTB.HyphenationFunction = g_pHyphenateFunction
  RTB.DisabledEvents = RTB.DisabledEvents And Not deKeyboardEvents
  ' DIN A4: 210 mm
  On Error Resume Next     ' might fail if the system does not have any printer
  Call RTB.SetTargetDevice(Printer.hDC, Me.ScaleX(210, vbMillimeters, vbTwips))
End Sub

Private Sub Form_Paint()
  If RTB.Left > 0 Then
    ' TODO: Draw a border and a drop shadow
  End If
End Sub

Private Sub Form_Resize()
  Dim cx As Long

  ' DIN A4: 210 mm
  cx = Me.ScaleX(210, vbMillimeters, Me.ScaleMode) + Me.ScaleX(2, vbPixels, Me.ScaleMode)
  If cx < Me.ScaleWidth Then
    RTB.Move (Me.ScaleWidth - cx) / 2, Me.ScaleY(10, vbPixels, Me.ScaleMode), cx, Me.ScaleHeight - Me.ScaleY(10, vbPixels, Me.ScaleMode)
  Else
    RTB.Move 0, 0, Me.ScaleWidth, Me.ScaleHeight
  End If
  Me.Refresh
End Sub

Private Sub Form_Unload(Cancel As Integer)
  If Not (mdiFrame Is Nothing) Then
    mdiFrame.ClosingDocument windowID
    Set mdiFrame = Nothing
  End If
End Sub

Private Sub RTB_ContextMenu(ByVal menuType As RichTextBoxCtlLibUCtl.MenuTypeConstants, ByVal TextRange As RichTextBoxCtlLibUCtl.IRichTextRange, ByVal SelectionType As RichTextBoxCtlLibUCtl.SelectionTypeConstants, ByVal OLEObject As RichTextBoxCtlLibUCtl.IRichOLEObject, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As RichTextBoxCtlLibUCtl.HitTestConstants, ByVal isRightMouseDrop As Boolean, ByVal draggedData As RichTextBoxCtlLibUCtl.IOLEDataObject, hMenuToDisplay As stdole.OLE_HANDLE)
  Const MFS_GRAYED = &H3
  Const MFS_DISABLED = MFS_GRAYED
  Const MFS_CHECKED = &H8
  Const MFS_HILITE = &H80
  Const MFS_ENABLED = &H0
  Const MFS_UNCHECKED = &H0
  Const MFS_UNHILITE = &H0
  Const MFS_DEFAULT = &H1000&
  Const MFT_STRING = &H0
  Const MFT_MENUBARBREAK = &H20
  Const MFT_MENUBREAK = &H40
  Const MFT_RADIOCHECK = &H200
  Const MFT_SEPARATOR = &H800
  Const MFT_RIGHTORDER = &H2000&
  Const MFT_RIGHTJUSTIFY = &H4000&
  Const MIIM_ID = &H2
  Const MIIM_STATE = &H1
  Const MIIM_TYPE = &H10
  Const TPM_LEFTALIGN = &H0&
  Const TPM_RETURNCMD = &H100
  Const TPM_TOPALIGN = &H0&
  Const TPM_VERPOSANIMATION = &H1000&
  Const TPM_VERTICAL = &H40&
  Dim i As Long
  Dim commandID As Long
  Dim hMenu As Long
  Dim miiData As MENUITEMINFO
  Dim ptMenu As POINT
  Dim verbDetails As OLEVERBDETAILS
  Dim verbs As Variant

  If ((SelectionType And setObject) = setObject) And Not (OLEObject Is Nothing) Then
    If OLEObject.GetVerbs(verbs) > 0 Then
      hMenu = CreatePopupMenu
      If hMenu Then
        With miiData
          .cbSize = LenB(miiData)
          .fMask = MIIM_TYPE Or MIIM_ID Or MIIM_STATE
          For i = LBound(verbs) To UBound(verbs)
            verbDetails = verbs(i)
            .fType = verbDetails.MenuFlags And (MFT_STRING Or MFT_MENUBARBREAK Or MFT_MENUBREAK Or MFT_RADIOCHECK Or MFT_SEPARATOR Or MFT_RIGHTORDER Or MFT_RIGHTJUSTIFY)
            .fState = verbDetails.MenuFlags And (MFS_GRAYED Or MFS_DISABLED Or MFS_CHECKED Or MFS_HILITE Or MFS_ENABLED Or MFS_UNCHECKED Or MFS_UNHILITE Or MFS_DEFAULT)
            .wID = i - LBound(verbs) + 1
            .dwTypeData = StrPtr(verbDetails.Name)
            .cch = Len(verbDetails.Name)
            InsertMenuItem hMenu, i - LBound(verbs) + 1, 1, miiData
          Next i
          
          ptMenu.x = ScaleX(x, vbTwips, vbPixels)
          ptMenu.y = ScaleY(y, vbTwips, vbPixels)
          ClientToScreen RTB.hWnd, ptMenu
          commandID = TrackPopupMenuEx(hMenu, TPM_LEFTALIGN Or TPM_TOPALIGN Or TPM_VERTICAL Or TPM_VERPOSANIMATION Or TPM_RETURNCMD, ptMenu.x, ptMenu.y, RTB.hWnd, 0)
          If commandID <> 0 Then
            OLEObject.ExecuteVerb verbs(commandID - 1).ID
          End If
        End With
        DestroyMenu hMenu
      End If
    End If
  End If
End Sub

Private Sub RTB_GotFocus()
  Timer1.Enabled = True
  RTB_SelectionChanged RTB.SelectedTextRange, SetText
End Sub

Private Sub RTB_RecreatedControlWindow(ByVal hWnd As Long)
  RTB.HyphenationFunction = g_pHyphenateFunction
  RTB.DisabledEvents = RTB.DisabledEvents And Not deKeyboardEvents
  ' DIN A4: 210 mm
  On Error Resume Next     ' might fail if the system does not have any printer
  Call RTB.SetTargetDevice(Printer.hDC, Me.ScaleX(210, vbMillimeters, vbTwips))
End Sub

Private Sub RTB_SelectionChanged(ByVal SelectedTextRange As RichTextBoxCtlLibUCtl.IRichTextRange, ByVal SelectionType As RichTextBoxCtlLibUCtl.SelectionTypeConstants)
  Dim bIsInTable As Boolean
  Dim tbl As Table
  Dim tblRow As TableRow
  Dim tblCell As TableCell

  With SelectedTextRange.Font
    mdiFrame.CheckTriStateCommand CMDID_FONT_BOLD, BooleanPropertyValueToSelectionState(.Bold)
    mdiFrame.CheckTriStateCommand CMDID_FONT_ITALIC, BooleanPropertyValueToSelectionState(.Italic)
    mdiFrame.CheckTriStateCommand CMDID_FONT_UNDERLINE, UnderlineTypeToSelectionState(.UnderlineType)
    mdiFrame.CheckTriStateCommand CMDID_FONT_STRIKETHROUGH, BooleanPropertyValueToSelectionState(.Strikethrough)
    mdiFrame.CheckTriStateCommand CMDID_FONT_SUBSCRIPT, BooleanPropertyValueToSelectionState(.Subscript)
    mdiFrame.CheckTriStateCommand CMDID_FONT_SUPERSCRIPT, BooleanPropertyValueToSelectionState(.Superscript)
    If tempFont Is Nothing Then
      mdiFrame.UpdateFontName .Name
      mdiFrame.UpdateFontSize .Size
    End If
  End With
  With SelectedTextRange.Paragraph
    mdiFrame.CheckCommand CMDID_FORMAT_ALIGNLEFT, .HAlignment = halLeft
    mdiFrame.CheckCommand CMDID_FORMAT_ALIGNCENTER, .HAlignment = halCenter
    mdiFrame.CheckCommand CMDID_FORMAT_ALIGNRIGHT, .HAlignment = halRight
    mdiFrame.CheckCommand CMDID_FORMAT_ALIGNJUSTIFY, .HAlignment = halJustify
    mdiFrame.CheckCommand CMDID_FORMAT_BULLETLIST, (.ListLevelIndex > 0) And ((.ListNumberingStyle And lnsTypeMask) = lnsBullet)
    mdiFrame.CheckCommand CMDID_FORMAT_NUMBEREDLIST, (.ListLevelIndex > 0) And ((.ListNumberingStyle And lnsTypeMask) = lnsArabicNumbers)
    mdiFrame.EnableDisableCommand CMDID_FORMAT_INCREASEINDENT, (.ListLevelIndex > 0) And CaretIsAtLineStartOrWholeLineIsSelected
    mdiFrame.EnableDisableCommand CMDID_FORMAT_DECREASEINDENT, (.ListLevelIndex > 0) And CaretIsAtLineStartOrWholeLineIsSelected
  End With
  If RTB.RichEditAPIVersion >= reav70 Then
    mdiFrame.EnableDisableCommand CMDID_FORMAT_CONVERTTOHYPERLINK, SelectedTextRange.RangeLength > 0
  Else
    mdiFrame.EnableDisableCommand CMDID_FORMAT_CONVERTTOHYPERLINK, False
  End If
  bIsInTable = SelectedTextRange.IsWithinTable(, , tblCell)
  If bIsInTable And Not (tblCell Is Nothing) Then
    mdiFrame.CheckCommand CMDID_TABLE_ALIGNTOP, tblCell.VAlignment = valTop
    mdiFrame.CheckCommand CMDID_TABLE_ALIGNCENTER, tblCell.VAlignment = valCenter
    mdiFrame.CheckCommand CMDID_TABLE_ALIGNBOTTOM, tblCell.VAlignment = valBottom
  Else
    mdiFrame.CheckCommand CMDID_TABLE_ALIGNTOP, False
    mdiFrame.CheckCommand CMDID_TABLE_ALIGNCENTER, False
    mdiFrame.CheckCommand CMDID_TABLE_ALIGNBOTTOM, False
  End If
  mdiFrame.EnableDisableCommand CMDID_TABLE_DELETETABLE, bIsInTable
  mdiFrame.EnableDisableCommand CMDID_TABLE_INSERTROW, bIsInTable
  mdiFrame.EnableDisableCommand CMDID_TABLE_DELETEROW, bIsInTable
  mdiFrame.EnableDisableCommand CMDID_TABLE_INSERTCOLUMN, bIsInTable
  mdiFrame.EnableDisableCommand CMDID_TABLE_DELETECOLUMN, bIsInTable
  mdiFrame.EnableDisableCommand CMDID_TABLE_MERGECELLS, bIsInTable And (tblCell Is Nothing)
  mdiFrame.EnableDisableCommand CMDID_TABLE_CELLBACKCOLOR, bIsInTable
  mdiFrame.EnableDisableCommand CMDID_TABLE_ALIGNTOP, bIsInTable
  mdiFrame.EnableDisableCommand CMDID_TABLE_ALIGNCENTER, bIsInTable
  mdiFrame.EnableDisableCommand CMDID_TABLE_ALIGNBOTTOM, bIsInTable
  If RTB.RichEditAPIVersion >= reav80 Then
    mdiFrame.CheckTriStateCommand CMDID_MATH_TOGGLEMATHZONE, BooleanPropertyValueToSelectionState(SelectedTextRange.Font.IsMathZone)
    mdiFrame.EnableDisableCommand CMDID_MATH_TOGGLEMATHZONE, SelectedTextRange.RangeLength > 0
    mdiFrame.EnableDisableCommand CMDID_MATH_BUILDUP, (SelectedTextRange.Font.IsMathZone = bpvTrue) And (SelectedTextRange.RangeLength > 0)
    mdiFrame.EnableDisableCommand CMDID_MATH_BUILDDOWN, (SelectedTextRange.Font.IsMathZone = bpvTrue) And (SelectedTextRange.RangeLength > 0)
  Else
    mdiFrame.CheckTriStateCommand CMDID_MATH_TOGGLEMATHZONE, False
    mdiFrame.EnableDisableCommand CMDID_MATH_TOGGLEMATHZONE, False
    mdiFrame.EnableDisableCommand CMDID_MATH_BUILDUP, False
    mdiFrame.EnableDisableCommand CMDID_MATH_BUILDDOWN, False
    mdiFrame.EnableDisableCommand CMDID_MATH_INSERTROOTSIGN, False
    mdiFrame.EnableDisableCommand CMDID_MATH_INSERTSUMSIGN, False
    mdiFrame.EnableDisableCommand CMDID_MATH_INSERTPRODUCTSIGN, False
    mdiFrame.EnableDisableCommand CMDID_MATH_INSERTINTSIGN, False
    mdiFrame.EnableDisableCommand CMDID_MATH_INSERTLIMES, False
    mdiFrame.EnableDisableCommand CMDID_MATH_INSERTMATRIX, False
  End If
End Sub

Private Sub RTB_LostFocus()
  Timer1.Enabled = False
End Sub

Private Sub Timer1_Timer()
  If Not (mdiFrame Is Nothing) Then
    mdiFrame.EnableDisableCommand CMDID_FILE_SAVE, RTB.Modified
    mdiFrame.EnableDisableCommand CMDID_EDIT_PASTE, RTB.SelectedTextRange.CanPaste
    mdiFrame.EnableDisableCommand CMDID_EDIT_CUT, RTB.CanCopy
    mdiFrame.EnableDisableCommand CMDID_EDIT_COPY, RTB.CanCopy
    mdiFrame.EnableDisableCommand CMDID_EDIT_UNDO, RTB.CanUndo
    mdiFrame.EnableDisableCommand CMDID_EDIT_REDO, RTB.CanRedo
  End If
End Sub

Private Function BooleanPropertyValueToSelectionState(ByVal boolPropVal As BooleanPropertyValueConstants) As SelectionStateConstants
  Select Case boolPropVal
    Case bpvUndefined
      BooleanPropertyValueToSelectionState = ssIndeterminate
    Case bpvFalse
      BooleanPropertyValueToSelectionState = ssUnchecked
    Case bpvTrue
      BooleanPropertyValueToSelectionState = ssChecked
  End Select
End Function

Private Function UnderlineTypeToSelectionState(ByVal underline As UnderlineTypeConstants) As SelectionStateConstants
  Select Case underline
    Case utUndefined
      UnderlineTypeToSelectionState = ssIndeterminate
    Case utNone
      UnderlineTypeToSelectionState = ssUnchecked
    Case Else
      UnderlineTypeToSelectionState = ssChecked
  End Select
End Function

Private Function CaretIsAtLineStartOrWholeLineIsSelected() As Boolean
  Dim ret As Boolean
  Dim rng As TextRange
  
  With RTB.SelectedTextRange
    Set rng = .Clone
    If rng.RangeLength = 0 Then
      ' an insertion point (Caret)
      Call rng.MoveStartToStartOfUnit(uLine, True)
      If rng.RangeStart = .RangeStart Then
        ' Caret is at line start
        ret = True
      End If
    Else
      ' Is the line with the selection end selected partially or completly?
      Call rng.MoveEndToEndOfUnit(uLine, True)
      ' allow both cases: including the line feed, or without it
      If (rng.RangeEnd = .RangeEnd) Or (rng.RangeEnd = .RangeEnd + 1) Then
        ' the selection ends at the line's end
        ret = True
      End If
    End If
  End With
  CaretIsAtLineStartOrWholeLineIsSelected = ret
End Function
