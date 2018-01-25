VERSION 5.00
Object = "{FCCB83BF-E483-4317-9FF2-A460758238B5}#1.4#0"; "CBLCtlsU.ocx"
Object = "{B6CC61F6-3F1A-4B00-9918-13F66F185263}#1.1#0"; "LblCtlsU.ocx"
Object = "{A8F9B8E7-E699-4FCE-A647-72C877F8E632}#1.9#0"; "EditCtlsU.ocx"
Object = "{2AFA7915-463D-4B61-AEB7-41B1236C143E}#1.9#0"; "BtnCtlsU.ocx"
Begin VB.Form frmFind 
   BorderStyle     =   3  'Fester Dialog
   Caption         =   "Search"
   ClientHeight    =   2160
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6105
   Icon            =   "frmFind.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   144
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   407
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'Fenstermitte
   Begin CBLCtlsLibUCtl.ComboBox cmbSearchDirection 
      Height          =   315
      Left            =   3360
      TabIndex        =   8
      Top             =   1080
      Width           =   975
      _cx             =   1720
      _cy             =   556
      AcceptNumbersOnly=   0   'False
      Appearance      =   3
      AutoHorizontalScrolling=   -1  'True
      BackColor       =   -2147483643
      BorderStyle     =   0
      CharacterConversion=   0
      DisabledEvents  =   267503
      DontRedraw      =   0   'False
      DoOEMConversion =   0   'False
      DragDropDownTime=   -1
      DropDownKey     =   0
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   -2147483640
      HasStrings      =   -1  'True
      HoverTime       =   -1
      IMEMode         =   -1
      IntegralHeight  =   -1  'True
      ItemHeight      =   -1
      ListAlwaysShowVerticalScrollBar=   0   'False
      ListBackColor   =   -2147483643
      ListDragScrollTimeBase=   -1
      ListForeColor   =   -2147483640
      ListHeight      =   -1
      ListInsertMarkColor=   0
      ListScrollableWidth=   0
      ListWidth       =   0
      Locale          =   1024
      MaxTextLength   =   -1
      MinVisibleItems =   30
      MousePointer    =   0
      OwnerDrawItems  =   0
      ProcessContextMenuKeys=   -1  'True
      RegisterForOLEDragDrop=   0   'False
      RightToLeft     =   0
      SelectionFieldHeight=   -1
      Sorted          =   0   'False
      Style           =   1
      SupportOLEDragImages=   -1  'True
      UseSystemFont   =   -1  'True
      CueBanner       =   "frmFind.frx":000C
      Text            =   "frmFind.frx":002C
   End
   Begin BtnCtlsLibUCtl.CommandButton cmdReplace 
      Height          =   375
      Left            =   4440
      TabIndex        =   12
      Top             =   1080
      Width           =   1575
      _cx             =   2778
      _cy             =   661
      Appearance      =   0
      BackColor       =   -2147483633
      BorderStyle     =   0
      ButtonType      =   0
      ContentType     =   0
      DisabledEvents  =   1289
      DontRedraw      =   0   'False
      DropDownArrowHeight=   0
      DropDownArrowWidth=   15
      DropDownGlyph   =   54
      DropDownOnRight =   -1  'True
      DropDownPushed  =   0   'False
      DropDownStyle   =   1
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   -2147483630
      HAlignment      =   1
      HoverTime       =   -1
      IconAlignment   =   0
      IconMarginBottom=   0
      IconMarginLeft  =   0
      IconMarginRight =   0
      IconMarginTop   =   0
      KeepDropDownArrowAspectRatio=   -1  'True
      MousePointer    =   0
      MultiLine       =   -1  'True
      ProcessContextMenuKeys=   -1  'True
      Pushed          =   0   'False
      RegisterForOLEDragDrop=   0   'False
      RightToLeft     =   0
      ShowRightsElevationIcon=   0   'False
      ShowSplitLine   =   -1  'True
      Style           =   0
      SupportOLEDragImages=   -1  'True
      TextMarginBottom=   1
      TextMarginLeft  =   1
      TextMarginRight =   1
      TextMarginTop   =   1
      UseImprovedImageListSupport=   0   'False
      UseSystemFont   =   -1  'True
      VAlignment      =   1
      IconIndexes     =   "frmFind.frx":004C
      CommandLinkNote =   "frmFind.frx":0084
      Text            =   "frmFind.frx":00A4
   End
   Begin BtnCtlsLibUCtl.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Height          =   375
      Left            =   4440
      TabIndex        =   11
      Top             =   570
      Width           =   1575
      _cx             =   2778
      _cy             =   661
      Appearance      =   0
      BackColor       =   -2147483633
      BorderStyle     =   0
      ButtonType      =   0
      ContentType     =   0
      DisabledEvents  =   1289
      DontRedraw      =   0   'False
      DropDownArrowHeight=   0
      DropDownArrowWidth=   15
      DropDownGlyph   =   54
      DropDownOnRight =   -1  'True
      DropDownPushed  =   0   'False
      DropDownStyle   =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   -2147483630
      HAlignment      =   1
      HoverTime       =   -1
      IconAlignment   =   0
      IconMarginBottom=   0
      IconMarginLeft  =   0
      IconMarginRight =   0
      IconMarginTop   =   0
      KeepDropDownArrowAspectRatio=   -1  'True
      MousePointer    =   0
      MultiLine       =   -1  'True
      ProcessContextMenuKeys=   -1  'True
      Pushed          =   0   'False
      RegisterForOLEDragDrop=   0   'False
      RightToLeft     =   0
      ShowRightsElevationIcon=   0   'False
      ShowSplitLine   =   -1  'True
      Style           =   0
      SupportOLEDragImages=   -1  'True
      TextMarginBottom=   1
      TextMarginLeft  =   1
      TextMarginRight =   1
      TextMarginTop   =   1
      UseImprovedImageListSupport=   0   'False
      UseSystemFont   =   -1  'True
      VAlignment      =   1
      IconIndexes     =   "frmFind.frx":00D4
      CommandLinkNote =   "frmFind.frx":010C
      Text            =   "frmFind.frx":012C
   End
   Begin BtnCtlsLibUCtl.CommandButton cmdFindNext 
      Default         =   -1  'True
      Height          =   375
      Left            =   4440
      TabIndex        =   10
      Top             =   150
      Width           =   1575
      _cx             =   2778
      _cy             =   661
      Appearance      =   0
      BackColor       =   -2147483633
      BorderStyle     =   0
      ButtonType      =   0
      ContentType     =   0
      DisabledEvents  =   1289
      DontRedraw      =   0   'False
      DropDownArrowHeight=   0
      DropDownArrowWidth=   15
      DropDownGlyph   =   54
      DropDownOnRight =   -1  'True
      DropDownPushed  =   0   'False
      DropDownStyle   =   1
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   -2147483630
      HAlignment      =   1
      HoverTime       =   -1
      IconAlignment   =   0
      IconMarginBottom=   0
      IconMarginLeft  =   0
      IconMarginRight =   0
      IconMarginTop   =   0
      KeepDropDownArrowAspectRatio=   -1  'True
      MousePointer    =   0
      MultiLine       =   -1  'True
      ProcessContextMenuKeys=   -1  'True
      Pushed          =   0   'False
      RegisterForOLEDragDrop=   0   'False
      RightToLeft     =   0
      ShowRightsElevationIcon=   0   'False
      ShowSplitLine   =   -1  'True
      Style           =   0
      SupportOLEDragImages=   -1  'True
      TextMarginBottom=   1
      TextMarginLeft  =   1
      TextMarginRight =   1
      TextMarginTop   =   1
      UseImprovedImageListSupport=   0   'False
      UseSystemFont   =   -1  'True
      VAlignment      =   1
      IconIndexes     =   "frmFind.frx":015A
      CommandLinkNote =   "frmFind.frx":0192
      Text            =   "frmFind.frx":01B2
   End
   Begin BtnCtlsLibUCtl.CheckBox chkCaseSensitiveSearch 
      Height          =   255
      Left            =   2040
      TabIndex        =   9
      Top             =   1800
      Width           =   2295
      _cx             =   4048
      _cy             =   450
      Appearance      =   0
      AutoToggleCheckMark=   -1  'True
      BackColor       =   -2147483633
      BorderStyle     =   0
      CheckMarkOnRight=   0   'False
      ContentType     =   0
      DisabledEvents  =   267
      DontRedraw      =   0   'False
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   -2147483630
      HAlignment      =   0
      HoverTime       =   -1
      IconAlignment   =   0
      IconMarginBottom=   0
      IconMarginLeft  =   0
      IconMarginRight =   0
      IconMarginTop   =   0
      MousePointer    =   0
      MultiLine       =   -1  'True
      ProcessContextMenuKeys=   -1  'True
      Pushed          =   0   'False
      PushLike        =   0   'False
      RegisterForOLEDragDrop=   0   'False
      RightToLeft     =   0
      SelectionState  =   0
      Style           =   0
      SupportOLEDragImages=   -1  'True
      TextMarginBottom=   1
      TextMarginLeft  =   1
      TextMarginRight =   1
      TextMarginTop   =   1
      TriState        =   0   'False
      UseImprovedImageListSupport=   0   'False
      UseSystemFont   =   -1  'True
      VAlignment      =   1
      IconIndexes     =   "frmFind.frx":01E6
      Text            =   "frmFind.frx":021E
   End
   Begin BtnCtlsLibUCtl.CheckBox chkSearchForWholeWordOnly 
      Height          =   255
      Left            =   2040
      TabIndex        =   7
      Top             =   1440
      Width           =   2295
      _cx             =   4048
      _cy             =   450
      Appearance      =   0
      AutoToggleCheckMark=   -1  'True
      BackColor       =   -2147483633
      BorderStyle     =   0
      CheckMarkOnRight=   0   'False
      ContentType     =   0
      DisabledEvents  =   267
      DontRedraw      =   0   'False
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   -2147483630
      HAlignment      =   0
      HoverTime       =   -1
      IconAlignment   =   0
      IconMarginBottom=   0
      IconMarginLeft  =   0
      IconMarginRight =   0
      IconMarginTop   =   0
      MousePointer    =   0
      MultiLine       =   -1  'True
      ProcessContextMenuKeys=   -1  'True
      Pushed          =   0   'False
      PushLike        =   0   'False
      RegisterForOLEDragDrop=   0   'False
      RightToLeft     =   0
      SelectionState  =   0
      Style           =   0
      SupportOLEDragImages=   -1  'True
      TextMarginBottom=   1
      TextMarginLeft  =   1
      TextMarginRight =   1
      TextMarginTop   =   1
      TriState        =   0   'False
      UseImprovedImageListSupport=   0   'False
      UseSystemFont   =   -1  'True
      VAlignment      =   1
      IconIndexes     =   "frmFind.frx":026A
      Text            =   "frmFind.frx":02A2
   End
   Begin BtnCtlsLibUCtl.Frame Frame1 
      Height          =   975
      Left            =   120
      TabIndex        =   4
      Top             =   1080
      Width           =   1815
      _cx             =   3201
      _cy             =   1720
      Appearance      =   0
      BackColor       =   -2147483633
      BorderStyle     =   0
      BorderVisible   =   -1  'True
      ContentType     =   0
      DisabledEvents  =   3
      DontRedraw      =   0   'False
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   -2147483630
      HAlignment      =   0
      HoverTime       =   -1
      IconAlignment   =   0
      IconMarginBottom=   0
      IconMarginLeft  =   0
      IconMarginRight =   0
      IconMarginTop   =   0
      MousePointer    =   0
      RegisterForOLEDragDrop=   0   'False
      RightToLeft     =   0
      Style           =   0
      SupportOLEDragImages=   -1  'True
      UseImprovedImageListSupport=   0   'False
      UseSystemFont   =   -1  'True
      IconIndexes     =   "frmFind.frx":02F8
      Text            =   "frmFind.frx":0330
      Begin BtnCtlsLibUCtl.OptionButton optSearchInDocument 
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   1575
         _cx             =   2778
         _cy             =   450
         Appearance      =   0
         BackColor       =   -2147483633
         BorderStyle     =   0
         ContentType     =   0
         DisabledEvents  =   267
         DontRedraw      =   0   'False
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   -2147483630
         HAlignment      =   0
         HoverTime       =   -1
         IconAlignment   =   0
         IconMarginBottom=   0
         IconMarginLeft  =   0
         IconMarginRight =   0
         IconMarginTop   =   0
         MousePointer    =   0
         MultiLine       =   -1  'True
         OptionMarkOnRight=   0   'False
         ProcessContextMenuKeys=   -1  'True
         Pushed          =   0   'False
         PushLike        =   0   'False
         RegisterForOLEDragDrop=   0   'False
         RightToLeft     =   0
         Selected        =   -1  'True
         Style           =   0
         SupportOLEDragImages=   -1  'True
         TextMarginBottom=   1
         TextMarginLeft  =   1
         TextMarginRight =   1
         TextMarginTop   =   1
         UseImprovedImageListSupport=   0   'False
         UseSystemFont   =   -1  'True
         VAlignment      =   1
         IconIndexes     =   "frmFind.frx":0362
         Text            =   "frmFind.frx":039A
      End
      Begin BtnCtlsLibUCtl.OptionButton optSearchInSelection 
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   600
         Width           =   1575
         _cx             =   2778
         _cy             =   450
         Appearance      =   0
         BackColor       =   -2147483633
         BorderStyle     =   0
         ContentType     =   0
         DisabledEvents  =   267
         DontRedraw      =   0   'False
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   -2147483630
         HAlignment      =   0
         HoverTime       =   -1
         IconAlignment   =   0
         IconMarginBottom=   0
         IconMarginLeft  =   0
         IconMarginRight =   0
         IconMarginTop   =   0
         MousePointer    =   0
         MultiLine       =   -1  'True
         OptionMarkOnRight=   0   'False
         ProcessContextMenuKeys=   -1  'True
         Pushed          =   0   'False
         PushLike        =   0   'False
         RegisterForOLEDragDrop=   0   'False
         RightToLeft     =   0
         Selected        =   0   'False
         Style           =   0
         SupportOLEDragImages=   -1  'True
         TextMarginBottom=   1
         TextMarginLeft  =   1
         TextMarginRight =   1
         TextMarginTop   =   1
         UseImprovedImageListSupport=   0   'False
         UseSystemFont   =   -1  'True
         VAlignment      =   1
         IconIndexes     =   "frmFind.frx":03CC
         Text            =   "frmFind.frx":0404
      End
   End
   Begin BtnCtlsLibUCtl.CheckBox chkReplaceWith 
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   630
      Width           =   1215
      _cx             =   2143
      _cy             =   450
      Appearance      =   0
      AutoToggleCheckMark=   -1  'True
      BackColor       =   -2147483633
      BorderStyle     =   0
      CheckMarkOnRight=   0   'False
      ContentType     =   0
      DisabledEvents  =   267
      DontRedraw      =   0   'False
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   -2147483630
      HAlignment      =   0
      HoverTime       =   -1
      IconAlignment   =   0
      IconMarginBottom=   0
      IconMarginLeft  =   0
      IconMarginRight =   0
      IconMarginTop   =   0
      MousePointer    =   0
      MultiLine       =   -1  'True
      ProcessContextMenuKeys=   -1  'True
      Pushed          =   0   'False
      PushLike        =   0   'False
      RegisterForOLEDragDrop=   0   'False
      RightToLeft     =   0
      SelectionState  =   0
      Style           =   0
      SupportOLEDragImages=   -1  'True
      TextMarginBottom=   1
      TextMarginLeft  =   1
      TextMarginRight =   1
      TextMarginTop   =   1
      TriState        =   0   'False
      UseImprovedImageListSupport=   0   'False
      UseSystemFont   =   -1  'True
      VAlignment      =   1
      IconIndexes     =   "frmFind.frx":0440
      Text            =   "frmFind.frx":0478
   End
   Begin EditCtlsLibUCtl.TextBox txtSearchFor 
      Height          =   315
      Left            =   1440
      TabIndex        =   1
      Top             =   180
      Width           =   2895
      _cx             =   5106
      _cy             =   556
      AcceptNumbersOnly=   0   'False
      AcceptTabKey    =   0   'False
      AllowDragDrop   =   -1  'True
      AlwaysShowSelection=   0   'False
      Appearance      =   1
      AutoScrolling   =   2
      BackColor       =   -2147483643
      BorderStyle     =   0
      CancelIMECompositionOnSetFocus=   0   'False
      CharacterConversion=   0
      CompleteIMECompositionOnKillFocus=   0   'False
      DisabledBackColor=   -1
      DisabledEvents  =   3075
      DisabledForeColor=   -1
      DisplayCueBannerOnFocus=   0   'False
      DontRedraw      =   0   'False
      DoOEMConversion =   0   'False
      DragScrollTimeBase=   -1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   -2147483640
      FormattingRectangleHeight=   0
      FormattingRectangleLeft=   0
      FormattingRectangleTop=   0
      FormattingRectangleWidth=   0
      HAlignment      =   0
      HoverTime       =   -1
      IMEMode         =   -1
      InsertMarkColor =   0
      InsertSoftLineBreaks=   0   'False
      LeftMargin      =   -1
      MaxTextLength   =   -1
      Modified        =   0   'False
      MousePointer    =   0
      MultiLine       =   0   'False
      OLEDragImageStyle=   0
      PasswordChar    =   0
      ProcessContextMenuKeys=   -1  'True
      ReadOnly        =   0   'False
      RegisterForOLEDragDrop=   0   'False
      RightMargin     =   -1
      RightToLeft     =   0
      ScrollBars      =   0
      SelectedTextMousePointer=   0
      SupportOLEDragImages=   -1  'True
      TabWidth        =   -1
      UseCustomFormattingRectangle=   0   'False
      UsePasswordChar =   0   'False
      UseSystemFont   =   -1  'True
      CueBanner       =   "frmFind.frx":04B4
      Text            =   "frmFind.frx":050C
   End
   Begin EditCtlsLibUCtl.TextBox txtReplaceWith 
      Height          =   315
      Left            =   1440
      TabIndex        =   3
      Top             =   600
      Width           =   2895
      _cx             =   5106
      _cy             =   556
      AcceptNumbersOnly=   0   'False
      AcceptTabKey    =   0   'False
      AllowDragDrop   =   -1  'True
      AlwaysShowSelection=   0   'False
      Appearance      =   1
      AutoScrolling   =   2
      BackColor       =   -2147483643
      BorderStyle     =   0
      CancelIMECompositionOnSetFocus=   0   'False
      CharacterConversion=   0
      CompleteIMECompositionOnKillFocus=   0   'False
      DisabledBackColor=   -1
      DisabledEvents  =   3075
      DisabledForeColor=   -1
      DisplayCueBannerOnFocus=   0   'False
      DontRedraw      =   0   'False
      DoOEMConversion =   0   'False
      DragScrollTimeBase=   -1
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   -2147483640
      FormattingRectangleHeight=   0
      FormattingRectangleLeft=   0
      FormattingRectangleTop=   0
      FormattingRectangleWidth=   0
      HAlignment      =   0
      HoverTime       =   -1
      IMEMode         =   -1
      InsertMarkColor =   0
      InsertSoftLineBreaks=   0   'False
      LeftMargin      =   -1
      MaxTextLength   =   -1
      Modified        =   0   'False
      MousePointer    =   0
      MultiLine       =   0   'False
      OLEDragImageStyle=   0
      PasswordChar    =   0
      ProcessContextMenuKeys=   -1  'True
      ReadOnly        =   0   'False
      RegisterForOLEDragDrop=   0   'False
      RightMargin     =   -1
      RightToLeft     =   0
      ScrollBars      =   0
      SelectedTextMousePointer=   0
      SupportOLEDragImages=   -1  'True
      TabWidth        =   -1
      UseCustomFormattingRectangle=   0   'False
      UsePasswordChar =   0   'False
      UseSystemFont   =   -1  'True
      CueBanner       =   "frmFind.frx":052C
      Text            =   "frmFind.frx":0580
   End
   Begin LblCtlsLibUCtl.WindowlessLabel WindowlessLabel2 
      Height          =   195
      Left            =   2040
      TabIndex        =   13
      Top             =   1140
      Width           =   1230
      _cx             =   2170
      _cy             =   344
      Appearance      =   0
      AutoSize        =   1
      BackColor       =   -2147483633
      BackStyle       =   0
      BorderStyle     =   0
      ClipLastLine    =   -1  'True
      DisabledEvents  =   4099
      DontRedraw      =   0   'False
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   -2147483630
      HAlignment      =   0
      HoverTime       =   -1
      MousePointer    =   0
      RegisterForOLEDragDrop=   0   'False
      RightToLeft     =   0
      SupportOLEDragImages=   -1  'True
      TextTruncationStyle=   -1
      UseMnemonic     =   -1  'True
      UseSystemFont   =   -1  'True
      VAlignment      =   0
      WordWrapping    =   1
      Text            =   "frmFind.frx":05A0
   End
   Begin LblCtlsLibUCtl.WindowlessLabel WindowlessLabel1 
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   780
      _cx             =   1376
      _cy             =   344
      Appearance      =   0
      AutoSize        =   1
      BackColor       =   -2147483633
      BackStyle       =   0
      BorderStyle     =   0
      ClipLastLine    =   -1  'True
      DisabledEvents  =   4099
      DontRedraw      =   0   'False
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   -2147483630
      HAlignment      =   0
      HoverTime       =   -1
      MousePointer    =   0
      RegisterForOLEDragDrop=   0   'False
      RightToLeft     =   0
      SupportOLEDragImages=   -1  'True
      TextTruncationStyle=   -1
      UseMnemonic     =   -1  'True
      UseSystemFont   =   -1  'True
      VAlignment      =   0
      WordWrapping    =   1
      Text            =   "frmFind.frx":05E4
   End
End
Attribute VB_Name = "frmFind"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

  Private document As IDocument
  Private SelectedTextRange As TextRange
  Private doReplace As Boolean


Private Sub chkReplaceWith_SelectionStateChanged(ByVal previousSelectionState As BtnCtlsLibUCtl.SelectionStateConstants, ByVal newSelectionState As BtnCtlsLibUCtl.SelectionStateConstants)
  txtReplaceWith.Enabled = (newSelectionState = ssChecked)
  cmdReplace.Enabled = ((newSelectionState = ssChecked) And (txtReplaceWith.TextLength > 0))
  If txtReplaceWith.Enabled Then
    Call txtReplaceWith.SetFocus
  End If
End Sub

Private Sub cmdCancel_Click(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  Call Me.Hide
End Sub

Private Sub cmdFindNext_Click(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  Dim rng As TextRange
  Dim searchDirection As SearchDirectionConstants
  Dim searchMode As SearchModeConstants

  If optSearchInSelection.Selected Then
    searchDirection = sdWithinRange
    Set rng = SelectedTextRange.Clone
  Else
    searchDirection = cmbSearchDirection.SelectedItem.ItemData
    Set rng = document.DocumentTextRange.Clone
    If searchDirection = sdForward Then
      rng.RangeStart = document.SelectedTextRange.RangeEnd
    ElseIf searchDirection = sdBackward Then
      rng.RangeEnd = document.SelectedTextRange.RangeStart
    End If
  End If
  
  searchMode = smDefault
  If chkSearchForWholeWordOnly.SelectionState = ssChecked Then searchMode = searchMode Or smMatchWord
  If chkCaseSensitiveSearch.SelectionState = ssChecked Then searchMode = searchMode Or smMatchCase
  If rng.FindText(txtSearchFor.Text, searchDirection, searchMode) > 0 Then
    Call rng.Select
    If optSearchInSelection.Selected Then
      SelectedTextRange.RangeStart = rng.RangeEnd
    End If
  Else
    Call MsgBox("The specified text could not be found.")
  End If
End Sub

Private Sub cmdReplace_Click(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  If doReplace Then
    document.SelectedTextRange.Text = txtReplaceWith.Text
  Else
    doReplace = True
  End If
  Call cmdFindNext.Click
End Sub

Private Sub Form_Activate()
  Call txtSearchFor.SetSelection(0, -1)
End Sub

Private Sub Form_Load()
  With cmbSearchDirection.ComboItems
    Set cmbSearchDirection.SelectedItem = .Add("Down", , sdForward)
    Call .Add("Up", , sdBackward)
  End With
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If UnloadMode = 0 Then
    Cancel = True
    Call Me.Hide
  End If
End Sub

Private Sub optSearchInSelection_SelectionChanged(ByVal previousSelectionState As Boolean, ByVal newSelectionState As Boolean)
  If newSelectionState Then
    Set cmbSearchDirection.SelectedItem = cmbSearchDirection.ComboItems(0)
  End If
  cmbSearchDirection.Enabled = Not newSelectionState
End Sub

Private Sub txtReplaceWith_TextChanged()
  cmdReplace.Enabled = ((chkReplaceWith.SelectionState = ssChecked) And (txtReplaceWith.TextLength > 0))
End Sub

Private Sub txtSearchFor_TextChanged()
  cmdFindNext.Enabled = (txtSearchFor.TextLength > 0)
End Sub


Public Sub ShowForm(ByVal mdiFrame As IMDIFrame, ByVal doc As IDocument)
  Dim lineIndex As Long
  Dim str As String
  Dim rng As TextRange
  
  doReplace = False
  Set document = doc
  
  Set SelectedTextRange = doc.SelectedTextRange.Clone
  str = SelectedTextRange.Text
  If str <> "" Then
    txtSearchFor.Text = str
  End If
  chkReplaceWith.SelectionState = ssUnchecked
  Set rng = SelectedTextRange.Clone
  lineIndex = rng.UnitIndex(uLine)
  Call rng.Collapse(False)
  If lineIndex < rng.UnitIndex(uLine) Then
    optSearchInSelection.Selected = True
  Else
    optSearchInDocument.Selected = True
  End If
  optSearchInSelection.Enabled = (SelectedTextRange.RangeLength > 0)
  
  Call Me.Show(, mdiFrame)
  Call txtSearchFor.SetFocus
End Sub
