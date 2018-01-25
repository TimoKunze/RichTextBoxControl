VERSION 5.00
Object = "{5D0D0ABC-4898-4E46-AB48-291074A737A1}#1.0#0"; "TBarCtlsU.ocx"
Object = "{FCCB83BF-E483-4317-9FF2-A460758238B5}#1.0#0"; "CBLCtlsU.ocx"
Begin VB.MDIForm frmMain 
   BackColor       =   &H8000000C&
   Caption         =   "RichTextBox 1.0 - Wordpad Sample"
   ClientHeight    =   6855
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   13710
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.PictureBox picContainer 
      Align           =   1  'Oben ausrichten
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   0
      ScaleHeight     =   61
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   910
      TabIndex        =   1
      Top             =   690
      Width           =   13710
      Begin TBarCtlsLibUCtl.ToolBar tbFormat 
         Height          =   375
         Left            =   5760
         TabIndex        =   4
         Top             =   120
         Width           =   1815
         _cx             =   3201
         _cy             =   661
         AllowCustomization=   0   'False
         AlwaysDisplayButtonText=   0   'False
         AnchorHighlighting=   0   'False
         Appearance      =   0
         BackStyle       =   0
         BorderStyle     =   0
         ButtonHeight    =   0
         ButtonStyle     =   1
         ButtonTextPosition=   0
         ButtonWidth     =   0
         DisabledEvents  =   917995
         DisplayMenuDivider=   0   'False
         DisplayPartiallyClippedButtons=   0   'False
         DontRedraw      =   0   'False
         DragClickTime   =   -1
         DragDropCustomizationModifierKey=   0
         DropDownGap     =   -1
         Enabled         =   -1  'True
         FirstButtonIndentation=   0
         FocusOnClick    =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         HighlightColor  =   -1
         HorizontalButtonPadding=   -1
         HorizontalButtonSpacing=   0
         HorizontalIconCaptionGap=   -1
         HorizontalTextAlignment=   0
         HoverTime       =   -1
         InsertMarkColor =   0
         MaximumButtonWidth=   0
         MaximumTextRows =   1
         MenuBarTheme    =   1
         MenuMode        =   0   'False
         MinimumButtonWidth=   0
         MousePointer    =   0
         MultiColumn     =   0   'False
         NormalDropDownButtonStyle=   0
         OLEDragImageStyle=   0
         Orientation     =   0
         ProcessContextMenuKeys=   -1  'True
         RaiseCustomDrawEventOnEraseBackground=   0   'False
         RegisterForOLEDragDrop=   0
         RightToLeft     =   0
         ShadowColor     =   -1
         ShowToolTips    =   -1  'True
         SupportOLEDragImages=   -1  'True
         UseMnemonics    =   -1  'True
         UseSystemFont   =   -1  'True
         VerticalButtonPadding=   -1
         VerticalButtonSpacing=   0
         VerticalTextAlignment=   0
         WrapButtons     =   0   'False
      End
      Begin TBarCtlsLibUCtl.ToolBar tbCommon 
         Height          =   375
         Left            =   3360
         TabIndex        =   3
         Top             =   240
         Width           =   1935
         _cx             =   3413
         _cy             =   661
         AllowCustomization=   0   'False
         AlwaysDisplayButtonText=   0   'False
         AnchorHighlighting=   0   'False
         Appearance      =   0
         BackStyle       =   0
         BorderStyle     =   0
         ButtonHeight    =   0
         ButtonStyle     =   1
         ButtonTextPosition=   0
         ButtonWidth     =   0
         DisabledEvents  =   917995
         DisplayMenuDivider=   0   'False
         DisplayPartiallyClippedButtons=   0   'False
         DontRedraw      =   0   'False
         DragClickTime   =   -1
         DragDropCustomizationModifierKey=   0
         DropDownGap     =   -1
         Enabled         =   -1  'True
         FirstButtonIndentation=   0
         FocusOnClick    =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         HighlightColor  =   -1
         HorizontalButtonPadding=   -1
         HorizontalButtonSpacing=   0
         HorizontalIconCaptionGap=   -1
         HorizontalTextAlignment=   1
         HoverTime       =   -1
         InsertMarkColor =   0
         MaximumButtonWidth=   0
         MaximumTextRows =   1
         MenuBarTheme    =   1
         MenuMode        =   0   'False
         MinimumButtonWidth=   0
         MousePointer    =   0
         MultiColumn     =   0   'False
         NormalDropDownButtonStyle=   0
         OLEDragImageStyle=   0
         Orientation     =   0
         ProcessContextMenuKeys=   -1  'True
         RaiseCustomDrawEventOnEraseBackground=   0   'False
         RegisterForOLEDragDrop=   0
         RightToLeft     =   0
         ShadowColor     =   -1
         ShowToolTips    =   -1  'True
         SupportOLEDragImages=   -1  'True
         UseMnemonics    =   -1  'True
         UseSystemFont   =   -1  'True
         VerticalButtonPadding=   -1
         VerticalButtonSpacing=   0
         VerticalTextAlignment=   0
         WrapButtons     =   0   'False
      End
      Begin TBarCtlsLibUCtl.ToolBar tbMenu 
         Height          =   375
         Left            =   600
         TabIndex        =   2
         Top             =   120
         Width           =   1815
         _cx             =   3201
         _cy             =   661
         AllowCustomization=   0   'False
         AlwaysDisplayButtonText=   -1  'True
         AnchorHighlighting=   0   'False
         Appearance      =   0
         BackStyle       =   0
         BorderStyle     =   0
         ButtonHeight    =   0
         ButtonStyle     =   1
         ButtonTextPosition=   1
         ButtonWidth     =   0
         DisabledEvents  =   917995
         DisplayMenuDivider=   0   'False
         DisplayPartiallyClippedButtons=   0   'False
         DontRedraw      =   0   'False
         DragClickTime   =   -1
         DragDropCustomizationModifierKey=   0
         DropDownGap     =   -1
         Enabled         =   -1  'True
         FirstButtonIndentation=   0
         FocusOnClick    =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         HighlightColor  =   -1
         HorizontalButtonPadding=   -1
         HorizontalButtonSpacing=   0
         HorizontalIconCaptionGap=   -1
         HorizontalTextAlignment=   1
         HoverTime       =   -1
         InsertMarkColor =   0
         MaximumButtonWidth=   0
         MaximumTextRows =   1
         MenuBarTheme    =   1
         MenuMode        =   -1  'True
         MinimumButtonWidth=   0
         MousePointer    =   0
         MultiColumn     =   0   'False
         NormalDropDownButtonStyle=   1
         OLEDragImageStyle=   0
         Orientation     =   0
         ProcessContextMenuKeys=   -1  'True
         RaiseCustomDrawEventOnEraseBackground=   0   'False
         RegisterForOLEDragDrop=   0
         RightToLeft     =   0
         ShadowColor     =   -1
         ShowToolTips    =   0   'False
         SupportOLEDragImages=   -1  'True
         UseMnemonics    =   -1  'True
         UseSystemFont   =   -1  'True
         VerticalButtonPadding=   -1
         VerticalButtonSpacing=   0
         VerticalTextAlignment=   1
         WrapButtons     =   0   'False
      End
      Begin CBLCtlsLibUCtl.ComboBox cmbFontSize 
         Height          =   315
         Left            =   7680
         TabIndex        =   5
         Top             =   480
         Width           =   735
         _cx             =   1296
         _cy             =   556
         AcceptNumbersOnly=   0   'False
         Appearance      =   3
         AutoHorizontalScrolling=   -1  'True
         BackColor       =   -2147483643
         BorderStyle     =   0
         CharacterConversion=   0
         DisabledEvents  =   267489
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
         SelectionFieldHeight=   14
         Sorted          =   0   'False
         Style           =   0
         SupportOLEDragImages=   -1  'True
         UseSystemFont   =   -1  'True
         CueBanner       =   "frmMain.frx":0000
         Text            =   "frmMain.frx":0020
      End
      Begin TBarCtlsLibUCtl.ToolBar tbFont 
         Height          =   375
         Left            =   2040
         TabIndex        =   6
         Top             =   480
         Width           =   1815
         _cx             =   3201
         _cy             =   661
         AllowCustomization=   0   'False
         AlwaysDisplayButtonText=   0   'False
         AnchorHighlighting=   0   'False
         Appearance      =   0
         BackStyle       =   0
         BorderStyle     =   0
         ButtonHeight    =   0
         ButtonStyle     =   1
         ButtonTextPosition=   0
         ButtonWidth     =   0
         DisabledEvents  =   917995
         DisplayMenuDivider=   0   'False
         DisplayPartiallyClippedButtons=   0   'False
         DontRedraw      =   0   'False
         DragClickTime   =   -1
         DragDropCustomizationModifierKey=   0
         DropDownGap     =   -1
         Enabled         =   -1  'True
         FirstButtonIndentation=   0
         FocusOnClick    =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         HighlightColor  =   -1
         HorizontalButtonPadding=   -1
         HorizontalButtonSpacing=   0
         HorizontalIconCaptionGap=   -1
         HorizontalTextAlignment=   1
         HoverTime       =   -1
         InsertMarkColor =   0
         MaximumButtonWidth=   0
         MaximumTextRows =   1
         MenuBarTheme    =   1
         MenuMode        =   0   'False
         MinimumButtonWidth=   0
         MousePointer    =   0
         MultiColumn     =   0   'False
         NormalDropDownButtonStyle=   0
         OLEDragImageStyle=   0
         Orientation     =   0
         ProcessContextMenuKeys=   -1  'True
         RaiseCustomDrawEventOnEraseBackground=   0   'False
         RegisterForOLEDragDrop=   0
         RightToLeft     =   0
         ShadowColor     =   -1
         ShowToolTips    =   -1  'True
         SupportOLEDragImages=   -1  'True
         UseMnemonics    =   -1  'True
         UseSystemFont   =   -1  'True
         VerticalButtonPadding=   -1
         VerticalButtonSpacing=   0
         VerticalTextAlignment=   0
         WrapButtons     =   0   'False
      End
      Begin CBLCtlsLibUCtl.ComboBox cmbFontName 
         Height          =   360
         Left            =   5040
         TabIndex        =   7
         Top             =   600
         Width           =   2295
         _cx             =   4048
         _cy             =   635
         AcceptNumbersOnly=   0   'False
         Appearance      =   3
         AutoHorizontalScrolling=   -1  'True
         BackColor       =   -2147483643
         BorderStyle     =   0
         CharacterConversion=   0
         DisabledEvents  =   267369
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
         ListHeight      =   400
         ListInsertMarkColor=   0
         ListScrollableWidth=   0
         ListWidth       =   0
         Locale          =   1024
         MaxTextLength   =   -1
         MinVisibleItems =   30
         MousePointer    =   0
         OwnerDrawItems  =   2
         ProcessContextMenuKeys=   -1  'True
         RegisterForOLEDragDrop=   0   'False
         RightToLeft     =   0
         SelectionFieldHeight=   14
         Sorted          =   -1  'True
         Style           =   1
         SupportOLEDragImages=   -1  'True
         UseSystemFont   =   -1  'True
         CueBanner       =   "frmMain.frx":0040
         Text            =   "frmMain.frx":0060
      End
      Begin TBarCtlsLibUCtl.ToolBar tbTable 
         Height          =   375
         Left            =   240
         TabIndex        =   8
         Top             =   480
         Width           =   1815
         _cx             =   3201
         _cy             =   661
         AllowCustomization=   0   'False
         AlwaysDisplayButtonText=   0   'False
         AnchorHighlighting=   0   'False
         Appearance      =   0
         BackStyle       =   0
         BorderStyle     =   0
         ButtonHeight    =   0
         ButtonStyle     =   1
         ButtonTextPosition=   0
         ButtonWidth     =   0
         DisabledEvents  =   917995
         DisplayMenuDivider=   0   'False
         DisplayPartiallyClippedButtons=   0   'False
         DontRedraw      =   0   'False
         DragClickTime   =   -1
         DragDropCustomizationModifierKey=   0
         DropDownGap     =   -1
         Enabled         =   -1  'True
         FirstButtonIndentation=   0
         FocusOnClick    =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         HighlightColor  =   -1
         HorizontalButtonPadding=   -1
         HorizontalButtonSpacing=   0
         HorizontalIconCaptionGap=   -1
         HorizontalTextAlignment=   0
         HoverTime       =   -1
         InsertMarkColor =   0
         MaximumButtonWidth=   0
         MaximumTextRows =   1
         MenuBarTheme    =   1
         MenuMode        =   0   'False
         MinimumButtonWidth=   0
         MousePointer    =   0
         MultiColumn     =   0   'False
         NormalDropDownButtonStyle=   0
         OLEDragImageStyle=   0
         Orientation     =   0
         ProcessContextMenuKeys=   -1  'True
         RaiseCustomDrawEventOnEraseBackground=   0   'False
         RegisterForOLEDragDrop=   0
         RightToLeft     =   0
         ShadowColor     =   -1
         ShowToolTips    =   -1  'True
         SupportOLEDragImages=   -1  'True
         UseMnemonics    =   -1  'True
         UseSystemFont   =   -1  'True
         VerticalButtonPadding=   -1
         VerticalButtonSpacing=   0
         VerticalTextAlignment=   0
         WrapButtons     =   0   'False
      End
      Begin TBarCtlsLibUCtl.ToolBar tbMath 
         Height          =   375
         Left            =   9000
         TabIndex        =   9
         Top             =   120
         Width           =   1815
         _cx             =   3201
         _cy             =   661
         AllowCustomization=   0   'False
         AlwaysDisplayButtonText=   0   'False
         AnchorHighlighting=   0   'False
         Appearance      =   0
         BackStyle       =   0
         BorderStyle     =   0
         ButtonHeight    =   0
         ButtonStyle     =   1
         ButtonTextPosition=   0
         ButtonWidth     =   0
         DisabledEvents  =   917995
         DisplayMenuDivider=   0   'False
         DisplayPartiallyClippedButtons=   0   'False
         DontRedraw      =   0   'False
         DragClickTime   =   -1
         DragDropCustomizationModifierKey=   0
         DropDownGap     =   -1
         Enabled         =   -1  'True
         FirstButtonIndentation=   0
         FocusOnClick    =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         HighlightColor  =   -1
         HorizontalButtonPadding=   -1
         HorizontalButtonSpacing=   0
         HorizontalIconCaptionGap=   -1
         HorizontalTextAlignment=   0
         HoverTime       =   -1
         InsertMarkColor =   0
         MaximumButtonWidth=   0
         MaximumTextRows =   1
         MenuBarTheme    =   1
         MenuMode        =   0   'False
         MinimumButtonWidth=   0
         MousePointer    =   0
         MultiColumn     =   0   'False
         NormalDropDownButtonStyle=   0
         OLEDragImageStyle=   0
         Orientation     =   0
         ProcessContextMenuKeys=   -1  'True
         RaiseCustomDrawEventOnEraseBackground=   0   'False
         RegisterForOLEDragDrop=   0
         RightToLeft     =   0
         ShadowColor     =   -1
         ShowToolTips    =   -1  'True
         SupportOLEDragImages=   -1  'True
         UseMnemonics    =   -1  'True
         UseSystemFont   =   -1  'True
         VerticalButtonPadding=   -1
         VerticalButtonSpacing=   0
         VerticalTextAlignment=   0
         WrapButtons     =   0   'False
      End
      Begin TBarCtlsLibUCtl.ToolBar tbObject 
         Height          =   375
         Left            =   10800
         TabIndex        =   10
         Top             =   480
         Width           =   1815
         _cx             =   3201
         _cy             =   661
         AllowCustomization=   0   'False
         AlwaysDisplayButtonText=   0   'False
         AnchorHighlighting=   0   'False
         Appearance      =   0
         BackStyle       =   0
         BorderStyle     =   0
         ButtonHeight    =   0
         ButtonStyle     =   1
         ButtonTextPosition=   0
         ButtonWidth     =   0
         DisabledEvents  =   917995
         DisplayMenuDivider=   0   'False
         DisplayPartiallyClippedButtons=   0   'False
         DontRedraw      =   0   'False
         DragClickTime   =   -1
         DragDropCustomizationModifierKey=   0
         DropDownGap     =   -1
         Enabled         =   -1  'True
         FirstButtonIndentation=   0
         FocusOnClick    =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         HighlightColor  =   -1
         HorizontalButtonPadding=   -1
         HorizontalButtonSpacing=   0
         HorizontalIconCaptionGap=   -1
         HorizontalTextAlignment=   0
         HoverTime       =   -1
         InsertMarkColor =   0
         MaximumButtonWidth=   0
         MaximumTextRows =   1
         MenuBarTheme    =   1
         MenuMode        =   0   'False
         MinimumButtonWidth=   0
         MousePointer    =   0
         MultiColumn     =   0   'False
         NormalDropDownButtonStyle=   0
         OLEDragImageStyle=   0
         Orientation     =   0
         ProcessContextMenuKeys=   -1  'True
         RaiseCustomDrawEventOnEraseBackground=   0   'False
         RegisterForOLEDragDrop=   0
         RightToLeft     =   0
         ShadowColor     =   -1
         ShowToolTips    =   -1  'True
         SupportOLEDragImages=   -1  'True
         UseMnemonics    =   -1  'True
         UseSystemFont   =   -1  'True
         VerticalButtonPadding=   -1
         VerticalButtonSpacing=   0
         VerticalTextAlignment=   0
         WrapButtons     =   0   'False
      End
   End
   Begin TBarCtlsLibUCtl.ReBar ReBarU 
      Align           =   1  'Oben ausrichten
      Height          =   690
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   13710
      _cx             =   24183
      _cy             =   1217
      AllowBandReordering=   -1  'True
      Appearance      =   0
      AutoUpdateLayout=   -1  'True
      BackColor       =   -1
      BorderStyle     =   0
      DisabledEvents  =   491
      DisplayBandSeparators=   -1  'True
      DisplaySplitter =   0   'False
      DontRedraw      =   0   'False
      Enabled         =   -1  'True
      FixedBandHeight =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   -1
      HighlightColor  =   -1
      HoverTime       =   -1
      MousePointer    =   0
      Orientation     =   0
      ReplaceMDIFrameMenu=   2
      RegisterForOLEDragDrop=   0
      RightToLeft     =   0
      ShadowColor     =   -1
      SupportOLEDragImages=   -1  'True
      ToggleOnDoubleClick=   -1  'True
      UseSystemFont   =   -1  'True
      VerticalSizingGripsOnVerticalOrientation=   0   'False
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

  Implements IMDIFrame
  Implements ISubclassedWindow


  Private Const CLR_DEFAULT = &HFF000000


  Private Type BITMAPINFOHEADER '40 bytes
    biSize As Long
    biWidth As Long
    biHeight As Long
    biPlanes As Integer
    biBitCount As Integer
    biCompression As Long
    biSizeImage As Long
    biXPelsPerMeter As Long
    biYPelsPerMeter As Long
    biClrUsed As Long
    biClrImportant As Long
  End Type

  Private Type RGBQUAD
    rgbBlue As Byte
    rgbGreen As Byte
    rgbRed As Byte
    rgbReserved As Byte
  End Type

  Private Type BITMAPINFO
    bmiHeader As BITMAPINFOHEADER
    bmiColors As RGBQUAD
  End Type

  Private Type DLLVERSIONINFO
    cbSize As Long
    dwMajor As Long
    dwMinor As Long
    dwBuildNumber As Long
    dwPlatformId As Long
  End Type

  Private Type ICONINFO
    fIcon As Long
    xHotspot As Long
    yHotspot As Long
    hbmMask As Long
    hbmColor As Long
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

  Private Type POINT
    x As Long
    y As Long
  End Type

  Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
  End Type

  Private Type TOOLBARBUTTONINFO
    ID As Long
    IconIndex As Long
    Text As String
    hSubMenu As Long
  End Type

  Private Type TPMPARAMS
    cbSize As Long
    rcExclude As RECT
  End Type


  Private BANDID_MENUBAR As Long
  Private BANDID_TOOLBAR_COMMON As Long
  Private BANDID_TOOLBAR_FONT As Long
  Private BANDID_TOOLBAR_FORMAT As Long
  Private BANDID_TOOLBAR_TABLE As Long
  Private BANDID_TOOLBAR_MATH As Long
  Private BANDID_TOOLBAR_OBJECT As Long

  Private bComctl32Version600OrNewer As Boolean
  Private buttonsCommon(0 To 10) As TOOLBARBUTTONINFO
  Private buttonsFormat(0 To 6) As TOOLBARBUTTONINFO
  Private buttonsFont(0 To 7) As TOOLBARBUTTONINFO
  Private buttonsTable(0 To 10) As TOOLBARBUTTONINFO
  Private buttonsMath(0 To 8) As TOOLBARBUTTONINFO
  Private buttonsObject(0 To 1) As TOOLBARBUTTONINFO
  Private WithEvents cellBackColorPicker As clsColorPicker
Attribute cellBackColorPicker.VB_VarHelpID = -1
  Private WithEvents textBackColorPicker As clsColorPicker
Attribute textBackColorPicker.VB_VarHelpID = -1
  Private WithEvents textForeColorPicker As clsColorPicker
Attribute textForeColorPicker.VB_VarHelpID = -1
  Private WithEvents ctlHostWindow As TBarCtlsLibUCtl.ControlHostWindow
Attribute ctlHostWindow.VB_VarHelpID = -1
  Private hImgLst As Long
  Private hMenuRecentFiles As Long
  Private hMenuInsertImage As Long
  Private hMenuInsertObject As Long
  Private menuButtons(0 To 3) As TOOLBARBUTTONINFO
  Private menusToDestroy As Collection
  Private nextDocumentIndex As Long
  Private openChildren As Collection
  Private previousColor As Long
  Private recentFiles(0 To 3) As String
  Private themeableOS As Boolean


  Private Declare Function ClientToScreen Lib "user32.dll" (ByVal hWnd As Long, lpPoint As POINT) As Long
  Private Declare Function CreateCompatibleDC Lib "gdi32.dll" (ByVal hDC As Long) As Long
  Private Declare Function CreateIconIndirect Lib "user32.dll" (piconinfo As ICONINFO) As Long
  Private Declare Function CreatePopupMenu Lib "user32.dll" () As Long
  Private Declare Function CreateSolidBrush Lib "gdi32.dll" (ByVal crColor As Long) As Long
  Private Declare Function DeleteDC Lib "gdi32.dll" (ByVal hDC As Long) As Long
  Private Declare Function DeleteMenu Lib "user32.dll" (ByVal hMenu As Long, ByVal uPosition As Long, ByVal uFlags As Long) As Long
  Private Declare Function DeleteObject Lib "gdi32.dll" (ByVal hObject As Long) As Long
  Private Declare Function DestroyIcon Lib "user32.dll" (ByVal hIcon As Long) As Long
  Private Declare Function DestroyMenu Lib "user32.dll" (ByVal hMenu As Long) As Long
  Private Declare Function DllGetVersion_comctl32 Lib "comctl32.dll" Alias "DllGetVersion" (Data As DLLVERSIONINFO) As Long
  Private Declare Function DrawFocusRect Lib "user32.dll" (ByVal hDC As Long, ByRef lpRect As RECT) As Long
  Private Declare Function DrawText Lib "user32.dll" Alias "DrawTextW" (ByVal hDC As Long, ByVal lpStr As Long, ByVal nCount As Long, ByRef lpRect As RECT, ByVal wFormat As Long) As Long
  Private Declare Function EnableMenuItem Lib "user32.dll" (ByVal hMenu As Long, ByVal uIDEnableItem As Long, ByVal uEnable As Long) As Long
  Private Declare Function ExtTextOut Lib "gdi32.dll" Alias "ExtTextOutW" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal fuOptions As Long, lprc As RECTANGLE, ByVal lpString As Long, ByVal cbCount As Long, ByVal lpDx As Long) As Long
  Private Declare Function FillRect Lib "user32.dll" (ByVal hDC As Long, ByRef lpRect As RECT, ByVal hBrush As Long) As Long
  Private Declare Function FreeLibrary Lib "kernel32.dll" (ByVal hLibModule As Long) As Long
  Private Declare Function GetDC Lib "user32.dll" (ByVal hWnd As Long) As Long
  Private Declare Function GetDIBits Lib "gdi32.dll" (ByVal hDC As Long, ByVal hBitmap As Long, ByVal uStartScan As Long, ByVal cScanLines As Long, ByVal lpvBits As Long, lpbi As BITMAPINFO, ByVal uUsage As Long) As Long
  Private Declare Function GetIconInfo Lib "user32.dll" (ByVal hIcon As Long, ByRef piconinfo As ICONINFO) As Long
  Private Declare Function GetMenuItemCount Lib "user32.dll" (ByVal hMenu As Long) As Long
  Private Declare Function GetMenuItemID Lib "user32.dll" (ByVal hMenu As Long, ByVal nPos As Long) As Long
  Private Declare Function GetObjectAPI Lib "gdi32.dll" Alias "GetObjectW" (ByVal hgdiobj As Long, ByVal cbBuffer As Long, lpvObject As Any) As Long
  Private Declare Function GetObjectType Lib "gdi32.dll" (ByVal hgdiobj As Long) As Long
  Private Declare Function GetSysColor Lib "user32.dll" (ByVal nIndex As Long) As Long
  Private Declare Function ImageList_AddIcon Lib "comctl32.dll" (ByVal himl As Long, ByVal hIcon As Long) As Long
  Private Declare Function ImageList_Create Lib "comctl32.dll" (ByVal cx As Long, ByVal cy As Long, ByVal flags As Long, ByVal cInitial As Long, ByVal cGrow As Long) As Long
  Private Declare Function ImageList_Destroy Lib "comctl32.dll" (ByVal himl As Long) As Long
  Private Declare Function ImageList_GetIcon Lib "comctl32.dll" (ByVal himl As Long, ByVal i As Long, ByVal flags As Long) As Long
  Private Declare Function ImageList_GetImageCount Lib "comctl32.dll" (ByVal himl As Long) As Long
  Private Declare Function ImageList_ReplaceIcon Lib "comctl32.dll" (ByVal himl As Long, ByVal i As Long, ByVal hIcon As Long) As Long
  Private Declare Sub InitCommonControls Lib "comctl32.dll" ()
  Private Declare Function InsertMenuItem Lib "user32.dll" Alias "InsertMenuItemW" (ByVal hMenu As Long, ByVal uItem As Long, ByVal fByPosition As Long, lpmii As MENUITEMINFO) As Long
  Private Declare Function LoadImage Lib "user32.dll" Alias "LoadImageW" (ByVal hInst As Long, ByVal lpszName As Long, ByVal uType As Long, ByVal cxDesired As Long, ByVal cyDesired As Long, ByVal fuLoad As Long) As Long
  Private Declare Function LoadLibrary Lib "kernel32.dll" Alias "LoadLibraryW" (ByVal lpLibFileName As Long) As Long
  Private Declare Function lstrlen Lib "kernel32.dll" Alias "lstrlenW" (ByVal lpString As Long) As Long
  Private Declare Function MapWindowPoints Lib "user32.dll" (ByVal hWndFrom As Long, ByVal hWndTo As Long, ByVal lpPoints As Long, ByVal cPoints As Long) As Long
  Private Declare Function OleTranslateColor Lib "oleaut32.dll" (ByVal lOleColor As Long, ByVal lHPalette As Long, ByRef lColorRef As Long) As Long
  Private Declare Function PathFileExists Lib "shlwapi.dll" Alias "PathFileExistsW" (ByVal pszPath As Long) As Long
  Private Declare Sub PathStripPath Lib "shlwapi.dll" Alias "PathStripPathW" (ByVal pszPath As Long)
  Private Declare Function ReleaseDC Lib "user32.dll" (ByVal hWnd As Long, ByVal hDC As Long) As Long
  Private Declare Function SelectObject Lib "gdi32.dll" (ByVal hDC As Long, ByVal hObject As Long) As Long
  Private Declare Function SendMessageAsLong Lib "user32.dll" Alias "SendMessageW" (ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
  Private Declare Function SetBkColor Lib "gdi32.dll" (ByVal hDC As Long, ByVal crColor As Long) As Long
  Private Declare Function SetDIBits Lib "gdi32.dll" (ByVal hDC As Long, ByVal hBMP As Long, ByVal uStartScan As Long, ByVal cScanLines As Long, ByVal lpvBits As Long, lpbi As BITMAPINFO, ByVal fuColorUse As Long) As Long
  Private Declare Function SetGraphicsMode Lib "gdi32.dll" (ByVal hDC As Long, ByVal iMode As Long) As Long
  Private Declare Function SetTextColor Lib "gdi32.dll" (ByVal hDC As Long, ByVal crColor As Long) As Long
  Private Declare Function SetWindowTheme Lib "uxtheme.dll" (ByVal hWnd As Long, ByVal pSubAppName As Long, ByVal pSubIDList As Long) As Long
  Private Declare Function TrackPopupMenuEx Lib "user32.dll" (ByVal hMenu As Long, ByVal fuFlags As Long, ByVal x As Long, ByVal y As Long, ByVal hWnd As Long, ByRef lptpm As TPMPARAMS) As Long
  
  Private Declare Function HunSpellFreeHyphenationDictionary Lib "SpellCheck.dll" (ByVal langid As Long) As Long


Private Sub IMDIFrame_ActivatedDocument(ByVal doc As IDocument)
  If Not (doc Is Nothing) Then
    On Error Resume Next
    #If False Then
      doc.CellBackColor = cellBackColorPicker.SelectedColor
      doc.TextBackColor = textBackColorPicker.SelectedColor
      doc.TextForeColor = textForeColorPicker.SelectedColor
    #Else
      cellBackColorPicker.SelectedColor = doc.CellBackColor
      textBackColorPicker.SelectedColor = doc.TextBackColor
      textForeColorPicker.SelectedColor = doc.TextForeColor
      UpdateColorInColorIcon cellBackColorPicker.SelectedColor, hImgLst, 38, 39
      UpdateColorInColorIcon textBackColorPicker.SelectedColor, hImgLst, 24, 26
      UpdateColorInColorIcon textForeColorPicker.SelectedColor, hImgLst, 25, 27
    #End If
  End If
End Sub

Private Sub IMDIFrame_CheckCommand(ByVal commandID As Long, ByVal check As Boolean)
  On Error Resume Next
  tbCommon.Buttons(commandID, btitID).SelectionState = IIf(check, ssChecked, ssUnchecked)
  tbFormat.Buttons(commandID, btitID).SelectionState = IIf(check, ssChecked, ssUnchecked)
  tbFont.Buttons(commandID, btitID).SelectionState = IIf(check, ssChecked, ssUnchecked)
  tbTable.Buttons(commandID, btitID).SelectionState = IIf(check, ssChecked, ssUnchecked)
  tbMath.Buttons(commandID, btitID).SelectionState = IIf(check, ssChecked, ssUnchecked)
  tbObject.Buttons(commandID, btitID).SelectionState = IIf(check, ssChecked, ssUnchecked)
End Sub

Private Sub IMDIFrame_CheckTriStateCommand(ByVal commandID As Long, ByVal selState As TBarCtlsLibUCtl.SelectionStateConstants)
  On Error Resume Next
  tbCommon.Buttons(commandID, btitID).SelectionState = selState
  tbFormat.Buttons(commandID, btitID).SelectionState = selState
  tbFont.Buttons(commandID, btitID).SelectionState = selState
  tbTable.Buttons(commandID, btitID).SelectionState = selState
  tbMath.Buttons(commandID, btitID).SelectionState = selState
  tbObject.Buttons(commandID, btitID).SelectionState = selState
End Sub

Private Sub IMDIFrame_ClosingDocument(ByVal windowID As Long)
  Dim frm As Form

  On Error GoTo InvalidWindowID
  Set frm = openChildren.Item("k" & windowID)
  If Not (frm Is Nothing) Then
    openChildren.Remove "k" & windowID
    Set frm = Nothing
  End If
  On Error Resume Next
  ReBarU.Bands(BANDID_TOOLBAR_FORMAT, bitID).Visible = (openChildren.Count > 0)
  ReBarU.Bands(BANDID_TOOLBAR_FONT, bitID).Visible = (openChildren.Count > 0)
  ReBarU.Bands(BANDID_TOOLBAR_OBJECT, bitID).Visible = (openChildren.Count > 0)
  ReBarU.Bands(BANDID_TOOLBAR_TABLE, bitID).Visible = (openChildren.Count > 0)
  ReBarU.Bands(BANDID_TOOLBAR_MATH, bitID).Visible = (openChildren.Count > 0)
  If openChildren.Count = 0 Then tbMenu.CommandEnabled(CMDID_FILE_SAVE) = False
  tbMenu.CommandEnabled(CMDID_FILE_PRINT) = (openChildren.Count > 0)
  If openChildren.Count = 0 Then tbCommon.CommandEnabled(CMDID_FILE_SAVE) = False
  tbCommon.CommandEnabled(CMDID_FILE_PRINT) = (openChildren.Count > 0)
  tbCommon.CommandEnabled(CMDID_EDIT_FIND) = (openChildren.Count > 0)
  tbCommon.CommandEnabled(CMDID_EDIT_SPELLCHECKING) = (openChildren.Count > 0)

InvalidWindowID:
  On Error Resume Next
  tbMenu.Buttons(CMDID_WINDOW, btitID).Visible = (openChildren.Count > 0)
  ReBarU.Bands(BANDID_MENUBAR, bitID).IdealWidth = tbMenu.IdealWidth
End Sub

Private Sub IMDIFrame_EnableDisableCommand(ByVal commandID As Long, ByVal enable As Boolean)
  On Error Resume Next
  tbMenu.CommandEnabled(commandID) = enable
  tbCommon.CommandEnabled(commandID) = enable
  tbFormat.CommandEnabled(commandID) = enable
  tbFont.CommandEnabled(commandID) = enable
  tbTable.CommandEnabled(commandID) = enable
  tbMath.CommandEnabled(commandID) = enable
  tbObject.CommandEnabled(commandID) = enable
End Sub

Private Sub IMDIFrame_OpenedDocument(ByVal frm As Form, windowID As Long)
  If Not (frm Is Nothing) Then
    windowID = nextDocumentIndex
    nextDocumentIndex = nextDocumentIndex + 1

    openChildren.Add frm, "k" & windowID

    On Error Resume Next
    tbMenu.Buttons(CMDID_WINDOW, btitID).Visible = True
    ReBarU.Bands(BANDID_MENUBAR, bitID).IdealWidth = tbMenu.IdealWidth
    ReBarU.Bands(BANDID_TOOLBAR_FORMAT, bitID).Visible = True
    ReBarU.Bands(BANDID_TOOLBAR_FONT, bitID).Visible = True
    ReBarU.Bands(BANDID_TOOLBAR_OBJECT, bitID).Visible = True
    ReBarU.Bands(BANDID_TOOLBAR_TABLE, bitID).Visible = True
    ReBarU.Bands(BANDID_TOOLBAR_MATH, bitID).Visible = True

    tbMenu.CommandEnabled(CMDID_FILE_PRINT) = (openChildren.Count > 0)
    tbCommon.CommandEnabled(CMDID_FILE_PRINT) = (openChildren.Count > 0)
    tbCommon.CommandEnabled(CMDID_EDIT_FIND) = (openChildren.Count > 0)
    tbCommon.CommandEnabled(CMDID_EDIT_SPELLCHECKING) = (openChildren.Count > 0)
  End If
End Sub

Private Sub IMDIFrame_UpdateFontName(ByVal fontName As String)
  cmbFontName.Tag = "IgnoreEvent"
  Set cmbFontName.SelectedItem = cmbFontName.FindItemByText(fontName)
  cmbFontName.Tag = ""
End Sub

Private Sub IMDIFrame_UpdateFontSize(ByVal fontSize As Single)
  Dim i As Long
  
  If fontSize < 0 Then
    cmbFontSize.Text = ""
  Else
    With cmbFontSize.ComboItems
      For i = 0 To .Count - 1
        If .Item(i).ItemData = CLng(fontSize * 10) Then
          Set cmbFontSize.SelectedItem = .Item(i)
          Exit For
        End If
      Next i
      If i = .Count Then
        If fontSize - CLng(fontSize) = 0 Then
          cmbFontSize.Text = CStr(CLng(fontSize))
        Else
          cmbFontSize.Text = Format$(fontSize, "0.0")
        End If
      End If
    End With
  End If
End Sub

Private Function ISubclassedWindow_HandleMessage(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long, ByVal eSubclassID As EnumSubclassID, bCallDefProc As Boolean) As Long
  Dim lRet As Long

  On Error GoTo StdHandler_Error
  Select Case eSubclassID
    Case EnumSubclassID.escidFrmMain
      lRet = HandleMessage_Form(hWnd, uMsg, wParam, lParam, bCallDefProc)
    Case Else
      Debug.Print "frmMain.ISubclassedWindow_HandleMessage: Unknown Subclassing ID " & CStr(eSubclassID)
  End Select

StdHandler_Ende:
  ISubclassedWindow_HandleMessage = lRet
  Exit Function

StdHandler_Error:
  Debug.Print "Error in frmMain.ISubclassedWindow_HandleMessage (SubclassID=" & CStr(eSubclassID) & ": ", Err.Number, Err.Description
  Resume StdHandler_Ende
End Function

Private Function HandleMessage_Form(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long, bCallDefProc As Boolean) As Long
  Const WM_NOTIFYFORMAT = &H55
  Const WM_USER = &H400
  Const OCM__BASE = WM_USER + &H1C00
  Dim lRet As Long

  On Error GoTo StdHandler_Error
  Select Case uMsg
    Case WM_NOTIFYFORMAT
      ' give the control a chance to request Unicode notifications
      lRet = SendMessageAsLong(wParam, OCM__BASE + uMsg, wParam, lParam)

      bCallDefProc = False
  End Select

StdHandler_Ende:
  HandleMessage_Form = lRet
  Exit Function

StdHandler_Error:
  Debug.Print "Error in frmMain.HandleMessage_Form: ", Err.Number, Err.Description
  Resume StdHandler_Ende
End Function


Private Sub cmbFontName_FreeItemData(ByVal comboItem As CBLCtlsLibUCtl.IComboBoxItem)
  Const OBJ_FONT = 6
  Dim h As Long

  h = comboItem.ItemData
  If GetObjectType(h) = OBJ_FONT Then DeleteObject h
End Sub

Private Sub cmbFontName_ItemMouseEnter(ByVal comboItem As CBLCtlsLibUCtl.IComboBoxItem, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As CBLCtlsLibUCtl.HitTestConstants)
  Dim doc As IDocument

  Set doc = Me.ActiveForm
  If Not (doc Is Nothing) And Not (comboItem Is Nothing) Then
    doc.SetFontName comboItem.Text
  End If
End Sub

Private Sub cmbFontName_ListCloseUp()
  Dim doc As IDocument

  Set doc = Me.ActiveForm
  If Not (doc Is Nothing) Then
    doc.EndFontChangePreview False
    doc.SetFontName cmbFontName.Text
  End If
End Sub

Private Sub cmbFontName_ListDropDown()
  Dim doc As IDocument

  Set doc = Me.ActiveForm
  If Not (doc Is Nothing) Then
    doc.StartFontChangePreview
  End If
End Sub

Private Sub cmbFontName_MeasureItem(ByVal comboItem As CBLCtlsLibUCtl.IComboBoxItem, ItemHeight As stdole.OLE_YSIZE_PIXELS)
  Const DT_CALCRECT = &H400
  Const DT_LEFT = &H0
  Const DT_RTLREADING = &H20000
  Const DT_SINGLELINE = &H20
  Const DT_VCENTER = &H4
  Const OBJ_FONT = 6
  Dim flags As Long
  Dim hDC As Long
  Dim hDCCompatible As Long
  Dim hFont As Long
  Dim hOldFont As Long
  Dim rc As RECT

  If Not (comboItem Is Nothing) Then
    hDCCompatible = GetDC(0)
    If hDCCompatible Then
      hDC = CreateCompatibleDC(hDCCompatible)
      If hDC Then
        ' measure item text
        hFont = comboItem.ItemData
        If GetObjectType(hFont) = OBJ_FONT Then
          hOldFont = SelectObject(hDC, hFont)
        End If
        flags = DT_LEFT Or DT_SINGLELINE
        DrawText hDC, StrPtr(comboItem.Text), -1, rc, flags Or DT_CALCRECT
        ItemHeight = rc.Bottom - rc.Top

        If hOldFont Then
          SelectObject hDC, hOldFont
        End If
        DeleteDC hDC
      End If
      ReleaseDC 0, hDCCompatible
    End If
  End If
End Sub

Private Sub cmbFontName_OwnerDrawItem(ByVal comboItem As CBLCtlsLibUCtl.IComboBoxItem, ByVal requiredAction As CBLCtlsLibUCtl.OwnerDrawActionConstants, ByVal itemState As CBLCtlsLibUCtl.OwnerDrawItemStateConstants, ByVal hDC As Long, drawingRectangle As CBLCtlsLibUCtl.RECTANGLE)
  Const COLOR_BTNFACE = 15
  Const COLOR_3DFACE = COLOR_BTNFACE
  Const COLOR_BTNTEXT = 18
  Const COLOR_HIGHLIGHT = 13
  Const COLOR_HIGHLIGHTTEXT = 14
  Const DT_LEFT = &H0
  Const DT_RTLREADING = &H20000
  Const DT_SINGLELINE = &H20
  Const DT_VCENTER = &H4
  Const OBJ_FONT = 6
  Const WM_GETFONT = &H31
  Dim backClr As Long
  Dim flags As Long
  Dim foreClr As Long
  Dim hBrush As Long
  Dim hFont As Long
  Dim hOldFont As Long
  Dim oldBkColor As Long
  Dim oldTextColor As Long
  Dim rc As RECT

  If Not (comboItem Is Nothing) Then
    If itemState And OwnerDrawItemStateConstants.odisSelected Then
      If itemState And OwnerDrawItemStateConstants.odisFocus Then
        backClr = GetSysColor(COLOR_HIGHLIGHT)
        foreClr = GetSysColor(COLOR_HIGHLIGHTTEXT)
      Else
        backClr = GetSysColor(COLOR_3DFACE)
        foreClr = GetSysColor(COLOR_BTNTEXT)
      End If
    Else
      backClr = TranslateColor(cmbFontName.BackColor)
      foreClr = TranslateColor(cmbFontName.ForeColor)
    End If

    ' draw item background
    LSet rc = drawingRectangle
    hBrush = CreateSolidBrush(backClr)
    If hBrush Then
      FillRect hDC, rc, hBrush
      DeleteObject hBrush
    End If

    ' draw item text
    If itemState And OwnerDrawItemStateConstants.odisDrawingIntoSelectionField Then
      hFont = SendMessageAsLong(cmbFontName.hWnd, WM_GETFONT, 0, 0)
    Else
      hFont = comboItem.ItemData
    End If
    If GetObjectType(hFont) = OBJ_FONT Then
      hOldFont = SelectObject(hDC, hFont)
    End If
    oldBkColor = SetBkColor(hDC, backClr)
    oldTextColor = SetTextColor(hDC, foreClr)
    flags = DT_LEFT Or DT_VCENTER Or DT_SINGLELINE
    If cmbFontName.RightToLeft And RightToLeftConstants.rtlText Then
      flags = flags Or DT_RTLREADING
    End If
    DrawText hDC, StrPtr(comboItem.Text), -1, rc, flags

    SetTextColor hDC, oldTextColor
    SetBkColor hDC, oldBkColor
    If hOldFont Then
      SelectObject hDC, hOldFont
    End If

    ' draw the focus rectangle
    If (itemState And (OwnerDrawItemStateConstants.odisSelected Or OwnerDrawItemStateConstants.odisFocus Or OwnerDrawItemStateConstants.odisNoFocusRectangle)) = (OwnerDrawItemStateConstants.odisSelected Or OwnerDrawItemStateConstants.odisFocus) Then
      DrawFocusRect hDC, rc
    End If
  End If
End Sub

Private Sub cmbFontName_SelectionChanged(ByVal previousSelectedItem As CBLCtlsLibUCtl.IComboBoxItem, ByVal newSelectedItem As CBLCtlsLibUCtl.IComboBoxItem)
  Dim doc As IDocument

  If cmbFontName.Tag = "" Then
    Set doc = Me.ActiveForm
    If Not (doc Is Nothing) And Not (newSelectedItem Is Nothing) Then
      doc.SetFontName newSelectedItem.Text
      doc.SetFocus
    End If
  End If
End Sub

Private Sub cmbFontSize_ItemMouseEnter(ByVal comboItem As CBLCtlsLibUCtl.IComboBoxItem, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As CBLCtlsLibUCtl.HitTestConstants)
  Dim doc As IDocument

  Set doc = Me.ActiveForm
  If Not (doc Is Nothing) And Not (comboItem Is Nothing) Then
    doc.SetFontSize comboItem.ItemData / 10
  End If
End Sub

Private Sub cmbFontSize_KeyPress(KeyAscii As Integer)
  Dim doc As IDocument

  If KeyAscii = vbKeyReturn Then
    Set doc = Me.ActiveForm
    If Not (doc Is Nothing) Then
      On Error Resume Next
      doc.SetFontSize CSng(cmbFontSize.Text)
      doc.SetFocus
    End If
  End If
End Sub

Private Sub cmbFontSize_ListCloseUp()
  Dim doc As IDocument

  Set doc = Me.ActiveForm
  If Not (doc Is Nothing) Then
    doc.EndFontChangePreview False
    doc.SetFontSize CSng(cmbFontSize.Text)
  End If
End Sub

Private Sub cmbFontSize_ListDropDown()
  Dim doc As IDocument

  Set doc = Me.ActiveForm
  If Not (doc Is Nothing) Then
    doc.StartFontChangePreview
  End If
End Sub

Private Sub cmbFontSize_SelectionChanged(ByVal previousSelectedItem As CBLCtlsLibUCtl.IComboBoxItem, ByVal newSelectedItem As CBLCtlsLibUCtl.IComboBoxItem)
  Dim doc As IDocument

  Set doc = Me.ActiveForm
  If Not (doc Is Nothing) And Not (newSelectedItem Is Nothing) Then
    doc.SetFontSize newSelectedItem.ItemData / 10
    doc.SetFocus
  End If
End Sub

Private Sub MDIForm_Initialize()
  Const ILC_COLOR24 = &H18
  Const ILC_COLOR32 = &H20
  Const ILC_MASK = &H1
  Const IMAGE_ICON = 1
  Const LR_DEFAULTSIZE = &H40
  Const LR_LOADFROMFILE = &H10
  Dim DLLVerData As DLLVERSIONINFO
  Dim fileName As String
  Dim hIcon As Long
  Dim hMod As Long
  Dim i As Long
  Dim iconsDir As String
  Dim iconSize As Long
  Dim tableBackColors As Variant
  Dim tableBackColorNames As Variant

  InitCommonControls

  hMod = LoadLibrary(StrPtr("uxtheme.dll"))
  If hMod Then
    themeableOS = True
    FreeLibrary hMod
  End If

  With DLLVerData
    .cbSize = LenB(DLLVerData)
    DllGetVersion_comctl32 DLLVerData
    bComctl32Version600OrNewer = (.dwMajor >= 6)
  End With

  iconSize = 16
  hImgLst = ImageList_Create(iconSize, iconSize, IIf(bComctl32Version600OrNewer, ILC_COLOR32, ILC_COLOR24) Or ILC_MASK, 51, 0)
  If Right$(App.Path, 3) = "bin" Then
    iconsDir = App.Path & "\..\res\"
  Else
    iconsDir = App.Path & "\res\"
  End If
  iconsDir = iconsDir & iconSize & "x" & iconSize & IIf(bComctl32Version600OrNewer, "x32bpp\", "x8bpp\")
  iconsDir = iconsDir & "normal\"
  For i = 1 To 11
    fileName = Choose(i, "New", "Open", "Save", "Print", "Cut", "Copy", "Paste", "Find", "Undo", "Redo", "SpellChecking") & ".ico"
    hIcon = LoadImage(0, StrPtr(iconsDir & fileName), IMAGE_ICON, iconSize, iconSize, LR_LOADFROMFILE Or LR_DEFAULTSIZE)
    If hIcon Then
      ImageList_AddIcon hImgLst, hIcon
      DestroyIcon hIcon
    End If
  Next i
  For i = 1 To 7
    fileName = Choose(i, "AlignLeft", "AlignCenter", "AlignRight", "AlignJustify", "BulletList", "NumberedList", "Convert to Hyperlink") & ".ico"
    hIcon = LoadImage(0, StrPtr(iconsDir & fileName), IMAGE_ICON, iconSize, iconSize, LR_LOADFROMFILE Or LR_DEFAULTSIZE)
    If hIcon Then
      ImageList_AddIcon hImgLst, hIcon
      DestroyIcon hIcon
    End If
  Next i
  For i = 1 To 10
    ' load the icons for colors twice - the second ones will be replaced by dynamic ones which contain the
    ' currently selected color
    fileName = Choose(i, "Bold", "Italic", "Underline", "Strike", "Subscript", "Superscript", "Text BackColor", "Text ForeColor", "Text BackColor", "Text ForeColor") & ".ico"
    hIcon = LoadImage(0, StrPtr(iconsDir & fileName), IMAGE_ICON, iconSize, iconSize, LR_LOADFROMFILE Or LR_DEFAULTSIZE)
    If hIcon Then
      ImageList_AddIcon hImgLst, hIcon
      DestroyIcon hIcon
    End If
  Next i
  For i = 1 To 12
    ' load the icons for colors twice - the second ones will be replaced by dynamic ones which contain the
    ' currently selected color
    FileName = Choose(i, "InsertTable", "DeleteTable", "InsertTableRow", "DeleteTableRow", "InsertTableColumn", "DeleteTableColumn", "MergeCells", "AlignTop", "AlignMiddle", "AlignBottom", "Cell BackColor", "Cell BackColor") & ".ico"
    hIcon = LoadImage(0, StrPtr(iconsDir & FileName), IMAGE_ICON, iconSize, iconSize, LR_LOADFROMFILE Or LR_DEFAULTSIZE)
    If hIcon Then
      ImageList_AddIcon hImgLst, hIcon
      DestroyIcon hIcon
    End If
  Next i
  For i = 1 To 9
    FileName = Choose(i, "Toggle Math Zone", "Build Up Math", "Build Down Math", "Insert Root", "Insert Sum", "Insert Product", "Insert Integral", "Insert Limes", "Insert Matrix") & ".ico"
    hIcon = LoadImage(0, StrPtr(iconsDir & FileName), IMAGE_ICON, iconSize, iconSize, LR_LOADFROMFILE Or LR_DEFAULTSIZE)
    If hIcon Then
      ImageList_AddIcon hImgLst, hIcon
      DestroyIcon hIcon
    End If
  Next i
  For i = 1 To 2
    fileName = Choose(i, "Insert Image", "Insert Object") & ".ico"
    hIcon = LoadImage(0, StrPtr(iconsDir & fileName), IMAGE_ICON, iconSize, iconSize, LR_LOADFROMFILE Or LR_DEFAULTSIZE)
    If hIcon Then
      ImageList_AddIcon hImgLst, hIcon
      DestroyIcon hIcon
    End If
  Next i

  Set textBackColorPicker = New clsColorPicker
  textBackColorPicker.DefaultColorText = "None"
  textBackColorPicker.IgnoreFirstLeftMouseButtonUp = True
  textBackColorPicker.TrackSelection = True
  Set textForeColorPicker = New clsColorPicker
  textForeColorPicker.DefaultColorText = "Automatic"
  textForeColorPicker.IgnoreFirstLeftMouseButtonUp = True
  textForeColorPicker.TrackSelection = True
  Set cellBackColorPicker = New clsColorPicker
  cellBackColorPicker.DefaultColorText = "None"
  cellBackColorPicker.MoreColorsText = ""
  cellBackColorPicker.IgnoreFirstLeftMouseButtonUp = True
  cellBackColorPicker.TrackSelection = True
  tableBackColors = Array(RGB(0, 0, 0), RGB(0, 0, 255), RGB(0, 255, 255), RGB(0, 255, 0), RGB(255, 0, 255), RGB(255, 0, 0), RGB(255, 255, 0), RGB(255, 255, 255), RGB(0, 0, 128), RGB(0, 128, 128), RGB(0, 128, 0), RGB(128, 0, 128), RGB(128, 0, 0), RGB(128, 128, 0), RGB(128, 128, 128), RGB(192, 192, 192))
  tableBackColorNames = Array("Black", "Blue", "Cyan", "Lime", "Magenta", "Red", "Yellow", "White", "Navy Blue", "Teal", "Green", "Purple", "Maroon", "Olive", "Gray", "Silver")
  cellBackColorPicker.SetPredefinedColors tableBackColors, tableBackColorNames

  Set openChildren = New Collection
  nextDocumentIndex = 1
End Sub

Private Sub MDIForm_Load()
  Const WM_GETFONT = &H31
  Dim fontSize As Long
  Dim hDefaultFont As Long

  Subclass

  hDefaultFont = SendMessageAsLong(cmbFontName.hWnd, WM_GETFONT, 0, 0)
  GetObjectAPI hDefaultFont, LenB(lfDefault), lfDefault
  InsertFonts

  Set menusToDestroy = New Collection
  CreateMenus
  For fontSize = 6 To 10 Step 1
    cmbFontSize.ComboItems.Add CStr(fontSize), , fontSize * 10
  Next fontSize
  cmbFontSize.ComboItems.Add Format$(10.5, "0.0"), , 105
  For fontSize = 11 To 16 Step 1
    cmbFontSize.ComboItems.Add CStr(fontSize), , fontSize * 10
  Next fontSize
  For fontSize = 18 To 28 Step 2
    cmbFontSize.ComboItems.Add CStr(fontSize), , fontSize * 10
  Next fontSize
  For fontSize = 32 To 48 Step 4
    cmbFontSize.ComboItems.Add CStr(fontSize), , fontSize * 10
  Next fontSize
  For fontSize = 54 To 72 Step 6
    cmbFontSize.ComboItems.Add CStr(fontSize), , fontSize * 10
  Next fontSize
  For fontSize = 80 To 96 Step 8
    cmbFontSize.ComboItems.Add CStr(fontSize), , fontSize * 10
  Next fontSize
  Set cmbFontSize.SelectedItem = cmbFontSize.ComboItems(2)

  Me.Show

  InsertButtons
  InsertBands

  ExecuteCommand CMDID_FILE_NEW
End Sub

Private Sub MDIForm_Terminate()
  If hImgLst Then Call ImageList_Destroy(hImgLst)
  Set ctlHostWindow = Nothing

  If g_hModSpellCheck Then
    Call HunSpellFreeHyphenationDictionary(1031)
    Call HunSpellFreeHyphenationDictionary(1033)
    Call FreeLibrary(g_hModSpellCheck)
    g_hModSpellCheck = 0
    g_pHyphenateFunction = 0
  End If
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
  Const OBJ_FONT = 6
  Dim comboItem As ComboBoxItem
  Dim h As Long
  Dim i As Long

  ' it is important to clean up contained windows on destruction
  tbFont.Buttons.Item(CMDID_FONT_FONTNAME, btitID).SetContainedWindow 0
  tbFont.Buttons.Item(CMDID_FONT_FONTSIZE, btitID).SetContainedWindow 0

  For i = 1 To menusToDestroy.Count
    DestroyMenu CLng(menusToDestroy(i))
  Next i
  ' The FreeItemData won't get fired on program termination (actually it gets fired, but we won't
  ' receive it anymore). So ensure the fonts get freed.
  For Each comboItem In cmbFontName.ComboItems
    h = comboItem.ItemData
    If GetObjectType(h) = OBJ_FONT Then
      DeleteObject h
    End If
  Next comboItem
  Set textBackColorPicker = Nothing
  Set textForeColorPicker = Nothing
  Set cellBackColorPicker = Nothing
  If Not UnSubclassWindow(Me.hWnd, EnumSubclassID.escidFrmMain) Then
    Debug.Print "UnSubclassing failed!"
  End If
End Sub

Private Sub tbCommon_ButtonGetInfoTipText(ByVal toolButton As TBarCtlsLibUCtl.IToolBarButton, ByVal maxInfoTipLength As Long, infoTipText As String, abortToolTip As Boolean)
  Dim doc As IDocument

  If toolButton.ButtonType <> btyPlaceholder Then
    Set doc = Me.ActiveForm
    If Not (doc Is Nothing) Then
      infoTipText = doc.GetCommandToolTip(toolButton.ID)
    End If
    If infoTipText = "" Then
      infoTipText = buttonsCommon(toolButton.ButtonData).Text
    End If
  End If
End Sub

Private Sub tbCommon_DropDown(ByVal toolButton As TBarCtlsLibUCtl.IToolBarButton, buttonRectangle As TBarCtlsLibUCtl.RECTANGLE, furtherProcessing As TBarCtlsLibUCtl.DropDownReturnValuesConstants)
  Const TPM_LEFTALIGN = &H0&
  Const TPM_TOPALIGN = &H0&
  Const TPM_VERPOSANIMATION = &H1000&
  Const TPM_VERTICAL = &H40&
  Dim ptMenu As POINT
  Dim tpmexParams As TPMPARAMS

  If Not (toolButton Is Nothing) Then
    If toolButton.ID = CMDID_FILE_OPEN Then
      ptMenu.x = buttonRectangle.Left
      ptMenu.y = buttonRectangle.Bottom
      ClientToScreen tbCommon.hWnd, ptMenu
      With tpmexParams
        .cbSize = LenB(tpmexParams)
        LSet .rcExclude = buttonRectangle
      End With
      TrackPopupMenuEx hMenuRecentFiles, TPM_LEFTALIGN Or TPM_TOPALIGN Or TPM_VERTICAL Or TPM_VERPOSANIMATION, ptMenu.x, ptMenu.y, tbCommon.hWnd, tpmexParams
    End If
  End If
End Sub

Private Sub tbCommon_ExecuteCommand(ByVal commandID As Long, ByVal toolButton As TBarCtlsLibUCtl.IToolBarButton, ByVal commandOrigin As TBarCtlsLibUCtl.CommandOriginConstants, forwardMessage As Boolean)
  ExecuteCommand commandID
  forwardMessage = False
End Sub

Private Sub tbFont_ButtonGetInfoTipText(ByVal toolButton As TBarCtlsLibUCtl.IToolBarButton, ByVal maxInfoTipLength As Long, infoTipText As String, abortToolTip As Boolean)
  Dim doc As IDocument

  If toolButton.ButtonType <> btyPlaceholder Then
    Set doc = Me.ActiveForm
    If Not (doc Is Nothing) Then
      infoTipText = doc.GetCommandToolTip(toolButton.ID)
    End If
    If infoTipText = "" Then
      infoTipText = buttonsFont(toolButton.ButtonData).Text
    End If
  End If
End Sub

Private Sub tbFont_DropDown(ByVal toolButton As TBarCtlsLibUCtl.IToolBarButton, buttonRectangle As TBarCtlsLibUCtl.RECTANGLE, furtherProcessing As TBarCtlsLibUCtl.DropDownReturnValuesConstants)
  Dim doc As IDocument

  If Not (toolButton Is Nothing) Then
    Set doc = Me.ActiveForm
    If Not (doc Is Nothing) Then
      If toolButton.ID = CMDID_FONT_TEXTBACKCOLOR Then
        previousColor = doc.SelectionTextBackColor
        textBackColorPicker.hWndParent = Me.hWnd
        textBackColorPicker.hwndOwner = tbFont.hWnd
        textBackColorPicker.SelectedColor = IIf(previousColor = -1, textBackColorPicker.DefaultColor, previousColor)
        MapWindowPoints tbFont.hWnd, 0, VarPtr(buttonRectangle), 2

        doc.StartFontChangePreview
        If Not textBackColorPicker.Popup(buttonRectangle) Then
          ' cancelled
          doc.EndFontChangePreview False
          doc.TextBackColor = previousColor
        Else
          UpdateColorInColorIcon textBackColorPicker.SelectedColor, hImgLst, 24, 26
          doc.EndFontChangePreview False
          doc.TextBackColor = IIf(textBackColorPicker.SelectedColor = CLR_DEFAULT, -1, textBackColorPicker.SelectedColor)
        End If

      ElseIf toolButton.ID = CMDID_FONT_TEXTFORECOLOR Then
        previousColor = doc.SelectionTextForeColor
        textForeColorPicker.hWndParent = Me.hWnd
        textForeColorPicker.hwndOwner = tbFont.hWnd
        textForeColorPicker.SelectedColor = IIf(previousColor = -1, textForeColorPicker.DefaultColor, previousColor)
        MapWindowPoints tbFont.hWnd, 0, VarPtr(buttonRectangle), 2

        doc.StartFontChangePreview
        If Not textForeColorPicker.Popup(buttonRectangle) Then
          ' cancelled
          doc.EndFontChangePreview False
          doc.TextForeColor = previousColor
        Else
          UpdateColorInColorIcon textForeColorPicker.SelectedColor, hImgLst, 25, 27
          doc.EndFontChangePreview False
          doc.TextForeColor = IIf(textForeColorPicker.SelectedColor = CLR_DEFAULT, -1, textForeColorPicker.SelectedColor)
        End If
      End If
    End If
  End If
End Sub

Private Sub tbFont_ExecuteCommand(ByVal commandID As Long, ByVal toolButton As TBarCtlsLibUCtl.IToolBarButton, ByVal commandOrigin As TBarCtlsLibUCtl.CommandOriginConstants, forwardMessage As Boolean)
  ExecuteCommand commandID
  forwardMessage = False
End Sub

Private Sub tbFormat_ButtonGetInfoTipText(ByVal toolButton As TBarCtlsLibUCtl.IToolBarButton, ByVal maxInfoTipLength As Long, infoTipText As String, abortToolTip As Boolean)
  Dim doc As IDocument

  If toolButton.ButtonType <> btyPlaceholder Then
    Set doc = Me.ActiveForm
    If Not (doc Is Nothing) Then
      infoTipText = doc.GetCommandToolTip(toolButton.ID)
    End If
    If infoTipText = "" Then
      infoTipText = buttonsFormat(toolButton.ButtonData).Text
    End If
  End If
End Sub

Private Sub tbFormat_ExecuteCommand(ByVal commandID As Long, ByVal toolButton As TBarCtlsLibUCtl.IToolBarButton, ByVal commandOrigin As TBarCtlsLibUCtl.CommandOriginConstants, forwardMessage As Boolean)
  ExecuteCommand commandID
  forwardMessage = False
End Sub

Private Sub tbMath_ButtonGetInfoTipText(ByVal toolButton As TBarCtlsLibUCtl.IToolBarButton, ByVal maxInfoTipLength As Long, infoTipText As String, abortToolTip As Boolean)
  Dim doc As IDocument

  If toolButton.ButtonType <> btyPlaceholder Then
    Set doc = Me.ActiveForm
    If Not (doc Is Nothing) Then
      infoTipText = doc.GetCommandToolTip(toolButton.ID)
    End If
    If infoTipText = "" Then
      infoTipText = buttonsMath(toolButton.ButtonData).Text
    End If
  End If
End Sub

Private Sub tbMath_ExecuteCommand(ByVal commandID As Long, ByVal toolButton As TBarCtlsLibUCtl.IToolBarButton, ByVal commandOrigin As TBarCtlsLibUCtl.CommandOriginConstants, forwardMessage As Boolean)
  ExecuteCommand commandID
  forwardMessage = False
End Sub

Private Sub tbMenu_ButtonGetDropDownMenu(ByVal toolButton As TBarCtlsLibUCtl.IToolBarButton, hMenu As Long)
  Dim i As Long

  For i = LBound(menuButtons) To UBound(menuButtons)
    If menuButtons(i).ID = toolButton.ID Then
      hMenu = menuButtons(i).hSubMenu
      If menuButtons(i).ID = CMDID_WINDOW Then
        UpdateWindowMenu hMenu
      End If
      Exit For
    End If
  Next i
End Sub

Private Sub tbMenu_ExecuteCommand(ByVal commandID As Long, ByVal toolButton As TBarCtlsLibUCtl.IToolBarButton, ByVal commandOrigin As TBarCtlsLibUCtl.CommandOriginConstants, forwardMessage As Boolean)
  If commandOrigin <> coButton Then
    ExecuteCommand commandID
    forwardMessage = False
  End If
End Sub

Private Sub tbMenu_OpenPopupMenu(ByVal hMenu As Long, ByVal parentMenuItemIndex As Long, ByVal isSystemMenu As Boolean)
  Const MF_BYCOMMAND = &H0
  Const MF_ENABLED = &H0
  Const MF_GRAYED = &H1
  Dim i As Long
  Dim commandID As Long

  For i = 0 To GetMenuItemCount(hMenu) - 1
    commandID = GetMenuItemID(hMenu, i)
    If commandID >= 0 Then
      EnableMenuItem hMenu, commandID, MF_BYCOMMAND Or IIf(tbMenu.CommandEnabled(commandID), MF_ENABLED, MF_GRAYED)
    End If
  Next i
End Sub

Private Sub tbObject_ButtonGetInfoTipText(ByVal toolButton As TBarCtlsLibUCtl.IToolBarButton, ByVal maxInfoTipLength As Long, infoTipText As String, abortToolTip As Boolean)
  Dim doc As IDocument

  If toolButton.ButtonType <> btyPlaceholder Then
    Set doc = Me.ActiveForm
    If Not (doc Is Nothing) Then
      infoTipText = doc.GetCommandToolTip(toolButton.ID)
    End If
    If infoTipText = "" Then
      infoTipText = buttonsObject(toolButton.ButtonData).Text
    End If
  End If
End Sub

Private Sub tbObject_DropDown(ByVal toolButton As TBarCtlsLibUCtl.IToolBarButton, buttonRectangle As TBarCtlsLibUCtl.RECTANGLE, furtherProcessing As TBarCtlsLibUCtl.DropDownReturnValuesConstants)
  Const TPM_LEFTALIGN = &H0&
  Const TPM_TOPALIGN = &H0&
  Const TPM_VERPOSANIMATION = &H1000&
  Const TPM_VERTICAL = &H40&
  Dim ptMenu As POINT
  Dim tpmexParams As TPMPARAMS

  If Not (toolButton Is Nothing) Then
    If toolButton.ID = CMDID_OBJECT_INSERTIMAGE Then
      ptMenu.x = buttonRectangle.Left
      ptMenu.y = buttonRectangle.Bottom
      ClientToScreen tbObject.hWnd, ptMenu
      With tpmexParams
        .cbSize = LenB(tpmexParams)
        LSet .rcExclude = buttonRectangle
      End With
      TrackPopupMenuEx hMenuInsertImage, TPM_LEFTALIGN Or TPM_TOPALIGN Or TPM_VERTICAL Or TPM_VERPOSANIMATION, ptMenu.x, ptMenu.y, tbObject.hWnd, tpmexParams
    ElseIf toolButton.ID = CMDID_OBJECT_INSERTOBJECT Then
      ptMenu.x = buttonRectangle.Left
      ptMenu.y = buttonRectangle.Bottom
      ClientToScreen tbObject.hWnd, ptMenu
      With tpmexParams
        .cbSize = LenB(tpmexParams)
        LSet .rcExclude = buttonRectangle
      End With
      TrackPopupMenuEx hMenuInsertObject, TPM_LEFTALIGN Or TPM_TOPALIGN Or TPM_VERTICAL Or TPM_VERPOSANIMATION, ptMenu.x, ptMenu.y, tbObject.hWnd, tpmexParams
    End If
  End If
End Sub

Private Sub tbObject_ExecuteCommand(ByVal commandID As Long, ByVal toolButton As TBarCtlsLibUCtl.IToolBarButton, ByVal commandOrigin As TBarCtlsLibUCtl.CommandOriginConstants, forwardMessage As Boolean)
  ExecuteCommand commandID
  forwardMessage = False
End Sub

Private Sub tbTable_ButtonGetInfoTipText(ByVal toolButton As TBarCtlsLibUCtl.IToolBarButton, ByVal maxInfoTipLength As Long, infoTipText As String, abortToolTip As Boolean)
  Dim doc As IDocument

  If toolButton.ButtonType <> btyPlaceholder Then
    Set doc = Me.ActiveForm
    If Not (doc Is Nothing) Then
      infoTipText = doc.GetCommandToolTip(toolButton.ID)
    End If
    If infoTipText = "" Then
      infoTipText = buttonsTable(toolButton.ButtonData).Text
    End If
  End If
End Sub

Private Sub tbTable_DropDown(ByVal toolButton As TBarCtlsLibUCtl.IToolBarButton, buttonRectangle As TBarCtlsLibUCtl.RECTANGLE, furtherProcessing As TBarCtlsLibUCtl.DropDownReturnValuesConstants)
  Dim doc As IDocument
  Dim tablePicker As clsTablePicker

  If Not (toolButton Is Nothing) Then
    Set doc = Me.ActiveForm
    If Not (doc Is Nothing) Then
      If toolButton.ID = CMDID_TABLE_INSERTTABLE Then
        Set tablePicker = New clsTablePicker
        tablePicker.IgnoreFirstLeftMouseButtonUp = True
        tablePicker.hWndParent = Me.hWnd
        tablePicker.hwndOwner = tbTable.hWnd
        MapWindowPoints tbTable.hWnd, 0, VarPtr(buttonRectangle), 2

        If Not tablePicker.Popup(buttonRectangle) Then
          ' cancelled
        Else
          doc.ExecuteCommand CMDID_TABLE_INSERTTABLE, Array(tablePicker.SelectedColumnCount, tablePicker.SelectedRowCount)
        End If

      ElseIf toolButton.ID = CMDID_TABLE_CELLBACKCOLOR Then
        previousColor = doc.SelectionCellBackColor
        cellBackColorPicker.hWndParent = Me.hWnd
        cellBackColorPicker.hwndOwner = tbTable.hWnd
        cellBackColorPicker.SelectedColor = IIf(previousColor = -1, cellBackColorPicker.DefaultColor, previousColor)
        MapWindowPoints tbTable.hWnd, 0, VarPtr(buttonRectangle), 2

        If Not cellBackColorPicker.Popup(buttonRectangle) Then
          ' cancelled
          doc.CellBackColor = previousColor
        Else
          UpdateColorInColorIcon cellBackColorPicker.SelectedColor, hImgLst, 38, 39
          doc.CellBackColor = IIf(cellBackColorPicker.SelectedColor = CLR_DEFAULT, -1, cellBackColorPicker.SelectedColor)
        End If
      End If
    End If
  End If
End Sub

Private Sub tbTable_ExecuteCommand(ByVal commandID As Long, ByVal toolButton As TBarCtlsLibUCtl.IToolBarButton, ByVal commandOrigin As TBarCtlsLibUCtl.CommandOriginConstants, forwardMessage As Boolean)
  ExecuteCommand commandID
  forwardMessage = False
End Sub

Private Sub cellBackColorPicker_HotTrackingColor(ByVal hotColor As Long, ByVal colorIsValid As Boolean)
  Dim doc As IDocument

  Set doc = ActiveForm
  If Not (doc Is Nothing) Then
    If colorIsValid Then
      doc.CellBackColor = IIf(hotColor = CLR_DEFAULT, -1, hotColor)
    Else
      doc.CellBackColor = previousColor
    End If
  End If
End Sub

Private Sub textBackColorPicker_HotTrackingColor(ByVal hotColor As Long, ByVal colorIsValid As Boolean)
  Dim doc As IDocument

  Set doc = ActiveForm
  If Not (doc Is Nothing) Then
    If colorIsValid Then
      doc.TextBackColor = IIf(hotColor = CLR_DEFAULT, -1, hotColor)
    Else
      doc.TextBackColor = previousColor
    End If
  End If
End Sub

Private Sub textForeColorPicker_HotTrackingColor(ByVal hotColor As Long, ByVal colorIsValid As Boolean)
  Dim doc As IDocument

  Set doc = ActiveForm
  If Not (doc Is Nothing) Then
    If colorIsValid Then
      doc.TextForeColor = IIf(hotColor = CLR_DEFAULT, -1, hotColor)
    Else
      doc.TextForeColor = previousColor
    End If
  End If
End Sub


Private Sub CreateMenus()
  Const MFS_CHECKED = &H8
  Const MFS_UNCHECKED = &H0
  Const MFT_SEPARATOR = &H800
  Const MFT_STRING = &H0
  Const MIIM_ID = &H2
  Const MIIM_STATE = &H1
  Const MIIM_SUBMENU = &H4
  Const MIIM_TYPE = &H10
  Dim hMenu As Long
  Dim miiData As MENUITEMINFO
  Dim strMenuItem As String

  hMenu = CreatePopupMenu
  menuButtons(0).hSubMenu = hMenu
  If hMenu Then
    menusToDestroy.Add hMenu
    With miiData
      .cbSize = LenB(miiData)
      .fMask = MIIM_TYPE Or MIIM_ID Or MIIM_SUBMENU
      .fType = MFT_STRING

      .wID = CMDID_FILE_NEW
      strMenuItem = "&New" & vbTab & "Ctrl+N"
      .dwTypeData = StrPtr(strMenuItem)
      .cch = Len(strMenuItem)
      InsertMenuItem hMenu, 1, 1, miiData
      .wID = CMDID_FILE_OPEN
      strMenuItem = "&Open..." & vbTab & "Ctrl+O"
      .dwTypeData = StrPtr(strMenuItem)
      .cch = Len(strMenuItem)
      InsertMenuItem hMenu, 2, 1, miiData
      .wID = CMDID_FILE_SAVE
      strMenuItem = "&Save" & vbTab & "Ctrl+S"
      .dwTypeData = StrPtr(strMenuItem)
      .cch = Len(strMenuItem)
      InsertMenuItem hMenu, 3, 1, miiData
      .wID = CMDID_FILE_PRINT
      strMenuItem = "&Print"
      .dwTypeData = StrPtr(strMenuItem)
      .cch = Len(strMenuItem)
      InsertMenuItem hMenu, 4, 1, miiData
      hMenuRecentFiles = CreatePopupMenu
      .wID = CMDID_FILE_RECENT
      .hSubMenu = hMenuRecentFiles
      strMenuItem = "&Recent Files"
      .dwTypeData = StrPtr(strMenuItem)
      .cch = Len(strMenuItem)
      InsertMenuItem hMenu, 5, 1, miiData
      .wID = CMDID_FILE_EXIT
      .hSubMenu = 0
      strMenuItem = "&Exit" & vbTab & "Alt+F4"
      .dwTypeData = StrPtr(strMenuItem)
      .cch = Len(strMenuItem)
      InsertMenuItem hMenu, 6, 1, miiData

      .fType = MFT_SEPARATOR
      InsertMenuItem hMenu, 5, 1, miiData
      InsertMenuItem hMenu, 4, 1, miiData
      InsertMenuItem hMenu, 3, 1, miiData

      Call UpdateRecentFilesMenu
    End With
  End If

  hMenu = CreatePopupMenu
  menuButtons(1).hSubMenu = hMenu
  If hMenu Then
    menusToDestroy.Add hMenu
    With miiData
      .cbSize = LenB(miiData)
      .fMask = MIIM_TYPE Or MIIM_ID Or MIIM_SUBMENU
      .fType = MFT_STRING

      .wID = CMDID_EDIT_UNDO
      strMenuItem = "&Undo" & vbTab & "Ctrl+Z"
      .dwTypeData = StrPtr(strMenuItem)
      .cch = Len(strMenuItem)
      InsertMenuItem hMenu, 1, 1, miiData
      .wID = CMDID_EDIT_REDO
      strMenuItem = "&Redo" & vbTab & "Ctrl+Y"
      .dwTypeData = StrPtr(strMenuItem)
      .cch = Len(strMenuItem)
      InsertMenuItem hMenu, 2, 1, miiData
      .wID = CMDID_EDIT_CUT
      strMenuItem = "Cu&t" & vbTab & "Ctrl+X"
      .dwTypeData = StrPtr(strMenuItem)
      .cch = Len(strMenuItem)
      InsertMenuItem hMenu, 3, 1, miiData
      .wID = CMDID_EDIT_COPY
      strMenuItem = "&Copy" & vbTab & "Ctrl+C"
      .dwTypeData = StrPtr(strMenuItem)
      .cch = Len(strMenuItem)
      InsertMenuItem hMenu, 4, 1, miiData
      .wID = CMDID_EDIT_PASTE
      strMenuItem = "&Paste" & vbTab & "Ctrl+V"
      .dwTypeData = StrPtr(strMenuItem)
      .cch = Len(strMenuItem)
      InsertMenuItem hMenu, 5, 1, miiData

      .fType = MFT_SEPARATOR
      InsertMenuItem hMenu, 2, 1, miiData
    End With
  End If

  hMenu = CreatePopupMenu
  menuButtons(2).hSubMenu = hMenu
  If hMenu Then
    menusToDestroy.Add hMenu
    With miiData
      .cbSize = LenB(miiData)
      .fMask = MIIM_TYPE Or MIIM_ID Or MIIM_SUBMENU
      .fType = MFT_STRING

      .wID = CMDID_WINDOW_CASCADE
      strMenuItem = "&Cascade"
      .dwTypeData = StrPtr(strMenuItem)
      .cch = Len(strMenuItem)
      InsertMenuItem hMenu, 1, 1, miiData
      .wID = CMDID_WINDOW_TILE
      strMenuItem = "&Tile"
      .dwTypeData = StrPtr(strMenuItem)
      .cch = Len(strMenuItem)
      InsertMenuItem hMenu, 2, 1, miiData
      .wID = CMDID_WINDOW_ARRANGEICONS
      strMenuItem = "Arrange &Icons"
      .dwTypeData = StrPtr(strMenuItem)
      .cch = Len(strMenuItem)
      InsertMenuItem hMenu, 3, 1, miiData
    End With
  End If

  hMenu = CreatePopupMenu
  menuButtons(3).hSubMenu = hMenu
  If hMenu Then
    menusToDestroy.Add hMenu
    With miiData
      .cbSize = LenB(miiData)
      .fMask = MIIM_TYPE Or MIIM_ID Or MIIM_SUBMENU
      .fType = MFT_STRING

      .wID = CMDID_HELP_ABOUT
      strMenuItem = "&About"
      .dwTypeData = StrPtr(strMenuItem)
      .cch = Len(strMenuItem)
      InsertMenuItem hMenu, 1, 1, miiData
    End With
  End If

  hMenuInsertImage = CreatePopupMenu
  If hMenuInsertImage Then
    menusToDestroy.Add hMenuInsertImage
    With miiData
      .cbSize = LenB(miiData)
      .fMask = MIIM_TYPE Or MIIM_ID Or MIIM_SUBMENU
      .fType = MFT_STRING

      .wID = CMDID_OBJECT_SETBKIMAGE
      strMenuItem = "&Choose Background Image"
      .dwTypeData = StrPtr(strMenuItem)
      .cch = Len(strMenuItem)
      InsertMenuItem hMenuInsertImage, 1, 1, miiData
    End With
  End If

  hMenuInsertObject = CreatePopupMenu
  If hMenuInsertObject Then
    menusToDestroy.Add hMenuInsertObject
    With miiData
      .cbSize = LenB(miiData)
      .fMask = MIIM_TYPE Or MIIM_ID Or MIIM_SUBMENU
      .fType = MFT_STRING

      .wID = CMDID_OBJECT_INSERTOBJECTFROMFILE
      strMenuItem = "&Insert From File"
      .dwTypeData = StrPtr(strMenuItem)
      .cch = Len(strMenuItem)
      InsertMenuItem hMenuInsertObject, 1, 1, miiData
    End With
  End If
End Sub

Private Sub UpdateRecentFilesMenu(Optional ByVal mostRecentFile As String = "")
  Const MF_BYPOSITION = &H400
  Const MFS_DISABLED = &H3
  Const MFT_STRING = &H0
  Const MIIM_ID = &H2
  Const MIIM_STATE = &H1
  Const MIIM_TYPE = &H10
  Dim hMenu As Long
  Dim miiData As MENUITEMINFO
  Dim strMenuItem As String
  Dim i As Long
  Dim j As Long

  If mostRecentFile <> "" Then
    If StrComp(recentFiles(LBound(recentFiles)), mostRecentFile, vbTextCompare) <> 0 Then
      ' remove mostRecentFile from the list
      For i = LBound(recentFiles) To UBound(recentFiles)
        If StrComp(recentFiles(i), mostRecentFile, vbTextCompare) = 0 Then
          For j = i + 1 To UBound(recentFiles)
            recentFiles(j - 1) = recentFiles(j)
          Next j
        End If
      Next i
      ' now move all entries down by one to create an empty slot at position 1
      For i = UBound(recentFiles) To LBound(recentFiles) + 1 Step -1
        recentFiles(i) = recentFiles(i - 1)
      Next i
      recentFiles(LBound(recentFiles)) = mostRecentFile
    End If
  End If

  ' remove files that do not exist
  For i = LBound(recentFiles) To UBound(recentFiles)
    If recentFiles(i) <> "" Then
      If PathFileExists(StrPtr(recentFiles(i))) = 0 Then
        For j = i + 1 To UBound(recentFiles)
          recentFiles(j - 1) = recentFiles(j)
        Next j
        ' clear the last entry
        recentFiles(UBound(recentFiles)) = ""
      End If
    End If
  Next i

  If hMenuRecentFiles Then
    ' clear the menu
    While GetMenuItemCount(hMenuRecentFiles) > 0
      DeleteMenu hMenuRecentFiles, 0, MF_BYPOSITION
    Wend
    With miiData
      .cbSize = LenB(miiData)
      .fMask = MIIM_TYPE Or MIIM_ID
      .fType = MFT_STRING

      .wID = CMDID_FILE_RECENT_FIRST - 1
      For i = LBound(recentFiles) To UBound(recentFiles)
        If recentFiles(i) = "" Then
          ' we've reached the end
          Exit For
        End If
        strMenuItem = recentFiles(i)
        Call PathStripPath(StrPtr(strMenuItem))
        .dwTypeData = StrPtr(strMenuItem)
        .cch = lstrlen(StrPtr(strMenuItem))
        .wID = .wID + 1
        InsertMenuItem hMenuRecentFiles, i - LBound(recentFiles) + 1, MF_BYPOSITION, miiData
      Next i

      If GetMenuItemCount(hMenuRecentFiles) = 0 Then
        .fMask = .fMask Or MIIM_STATE
        .fState = MFS_DISABLED
        .wID = CMDID_FILE_RECENT_FIRST
        strMenuItem = "(Empty)"
        .dwTypeData = StrPtr(strMenuItem)
        .cch = Len(strMenuItem)
        InsertMenuItem hMenuRecentFiles, 1, MF_BYPOSITION, miiData
        tbMenu.CommandEnabled(CMDID_FILE_RECENT_FIRST) = False
      Else
        tbMenu.CommandEnabled(CMDID_FILE_RECENT_FIRST) = True
      End If
    End With
  End If
End Sub

Private Sub ExecuteCommand(ByVal commandID As Long)
  Dim doc As IDocument
  Dim frm As Form
  Dim commonDlg As clsCommonDialog
  Dim newFileName As String

  Select Case commandID
    Case CMDID_FILE_EXIT
      Unload Me
      
    Case CMDID_FILE_NEW
      Set frm = New frmDocument
      Set frm.mdiFrame = Me
      frm.Show
      frm.WindowState = vbMaximized
      Set frm = Nothing
    Case CMDID_FILE_OPEN
      Set commonDlg = New clsCommonDialog
      With commonDlg
        Call .AddFilter("Rich Text Documents", "*.rtf")
        Call .AddFilter("Text Files", "*.txt")
        Call .AddFilter("All Files", "*.*")
        .ActiveFilter = 0
        .DefaultExtension = "*.rtf"
        .hWndParent = Me.hWnd
        .OpenFlags = OFN_ENABLESIZING Or OFN_FILEMUSTEXIST Or OFN_LONGNAMES Or OFN_PATHMUSTEXIST
        .Caption = "Open Document"
        If Not .ShowOpen(newFileName) Then
          newFileName = ""
        End If
      End With
      If newFileName <> "" Then
        Set frm = New frmDocument
        Set frm.mdiFrame = Me
        Set doc = frm
        If doc.OpenFile(newFileName) Then
          frm.Show
          frm.WindowState = vbMaximized
          Call UpdateRecentFilesMenu(newFileName)
        Else
          Unload frm
          MsgBox "Error while loading file.", vbCritical
        End If
        Set frm = Nothing
      End If
    Case CMDID_FILE_SAVE
      Set doc = Me.ActiveForm
      If Not (doc Is Nothing) Then
        If doc.FileName = "" Then
          Set commonDlg = New clsCommonDialog
          With commonDlg
            Call .AddFilter("Rich Text Documents", "*.rtf")
            Call .AddFilter("Text Files", "*.txt")
            Call .AddFilter("All Files", "*.*")
            .ActiveFilter = 0
            .DefaultExtension = "*.rtf"
            .DefaultFile = Me.ActiveForm.Caption & ".rtf"
            .hWndParent = Me.hWnd
            .OpenFlags = OFN_ENABLESIZING Or OFN_LONGNAMES Or OFN_PATHMUSTEXIST Or OFN_OVERWRITEPROMPT
            .Caption = "Save Document"
            If Not .ShowSave(newFileName) Then
              newFileName = ""
            End If
          End With
        Else
          newFileName = doc.FileName
        End If
        If newFileName <> "" Then
          If doc.SaveFile(newFileName) Then
            Call UpdateRecentFilesMenu(newFileName)
          Else
            MsgBox "Error while saving file.", vbCritical
          End If
        End If
      End If
      
    Case CMDID_WINDOW_CASCADE
      Me.Arrange vbCascade
    Case CMDID_WINDOW_TILE
      Me.Arrange vbTileHorizontal
    Case CMDID_WINDOW_ARRANGEICONS
      Me.Arrange vbArrangeIcons
      
    Case Else
      If commandID >= CMDID_WINDOW_LISTSTART And commandID <= CMDID_WINDOW_LISTSTART + openChildren.Count Then
        Set frm = openChildren.Item(commandID - CMDID_WINDOW_LISTSTART)
        If Not (frm Is Nothing) Then
          frm.ZOrder 0
          If frm.WindowState = vbMinimized Then
            frm.WindowState = vbNormal
          End If
        End If
      ElseIf commandID >= CMDID_FILE_RECENT_FIRST And commandID <= CMDID_FILE_RECENT_FIRST + (UBound(recentFiles) - LBound(recentFiles) + 1) Then
        Set frm = New frmDocument
        Set frm.mdiFrame = Me
        Set doc = frm
        If doc.OpenFile(recentFiles(commandID - CMDID_FILE_RECENT_FIRST)) Then
          frm.Show
          frm.WindowState = vbMaximized
          Call UpdateRecentFilesMenu(recentFiles(commandID - CMDID_FILE_RECENT_FIRST))
        Else
          Unload frm
          MsgBox "Error while loading file.", vbCritical
        End If
        Set frm = Nothing
      Else
        Set doc = Me.ActiveForm
        If Not (doc Is Nothing) Then
          doc.ExecuteCommand commandID
        End If
      End If
  End Select
  'Me.ActiveForm.SetFocus
End Sub

Private Sub InsertBands()
  Dim leftBorder As Long
  Dim rightBorder As Long

  If themeableOS Then
    ' for Windows Vista
    SetWindowTheme ReBarU.hWnd, StrPtr("explorer"), 0
    'SetWindowTheme ReBarU.hWnd, StrPtr("Communications"), 0
  End If

  With ReBarU.Bands
    With .Add(, tbMenu.hWnd, , SizingGripVisibilityConstants.sgvAlways, KeepInFirstRow:=True, IdealWidth:=tbMenu.IdealWidth, MinimumHeight:=22, MaximumHeight:=22)
      .UseChevron = True
      BANDID_MENUBAR = .ID
    End With
    With .Add(, tbCommon.hWnd, True, SizingGripVisibilityConstants.sgvAlways, IdealWidth:=tbCommon.IdealWidth, MinimumHeight:=22, MaximumHeight:=22)
      .UseChevron = True
      BANDID_TOOLBAR_COMMON = .ID
    End With
    With .Add(, tbFont.hWnd, , SizingGripVisibilityConstants.sgvAlways, IdealWidth:=tbFont.IdealWidth, MinimumHeight:=22, MaximumHeight:=22)
      .UseChevron = True
      BANDID_TOOLBAR_FONT = .ID
      .Maximize
    End With
    With .Add(, tbFormat.hWnd, True, SizingGripVisibilityConstants.sgvAlways, IdealWidth:=tbFormat.IdealWidth, MinimumHeight:=22, MaximumHeight:=22)
      .UseChevron = True
      BANDID_TOOLBAR_FORMAT = .ID
    End With
    With .Add(, tbObject.hWnd, , SizingGripVisibilityConstants.sgvAlways, IdealWidth:=tbObject.IdealWidth, MinimumHeight:=22, MaximumHeight:=22)
      .UseChevron = True
      BANDID_TOOLBAR_OBJECT = .ID
    End With
    With .Add(, tbTable.hWnd, , SizingGripVisibilityConstants.sgvAlways, IdealWidth:=tbTable.IdealWidth, MinimumHeight:=22, MaximumHeight:=22)
      .UseChevron = True
      BANDID_TOOLBAR_TABLE = .ID
    End With
    With .Add(, tbMath.hWnd, , SizingGripVisibilityConstants.sgvAlways, IdealWidth:=tbMath.IdealWidth, MinimumHeight:=22, MaximumHeight:=22)
      .UseChevron = True
      BANDID_TOOLBAR_MATH = .ID
      .Maximize
    End With

    With .Item(BANDID_TOOLBAR_COMMON, bitID)
      .GetBorderSizes leftBorder, 0, rightBorder, 0
      .CurrentWidth = tbCommon.IdealWidth + leftBorder + rightBorder
    End With
    With .Item(BANDID_TOOLBAR_FONT, bitID)
      .GetBorderSizes leftBorder, 0, rightBorder, 0
      .CurrentWidth = tbFont.IdealWidth + leftBorder + rightBorder
    End With
    With .Item(BANDID_TOOLBAR_FORMAT, bitID)
      .GetBorderSizes leftBorder, 0, rightBorder, 0
      .CurrentWidth = tbFormat.IdealWidth + leftBorder + rightBorder
    End With
    With .Item(BANDID_TOOLBAR_OBJECT, bitID)
      .GetBorderSizes leftBorder, 0, rightBorder, 0
      .CurrentWidth = tbObject.IdealWidth + leftBorder + rightBorder
    End With
    With .Item(BANDID_TOOLBAR_TABLE, bitID)
      .GetBorderSizes leftBorder, 0, rightBorder, 0
      .CurrentWidth = tbTable.IdealWidth + leftBorder + rightBorder
    End With
  End With
  Set ReBarU.MDIFrameMenuBand = ReBarU.Bands(BANDID_MENUBAR, bitID)
  picContainer.Visible = False
End Sub

Private Sub InsertButtons()
  Dim i As Long
  Dim isPartOfGroup As Boolean

  If themeableOS Then
    ' for Windows Vista
    SetWindowTheme tbMenu.hWnd, StrPtr("explorer"), 0
    SetWindowTheme tbCommon.hWnd, StrPtr("explorer"), 0
    SetWindowTheme tbFormat.hWnd, StrPtr("explorer"), 0
    SetWindowTheme tbFont.hWnd, StrPtr("explorer"), 0
    SetWindowTheme tbTable.hWnd, StrPtr("explorer"), 0
    SetWindowTheme tbMath.hWnd, StrPtr("explorer"), 0
    SetWindowTheme tbObject.hWnd, StrPtr("explorer"), 0
  End If

  For i = LBound(menuButtons) To UBound(menuButtons)
    menuButtons(i).ID = Choose(i + 1, CMDID_FILE, CMDID_EDIT, CMDID_WINDOW, CMDID_HELP)
    menuButtons(i).Text = Choose(i + 1, "&File", "&Edit", "&Window", "&?")
  Next i

  With tbMenu
    With .Buttons
      .RemoveAll
      For i = LBound(menuButtons) To UBound(menuButtons)
        .Add menuButtons(i).ID, Text:=menuButtons(i).Text, DropDownStyle:=DropDownStyleConstants.ddstNormal, Visible:=(menuButtons(i).ID <> CMDID_WINDOW)
      Next i
    End With
    .RegisterHotkey mkCtrl, vbKeyN, CMDID_FILE_NEW
    .RegisterHotkey mkCtrl, vbKeyO, CMDID_FILE_OPEN
    .RegisterHotkey mkCtrl, vbKeyS, CMDID_FILE_SAVE
    .RegisterHotkey mkAlt, vbKeyF4, CMDID_FILE_EXIT
    .RegisterHotkey mkCtrl, vbKeyX, CMDID_EDIT_CUT
    .RegisterHotkey mkCtrl, vbKeyC, CMDID_EDIT_COPY
    .RegisterHotkey mkCtrl, vbKeyV, CMDID_EDIT_PASTE
    .RegisterHotkey mkCtrl, vbKeyA, CMDID_EDIT_SELECTALL
    .RegisterHotkey mkCtrl, vbKeyF, CMDID_EDIT_FIND
    .RegisterHotkey mkCtrl, vbKeyZ, CMDID_EDIT_UNDO
    .RegisterHotkey mkCtrl, vbKeyY, CMDID_EDIT_REDO
    .RegisterHotkey 0, vbKeyTab, CMDID_FORMAT_INCREASEINDENT
    .RegisterHotkey mkShift, vbKeyTab, CMDID_FORMAT_DECREASEINDENT
    .RegisterHotkey 0, vbKeyF1, CMDID_HELP_ABOUT
  End With

  'tbCommon.LoadDefaultImages siltSmallStandardBitmaps
  tbCommon.hImageList(ilNormalButtons) = hImgLst
  buttonsCommon(0).ID = CMDID_FILE_NEW
  buttonsCommon(0).Text = "&New"
  buttonsCommon(0).IconIndex = 0 'siiStandardFileNew
  buttonsCommon(1).ID = CMDID_FILE_OPEN
  buttonsCommon(1).Text = "&Open..."
  buttonsCommon(1).IconIndex = 1 'siiStandardFileOpen
  buttonsCommon(2).ID = CMDID_FILE_SAVE
  buttonsCommon(2).Text = "&Save"
  buttonsCommon(2).IconIndex = 2 'siiStandardFileSave
  buttonsCommon(3).ID = CMDID_FILE_PRINT
  buttonsCommon(3).Text = "&Print"
  buttonsCommon(3).IconIndex = 3 'siiStandardPrint
  buttonsCommon(4).ID = CMDID_EDIT_CUT
  buttonsCommon(4).Text = "Cu&t"
  buttonsCommon(4).IconIndex = 4 'siiStandardCut
  buttonsCommon(5).ID = CMDID_EDIT_COPY
  buttonsCommon(5).Text = "&Copy"
  buttonsCommon(5).IconIndex = 5 'siiStandardCopy
  buttonsCommon(6).ID = CMDID_EDIT_PASTE
  buttonsCommon(6).Text = "&Paste"
  buttonsCommon(6).IconIndex = 6 'siiStandardPaste
  buttonsCommon(7).ID = CMDID_EDIT_FIND
  buttonsCommon(7).Text = "&Find"
  buttonsCommon(7).IconIndex = 7 'siiStandardFind
  buttonsCommon(8).ID = CMDID_EDIT_UNDO
  buttonsCommon(8).Text = "Undo"
  buttonsCommon(8).IconIndex = 8 'siiStandardUndo
  buttonsCommon(9).ID = CMDID_EDIT_REDO
  buttonsCommon(9).Text = "Redo"
  buttonsCommon(9).IconIndex = 9 'siiStandardRedo
  buttonsCommon(10).ID = CMDID_EDIT_SPELLCHECKING
  buttonsCommon(10).Text = "Check Spelling"
  buttonsCommon(10).IconIndex = 10

  With tbCommon.Buttons
    .RemoveAll
    For i = LBound(buttonsCommon) To UBound(buttonsCommon)
      If buttonsCommon(i).ID = CMDID_EDIT_SPELLCHECKING Then
        .Add buttonsCommon(i).ID, IconIndex:=buttonsCommon(i).IconIndex, ButtonType:=btyCheckButton, ButtonData:=i
      Else
        .Add buttonsCommon(i).ID, IconIndex:=buttonsCommon(i).IconIndex, ButtonData:=i
      End If
    Next i
    ' Now insert a separator between Print and Cut command. NOTE: It's better to disable auto-sizing for separators
    .Add 0, .Item(CMDID_EDIT_CUT, btitID).Index, ButtonType:=btySeparator, AutoSize:=False
    .Item(CMDID_FILE_OPEN, btitID).DropDownStyle = ddstNormal
    ' Insert a separator before the Undo command.
    .Add 0, .Item(CMDID_EDIT_UNDO, btitID).Index, ButtonType:=btySeparator, AutoSize:=False
  End With

  tbFormat.hImageList(ilNormalButtons) = hImgLst
  buttonsFormat(0).ID = CMDID_FORMAT_ALIGNLEFT
  buttonsFormat(0).Text = "Align &Left"
  buttonsFormat(0).IconIndex = 11
  buttonsFormat(1).ID = CMDID_FORMAT_ALIGNCENTER
  buttonsFormat(1).Text = "C&enter"
  buttonsFormat(1).IconIndex = 12
  buttonsFormat(2).ID = CMDID_FORMAT_ALIGNRIGHT
  buttonsFormat(2).Text = "Align &Right"
  buttonsFormat(2).IconIndex = 13
  buttonsFormat(3).ID = CMDID_FORMAT_ALIGNJUSTIFY
  buttonsFormat(3).Text = "Align &Justified"
  buttonsFormat(3).IconIndex = 14
  buttonsFormat(4).ID = CMDID_FORMAT_BULLETLIST
  buttonsFormat(4).Text = "Li&st"
  buttonsFormat(4).IconIndex = 15
  buttonsFormat(5).ID = CMDID_FORMAT_NUMBEREDLIST
  buttonsFormat(5).Text = "&Numbered List"
  buttonsFormat(5).IconIndex = 16
  buttonsFormat(6).ID = CMDID_FORMAT_CONVERTTOHYPERLINK
  buttonsFormat(6).Text = "Convert to H&yperlink"
  buttonsFormat(6).IconIndex = 17

  With tbFormat.Buttons
    .RemoveAll
    For i = LBound(buttonsFormat) To UBound(buttonsFormat)
      If buttonsFormat(i).ID >= CMDID_FORMAT_BULLETLIST And buttonsFormat(i).ID <= CMDID_FORMAT_NUMBEREDLIST Then
        isPartOfGroup = False
      ElseIf buttonsFormat(i).ID >= CMDID_FORMAT_ALIGNLEFT And buttonsFormat(i).ID <= CMDID_FORMAT_ALIGNJUSTIFY Then
        isPartOfGroup = True
      Else
        isPartOfGroup = False
      End If
      If buttonsFormat(i).ID = CMDID_FORMAT_CONVERTTOHYPERLINK Then
        .Add buttonsFormat(i).ID, IconIndex:=buttonsFormat(i).IconIndex, ButtonData:=i
      Else
        .Add buttonsFormat(i).ID, ButtonType:=btyCheckButton, PartOfGroup:=isPartOfGroup, IconIndex:=buttonsFormat(i).IconIndex, ButtonData:=i
      End If
    Next i
    ' Now insert a separator between alignment and list options. NOTE: It's better to disable auto-sizing for separators
    .Add 0, .Item(CMDID_FORMAT_BULLETLIST, btitID).Index, ButtonType:=btySeparator, AutoSize:=False
    .Add 0, .Item(CMDID_FORMAT_CONVERTTOHYPERLINK, btitID).Index, ButtonType:=btySeparator, AutoSize:=False
    .Item(CMDID_FORMAT_ALIGNLEFT, btitID).SelectionState = ssChecked
  End With

  tbFont.hImageList(ilNormalButtons) = hImgLst
  buttonsFont(0).ID = CMDID_FONT_BOLD
  buttonsFont(0).Text = "&Bold"
  buttonsFont(0).IconIndex = 18
  buttonsFont(1).ID = CMDID_FONT_ITALIC
  buttonsFont(1).Text = "&Italic"
  buttonsFont(1).IconIndex = 19
  buttonsFont(2).ID = CMDID_FONT_UNDERLINE
  buttonsFont(2).Text = "&Underline"
  buttonsFont(2).IconIndex = 20
  buttonsFont(3).ID = CMDID_FONT_STRIKETHROUGH
  buttonsFont(3).Text = "Strike &Through"
  buttonsFont(3).IconIndex = 21
  buttonsFont(4).ID = CMDID_FONT_SUBSCRIPT
  buttonsFont(4).Text = "&Subscript"
  buttonsFont(4).IconIndex = 22
  buttonsFont(5).ID = CMDID_FONT_SUPERSCRIPT
  buttonsFont(5).Text = "Su&perscript"
  buttonsFont(5).IconIndex = 23
  buttonsFont(6).ID = CMDID_FONT_TEXTBACKCOLOR
  buttonsFont(6).Text = "Text B&ackground Color"
  buttonsFont(6).IconIndex = 26
  buttonsFont(7).ID = CMDID_FONT_TEXTFORECOLOR
  buttonsFont(7).Text = "Te&xt Color"
  buttonsFont(7).IconIndex = 27

  With tbFont.Buttons
    .RemoveAll
    For i = LBound(buttonsFont) To UBound(buttonsFont)
      isPartOfGroup = False
      If buttonsFont(i).ID >= CMDID_FONT_TEXTBACKCOLOR And buttonsFont(i).ID <= CMDID_FONT_TEXTFORECOLOR Then
        .Add buttonsFont(i).ID, IconIndex:=buttonsFont(i).IconIndex, ButtonData:=i
      Else
        .Add buttonsFont(i).ID, ButtonType:=btyCheckButton, IconIndex:=buttonsFont(i).IconIndex, ButtonData:=i
      End If
    Next i
    .Item(CMDID_FONT_TEXTBACKCOLOR, btitID).DropDownStyle = ddstNormal
    .Item(CMDID_FONT_TEXTFORECOLOR, btitID).DropDownStyle = ddstNormal
    ' insert a placeholder for the font combo box
    With .Add(CMDID_FONT_FONTNAME, 0, ButtonType:=btyPlaceholder, AutoSize:=False, Width:=cmbFontName.Width + 4)
      .SetContainedWindow cmbFontName.hWnd, , valTop
    End With
    ' insert a placeholder for the font size combo box
    With .Add(CMDID_FONT_FONTSIZE, 1, ButtonType:=btyPlaceholder, AutoSize:=False, Width:=cmbFontSize.Width + 4)
      .SetContainedWindow cmbFontSize.hWnd, , valTop
    End With
    ' Now insert a separator before the Bold command. NOTE: It's better to disable auto-sizing for separators
    .Add 0, .Item(CMDID_FONT_BOLD, btitID).Index, ButtonType:=btySeparator, AutoSize:=False
    ' Insert a separator before the Text Background Color command.
    .Add 0, .Item(CMDID_FONT_TEXTBACKCOLOR, btitID).Index, ButtonType:=btySeparator, AutoSize:=False
  End With

  tbObject.hImageList(ilNormalButtons) = hImgLst
  buttonsObject(0).ID = CMDID_OBJECT_INSERTIMAGE
  buttonsObject(0).Text = "Insert Image"
  buttonsObject(0).IconIndex = 49
  buttonsObject(1).ID = CMDID_OBJECT_INSERTOBJECT
  buttonsObject(1).Text = "Insert OLE Object"
  buttonsObject(1).IconIndex = 50

  With tbObject.Buttons
    .RemoveAll
    For i = LBound(buttonsObject) To UBound(buttonsObject)
      .Add buttonsObject(i).ID, IconIndex:=buttonsObject(i).IconIndex, ButtonData:=i
    Next i
    .Item(CMDID_OBJECT_INSERTIMAGE, btitID).DropDownStyle = ddstNormal
    .Item(CMDID_OBJECT_INSERTOBJECT, btitID).DropDownStyle = ddstNormal
  End With

  tbTable.hImageList(ilNormalButtons) = hImgLst
  buttonsTable(0).ID = CMDID_TABLE_INSERTTABLE
  buttonsTable(0).Text = "Insert Table"
  buttonsTable(0).IconIndex = 28
  buttonsTable(1).ID = CMDID_TABLE_DELETETABLE
  buttonsTable(1).Text = "Delete Table"
  buttonsTable(1).IconIndex = 29
  buttonsTable(2).ID = CMDID_TABLE_INSERTROW
  buttonsTable(2).Text = "Insert Row"
  buttonsTable(2).IconIndex = 30
  buttonsTable(3).ID = CMDID_TABLE_DELETEROW
  buttonsTable(3).Text = "Delete Row"
  buttonsTable(3).IconIndex = 31
  buttonsTable(4).ID = CMDID_TABLE_INSERTCOLUMN
  buttonsTable(4).Text = "Insert Column"
  buttonsTable(4).IconIndex = 32
  buttonsTable(5).ID = CMDID_TABLE_DELETECOLUMN
  buttonsTable(5).Text = "Delete Column"
  buttonsTable(5).IconIndex = 33
  buttonsTable(6).ID = CMDID_TABLE_MERGECELLS
  buttonsTable(6).Text = "Merge Cells"
  buttonsTable(6).IconIndex = 34
  buttonsTable(7).ID = CMDID_TABLE_ALIGNTOP
  buttonsTable(7).Text = "Align Top"
  buttonsTable(7).IconIndex = 35
  buttonsTable(8).ID = CMDID_TABLE_ALIGNCENTER
  buttonsTable(8).Text = "Align Center"
  buttonsTable(8).IconIndex = 36
  buttonsTable(9).ID = CMDID_TABLE_ALIGNBOTTOM
  buttonsTable(9).Text = "Align Bottom"
  buttonsTable(9).IconIndex = 37
  buttonsTable(10).ID = CMDID_TABLE_CELLBACKCOLOR
  buttonsTable(10).Text = "Cell Background Color"
  buttonsTable(10).IconIndex = 39

  With tbTable.Buttons
    .RemoveAll
    For i = LBound(buttonsTable) To UBound(buttonsTable)
      If buttonsTable(i).ID >= CMDID_TABLE_ALIGNTOP And buttonsTable(i).ID <= CMDID_TABLE_ALIGNBOTTOM Then
        .Add buttonsTable(i).ID, IconIndex:=buttonsTable(i).IconIndex, ButtonType:=btyCheckButton, PartOfGroup:=True, ButtonData:=i
      Else
        .Add buttonsTable(i).ID, IconIndex:=buttonsTable(i).IconIndex, ButtonData:=i
      End If
    Next i
    .Item(CMDID_TABLE_INSERTTABLE, btitID).DropDownStyle = ddstAlwaysWholeButton
    .Item(CMDID_TABLE_CELLBACKCOLOR, btitID).DropDownStyle = ddstNormal
    ' Now insert a separator before the alignment commands. NOTE: It's better to disable auto-sizing for separators
    .Add 0, .Item(CMDID_TABLE_ALIGNTOP, btitID).Index, ButtonType:=btySeparator, AutoSize:=False
    ' Insert a separator before the Cell Background Color command.
    .Add 0, .Item(CMDID_TABLE_CELLBACKCOLOR, btitID).Index, ButtonType:=btySeparator, AutoSize:=False
    IMDIFrame_EnableDisableCommand CMDID_TABLE_DELETETABLE, False
    IMDIFrame_EnableDisableCommand CMDID_TABLE_INSERTROW, False
    IMDIFrame_EnableDisableCommand CMDID_TABLE_DELETEROW, False
    IMDIFrame_EnableDisableCommand CMDID_TABLE_INSERTCOLUMN, False
    IMDIFrame_EnableDisableCommand CMDID_TABLE_DELETECOLUMN, False
    IMDIFrame_EnableDisableCommand CMDID_TABLE_MERGECELLS, False
    IMDIFrame_EnableDisableCommand CMDID_TABLE_ALIGNTOP, False
    IMDIFrame_EnableDisableCommand CMDID_TABLE_ALIGNCENTER, False
    IMDIFrame_EnableDisableCommand CMDID_TABLE_ALIGNBOTTOM, False
  End With

  tbMath.hImageList(ilNormalButtons) = hImgLst
  buttonsMath(0).ID = CMDID_MATH_TOGGLEMATHZONE
  buttonsMath(0).Text = "Display as Math"
  buttonsMath(0).IconIndex = 40
  buttonsMath(1).ID = CMDID_MATH_BUILDUP
  buttonsMath(1).Text = "Professional Form"
  buttonsMath(1).IconIndex = 41
  buttonsMath(2).ID = CMDID_MATH_BUILDDOWN
  buttonsMath(2).Text = "Linearize"
  buttonsMath(2).IconIndex = 42
  buttonsMath(3).ID = CMDID_MATH_INSERTROOTSIGN
  buttonsMath(3).Text = "Insert Root"
  buttonsMath(3).IconIndex = 43
  buttonsMath(4).ID = CMDID_MATH_INSERTSUMSIGN
  buttonsMath(4).Text = "Insert N-ary Sum"
  buttonsMath(4).IconIndex = 44
  buttonsMath(5).ID = CMDID_MATH_INSERTPRODUCTSIGN
  buttonsMath(5).Text = "Insert N-ary Product"
  buttonsMath(5).IconIndex = 45
  buttonsMath(6).ID = CMDID_MATH_INSERTINTSIGN
  buttonsMath(6).Text = "Insert Integral"
  buttonsMath(6).IconIndex = 46
  buttonsMath(7).ID = CMDID_MATH_INSERTLIMES
  buttonsMath(7).Text = "Insert Limes"
  buttonsMath(7).IconIndex = 47
  buttonsMath(8).ID = CMDID_MATH_INSERTMATRIX
  buttonsMath(8).Text = "Insert Matrix"
  buttonsMath(8).IconIndex = 48

  With tbMath.Buttons
    .RemoveAll
    For i = LBound(buttonsMath) To UBound(buttonsMath)
      If buttonsMath(i).ID = CMDID_MATH_TOGGLEMATHZONE Then
        .Add buttonsMath(i).ID, IconIndex:=buttonsMath(i).IconIndex, ButtonType:=btyCheckButton, ButtonData:=i
      Else
        .Add buttonsMath(i).ID, IconIndex:=buttonsMath(i).IconIndex, ButtonData:=i
      End If
    Next i
    ' Now insert a separator between Build Down and Insert Root command. NOTE: It's better to disable auto-sizing for separators
    .Add 0, .Item(CMDID_MATH_INSERTROOTSIGN, btitID).Index, ButtonType:=btySeparator, AutoSize:=False
  End With

  IMDIFrame_EnableDisableCommand CMDID_FILE_SAVE, False
  IMDIFrame_EnableDisableCommand CMDID_EDIT_CUT, False
  IMDIFrame_EnableDisableCommand CMDID_EDIT_COPY, False
  IMDIFrame_EnableDisableCommand CMDID_EDIT_PASTE, False
End Sub

Private Sub InsertFonts()
  Const DEFAULT_CHARSET = 1
  Const GM_ADVANCED = 2
  Dim hDC As Long
  Dim previousMode As Long

  With cmbFontName.ComboItems
    hDC = GetDC(0)
    previousMode = SetGraphicsMode(hDC, GM_ADVANCED)
    EnumFonts hDC, cmbFontName, DEFAULT_CHARSET
    SetGraphicsMode hDC, previousMode
    ReleaseDC 0, hDC
  End With
End Sub

' subclasses this Form
Private Sub Subclass()
  Const NF_REQUERY = 4
  Const WM_NOTIFYFORMAT = &H55

  #If UseSubClassing Then
    If Not SubclassWindow(Me.hWnd, Me, EnumSubclassID.escidFrmMain) Then
      Debug.Print "Subclassing failed!"
    End If
    ' tell the controls to negotiate the correct format with the form
    SendMessageAsLong ReBarU.hWnd, WM_NOTIFYFORMAT, Me.hWnd, NF_REQUERY
    SendMessageAsLong tbMenu.hWnd, WM_NOTIFYFORMAT, Me.hWnd, NF_REQUERY
    SendMessageAsLong tbCommon.hWnd, WM_NOTIFYFORMAT, Me.hWnd, NF_REQUERY
    SendMessageAsLong tbFormat.hWnd, WM_NOTIFYFORMAT, Me.hWnd, NF_REQUERY
    SendMessageAsLong tbFont.hWnd, WM_NOTIFYFORMAT, Me.hWnd, NF_REQUERY
    SendMessageAsLong tbTable.hWnd, WM_NOTIFYFORMAT, Me.hWnd, NF_REQUERY
    SendMessageAsLong tbMath.hWnd, WM_NOTIFYFORMAT, Me.hWnd, NF_REQUERY
    SendMessageAsLong tbObject.hWnd, WM_NOTIFYFORMAT, Me.hWnd, NF_REQUERY
  #End If
End Sub

Private Sub UpdateWindowMenu(ByVal hWindowMenu As Long)
  Const MF_BYPOSITION = &H400
  Const MFS_CHECKED = &H8
  Const MFS_UNCHECKED = &H0
  Const MFT_SEPARATOR = &H800
  Const MFT_STRING = &H0
  Const MIIM_ID = &H2
  Const MIIM_STATE = &H1
  Const MIIM_TYPE = &H10
  Dim endOfMenu As Long
  Dim frm As Form
  Dim i As Long
  Dim miiData As MENUITEMINFO
  Dim strMenuItem As String

  If hWindowMenu Then
    ' remove any menu items with an ID >= CMDID_WINDOW_LISTSTART
    For i = GetMenuItemCount(hWindowMenu) - 1 To 0 Step -1
      If GetMenuItemID(hWindowMenu, i) >= CMDID_WINDOW_LISTSTART Then
        DeleteMenu hWindowMenu, i, MF_BYPOSITION
      End If
    Next i

    If openChildren.Count > 0 Then
      endOfMenu = GetMenuItemCount(hWindowMenu) + 1

      ' insert a separator
      With miiData
        .cbSize = LenB(miiData)
        .fMask = MIIM_TYPE Or MIIM_ID
        .fType = MFT_SEPARATOR

        .wID = CMDID_WINDOW_LISTSTART
        InsertMenuItem hWindowMenu, endOfMenu, 1, miiData

        ' insert an item for each child window
        .fType = MFT_STRING
        For i = 1 To openChildren.Count
          Set frm = openChildren.Item(i)
          If Not (frm Is Nothing) Then
            If frm Is Me.ActiveForm Then
              .fMask = .fMask Or MIIM_STATE
              .fState = MFS_CHECKED
            Else
              .fMask = .fMask And Not MIIM_STATE
            End If

            .wID = CMDID_WINDOW_LISTSTART + i
            strMenuItem = frm.Caption
            .dwTypeData = StrPtr(strMenuItem)
            .cch = Len(strMenuItem)
            InsertMenuItem hWindowMenu, endOfMenu + i, 1, miiData
          End If
        Next i
      End With
    End If
  End If
End Sub

Private Function TranslateColor(ByVal clr As OLE_COLOR, Optional ByVal hPal As Long = 0) As Long
  Const CLR_INVALID = &HFFFF&

  If OleTranslateColor(clr, hPal, TranslateColor) Then TranslateColor = CLR_INVALID
End Function

Private Sub UpdateColorInColorIcon(ByVal SelectedColor As Long, ByVal hImgLst As Long, ByVal originalIconIndex As Long, ByVal coloredIconIndex As Long)
  Const BI_RGB = &H0
  Const DIB_RGB_COLORS = 0
  Const ETO_OPAQUE = &H2
  Const ILD_IMAGE = &H20
  Dim hIcon As Long
  Dim bmpInfo As BITMAPINFO
  Dim hMemDC As Long
  Dim image As ICONINFO
  Dim bits() As Long
  Dim l As Long
  Dim hPrevBmp As Long
  Dim rc As RECTANGLE

  hIcon = ImageList_GetIcon(hImgLst, originalIconIndex, ILD_IMAGE)
  If hIcon Then
    If SelectedColor >= 0 And SelectedColor <= &HFFFFFF Then
      hMemDC = CreateCompatibleDC(0)
      If hMemDC Then
        If GetIconInfo(hIcon, image) Then
          bmpInfo.bmiHeader.biSize = LenB(bmpInfo.bmiHeader)
          bmpInfo.bmiHeader.biWidth = 16
          bmpInfo.bmiHeader.biHeight = -16
          bmpInfo.bmiHeader.biPlanes = 1
          bmpInfo.bmiHeader.biBitCount = 32
          bmpInfo.bmiHeader.biCompression = BI_RGB
          ReDim bits(1 To 16 * 16) As Long
          If GetDIBits(hMemDC, image.hbmColor, 0, 16, VarPtr(bits(1)), bmpInfo, DIB_RGB_COLORS) Then
            For l = 193 To 256
              bits(l) = IIf(bComctl32Version600OrNewer, &HFF000000, &H0) Or SelectedColor
              CopyMemory VarPtr(bits(l)), VarPtr(SelectedColor) + 2, 1
              CopyMemory VarPtr(bits(l)) + 2, VarPtr(SelectedColor), 1
            Next l
            SetDIBits hMemDC, image.hbmColor, 0, 16, VarPtr(bits(1)), bmpInfo, DIB_RGB_COLORS
  
            If Not bComctl32Version600OrNewer Then
              ' I don't know how to use GetDIBits/SetDIBits with monochrome bitmaps, so select
              ' it into a DC instead and paint the color rectangle black
              rc.Top = 12
              rc.Bottom = 16
              rc.Right = 16
              hPrevBmp = SelectObject(hMemDC, image.hbmMask)
              SetBkColor hMemDC, vbBlack
              ExtTextOut hMemDC, 0, 0, ETO_OPAQUE, rc, 0, 0, 0
              SelectObject hMemDC, hPrevBmp
            End If
            DestroyIcon hIcon
            hIcon = CreateIconIndirect(image)
          End If
          If image.hbmColor Then DeleteObject image.hbmColor
          If image.hbmMask Then DeleteObject image.hbmMask
        End If
        DeleteDC hMemDC
      End If
    End If
    ImageList_ReplaceIcon hImgLst, coloredIconIndex, hIcon
    DestroyIcon hIcon
  End If
End Sub
