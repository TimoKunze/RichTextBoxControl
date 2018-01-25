VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCommonDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

  Private Enum CDFields
    DM_ORIENTATION = &H1
    DM_PAPERSIZE = &H2
    DM_PAPERLENGTH = &H4
    DM_PAPERWIDTH = &H8
    DM_SCALE = &H10
    DM_POSITION = &H20
    DM_COPIES = &H100
    DM_DEFAULTSOURCE = &H200
    DM_PRINTQUALITY = &H400
    DM_COLOR = &H800
    DM_DUPLEX = &H1000
    DM_YRESOLUTION = &H2000
    DM_TTOPTION = &H4000
    DM_COLLATE = &H8000
    DM_FORMNAME = &H10000
    DM_LOGPIXELS = &H20000
    DM_BITSPERPEL = &H40000
    DM_PELSWIDTH = &H80000
    DM_PELSHEIGHT = &H100000
    DM_DISPLAYFLAGS = &H200000
    DM_DISPLAYFREQUENCY = &H400000
    DM_ICMMETHOD = &H800000
    DM_ICMINTENT = &H1000000
    DM_MEDIATYPE = &H2000000
    DM_DITHERTYPE = &H4000000
    DM_PANNINGWIDTH = &H8000000
    DM_PANNINGHEIGHT = &H10000000
  End Enum


  Public Enum CDColorFlags
    CC_ANYCOLOR = &H100
    CC_ENABLEHOOK = &H10
    CC_ENABLETEMPLATE = &H20
    CC_ENABLETEMPLATEHANDLE = &H40
    CC_FULLOPEN = &H2
    CC_PREVENTFULLOPEN = &H4
    CC_RGBINIT = &H1
    CC_SHOWHELP = &H8
    CC_SOLIDCOLOR = &H80
  End Enum

  Public Enum CDDuplexConsts
    DMDUP_SIMPLEX = 1    ' no duplex
    DMDUP_VERTICAL = 2   ' tablet
    DMDUP_HORIZONTAL = 3 ' book
  End Enum

  Public Enum CDFontFlags
    CF_SCREENFONTS = &H1
    CF_PRINTERFONTS = &H2
    CF_BOTH = (CF_SCREENFONTS Or CF_PRINTERFONTS)
    CF_SHOWHELP = &H4
    CF_ENABLEHOOK = &H8
    CF_ENABLETEMPLATE = &H10
    CF_ENABLETEMPLATEHANDLE = &H20
    CF_INITTOLOGFONTSTRUCT = &H40
    CF_USESTYLE = &H80
    CF_EFFECTS = &H100
    CF_APPLY = &H200
    CF_ANSIONLY = &H400
    CF_SCRIPTSONLY = CF_ANSIONLY
    CF_NOVECTORFONTS = &H800
    CF_NOOEMFONTS = CF_NOVECTORFONTS
    CF_NOSIMULATIONS = &H1000
    CF_LIMITSIZE = &H2000
    CF_FIXEDPITCHONLY = &H4000
    CF_WYSIWYG = &H8000
    CF_FORCEFONTEXIST = &H10000
    CF_SCALABLEONLY = &H20000
    CF_TTONLY = &H40000
    CF_NOFACESEL = &H80000
    CF_NOSTYLESEL = &H100000
    CF_NOSIZESEL = &H200000
    CF_SELECTSCRIPT = &H400000
    CF_NOSCRIPTSEL = &H800000
    CF_NOVERTFONTS = &H1000000
  End Enum

  Public Enum CDFontTypeConsts
    SIMULATED_FONTTYPE = &H8000
    PRINTER_FONTTYPE = &H4000
    SCREEN_FONTTYPE = &H2000
    BOLD_FONTTYPE = &H100
    ITALIC_FONTTYPE = &H200
    REGULAR_FONTTYPE = &H400
    OPENTYPE_FONTTYPE = &H10000
    TYPE1_FONTTYPE = &H20000
    DSIG_FONTTYPE = &H40000
  End Enum

  Public Enum CDICMConsts
    DMICMMETHOD_NONE = 1   ' no ICM
    DMICMMETHOD_SYSTEM = 2 ' ICM will be done by system
    DMICMMETHOD_DRIVER = 3 ' ICM will be done by driver
    DMICMMETHOD_DEVICE = 4 ' ICM will be done by hardware
  End Enum

  Public Enum CDICMModeConsts
    DMICM_SATURATE = 1         ' maximum saturation
    DMICM_CONTRAST = 2         ' maximum contrast
    DMICM_COLORIMETRIC = 3     ' maximum color precision (correction)
    DMICM_ABS_COLORIMETRIC = 4 ' maximum color precision
  End Enum

  Public Enum CDMediaTypeConsts
    DMMEDIA_STANDARD = 1
    DMMEDIA_TRANSPARENCY = 2
    DMMEDIA_GLOSSY = 3
  End Enum

  Public Enum CDOpenFlagsConsts
    OFN_ALLOWMULTISELECT = &H200
    OFN_CREATEPROMPT = &H2000
    OFN_DONTADDTORECENT = &H2000000
    OFN_ENABLEHOOK = &H20
    OFN_ENABLEINCLUDENOTIFY = &H400000
    OFN_ENABLESIZING = &H800000
    OFN_ENABLETEMPLATE = &H40
    OFN_ENABLETEMPLATEHANDLE = &H80
    OFN_EXPLORER = &H80000
    OFN_EXTENSIONDIFFERENT = &H400
    OFN_FILEMUSTEXIST = &H1000
    OFN_FORCESHOWHIDDEN = &H10000000
    OFN_HIDEREADONLY = &H4
    OFN_LONGNAMES = &H200000
    OFN_NOCHANGEDIR = &H8
    OFN_NODEREFERENCELINKS = &H100000
    OFN_NOLONGNAMES = &H40000
    OFN_NONETWORKBUTTON = &H20000
    OFN_NOREADONLYRETURN = &H8000
    OFN_NOTESTFILECREATE = &H10000
    OFN_NOVALIDATE = &H100
    OFN_OVERWRITEPROMPT = &H2
    OFN_PATHMUSTEXIST = &H800
    OFN_READONLY = &H1
    OFN_SHAREAWARE = &H4000
    OFN_SHAREFALLTHROUGH = 2
    OFN_SHARENOWARN = 1
    OFN_SHAREWARN = 0
    OFN_SHOWHELP = &H10
  End Enum

  Public Enum CDOrientationConsts
    DMORIENT_PORTRAIT = 1
    DMORIENT_LANDSCAPE = 2
  End Enum

  Public Enum CDPageSetupFlags
    PSD_DEFAULTMINMARGINS = &H0
    PSD_DISABLEMARGINS = &H10
    PSD_DISABLEORIENTATION = &H100
    PSD_DISABLEPAGEPAINTING = &H80000
    PSD_DISABLEPAPER = &H200
    PSD_DISABLEPRINTER = &H20
    PSD_ENABLEPAGEPAINTHOOK = &H40000
    PSD_ENABLEPAGESETUPHOOK = &H2000
    PSD_ENABLEPAGESETUPTEMPLATE = &H8000
    PSD_ENABLEPAGESETUPTEMPLATEHANDLE = &H20000
    PSD_INHUNDREDTHSOFMILLIMETERS = &H8
    PSD_INTHOUSANDTHSOFINCHES = &H4
    PSD_INWININIINTLMEASURE = &H0
    PSD_MARGINS = &H2
    PSD_MINMARGINS = &H1
    PSD_NONETWORKBUTTON = &H200000
    PSD_NOWARNING = &H80
    PSD_RETURNDEFAULT = &H400
    PSD_SHOWHELP = &H800
  End Enum

  Public Enum CDPaperSizeConsts
    DMPAPER_LETTER = 1                         ' Letter 8 1/2 x 11 in
    DMPAPER_LETTERSMALL = 2                    ' Letter Small 8 1/2 x 11 in
    DMPAPER_TABLOID = 3                        ' Tabloid 11 x 17 in
    DMPAPER_LEDGER = 4                         ' Ledger 17 x 11 in
    DMPAPER_LEGAL = 5                          ' Legal 8 1/2 x 14 in
    DMPAPER_STATEMENT = 6                      ' Statement 5 1/2 x 8 1/2 in
    DMPAPER_EXECUTIVE = 7                      ' Executive 7 1/4 x 10 1/2 in
    DMPAPER_A3 = 8                             ' A3 297 x 420 mm
    DMPAPER_A4 = 9                             ' A4 210 x 297 mm
    DMPAPER_A4SMALL = 10                       ' A4 Small 210 x 297 mm
    DMPAPER_A5 = 11                            ' A5 148 x 210 mm
    DMPAPER_B4 = 12                            ' B4 (JIS) 250 x 354
    DMPAPER_B5 = 13                            ' B5 (JIS) 182 x 257 mm
    DMPAPER_FOLIO = 14                         ' Folio 8 1/2 x 13 in
    DMPAPER_QUARTO = 15                        ' Quarto 215 x 275 mm
    DMPAPER_10x14 = 16                         ' 10x14 in
    DMPAPER_11x17 = 17                         ' 11x17 in
    DMPAPER_NOTE = 18                          ' Note 8 1/2 x 11 in
    DMPAPER_ENV_9 = 19                         ' Envelope #9 3 7/8 x 8 7/8
    DMPAPER_ENV_10 = 20                        ' Envelope #10 4 1/8 x 9 1/2
    DMPAPER_ENV_11 = 21                        ' Envelope #11 4 1/2 x 10 3/8
    DMPAPER_ENV_12 = 22                        ' Envelope #12 4 \276 x 11
    DMPAPER_ENV_14 = 23                        ' Envelope #14 5 x 11 1/2
    DMPAPER_CSHEET = 24                        ' C size sheet
    DMPAPER_DSHEET = 25                        ' D size sheet
    DMPAPER_ESHEET = 26                        ' E size sheet
    DMPAPER_ENV_DL = 27                        ' Envelope DL 110 x 220mm
    DMPAPER_ENV_C5 = 28                        ' Envelope C5 162 x 229 mm
    DMPAPER_ENV_C3 = 29                        ' Envelope C3  324 x 458 mm
    DMPAPER_ENV_C4 = 30                        ' Envelope C4  229 x 324 mm
    DMPAPER_ENV_C6 = 31                        ' Envelope C6  114 x 162 mm
    DMPAPER_ENV_C65 = 32                       ' Envelope C65 114 x 229 mm
    DMPAPER_ENV_B4 = 33                        ' Envelope B4  250 x 353 mm
    DMPAPER_ENV_B5 = 34                        ' Envelope B5  176 x 250 mm
    DMPAPER_ENV_B6 = 35                        ' Envelope B6  176 x 125 mm
    DMPAPER_ENV_ITALY = 36                     ' Envelope 110 x 230 mm
    DMPAPER_ENV_MONARCH = 37                   ' Envelope Monarch 3.875 x 7.5 in
    DMPAPER_ENV_PERSONAL = 38                  ' 6 3/4 Envelope 3 5/8 x 6 1/2 in
    DMPAPER_FANFOLD_US = 39                    ' US Std Fanfold 14 7/8 x 11 in
    DMPAPER_FANFOLD_STD_GERMAN = 40            ' German Std Fanfold 8 1/2 x 12 in
    DMPAPER_FANFOLD_LGL_GERMAN = 41            ' German Legal Fanfold 8 1/2 x 13 in
    DMPAPER_ISO_B4 = 42                        ' B4 (ISO) 250 x 353 mm
    DMPAPER_JAPANESE_POSTCARD = 43             ' Japanese Postcard 100 x 148 mm
    DMPAPER_9X11 = 44                          ' 9 x 11 in
    DMPAPER_1 = &H11 = 45                      ' 10 x 11 in
    DMPAPER_15X11 = 46                         ' 15 x 11 in
    DMPAPER_ENV_INVITE = 47                    ' Envelope Invite 220 x 220 mm
    DMPAPER_RESERVED_48 = 48                   ' RESERVED--DO NOT USE
    DMPAPER_RESERVED_49 = 49                   ' RESERVED--DO NOT USE
    DMPAPER_LETTER_EXTRA = 50                  ' Letter Extra 9 \275 x 12 in
    DMPAPER_LEGAL_EXTRA = 51                   ' Legal Extra 9 \275 x 15 in
    DMPAPER_TABLOID_EXTRA = 52                 ' Tabloid Extra 11.69 x 18 in
    DMPAPER_A4_EXTRA = 53                      ' A4 Extra 9.27 x 12.69 in
    DMPAPER_LETTER_TRANSVERSE = 54             ' Letter Transverse 8 \275 x 11 in
    DMPAPER_A4_TRANSVERSE = 55                 ' A4 Transverse 210 x 297 mm
    DMPAPER_LETTER_EXTRA_TRANSVERSE = 56       ' Letter Extra Transverse 9\275 x 12 in
    DMPAPER_A_PLUS = 57                        ' SuperA/SuperA/A4 227 x 356 mm
    DMPAPER_B_PLUS = 58                        ' SuperB/SuperB/A3 305 x 487 mm
    DMPAPER_LETTER_PLUS = 59                   ' Letter Plus 8.5 x 12.69 in
    DMPAPER_A4_PLUS = 60                       ' A4 Plus 210 x 330 mm
    DMPAPER_A5_TRANSVERSE = 61                 ' A5 Transverse 148 x 210 mm
    DMPAPER_B5_TRANSVERSE = 62                 ' B5 (JIS) Transverse 182 x 257 mm
    DMPAPER_A3_EXTRA = 63                      ' A3 Extra 322 x 445 mm
    DMPAPER_A5_EXTRA = 64                      ' A5 Extra 174 x 235 mm
    DMPAPER_B5_EXTRA = 65                      ' B5 (ISO) Extra 201 x 276 mm
    DMPAPER_A2 = 66                            ' A2 420 x 594 mm
    DMPAPER_A3_TRANSVERSE = 67                 ' A3 Transverse 297 x 420 mm
    DMPAPER_A3_EXTRA_TRANSVERSE = 68           ' A3 Extra Transverse 322 x 445 mm
    DMPAPER_DBL_JAPANESE_POSTCARD = 69         ' Japanese Double Postcard 200 x 148 mm
    DMPAPER_A6 = 70                            ' A6 105 x 148 mm
    DMPAPER_JENV_KAKU2 = 71                    ' Japanese Envelope Kaku #2
    DMPAPER_JENV_KAKU3 = 72                    ' Japanese Envelope Kaku #3
    DMPAPER_JENV_CHOU3 = 73                    ' Japanese Envelope Chou #3
    DMPAPER_JENV_CHOU4 = 74                    ' Japanese Envelope Chou #4
    DMPAPER_LETTER_ROTATED = 75                ' Letter Rotated 11 x 8 1/2 11 in
    DMPAPER_A3_ROTATED = 76                    ' A3 Rotated 420 x 297 mm
    DMPAPER_A4_ROTATED = 77                    ' A4 Rotated 297 x 210 mm
    DMPAPER_A5_ROTATED = 78                    ' A5 Rotated 210 x 148 mm
    DMPAPER_B4_JIS_ROTATED = 79                ' B4 (JIS) Rotated 364 x 257 mm
    DMPAPER_B5_JIS_ROTATED = 80                ' B5 (JIS) Rotated 257 x 182 mm
    DMPAPER_JAPANESE_POSTCARD_ROTATED = 81     ' Japanese Postcard Rotated 148 x 100 mm
    DMPAPER_DBL_JAPANESE_POSTCARD_ROTATED = 82 ' Double Japanese Postcard Rotated 148 x 200 mm
    DMPAPER_A6_ROTATED = 83                    ' A6 Rotated 148 x 105 mm
    DMPAPER_JENV_KAKU2_ROTATED = 84            ' Japanese Envelope Kaku #2 Rotated
    DMPAPER_JENV_KAKU3_ROTATED = 85            ' Japanese Envelope Kaku #3 Rotated
    DMPAPER_JENV_CHOU3_ROTATED = 86            ' Japanese Envelope Chou #3 Rotated
    DMPAPER_JENV_CHOU4_ROTATED = 87            ' Japanese Envelope Chou #4 Rotated
    DMPAPER_B6_JIS = 88                        ' B6 (JIS) 128 x 182 mm
    DMPAPER_B6_JIS_ROTATED = 89                ' B6 (JIS) Rotated 182 x 128 mm
    DMPAPER_12X11 = 90                         ' 12 x 11 in
    DMPAPER_JENV_YOU4 = 91                     ' Japanese Envelope You #4
    DMPAPER_JENV_YOU4_ROTATED = 92             ' Japanese Envelope You #4 Rotated
    DMPAPER_P16K = 93                          ' PRC 16K 146 x 215 mm
    DMPAPER_P32K = 94                          ' PRC 32K 97 x 151 mm
    DMPAPER_P32KBIG = 95                       ' PRC 32K(Big) 97 x 151 mm
    DMPAPER_PENV_1 = 96                        ' PRC Envelope #1 102 x 165 mm
    DMPAPER_PENV_2 = 97                        ' PRC Envelope #2 102 x 176 mm
    DMPAPER_PENV_3 = 98                        ' PRC Envelope #3 125 x 176 mm
    DMPAPER_PENV_4 = 99                        ' PRC Envelope #4 110 x 208 mm
    DMPAPER_PENV_5 = 100                       ' PRC Envelope #5 110 x 220 mm
    DMPAPER_PENV_6 = 101                       ' PRC Envelope #6 120 x 230 mm
    DMPAPER_PENV_7 = 102                       ' PRC Envelope #7 160 x 230 mm
    DMPAPER_PENV_8 = 103                       ' PRC Envelope #8 120 x 309 mm
    DMPAPER_PENV_9 = 104                       ' PRC Envelope #9 229 x 324 mm
    DMPAPER_PENV_10 = 105                      ' PRC Envelope #10 324 x 458 mm
    DMPAPER_P16K_ROTATED = 106                 ' PRC 16K Rotated
    DMPAPER_P32K_ROTATED = 107                 ' PRC 32K Rotated
    DMPAPER_P32KBIG_ROTATED = 108              ' PRC 32K(Big) Rotated
    DMPAPER_PENV_1_ROTATED = 109               ' PRC Envelope #1 Rotated 165 x 102 mm
    DMPAPER_PENV_2_ROTATED = 110               ' PRC Envelope #2 Rotated 176 x 102 mm
    DMPAPER_PENV_3_ROTATED = 111               ' PRC Envelope #3 Rotated 176 x 125 mm
    DMPAPER_PENV_4_ROTATED = 112               ' PRC Envelope #4 Rotated 208 x 110 mm
    DMPAPER_PENV_5_ROTATED = 113               ' PRC Envelope #5 Rotated 220 x 110 mm
    DMPAPER_PENV_6_ROTATED = 114               ' PRC Envelope #6 Rotated 230 x 120 mm
    DMPAPER_PENV_7_ROTATED = 115               ' PRC Envelope #7 Rotated 230 x 160 mm
    DMPAPER_PENV_8_ROTATED = 116               ' PRC Envelope #8 Rotated 309 x 120 mm
    DMPAPER_PENV_9_ROTATED = 117               ' PRC Envelope #9 Rotated 324 x 229 mm
    DMPAPER_PENV_10_ROTATED = 118              ' PRC Envelope #10 Rotated 458 x 324 mm
  End Enum

  Public Enum CDPaperSourceConsts
    DMBIN_UPPER = 1
    DMBIN_ONLYONE = 1
    DMBIN_LOWER = 2
    DMBIN_MIDDLE = 3
    DMBIN_MANUAL = 4
    DMBIN_ENVELOPE = 5
    DMBIN_ENVMANUAL = 6
    DMBIN_AUTO = 7
    DMBIN_TRACTOR = 8
    DMBIN_SMALLFMT = 9
    DMBIN_LARGEFMT = 10
    DMBIN_LARGECAPACITY = 11
    DMBIN_CASSETTE = 14
    DMBIN_FORMSOURCE = 15
  End Enum

  Public Enum CDPrintDlgExResultActionIDs
    PD_RESULT_CANCEL = 0
    PD_RESULT_PRINT = 1
    PD_RESULT_APPLY = 2
  End Enum

  Public Enum CDPrintDialogFlags
    PD_ALLPAGES = &H0
    PD_SELECTION = &H1
    PD_PAGENUMS = &H2
    PD_NOSELECTION = &H4
    PD_NOPAGENUMS = &H8
    PD_COLLATE = &H10
    PD_PRINTTOFILE = &H20
    PD_PRINTSETUP = &H40
    PD_NOWARNING = &H80
    PD_RETURNDC = &H100
    PD_RETURNIC = &H200
    PD_RETURNDEFAULT = &H400
    PD_SHOWHELP = &H800
    PD_ENABLEPRINTHOOK = &H1000
    PD_ENABLESETUPHOOK = &H2000
    PD_ENABLEPRINTTEMPLATE = &H4000
    PD_ENABLESETUPTEMPLATE = &H8000
    PD_ENABLEPRINTTEMPLATEHANDLE = &H10000
    PD_ENABLESETUPTEMPLATEHANDLE = &H20000
    PD_USEDEVMODECOPIES = &H40000
    PD_USEDEVMODECOPIESANDCOLLATE = &H40000
    PD_DISABLEPRINTTOFILE = &H80000
    PD_HIDEPRINTTOFILE = &H100000
    PD_NONETWORKBUTTON = &H200000
    PD_CURRENTPAGE = &H400000
    PD_NOCURRENTPAGE = &H800000
    PD_EXCLUSIONFLAGS = &H1000000
    PD_USELARGETEMPLATE = &H10000000
  End Enum

  Public Enum CDPrintDialogExclusionFlags
    PD_EXCL_COPIESANDCOLLATE = (DM_COPIES Or DM_COLLATE)
  End Enum

  Public Enum CDPrintQualityConsts
    DMRES_DRAFT = -1
    DMRES_LOW = -2
    DMRES_MEDIUM = -3
    DMRES_HIGH = -4
  End Enum

  Public Enum CDTrueTypeConsts
    DMTT_BITMAP = 1           ' print TT fonts as graphics
    DMTT_DOWNLOAD = 2         ' download TT fonts as soft fonts
    DMTT_SUBDEV = 3           ' substitute device fonts for TT fonts
    DMTT_DOWNLOAD_OUTLINE = 4 ' download TT fonts as outline soft fonts
  End Enum

  Public Enum ExCDOpenFlagsConsts
    OFN_EX_NOPLACESBAR = &H1
  End Enum


  Private Const START_PAGE_GENERAL = &HFFFFFFFF


  Private Type DEVNAMES
    wDriverOffset As Integer
    wDeviceOffset As Integer
    wOutputOffset As Integer
    wDefault As Integer
  End Type

  Private Type DEVMODE
    dmDeviceName(1 To 64) As Byte
    dmSpecVersion As Integer
    dmDriverVersion As Integer
    dmSize As Integer
    dmDriverExtra As Integer
    dmFields As Long
    dmOrientation As Integer
    dmPaperSize As Integer
    dmPaperLength As Integer
    dmPaperWidth As Integer
    dmScale As Integer
    dmCopies As Integer
    dmDefaultSource As Integer
    dmPrintQuality As Integer
    dmColor As Integer
    dmDuplex As Integer
    dmYResolution As Integer
    dmTTOption As Integer
    dmCollate As Integer
    dmFormName(1 To 64) As Byte
    dmUnusedPadding As Integer
    dmBitsPerPel As Long
    dmPelsWidth As Long
    dmPelsHeight As Long
    dmDisplayFlags As Long
    dmDisplayFrequency As Long
    dmICMMethod As Long
    dmICMIntent As Long
    dmMediaType As Long
    dmDitherType As Long
    dmReserved1 As Long
    dmReserved2 As Long
    dmPanningWidth As Long
    dmPanningHeight As Long
  End Type

  Private Type LOGFONT
    lfHeight As Long
    lfWidth As Long
    lfEscapement As Long
    lfOrientation As Long
    lfWeight As Long
'    lfItalic As Byte
'    lfUnderline As Byte
'    lfStrikeOut As Byte
'    lfCharSet As Byte
    lf1 As Long
'    lfOutPrecision As Byte
'    lfClipPrecision As Byte
'    lfQuality As Byte
'    lfPitchAndFamily As Byte
    lf2 As Long
    lfFaceName(1 To 64) As Byte
  End Type

  Private Type OPENFILENAME
    lStructSize As Long
    hwndOwner As Long
    hInstance As Long
    lpstrFilter As Long
    lpstrCustomFilter As Long
    nMaxCustFilter As Long
    nFilterIndex As Long
    lpstrFile As Long
    nMaxFile As Long
    lpstrFileTitle As Long
    nMaxFileTitle As Long
    lpstrInitialDir As Long
    lpstrTitle As Long
    Flags As Long
    nFileOffset As Integer
    nFileExtension As Integer
    lpstrDefExt As Long
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As Long
    pvReserved As Long
    dwReserved As Long
    FlagsEx As Long
  End Type

  Private Type POINTAPI
    X As Long
    Y As Long
  End Type

  Private Type PRINTPAGERANGE
    nFromPage As Long
    nToPage As Long
  End Type

  Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
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

  Private Type TCHOOSEFONT
    lStructSize As Long
    hwndOwner As Long
    hDC As Long
    lpLogFont As Long
    iPointSize As Long
    Flags As Long
    rgbColors As Long
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As Long
    hInstance As Long
    lpszStyle As Long
    nFontType As Integer
    WRONG_ALIGNMENT As Integer
    nSizeMin As Long
    nSizeMax As Long
  End Type

  Private Type TPAGESETUPDLG
    lStructSize As Long
    hwndOwner As Long
    hDevMode As Long
    hDevNames As Long
    Flags As Long
    ptPaperSize As POINTAPI
    rtMinMargin As RECT
    rtMargin As RECT
    hInstance As Long
    lCustData As Long
    lpfnPageSetupHook As Long
    lpfnPagePaintHook As Long
    lpPageSetupTemplateName As Long
    hPageSetupTemplate As Long
  End Type

  Private Type TPRINTDLG
    lStructSize As Long
    hwndOwner As Long
    hDevMode As Long
    hDevNames As Long
    hDC As Long
    Flags As CDPrintDialogFlags
    nFromPage As Integer
    nToPage As Integer
    nMinPage As Integer
    nMaxPage As Integer
    nCopies As Integer
    ' right, now we have serious alignment problems.
    ' We're two bytes off for the next 8 fields...
    b(0 To 31) As Byte
    'hInstance As Long              '  0 - 3
    'lCustData As Long              '  4 - 7
    'lpfnPrintHook As Long          '  8 - 11
    'lpfnSetupHook As Long          ' 12 - 15
    'lpPrintTemplateName As Long    ' 16 - 19
    'lpSetupTemplateName As Long    ' 20 - 23
    'hPrintTemplate As Long         ' 24 - 27
    'hSetupTemplate As Long         ' 28 - 31
  End Type

  Private Type TPRINTDLGEX
    lStructSize As Long
    hwndOwner As Long
    hDevMode As Long
    hDevNames As Long
    hDC As Long
    Flags As CDPrintDialogFlags
    Flags2 As Long               ' reserved
    ExclusionFlags As Long
    nPageRanges As Long
    nMaxPageRanges As Long
    lpPageRanges As Long         ' array of PRINTPAGERANGEs
    nMinPage As Long
    nMaxPage As Long
    nCopies As Long
    hInstance As Long
    lpPrintTemplateName As Long
    lpCallback As Long
    nPropertyPages As Long
    lphPropertyPages As Long     ' array of PROPSHEETPAGE handles
    nStartPage As Long
    dwResultAction As Long
  End Type


  Private colFilters As Collection
  Private colRanges As Collection


  ' vars used to store the values of public properties

  Private propActiveFilter As Integer
  Private propCaption As String
  Private propColorFlags As CDColorFlags
  Private propCopies As Integer
  Private propCustomColors(15) As Long
  Private propDefaultColor As Long
  Private propDefaultExtension As String
  Private propDefaultFile As String
  Private propDeviceName As String
  Private propDriverName As String
  Private propDuplex As CDDuplexConsts
  Private propExclusionFlags As CDPrintDialogExclusionFlags
  Private propExOpenFlags As ExCDOpenFlagsConsts
  Private propFontColor As Long
  Private propFontFlags As CDFontFlags
  Private propFontStyle As String
  Private propFontType As CDFontTypeConsts
  Private prophWndParent As Long
  Private propICM As CDICMConsts
  Private propICMMode As CDICMModeConsts
  Private propInitDeviceData As Boolean
  Private propMaxFontSize As Long
  Private propMaxPage As Long
  Private propMediaType As CDMediaTypeConsts
  Private propMinFontSize As Long
  Private propMinPage As Long
  Private propMinPageBorderBottom As Long
  Private propMinPageBorderLeft As Long
  Private propMinPageBorderRight As Long
  Private propMinPageBorderTop As Long
  Private propMonochrome As Boolean
  Private propOpenFlags As CDOpenFlagsConsts
  Private propOrientation As CDOrientationConsts
  Private propOutputName As String
  Private propPageBorderBottom As Long
  Private propPageBorderLeft As Long
  Private propPageBorderRight As Long
  Private propPageBorderTop As Long
  Private propPageSetupFlags As CDPageSetupFlags
  Private propPaperLength As Long
  Private propPaperSize As CDPaperSizeConsts
  Private propPaperSizeName As String
  Private propPaperSource As CDPaperSourceConsts
  Private propPaperWidth As Long
  Private propPrintDlgFlags As CDPrintDialogFlags
  Private propPrinterhDC As Long
  Private propPrintQuality As CDPrintQualityConsts
  Private propSort As Boolean
  Private propStartDirection As String
  Private propTrueTypeMode As CDTrueTypeConsts
  Private propUseCustomColors As Boolean
  Private propUseDefaultColor As Boolean
  Private propxResolution As Integer
  Private propyResolution As Integer
  Private propZoomPercent As Integer


  Private Declare Function ChooseColor Lib "comdlg32.dll" Alias "ChooseColorW" (lpcc As TCHOOSECOLOR) As Long
  Private Declare Function ChooseFont Lib "comdlg32.dll" Alias "ChooseFontW" (lpcf As TCHOOSEFONT) As Long
  Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (ByVal pDest As Long, ByVal pSrc As Long, ByVal sz As Long)
  Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameW" (lpofn As OPENFILENAME) As Long
  Private Declare Function GetSaveFileName Lib "comdlg32.dll" Alias "GetSaveFileNameW" (lpofn As OPENFILENAME) As Long
  Private Declare Function GetSysColor Lib "user32.dll" (ByVal nIndex As Long) As Long
  Private Declare Function GlobalAlloc Lib "kernel32.dll" (ByVal uFlags As Long, ByVal dwBytes As Long) As Long
  Private Declare Function GlobalFree Lib "kernel32.dll" (ByVal hMem As Long) As Long
  Private Declare Function GlobalLock Lib "kernel32.dll" (ByVal hMem As Long) As Long
  Private Declare Function GlobalUnlock Lib "kernel32.dll" (ByVal hMem As Long) As Long
  Private Declare Function lstrcpyn Lib "kernel32.dll" Alias "lstrcpynW" (ByVal lpString1 As Long, ByVal lpString2 As Long, ByVal iMaxLength As Long) As Long
  Private Declare Function lstrlen Lib "kernel32.dll" Alias "lstrlenW" (ByVal lpString As Long) As Long
  Private Declare Function PageSetupDlg Lib "comdlg32.dll" Alias "PageSetupDlgW" (lppsd As TPAGESETUPDLG) As Long
  Private Declare Function PathAddBackslash Lib "shlwapi.dll" Alias "PathAddBackslashW" (ByVal lpszPath As Long) As Long
  Private Declare Function PrintDlg Lib "comdlg32.dll" Alias "PrintDlgW" (lppd As TPRINTDLG) As Long
  Private Declare Function PrintDlgEx Lib "comdlg32.dll" Alias "PrintDlgExW" (lppd As TPRINTDLGEX) As Long


Private Sub Class_Initialize()
  Set colFilters = New Collection
  Set colRanges = New Collection
End Sub

Private Sub Class_Terminate()
  Set colFilters = Nothing
  Set colRanges = Nothing
End Sub


Public Property Get ActiveFilter() As Integer
  ActiveFilter = propActiveFilter
End Property

Public Property Let ActiveFilter(ByVal New_ActiveFilter As Integer)
  propActiveFilter = New_ActiveFilter
End Property

Public Property Get Caption() As String
  Caption = propCaption
End Property

Public Property Let Caption(ByVal New_Caption As String)
  propCaption = New_Caption
End Property

Public Property Get ColorFlags() As CDColorFlags
  ColorFlags = propColorFlags
End Property

Public Property Let ColorFlags(ByVal New_ColorFlags As CDColorFlags)
  propColorFlags = New_ColorFlags
End Property

Public Property Get Copies() As Integer
  Copies = propCopies
End Property

Public Property Let Copies(ByVal New_Copies As Integer)
  propCopies = New_Copies
End Property

Public Property Get CustomColors(ByVal Index As Integer) As Long
  If Not IsWithinRange(Index, 0, 15) Then Exit Property

  CustomColors = propCustomColors(Index)
End Property

Public Property Let CustomColors(ByVal Index As Integer, ByVal New_CustomColors As Long)
  If Not IsWithinRange(Index, 0, 15) Then Exit Property

  propCustomColors(Index) = New_CustomColors
End Property

Public Property Get DefaultColor() As Long
  DefaultColor = propDefaultColor
End Property

Public Property Let DefaultColor(ByVal New_DefaultColor As Long)
  propDefaultColor = New_DefaultColor
End Property

Public Property Get DefaultExtension() As String
  DefaultExtension = propDefaultExtension
End Property

Public Property Let DefaultExtension(ByVal New_DefaultExtension As String)
  propDefaultExtension = New_DefaultExtension
End Property

Public Property Get DefaultFile() As String
  DefaultFile = propDefaultFile
End Property

Public Property Let DefaultFile(ByVal New_DefaultFile As String)
  propDefaultFile = New_DefaultFile
End Property

Public Property Get DeviceName() As String
  DeviceName = propDeviceName
End Property

Public Property Let DeviceName(ByVal New_DeviceName As String)
  propDeviceName = New_DeviceName
End Property

Public Property Get DriverName() As String
  DriverName = propDriverName
End Property
'
'Public Property Let DriverName(ByVal New_DriverName As String)
'  propDriverName = New_DriverName
'End Property

Public Property Get Duplex() As CDDuplexConsts
  Duplex = propDuplex
End Property

Public Property Let Duplex(ByVal New_Duplex As CDDuplexConsts)
  propDuplex = New_Duplex
End Property

Public Property Get ExclusionFlags() As CDPrintDialogExclusionFlags
  ExclusionFlags = propExclusionFlags
End Property

Public Property Let ExclusionFlags(ByVal New_ExclusionFlags As CDPrintDialogExclusionFlags)
  propExclusionFlags = New_ExclusionFlags
End Property

Public Property Get ExOpenFlags() As ExCDOpenFlagsConsts
  ExOpenFlags = propExOpenFlags
End Property

Public Property Let ExOpenFlags(ByVal New_ExOpenFlags As ExCDOpenFlagsConsts)
  propExOpenFlags = New_ExOpenFlags
End Property

Public Property Get FilterCount() As Integer
  FilterCount = colFilters.Count / 2
End Property

Public Property Get FilterDescription(ByVal iFilter As Integer) As String
  If Not IsWithinRange(iFilter, 1, Me.FilterCount) Then
    Err.Raise 9
    Exit Property
  End If

  FilterDescription = colFilters(iFilter * 2 - 1)
End Property

Public Property Get FilterMask(ByVal iFilter As Integer) As String
  If Not IsWithinRange(iFilter, 1, Me.FilterCount) Then
    Err.Raise 9
    Exit Property
  End If

  FilterMask = colFilters(iFilter * 2)
End Property

Public Property Get FontColor() As Long
  FontColor = propFontColor
End Property

Public Property Let FontColor(ByVal New_FontColor As Long)
  propFontColor = New_FontColor
End Property

Public Property Get FontFlags() As CDFontFlags
  FontFlags = propFontFlags
End Property

Public Property Let FontFlags(ByVal New_FontFlags As CDFontFlags)
  propFontFlags = New_FontFlags
End Property

Public Property Get FontStyle() As String
  FontStyle = propFontStyle
End Property

Public Property Let FontStyle(ByVal New_FontStyle As String)
  propFontStyle = New_FontStyle
End Property

Public Property Get FontType() As CDFontTypeConsts
  FontType = propFontType
End Property

Public Property Get hWndParent() As Long
  hWndParent = prophWndParent
End Property

Public Property Let hWndParent(ByVal New_hWndParent As Long)
  prophWndParent = New_hWndParent
End Property

Public Property Get ICM() As CDICMConsts
  ICM = propICM
End Property

Public Property Let ICM(ByVal New_ICM As CDICMConsts)
  propICM = New_ICM
End Property

Public Property Get ICMMode() As CDICMModeConsts
  ICMMode = propICMMode
End Property

Public Property Let ICMMode(ByVal New_ICMMode As CDICMModeConsts)
  propICMMode = New_ICMMode
End Property

Public Property Get InitDeviceData() As Boolean
  InitDeviceData = propInitDeviceData
End Property

Public Property Let InitDeviceData(ByVal New_InitDeviceData As Boolean)
  propInitDeviceData = New_InitDeviceData
End Property

Public Property Get MaxFontSize() As Long
  MaxFontSize = propMaxFontSize
End Property

Public Property Let MaxFontSize(ByVal New_MaxFontSize As Long)
  propMaxFontSize = New_MaxFontSize
End Property

Public Property Get MaxPage() As Long
  MaxPage = propMaxPage
End Property

Public Property Let MaxPage(ByVal New_MaxPage As Long)
  propMaxPage = New_MaxPage
End Property

Public Property Get MediaType() As CDMediaTypeConsts
  MediaType = propMediaType
End Property

Public Property Let MediaType(ByVal New_MediaType As CDMediaTypeConsts)
  propMediaType = New_MediaType
End Property

Public Property Get MinFontSize() As Long
  MinFontSize = propMinFontSize
End Property

Public Property Let MinFontSize(ByVal New_MinFontSize As Long)
  propMinFontSize = New_MinFontSize
End Property

Public Property Get MinPage() As Long
  MinPage = propMinPage
End Property

Public Property Let MinPage(ByVal New_MinPage As Long)
  propMinPage = New_MinPage
End Property

Public Property Get MinPageBorderBottom() As Long
  MinPageBorderBottom = propMinPageBorderBottom / 100
End Property

Public Property Let MinPageBorderBottom(ByVal New_MinPageBorderBottom As Long)
  propMinPageBorderBottom = New_MinPageBorderBottom * 100
End Property

Public Property Get MinPageBorderLeft() As Long
  MinPageBorderLeft = propMinPageBorderLeft / 100
End Property

Public Property Let MinPageBorderLeft(ByVal New_MinPageBorderLeft As Long)
  propMinPageBorderLeft = New_MinPageBorderLeft * 100
End Property

Public Property Get MinPageBorderRight() As Long
  MinPageBorderRight = propMinPageBorderRight / 100
End Property

Public Property Let MinPageBorderRight(ByVal New_MinPageBorderRight As Long)
  propMinPageBorderRight = New_MinPageBorderRight * 100
End Property

Public Property Get MinPageBorderTop() As Long
  MinPageBorderTop = propMinPageBorderTop / 100
End Property

Public Property Let MinPageBorderTop(ByVal New_MinPageBorderTop As Long)
  propMinPageBorderTop = New_MinPageBorderTop * 100
End Property

Public Property Get Monochrome() As Boolean
  Monochrome = propMonochrome
End Property

Public Property Let Monochrome(ByVal New_Monochrome As Boolean)
  propMonochrome = New_Monochrome
End Property

Public Property Get OpenFlags() As CDOpenFlagsConsts
  OpenFlags = propOpenFlags
End Property

Public Property Let OpenFlags(ByVal New_OpenFlags As CDOpenFlagsConsts)
  propOpenFlags = New_OpenFlags
End Property

Public Property Get Orientation() As CDOrientationConsts
  Orientation = propOrientation
End Property

Public Property Let Orientation(ByVal New_Orientation As CDOrientationConsts)
  propOrientation = New_Orientation
End Property

Public Property Get OutputName() As String
  OutputName = propOutputName
End Property
'
'Public Property Let OutputName(ByVal New_OutputName As String)
'  propOutputName = New_OutputName
'End Property

Public Property Get PageBorderBottom() As Long
  PageBorderBottom = propPageBorderBottom / 100
End Property

Public Property Let PageBorderBottom(ByVal New_PageBorderBottom As Long)
  propPageBorderBottom = New_PageBorderBottom * 100
End Property

Public Property Get PageBorderLeft() As Long
  PageBorderLeft = propPageBorderLeft / 100
End Property

Public Property Let PageBorderLeft(ByVal New_PageBorderLeft As Long)
  propPageBorderLeft = New_PageBorderLeft * 100
End Property

Public Property Get PageBorderRight() As Long
  PageBorderRight = propPageBorderRight / 100
End Property

Public Property Let PageBorderRight(ByVal New_PageBorderRight As Long)
  propPageBorderRight = New_PageBorderRight * 100
End Property

Public Property Get PageBorderTop() As Long
  PageBorderTop = propPageBorderTop / 100
End Property

Public Property Let PageBorderTop(ByVal New_PageBorderTop As Long)
  propPageBorderTop = New_PageBorderTop * 100
End Property

Public Property Get PageSetupFlags() As CDPageSetupFlags
  PageSetupFlags = propPageSetupFlags
End Property

Public Property Let PageSetupFlags(ByVal New_PageSetupFlags As CDPageSetupFlags)
  propPageSetupFlags = New_PageSetupFlags
End Property

Public Property Get PaperLength() As Long
  PaperLength = propPaperLength / 100
End Property

Public Property Let PaperLength(ByVal New_PaperLength As Long)
  propPaperLength = New_PaperLength * 10
End Property

Public Property Get PaperSize() As CDPaperSizeConsts
  PaperSize = propPaperSize
End Property

Public Property Get PaperSizeName() As String
  PaperSizeName = propPaperSizeName
End Property

Public Property Let PaperSize(ByVal New_PaperSize As CDPaperSizeConsts)
  propPaperSize = New_PaperSize
End Property

Public Property Get PaperSource() As CDPaperSourceConsts
  PaperSource = propPaperSource
End Property

Public Property Let PaperSource(ByVal New_PaperSource As CDPaperSourceConsts)
  propPaperSource = New_PaperSource
End Property

Public Property Get PaperWidth() As Long
  PaperWidth = propPaperWidth / 100
End Property

Public Property Let PaperWidth(ByVal New_PaperWidth As Long)
  propPaperWidth = New_PaperWidth * 10
End Property

Public Property Get PrintDlgFlags() As CDPrintDialogFlags
  PrintDlgFlags = propPrintDlgFlags
End Property

Public Property Let PrintDlgFlags(ByVal New_PrintDlgFlags As CDPrintDialogFlags)
  If propSort Then
    New_PrintDlgFlags = New_PrintDlgFlags Or CDPrintDialogFlags.PD_COLLATE
  Else
    New_PrintDlgFlags = New_PrintDlgFlags And Not CDPrintDialogFlags.PD_COLLATE
  End If

  propPrintDlgFlags = New_PrintDlgFlags
End Property

Public Property Get PrinterhDC() As Long
  PrinterhDC = propPrinterhDC
End Property

Public Property Let PrinterhDC(ByVal New_PrinterhDC As Long)
  propPrinterhDC = New_PrinterhDC
End Property

Public Property Get PrintQuality() As CDPrintQualityConsts
  PrintQuality = propPrintQuality
End Property

Public Property Let PrintQuality(ByVal New_PrintQuality As CDPrintQualityConsts)
  propPrintQuality = New_PrintQuality
End Property

Public Property Get RangeCount() As Integer
  RangeCount = colRanges.Count / 2
End Property

Public Property Get RangeFromPage(ByVal iRange As Integer) As String
  If Not IsWithinRange(iRange, 1, Me.RangeCount) Then
    Err.Raise 9
    Exit Property
  End If

  RangeFromPage = colRanges(iRange * 2 - 1)
End Property

Public Property Get RangeToPage(ByVal iRange As Integer) As String
  If Not IsWithinRange(iRange, 1, Me.RangeCount) Then
    Err.Raise 9
    Exit Property
  End If

  RangeToPage = colRanges(iRange * 2)
End Property

Public Property Get Sort() As Boolean
  Sort = propSort
End Property

Public Property Let Sort(ByVal New_Sort As Boolean)
  If New_Sort Then
    propPrintDlgFlags = propPrintDlgFlags Or CDPrintDialogFlags.PD_COLLATE
  Else
    propPrintDlgFlags = propPrintDlgFlags And Not CDPrintDialogFlags.PD_COLLATE
  End If

  propSort = New_Sort
End Property

Public Property Get StartDirection() As String
  StartDirection = propStartDirection
End Property

Public Property Let StartDirection(ByVal New_StartDirection As String)
  propStartDirection = New_StartDirection
End Property

Public Property Get TrueTypeMode() As CDTrueTypeConsts
  TrueTypeMode = propTrueTypeMode
End Property

Public Property Let TrueTypeMode(ByVal New_TrueTypeMode As CDTrueTypeConsts)
  propTrueTypeMode = New_TrueTypeMode
End Property

Public Property Get UseCustomColors() As Boolean
  UseCustomColors = propUseCustomColors
End Property

Public Property Let UseCustomColors(ByVal New_UseCustomColors As Boolean)
  propUseCustomColors = New_UseCustomColors
End Property

Public Property Get UseDefaultColor() As Boolean
  UseDefaultColor = propUseDefaultColor
End Property

Public Property Let UseDefaultColor(ByVal New_UseDefaultColor As Boolean)
  propUseDefaultColor = New_UseDefaultColor
End Property

Public Property Get xResolution() As Integer
  xResolution = propxResolution
End Property

Public Property Let xResolution(ByVal New_xResolution As Integer)
  propxResolution = New_xResolution
End Property

Public Property Get yResolution() As Integer
  yResolution = propyResolution
End Property

Public Property Let yResolution(ByVal New_yResolution As Integer)
  propyResolution = New_yResolution
End Property

Public Property Get ZoomPercent() As Integer
  ZoomPercent = propZoomPercent
End Property

Public Property Let ZoomPercent(ByVal New_ZoomPercent As Integer)
  propZoomPercent = New_ZoomPercent
End Property


Private Function GetDevNameString(ByVal pDevNames As Long, ByVal pOffset As Long) As String
  Dim ret As String
  Dim l As Long

  l = lstrlen(pDevNames + pOffset)
  If l > 0 Then
    ret = String$(l + 1, Chr$(0))
    lstrcpyn StrPtr(ret), pDevNames + pOffset, l + 1
    ret = Left$(ret, l)
  End If

  GetDevNameString = ret
End Function

' returns a string containing all filters
' format: <description> & Chr(0) & <mask> & Chr(0)...
' the string ends with two Chr(0)s
Private Function GetFilters() As String
  Dim Description As String
  Dim i As Long
  Dim Mask As String
  Dim ret As String

  For i = 1 To colFilters.Count Step 2
    Description = colFilters(i)
    Mask = colFilters(i + 1)

    ret = ret & Description & Chr$(0) & Mask & Chr$(0)
  Next i
  If ret <> "" Then ret = ret & Chr$(0)

  GetFilters = ret
End Function

Private Function IsArrayEmpty(Arr As Variant) As Boolean
  Dim tmp As Long

  On Error Resume Next
  tmp = LBound(Arr)
  IsArrayEmpty = Err
End Function

' checks whether <Num> is between <LoBound> and <HiBound>
Private Function IsWithinRange(ByVal Num As Long, ByVal LoBound As Long, ByVal HiBound As Long) As Boolean
  IsWithinRange = ((Num >= LoBound) And (Num <= HiBound))
End Function

' fills <propCustomColors> with the color white
Private Sub WhiteToCustomColors()
  Dim i As Long

  For i = LBound(propCustomColors) To UBound(propCustomColors)
    propCustomColors(i) = ColorConstants.vbWhite
  Next i
End Sub


' adds a filter
Public Sub AddFilter(ByVal Description As String, ByVal Mask As String, Optional ByVal After As Integer = -1)
  If After > 0 Then
    colFilters.Add Description, , , After * 2
    colFilters.Add Mask, , , After * 2 + 1
  Else
    colFilters.Add Description
    colFilters.Add Mask
  End If
End Sub

' adds a range
Public Sub AddRange(ByVal FromPage As Long, ByVal ToPage As Long, Optional ByVal After As Integer = -1)
  If After > 0 Then
    colRanges.Add FromPage, , , After * 2
    colRanges.Add ToPage, , , After * 2 + 1
  Else
    colRanges.Add FromPage
    colRanges.Add ToPage
  End If
End Sub

' removes all filters
Public Sub RemoveAllFilters()
  Set colFilters = New Collection
End Sub

' removes all ranges
Public Sub RemoveAllRanges()
  Set colRanges = New Collection
End Sub

' removes the filter <iFilter>
Public Sub RemoveFilter(ByVal iFilter As Integer)
  ' remove mask
  colFilters.Remove iFilter * 2
  ' remove description
  colFilters.Remove iFilter * 2 - 1
End Sub

' removes the range <iRange>
Public Sub RemoveRange(ByVal iRange As Integer)
  ' remove ToPage
  colRanges.Remove iRange * 2
  ' remove FromPage
  colRanges.Remove iRange * 2 - 1
End Sub

' displays a "Choose color..."-dialog and returns the selected color
Public Function ShowColor(Color As Long) As Boolean
  Dim Data As TCHOOSECOLOR
  Dim ret As Boolean

  ' prepare the API call
  With Data
    .Flags = propColorFlags
    .hInstance = App.hInstance
    .hwndOwner = prophWndParent
    If Not propUseCustomColors Then WhiteToCustomColors
    .lpCustColors = VarPtr(propCustomColors(0))
    If propUseDefaultColor Then
      .Flags = .Flags Or CDColorFlags.CC_RGBINIT
      .rgbResult = propDefaultColor
    End If
    .lStructSize = LenB(Data)

    ret = ChooseColor(Data)

    If ret Then
      ' return the selected color
      Color = .rgbResult
    End If
  End With

  ShowColor = ret
End Function

' displays a "Choose font..."-dialog and returns the selected font
Public Function ShowFont(SelectedFont As Font) As Boolean
  Dim Data As TCHOOSEFONT
  Dim FontData As LOGFONT
  Dim useLogFont As Boolean
  Dim ret As Boolean
  Dim fontName As String
  Dim b As Byte

  ' prepare the API call
  useLogFont = Not (SelectedFont Is Nothing)
  If useLogFont Then
    With FontData
      fontName = SelectedFont.Name
      lstrcpyn VarPtr(.lfFaceName(1)), StrPtr(fontName), 64
      '.lfCharSet = SelectedFont.CharSet
      b = SelectedFont.CharSet
      CopyMemory VarPtr(.lf1) + 3, VarPtr(b), 1
      .lfHeight = -(SelectedFont.Size * (20 / Screen.TwipsPerPixelY))
      b = Abs(SelectedFont.Italic)
      CopyMemory VarPtr(.lf1), VarPtr(b), 1
      b = Abs(SelectedFont.Strikethrough)
      CopyMemory VarPtr(.lf1) + 2, VarPtr(b), 1
      b = Abs(SelectedFont.underline)
      CopyMemory VarPtr(.lf1) + 1, VarPtr(b), 1
      .lfWeight = SelectedFont.Weight
    End With
  End If

  With Data
    .Flags = propFontFlags
    If useLogFont Then .Flags = .Flags Or CDFontFlags.CF_INITTOLOGFONTSTRUCT
    .hDC = propPrinterhDC
    .hInstance = App.hInstance
    .hwndOwner = prophWndParent
    If useLogFont Then
      .iPointSize = SelectedFont.Size * 10
      .lpLogFont = VarPtr(FontData)
    End If
    .lpszStyle = StrPtr(propFontStyle)
    .nSizeMax = propMaxFontSize
    .nSizeMin = propMinFontSize
    .rgbColors = propFontColor
    .lStructSize = LenB(Data)
  End With

  ret = ChooseFont(Data)

  If ret Then
    ' return the selected font
    With FontData
      CopyMemory VarPtr(b), VarPtr(.lf1) + 3, 1
      SelectedFont.CharSet = b
      CopyMemory VarPtr(b), VarPtr(.lf1), 1
      SelectedFont.Italic = b
      fontName = String$(65, Chr$(0))
      lstrcpyn StrPtr(fontName), VarPtr(.lfFaceName(1)), 64
      SelectedFont.Name = Left$(fontName, lstrlen(StrPtr(fontName)))
      CopyMemory VarPtr(b), VarPtr(.lf1) + 2, 1
      SelectedFont.Strikethrough = b
      CopyMemory VarPtr(b), VarPtr(.lf1) + 1, 1
      SelectedFont.underline = b
      SelectedFont.Weight = .lfWeight
      SelectedFont.Bold = (.lfWeight >= 700)
    End With

    With Data
      SelectedFont.Size = .iPointSize / 10
      FontColor = .rgbColors
      FontStyle = Left$(propFontStyle, lstrlen(StrPtr(propFontStyle)))
      propFontType = .nFontType
    End With
  End If

  ShowFont = ret
End Function

' displays an "Open file..."-dialog and returns the selected file
Public Function ShowOpen(File As String) As Boolean
  Dim Data As OPENFILENAME
  Dim pos As Long
  Dim ret As Boolean
  Dim str As String
  Dim filters As String
  Dim Folder As String
  Dim arrFiles() As String
  Dim i As Long

  ' prepare the API call
  str = propDefaultFile & String$(4096 - Len(propDefaultFile), Chr$(0))
  filters = GetFilters
  With Data
    .Flags = propOpenFlags
    .FlagsEx = propExOpenFlags
    .hInstance = App.hInstance
    .hwndOwner = prophWndParent
    .lpstrDefExt = StrPtr(propDefaultExtension)
    .lpstrFile = StrPtr(str)
    .lpstrFilter = StrPtr(filters)
    .lpstrInitialDir = StrPtr(propStartDirection)
    .lpstrTitle = StrPtr(propCaption)
    .nFilterIndex = propActiveFilter
    .nMaxFile = 4096
    .lStructSize = LenB(Data)

    ret = GetOpenFileName(Data)

    If ret Then
      ' return the selected file
      If propOpenFlags And CDOpenFlagsConsts.OFN_ALLOWMULTISELECT Then
        Do
          pos = lstrlen(StrPtr(str))
          If pos > 0 Then
            Mid$(str, pos + 1, 1) = "|"
            If pos + 1 = lstrlen(StrPtr(str)) Then
              ' reached end, remove last "|"
              str = Left$(str, pos)
              pos = 0
            End If
          End If
        Loop Until pos = 0
        pos = InStr(str, "|")
        If pos > 0 Then
          Folder = Left$(str, pos - 1) & Chr$(0) & Chr$(0)
          Call PathAddBackslash(StrPtr(Folder))
          Folder = Left$(Folder, lstrlen(StrPtr(Folder)))
          
          arrFiles = Split(Mid$(str, pos + 1), "|")
          For i = LBound(arrFiles) To UBound(arrFiles)
            arrFiles(i) = Folder & arrFiles(i)
          Next i
          File = Join$(arrFiles, "|")
        End If
      Else
        File = Left$(str, lstrlen(StrPtr(str)))
      End If
    End If
  End With

  ShowOpen = ret
End Function

' displays a "Setup page..."-dialog
Public Function ShowPageSetup() As Boolean
  Const DMCOLOR_COLOR = 2
  Const DMCOLOR_MONOCHROME = 1
  Const GMEM_FIXED = &H0
  Const GMEM_ZEROINIT = &H40
  Dim Data As TPAGESETUPDLG
  Dim DevModeData As DEVMODE
  Dim Fields As CDFields
  Dim hDevMode As Long
  Dim pDevNames As Long
  Dim phDevMode As Long
  Dim ret As Boolean
  Dim tDevNames As DEVNAMES

  ' prepare the API call
  If propInitDeviceData Then
    With DevModeData
      .dmSize = LenB(DevModeData)
      .dmCollate = Abs(propSort)
      .dmColor = IIf(propMonochrome, DMCOLOR_MONOCHROME, DMCOLOR_COLOR)
      .dmCopies = propCopies
      .dmDefaultSource = propPaperSource
      lstrcpyn VarPtr(.dmDeviceName(1)), StrPtr(propDeviceName), 64
      .dmDuplex = propDuplex
      Fields = CDFields.DM_COLLATE Or CDFields.DM_COLOR Or CDFields.DM_COPIES Or CDFields.DM_DEFAULTSOURCE Or CDFields.DM_DUPLEX Or CDFields.DM_ICMINTENT Or CDFields.DM_ICMMETHOD Or CDFields.DM_MEDIATYPE Or CDFields.DM_ORIENTATION Or CDFields.DM_PAPERLENGTH Or CDFields.DM_PAPERSIZE Or CDFields.DM_PAPERWIDTH Or CDFields.DM_PRINTQUALITY Or CDFields.DM_SCALE Or CDFields.DM_TTOPTION Or CDFields.DM_YRESOLUTION
      .dmFields = Fields
      .dmICMIntent = propICMMode
      .dmICMMethod = propICM
      .dmMediaType = propMediaType
      .dmOrientation = propOrientation
      .dmPaperLength = propPaperLength
      .dmPaperSize = propPaperSize
      .dmPaperWidth = propPaperWidth
      .dmPrintQuality = IIf(propxResolution > 0, propxResolution, propPrintQuality)
      .dmScale = propZoomPercent
      .dmTTOption = propTrueTypeMode
      .dmYResolution = propyResolution
    End With

    hDevMode = GlobalAlloc(GMEM_ZEROINIT Or GMEM_FIXED, LenB(DevModeData))
    If hDevMode Then
      phDevMode = GlobalLock(hDevMode)
      If phDevMode Then
        CopyMemory phDevMode, VarPtr(DevModeData), LenB(DevModeData)
        GlobalUnlock hDevMode
      End If
    End If
  End If

  ' ToDo: implement setting hDevNames

  With Data
    .Flags = propPageSetupFlags Or CDPageSetupFlags.PSD_INHUNDREDTHSOFMILLIMETERS
    .hDevMode = hDevMode
    .hInstance = App.hInstance
    .hwndOwner = prophWndParent
    With .rtMargin
      .Bottom = propPageBorderBottom
      .Left = propPageBorderLeft
      .Right = propPageBorderRight
      .Top = propPageBorderTop
    End With
    With .rtMinMargin
      .Bottom = propMinPageBorderBottom
      .Left = propMinPageBorderLeft
      .Right = propMinPageBorderRight
      .Top = propMinPageBorderTop
    End With
    .lStructSize = LenB(Data)
  End With

  ret = PageSetupDlg(Data)

  ' free the memory we used
  If hDevMode Then
    GlobalFree hDevMode
    hDevMode = 0
  End If

  If Data.hDevMode Then
    phDevMode = GlobalLock(Data.hDevMode)
    If phDevMode Then
      CopyMemory VarPtr(DevModeData), phDevMode, LenB(DevModeData)
      GlobalUnlock Data.hDevMode
    End If
  End If
  If ret Then
    ' return the selected settings
    With DevModeData
      propCopies = .dmCopies
      propDuplex = .dmDuplex
      propICM = .dmICMMethod
      propICMMode = .dmICMIntent
      propMediaType = .dmMediaType
      propMonochrome = .dmColor = DMCOLOR_MONOCHROME
      propOrientation = .dmOrientation
      propPaperLength = .dmPaperLength
      propPaperSize = .dmPaperSize
      propPaperSizeName = String$(65, Chr$(0))
      lstrcpyn StrPtr(propPaperSizeName), VarPtr(.dmFormName(1)), 64
      propPaperSizeName = Left$(propPaperSizeName, lstrlen(StrPtr(propPaperSizeName)))
      propPaperSource = .dmDefaultSource
      propPaperWidth = .dmPaperWidth
      propPrintQuality = .dmPrintQuality
      propSort = .dmCollate
      propTrueTypeMode = .dmTTOption
      propxResolution = .dmPrintQuality
      propyResolution = .dmYResolution
      propZoomPercent = .dmScale
    End With
    With Data
      If .hDevNames Then
        pDevNames = GlobalLock(.hDevNames)
        If pDevNames Then
          CopyMemory VarPtr(tDevNames), pDevNames, LenB(tDevNames)
          ' I don't know why we have to add 4 extra bytes. Alignment?
          propDriverName = GetDevNameString(pDevNames + 4, tDevNames.wDriverOffset)
          propDeviceName = GetDevNameString(pDevNames + 4, tDevNames.wDeviceOffset)
          propOutputName = GetDevNameString(pDevNames + 4, tDevNames.wOutputOffset)
          GlobalUnlock .hDevNames
        End If
      End If

      With .rtMargin
        propPageBorderBottom = .Bottom
        propPageBorderLeft = .Left
        propPageBorderRight = .Right
        propPageBorderTop = .Top
      End With
    End With
  End If

  ShowPageSetup = ret
End Function

' displays a "Choose printer..."-dialog and returns the selected printer
Public Function ShowPrinter() As Boolean
  Const DMCOLOR_COLOR = 2
  Const DMCOLOR_MONOCHROME = 1
  Const GMEM_FIXED = &H0
  Const GMEM_ZEROINIT = &H40
  Dim Data As TPRINTDLG
  Dim DevModeData As DEVMODE
  Dim Fields As CDFields
  Dim hDevMode As Long
  Dim pDevNames As Long
  Dim phDevMode As Long
  Dim ret As Boolean
  Dim tDevNames As DEVNAMES
  Dim tmp As Long

  ' prepare the API call
  If propInitDeviceData Then
    With DevModeData
      .dmSize = LenB(DevModeData)
      .dmCollate = Abs(propSort)
      .dmColor = IIf(propMonochrome, DMCOLOR_MONOCHROME, DMCOLOR_COLOR)
      .dmCopies = propCopies
      .dmDefaultSource = propPaperSource
      lstrcpyn VarPtr(.dmDeviceName(1)), StrPtr(propDeviceName), 64
      .dmDuplex = propDuplex
      Fields = CDFields.DM_COLLATE Or CDFields.DM_COLOR Or CDFields.DM_COPIES Or CDFields.DM_DEFAULTSOURCE Or CDFields.DM_DUPLEX Or CDFields.DM_ICMINTENT Or CDFields.DM_ICMMETHOD Or CDFields.DM_MEDIATYPE Or CDFields.DM_ORIENTATION Or CDFields.DM_PAPERLENGTH Or CDFields.DM_PAPERSIZE Or CDFields.DM_PAPERWIDTH Or CDFields.DM_PRINTQUALITY Or CDFields.DM_SCALE Or CDFields.DM_TTOPTION Or CDFields.DM_YRESOLUTION
      .dmFields = Fields
      .dmICMIntent = propICMMode
      .dmICMMethod = propICM
      .dmMediaType = propMediaType
      .dmOrientation = propOrientation
      .dmPaperLength = propPaperLength
      .dmPaperSize = propPaperSize
      .dmPaperWidth = propPaperWidth
      .dmPrintQuality = IIf(propxResolution > 0, propxResolution, propPrintQuality)
      .dmScale = propZoomPercent
      .dmTTOption = propTrueTypeMode
      .dmYResolution = propyResolution
    End With

    hDevMode = GlobalAlloc(GMEM_ZEROINIT Or GMEM_FIXED, LenB(DevModeData))
    If hDevMode Then
      phDevMode = GlobalLock(hDevMode)
      If phDevMode Then
        CopyMemory phDevMode, VarPtr(DevModeData), LenB(DevModeData)
        GlobalUnlock hDevMode
      End If
    End If
  End If

  ' ToDo: implement setting hDevNames

  With Data
    .Flags = propPrintDlgFlags
    tmp = App.hInstance
    CopyMemory VarPtr(.b(0)), VarPtr(tmp), 4
    .hwndOwner = prophWndParent
    .hDevMode = hDevMode
    .nCopies = propCopies
    .nMaxPage = propMaxPage
    .nMinPage = propMinPage
    On Error Resume Next
    .nFromPage = colRanges(1)
    .nToPage = colRanges(2)
    .lStructSize = Len(Data)
  End With

  ret = (PrintDlg(Data) = 1)

  ' free the memory we used
  If hDevMode Then
    GlobalFree hDevMode
    hDevMode = 0
  End If

  If Data.hDevMode Then
    phDevMode = GlobalLock(Data.hDevMode)
    If phDevMode Then
      CopyMemory VarPtr(DevModeData), phDevMode, LenB(DevModeData)
      GlobalUnlock Data.hDevMode
    End If
  End If
  If ret Then
    ' return the selected printer
    With DevModeData
      propCopies = .dmCopies
      propDuplex = .dmDuplex
      propICM = .dmICMMethod
      propICMMode = .dmICMIntent
      propMediaType = .dmMediaType
      propMonochrome = .dmColor = DMCOLOR_MONOCHROME
      propOrientation = .dmOrientation
      propPaperLength = .dmPaperLength
      propPaperSize = .dmPaperSize
      propPaperSizeName = Left$(.dmFormName, lstrlen(StrPtr(.dmFormName)))
      propPaperSource = .dmDefaultSource
      propPaperWidth = .dmPaperWidth
      propPrintQuality = .dmPrintQuality
      propSort = .dmCollate
      propTrueTypeMode = .dmTTOption
      propxResolution = .dmPrintQuality
      propyResolution = .dmYResolution
      propZoomPercent = .dmScale
    End With
    With Data
      If .hDevNames Then
        pDevNames = GlobalLock(.hDevNames)
        If pDevNames Then
          CopyMemory VarPtr(tDevNames), pDevNames, LenB(tDevNames)
          ' I don't know why we have to add 4 extra bytes. Alignment?
          propDriverName = GetDevNameString(pDevNames + 4, tDevNames.wDriverOffset)
          propDeviceName = GetDevNameString(pDevNames + 4, tDevNames.wDeviceOffset)
          propOutputName = GetDevNameString(pDevNames + 4, tDevNames.wOutputOffset)
          GlobalUnlock .hDevNames
        End If
      End If

      propCopies = .nCopies
      Me.RemoveAllRanges
      Me.AddRange .nFromPage, .nToPage
      propPrintDlgFlags = .Flags
      propPrinterhDC = .hDC
      propSort = .Flags And PD_COLLATE
    End With
  End If

  ShowPrinter = ret
End Function

' displays a "Choose printer..."-dialog and returns the selected printer
Public Function ShowPrinterEx(Optional ByVal iInitialPropPage As Long = START_PAGE_GENERAL, Optional ByVal arrAdditionalPropertyPages As Variant) As CDPrintDlgExResultActionIDs
  Const DMCOLOR_COLOR = 2
  Const DMCOLOR_MONOCHROME = 1
  Const GMEM_FIXED = &H0
  Const GMEM_ZEROINIT = &H40
  Const S_OK = &H0
  Dim Data As TPRINTDLGEX
  Dim DevModeData As DEVMODE
  Dim Fields As CDFields
  Dim hDevMode As Long
  Dim i As Long
  Dim pDevNames As Long
  Dim phDevMode As Long
  Dim Ranges() As PRINTPAGERANGE
  Dim ret As CDPrintDlgExResultActionIDs
  Dim ret2 As Boolean
  Dim tDevNames As DEVNAMES

  ' prepare the API call
  If propInitDeviceData Then
    With DevModeData
      .dmSize = LenB(DevModeData)
      .dmCollate = Abs(propSort)
      .dmColor = IIf(propMonochrome, DMCOLOR_MONOCHROME, DMCOLOR_COLOR)
      .dmCopies = propCopies
      .dmDefaultSource = propPaperSource
      lstrcpyn VarPtr(.dmDeviceName(1)), StrPtr(propDeviceName), 64
      .dmDuplex = propDuplex
      Fields = CDFields.DM_COLLATE Or CDFields.DM_COLOR Or CDFields.DM_COPIES Or CDFields.DM_DEFAULTSOURCE Or CDFields.DM_DUPLEX Or CDFields.DM_ICMINTENT Or CDFields.DM_ICMMETHOD Or CDFields.DM_MEDIATYPE Or CDFields.DM_ORIENTATION Or CDFields.DM_PAPERLENGTH Or CDFields.DM_PAPERSIZE Or CDFields.DM_PAPERWIDTH Or CDFields.DM_PRINTQUALITY Or CDFields.DM_SCALE Or CDFields.DM_TTOPTION Or CDFields.DM_YRESOLUTION
      .dmFields = Fields
      .dmICMIntent = propICMMode
      .dmICMMethod = propICM
      .dmMediaType = propMediaType
      .dmOrientation = propOrientation
      .dmPaperLength = propPaperLength
      .dmPaperSize = propPaperSize
      .dmPaperWidth = propPaperWidth
      .dmPrintQuality = IIf(propxResolution > 0, propxResolution, propPrintQuality)
      .dmScale = propZoomPercent
      .dmTTOption = propTrueTypeMode
      .dmYResolution = propyResolution
    End With

    hDevMode = GlobalAlloc(GMEM_ZEROINIT Or GMEM_FIXED, LenB(DevModeData))
    If hDevMode Then
      phDevMode = GlobalLock(hDevMode)
      If phDevMode Then
        CopyMemory phDevMode, VarPtr(DevModeData), LenB(DevModeData)
        GlobalUnlock hDevMode
      End If
    End If
  End If

  ' ToDo: implement setting hDevNames

  With Data
    .Flags = propPrintDlgFlags
    .ExclusionFlags = propExclusionFlags
    .hwndOwner = prophWndParent
    .hDevMode = hDevMode
    .nCopies = propCopies
    .nMaxPage = propMaxPage
    .nMinPage = propMinPage
    .nPageRanges = Me.RangeCount
    .nMaxPageRanges = 100
    ReDim Ranges(1 To .nMaxPageRanges) As PRINTPAGERANGE
    For i = LBound(Ranges) To .nPageRanges
      Ranges(i).nFromPage = Me.RangeFromPage(i)
      Ranges(i).nToPage = Me.RangeToPage(i)
    Next
    .lpPageRanges = VarPtr(Ranges(LBound(Ranges)))
    If IsArrayEmpty(arrAdditionalPropertyPages) Then
      .nPropertyPages = 0
    Else
      .nPropertyPages = UBound(arrAdditionalPropertyPages) - LBound(arrAdditionalPropertyPages) + 1
      .lphPropertyPages = VarPtr(arrAdditionalPropertyPages(LBound(arrAdditionalPropertyPages)))
    End If
    .nStartPage = iInitialPropPage
    .lStructSize = Len(Data)
  End With

  ret2 = (PrintDlgEx(Data) = S_OK)
  ret = Data.dwResultAction

  ' free the memory we used
  If hDevMode Then
    GlobalFree hDevMode
    hDevMode = 0
  End If

  If Data.hDevMode Then
    phDevMode = GlobalLock(Data.hDevMode)
    If phDevMode Then
      CopyMemory VarPtr(DevModeData), phDevMode, LenB(DevModeData)
      GlobalUnlock Data.hDevMode
    End If
  End If
  If ret2 Then
    ' return the selected printer
    With DevModeData
      propCopies = .dmCopies
      propDuplex = .dmDuplex
      propICM = .dmICMMethod
      propICMMode = .dmICMIntent
      propMediaType = .dmMediaType
      propMonochrome = .dmColor = DMCOLOR_MONOCHROME
      propOrientation = .dmOrientation
      propPaperLength = .dmPaperLength
      propPaperSize = .dmPaperSize
      propPaperSizeName = Left$(.dmFormName, lstrlen(StrPtr(.dmFormName)))
      propPaperSource = .dmDefaultSource
      propPaperWidth = .dmPaperWidth
      propPrintQuality = .dmPrintQuality
      propSort = .dmCollate
      propTrueTypeMode = .dmTTOption
      propxResolution = .dmPrintQuality
      propyResolution = .dmYResolution
      propZoomPercent = .dmScale
    End With
    With Data
      If .hDevNames Then
        pDevNames = GlobalLock(.hDevNames)
        If pDevNames Then
          CopyMemory VarPtr(tDevNames), pDevNames, LenB(tDevNames)
          ' I don't know why we have to add 4 extra bytes. Alignment?
          propDriverName = GetDevNameString(pDevNames + 4, tDevNames.wDriverOffset)
          propDeviceName = GetDevNameString(pDevNames + 4, tDevNames.wDeviceOffset)
          propOutputName = GetDevNameString(pDevNames + 4, tDevNames.wOutputOffset)
          GlobalUnlock .hDevNames
        End If
      End If

      propCopies = .nCopies
      Me.RemoveAllRanges
      For i = LBound(Ranges) To .nPageRanges
        Me.AddRange Ranges(i).nFromPage, Ranges(i).nToPage
      Next
      propPrintDlgFlags = .Flags
      propPrinterhDC = .hDC
      propSort = .Flags And PD_COLLATE
    End With
  End If

  ShowPrinterEx = ret
End Function

' displays a "Save file..."-dialog and returns the selected file
Public Function ShowSave(File As String) As Boolean
  Dim Data As OPENFILENAME
  Dim ret As Boolean
  Dim str As String
  Dim filters As String

  ' prepare the API call
  str = propDefaultFile & String$(4096 - Len(propDefaultFile), Chr$(0))
  filters = GetFilters
  With Data
    .Flags = propOpenFlags
    .FlagsEx = propExOpenFlags
    .hInstance = App.hInstance
    .hwndOwner = prophWndParent
    .lpstrDefExt = StrPtr(propDefaultExtension)
    .lpstrFile = StrPtr(str)
    .lpstrFilter = StrPtr(filters)
    .lpstrInitialDir = StrPtr(propStartDirection)
    .lpstrTitle = StrPtr(propCaption)
    .nFilterIndex = propActiveFilter
    .nMaxFile = 4096
    .lStructSize = LenB(Data)

    ret = GetSaveFileName(Data)

    If ret Then
      ' return the selected file
      File = Left$(str, lstrlen(StrPtr(str)))
    End If
  End With

  ShowSave = ret
End Function

' fills <propCustomColors> with the first 16 system colors
Public Sub SysColorsToCustomColors()
  Dim i As Integer

  For i = LBound(propCustomColors) To UBound(propCustomColors)
    propCustomColors(i) = GetSysColor(i)
  Next i
End Sub
