Attribute VB_Name = "basCommands"
Option Explicit

  Public Const CMDID_FILE = 100
  Public Const CMDID_FILE_NEW = 101
  Public Const CMDID_FILE_OPEN = 102
  Public Const CMDID_FILE_SAVE = 103
  Public Const CMDID_FILE_PRINT = 104
  Public Const CMDID_FILE_RECENT = 105
  Public Const CMDID_FILE_RECENT_FIRST = 10501
  Public Const CMDID_FILE_EXIT = 106
  
  Public Const CMDID_EDIT = 200
  Public Const CMDID_EDIT_CUT = 201
  Public Const CMDID_EDIT_COPY = 202
  Public Const CMDID_EDIT_PASTE = 203
  Public Const CMDID_EDIT_SELECTALL = 204
  Public Const CMDID_EDIT_UNDO = 205
  Public Const CMDID_EDIT_REDO = 206
  Public Const CMDID_EDIT_SPELLCHECKING = 207
  Public Const CMDID_EDIT_FIND = 208
  
  Public Const CMDID_FONT = 300
  Public Const CMDID_FONT_FONTNAME = 301
  Public Const CMDID_FONT_FONTSIZE = 302
  Public Const CMDID_FONT_BOLD = 303
  Public Const CMDID_FONT_ITALIC = 304
  Public Const CMDID_FONT_UNDERLINE = 305
  Public Const CMDID_FONT_STRIKETHROUGH = 306
  Public Const CMDID_FONT_SUBSCRIPT = 307
  Public Const CMDID_FONT_SUPERSCRIPT = 308
  Public Const CMDID_FONT_TEXTBACKCOLOR = 309
  Public Const CMDID_FONT_TEXTFORECOLOR = 310
  
  Public Const CMDID_FORMAT = 400
  Public Const CMDID_FORMAT_ALIGNLEFT = 401
  Public Const CMDID_FORMAT_ALIGNCENTER = 402
  Public Const CMDID_FORMAT_ALIGNRIGHT = 403
  Public Const CMDID_FORMAT_ALIGNJUSTIFY = 404
  Public Const CMDID_FORMAT_BULLETLIST = 405
  Public Const CMDID_FORMAT_NUMBEREDLIST = 406
  Public Const CMDID_FORMAT_INCREASEINDENT = 407
  Public Const CMDID_FORMAT_DECREASEINDENT = 408
  Public Const CMDID_FORMAT_CONVERTTOHYPERLINK = 409
  
  Public Const CMDID_OBJECT = 500
  Public Const CMDID_OBJECT_INSERTIMAGE = 501
  Public Const CMDID_OBJECT_SETBKIMAGE = 502
  Public Const CMDID_OBJECT_INSERTOBJECT = 503
  Public Const CMDID_OBJECT_INSERTOBJECTFROMFILE = 504
  
  Public Const CMDID_TABLE = 600
  Public Const CMDID_TABLE_INSERTTABLE = 601
  Public Const CMDID_TABLE_DELETETABLE = 602
  Public Const CMDID_TABLE_INSERTROW = 603
  Public Const CMDID_TABLE_INSERTCOLUMN = 604
  Public Const CMDID_TABLE_DELETEROW = 605
  Public Const CMDID_TABLE_DELETECOLUMN = 606
  Public Const CMDID_TABLE_MERGECELLS = 607
  Public Const CMDID_TABLE_ALIGNTOP = 608
  Public Const CMDID_TABLE_ALIGNCENTER = 609
  Public Const CMDID_TABLE_ALIGNBOTTOM = 610
  Public Const CMDID_TABLE_CELLBACKCOLOR = 611
  
  Public Const CMDID_MATH = 700
  Public Const CMDID_MATH_TOGGLEMATHZONE = 701
  Public Const CMDID_MATH_BUILDUP = 702
  Public Const CMDID_MATH_BUILDDOWN = 703
  Public Const CMDID_MATH_INSERTROOTSIGN = 704
  Public Const CMDID_MATH_INSERTSUMSIGN = 705
  Public Const CMDID_MATH_INSERTPRODUCTSIGN = 706
  Public Const CMDID_MATH_INSERTINTSIGN = 707
  Public Const CMDID_MATH_INSERTLIMES = 708
  Public Const CMDID_MATH_INSERTMATRIX = 709
  
  Public Const CMDID_HELP = 900
  Public Const CMDID_HELP_ABOUT = 901
  
  Public Const CMDID_WINDOW = 1000
  Public Const CMDID_WINDOW_CASCADE = 1001
  Public Const CMDID_WINDOW_TILE = 1002
  Public Const CMDID_WINDOW_ARRANGEICONS = 1003
  Public Const CMDID_WINDOW_LISTSTART = 1050