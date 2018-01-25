VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IMDIFrame"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Sub OpenedDocument(ByVal frm As Form, ByRef windowID As Long)
  '
End Sub

Public Sub ClosingDocument(ByVal windowID As Long)
  '
End Sub

Public Sub ActivatedDocument(ByVal doc As IDocument)
  '
End Sub

Public Sub CheckCommand(ByVal commandID As Long, ByVal check As Boolean)
  '
End Sub

Public Sub CheckTriStateCommand(ByVal commandID As Long, ByVal selState As TBarCtlsLibUCtl.SelectionStateConstants)
  '
End Sub

Public Sub EnableDisableCommand(ByVal commandID As Long, ByVal enable As Boolean)
  '
End Sub

Public Sub UpdateFontName(ByVal fontName As String)
  '
End Sub

Public Sub UpdateFontSize(ByVal fontSize As Single)
  '
End Sub
