VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IDocument"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Function OpenFile(ByVal File As String) As Boolean
  '
End Function

Public Function SaveFile(ByVal File As String) As Boolean
  '
End Function

Public Property Get FileName() As String
  '
End Property

Public Property Get SelectionCellBackColor() As Long
  '
End Property

Public Property Get CellBackColor() As Long
  '
End Property

Public Property Let CellBackColor(ByVal newValue As Long)
  '
End Property

Public Property Get SelectionTextBackColor() As Long
  '
End Property

Public Property Get TextBackColor() As Long
  '
End Property

Public Property Let TextBackColor(ByVal newValue As Long)
  '
End Property

Public Property Get SelectionTextForeColor() As Long
  '
End Property

Public Property Get TextForeColor() As Long
  '
End Property

Public Property Let TextForeColor(ByVal newValue As Long)
  '
End Property

Public Property Get DocumentTextRange() As TextRange
  '
End Property

Public Property Get SelectedTextRange() As TextRange
  '
End Property

Public Function GetCommandToolTip(ByVal commandID As Long) As String
  '
End Function

Public Sub ExecuteCommand(ByVal commandID As Long, Optional ByVal params As Variant)
  '
End Sub

Public Sub SetFocus()
  '
End Sub

Public Sub SetFontName(ByVal fontName As String)
  '
End Sub

Public Sub SetFontSize(ByVal fontSize As Single)
  '
End Sub

Public Sub StartFontChangePreview()
  '
End Sub

Public Sub EndFontChangePreview(ByVal apply As Boolean)
  '
End Sub
