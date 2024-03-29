VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IModel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Enum x2lOptions
    x2lNone = 0
    
    x2lCreateTableEnvironment = 1
    x2lBooktabs = 2
    x2lConvertMathChars = 4
    x2lAddNumCommand = 8
    x2lSkipOne = 16
    x2lCreateTabular = 32
End Enum

Public Property Get Events() As IModelEvents
'
End Property

Public Property Get Options() As x2lOptions
'
End Property
Public Property Let Options(ByVal Options As x2lOptions)
'
End Property

Public Property Get Indent() As Integer
'
End Property
Public Property Let Indent(ByVal Indent As Integer)
'
End Property

Public Property Get CellWidth() As Integer
'
End Property
Public Property Let CellWidth(ByVal CellWidth As Integer)
'
End Property

Public Property Get NumCommand() As String
'
End Property
Public Property Let NumCommand(ByVal CellWidth As String)
'
End Property

Public Property Get NumMid() As String
'
End Property
Public Property Let NumMid(ByVal CellWidth As String)
'
End Property

Public Property Get NumSuffix() As String
'
End Property
Public Property Let NumSuffix(ByVal CellWidth As String)
'
End Property

Public Property Get FileName() As String
'
End Property
Public Property Let FileName(ByVal FileName As String)
'
End Property
Public Property Get AbsoluteFileName() As String
'
End Property

Public Property Get RangeAddress() As String
'
End Property
Public Property Let RangeAddress(ByVal iRange As String)
'
End Property

Public Function GetConversionResult() As String
'
End Function

Public Sub InitDefault()
'
End Sub

Public Property Get Description() As String
'
End Property

Public Property Get Range() As Range
'
End Property
Public Property Set Range(ByVal pRange As Range)
'
End Property

