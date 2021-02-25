VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CellPoint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "VBAProjectDrawOutline.Libraries"
'@IgnoreModule HungarianNotation
Option Explicit

Private mRow As Long
Private mColumn As Long

Public Property Get Row() As Long
    Row = mRow
End Property

Public Property Let Row(ByVal lngRow As Long)
    mRow = lngRow
End Property

Public Property Get Column() As Long
    Column = mColumn
End Property

Public Property Let Column(ByVal lngColumn As Long)
    mColumn = lngColumn
End Property

