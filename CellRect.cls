VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CellRect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "VBAProjectDrawOutline.Libraries"
'@IgnoreModule HungarianNotation
Option Explicit

'@Ignore EncapsulatePublicField
Public Left As Long

'@Ignore EncapsulatePublicField
Public Top As Long

'@Ignore EncapsulatePublicField
Public Right As Long

'@Ignore EncapsulatePublicField
Public Bottom As Long

'@Ignore ProcedureNotUsed
Public Function EqualsTo(ByVal crRect As CellRect) As Boolean
    EqualsTo = (Left = crRect.Left And Top = crRect.Top And Right = crRect.Right And Bottom = crRect.Bottom)
End Function

