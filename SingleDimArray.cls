VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SingleDimArray"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "VBAProjectDrawOutline.Libraries"
'@IgnoreModule HungarianNotation
Option Explicit

Private mArray As Variant
Private mInitialized As Boolean

Public Sub Initialize()
    mArray = Null
    mInitialized = False
End Sub

Private Sub ExtendArray()
    If Not mInitialized Then
        ReDim mArray(0)
        mInitialized = True
    Else
        ReDim Preserve mArray(UBound(mArray) + 1)
    End If
End Sub

Public Sub Add(ByVal varObject As Variant)
    ExtendArray
    mArray(UBound(mArray)) = varObject
End Sub

Public Sub AddObject(ByVal varObject As Variant)
    ExtendArray
    Set mArray(UBound(mArray)) = varObject
End Sub

Public Function GetElement(ByVal lngIndex As Long) As Variant
    GetElement = mArray(lngIndex)
End Function

Public Function GetObjectElement(ByVal lngIndex As Long) As Variant
    Set GetObjectElement = mArray(lngIndex)
End Function

Public Function GetLength() As Long
    If mInitialized Then
        GetLength = UBound(mArray) - LBound(mArray) + 1
    Else
        GetLength = 0
    End If
End Function

Public Function GetLastIndex() As Long
    If mInitialized Then
        GetLastIndex = UBound(mArray)
    Else
        GetLastIndex = -1
    End If
End Function

