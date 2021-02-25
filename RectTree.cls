VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "RectTree"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "VBAProjectDrawOutline.Libraries"
'@IgnoreModule HungarianNotation
Option Explicit

Private mRect As CellRect
Private mChildren() As RectTree
Private mChildrenInitialized As Boolean

Private Sub Class_Initialize()
    Set mRect = New CellRect
    mChildrenInitialized = False
End Sub

Public Sub AddChild(ByVal rtChild As RectTree)
    Dim lngLastIndex As Long
    
    If mChildrenInitialized = False Then
        ReDim Preserve mChildren(0)
        mChildrenInitialized = True
    Else
        ReDim Preserve mChildren(UBound(mChildren) + 1)
    End If
    
    lngLastIndex = UBound(mChildren)
    Set mChildren(lngLastIndex) = rtChild
End Sub

Public Property Get Left() As Long
    Left = mRect.Left
End Property

Public Property Let Left(ByVal lngLeft As Long)
    mRect.Left = lngLeft
End Property

Public Property Get Top() As Long
    Top = mRect.Top
End Property

Public Property Let Top(ByVal longTop As Long)
    mRect.Top = longTop
End Property

Public Property Get Right() As Long
    Right = mRect.Right
End Property

Public Property Let Right(ByVal lngRight As Long)
    mRect.Right = lngRight
End Property

Public Property Get Bottom() As Long
    Bottom = mRect.Bottom
End Property

Public Property Let Bottom(ByVal lngBottom As Long)
    mRect.Bottom = lngBottom
End Property

Public Property Get HasChildren() As Boolean
    HasChildren = mChildrenInitialized
End Property

Public Property Get Children() As Variant
    Children = mChildren
End Property

Public Sub SetRect(ByVal crRect As CellRect)
    mRect.Left = crRect.Left
    mRect.Top = crRect.Top
    mRect.Right = crRect.Right
    mRect.Bottom = crRect.Bottom
End Sub

Public Function GetRect() As CellRect
    Set GetRect = New CellRect
    GetRect.Left = mRect.Left
    GetRect.Top = mRect.Top
    GetRect.Right = mRect.Right
    GetRect.Bottom = mRect.Bottom
End Function

