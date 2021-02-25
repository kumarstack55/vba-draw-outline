Attribute VB_Name = "DrawOutlineModule"
'@Folder "VBAProjectDrawOutline"
'@IgnoreModule UseMeaningfulName, HungarianNotation
Option Explicit

Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Const SleepWaitMs As Long = 0

' 選択範囲を varValueTable に格納する
Public Function GetValueTableFromSelection( _
        ByRef cpValueTableOffset As CellPoint, _
        ByRef varValueTable() As Variant _
        ) As Boolean
    Dim app As Application
    Dim ws As Worksheet
    Dim lngRow As Long
    Dim lngCol As Long
    Dim rngLeftTop As Range
    Dim rngRightBottom As Range
    Dim r As Range
    
    Set app = Application
    Set ws = ActiveSheet
    
    ' 選択範囲が複数存在している場合は処理しない
    If Selection.Areas.Count <> 1 Then
        GetValueTableFromSelection = False
        Exit Function
    End If

    ' 選択されている範囲を得る
    '@Ignore IndexedUnboundDefaultMemberAccess
    Set rngLeftTop = app.Selection(1)
    Set cpValueTableOffset = New CellPoint
    cpValueTableOffset.Row = rngLeftTop.Row
    cpValueTableOffset.Column = rngLeftTop.Column
    '@Ignore IndexedUnboundDefaultMemberAccess
    Set rngRightBottom = app.Selection(app.Selection.Count)
    
    ' 各セルの値を得る
    ReDim varValueTable(rngRightBottom.Row - rngLeftTop.Row, rngRightBottom.Column - rngLeftTop.Column)
    For lngRow = rngLeftTop.Row To rngRightBottom.Row
        For lngCol = rngLeftTop.Column To rngRightBottom.Column
            '@Ignore IndexedDefaultMemberAccess
            Set r = ws.Cells(lngRow, lngCol)
            ' 値の有無を格納する
            varValueTable(lngRow - rngLeftTop.Row, lngCol - rngLeftTop.Column) = r.Value <> vbNullString
        Next
    Next
    
    GetValueTableFromSelection = True
End Function

' parent の子を返す。
Private Sub GetCellRectChildren( _
        ByVal cpValueTableOffset As CellPoint, _
        ByRef varValueTable() As Variant, _
        ByVal crParent As CellRect, _
        ByRef sdaChildren As SingleDimArray _
        )
    Dim lngColIndex As Long
    Dim lngRowIndex As Long
    Dim lngColEnd As Long
    Dim crChild As CellRect
    
    Set sdaChildren = CreateSingleDimArray
    lngColEnd = crParent.Right
    
    ' 左列 + 1 から、上行から探索していく。
    For lngRowIndex = crParent.Top To crParent.Bottom
        lngColIndex = crParent.Left + 1
        Do While lngColIndex <= lngColEnd
            If CBool(varValueTable( _
                    lngRowIndex - cpValueTableOffset.Row, _
                    lngColIndex - cpValueTableOffset.Column)) Then
                ' 見つかったら、より右列は探索しない。
                lngColEnd = lngColIndex
                
                If Not (crChild Is Nothing) Then
                    ' 子要素の下行は確定した。
                    crChild.Bottom = lngRowIndex - 1
                End If
                
                ' 子要素の下行以外は確定した。
                Set crChild = CreateCellRect(lngColIndex, lngRowIndex, crParent.Right, -1)
                sdaChildren.AddObject crChild
            End If
        
            lngColIndex = lngColIndex + 1
        Loop
    Next
    If Not (crChild Is Nothing) Then
        ' 子要素の下行は確定した。
        crChild.Bottom = crParent.Bottom
    End If
End Sub

Public Sub AddChildren( _
       ByVal cpValueTableOffset As CellPoint, _
       ByRef varValueTable() As Variant, _
       ByVal rtParent As RectTree _
       )
    ' parent の子要素を parent に追加する
    Dim crChildren As SingleDimArray
    Dim rtChild As RectTree
    Dim lngIndex As Long
    Dim lngLastIndex As Long
    
    ' 子の座標を得る
    GetCellRectChildren cpValueTableOffset, varValueTable, rtParent.GetRect, crChildren
    
    ' 子を加える
    lngLastIndex = crChildren.GetLastIndex
    For lngIndex = 0 To lngLastIndex
        Set rtChild = New RectTree
        rtChild.SetRect crChildren.GetObjectElement(lngIndex)
        AddChildren cpValueTableOffset, varValueTable, rtChild
        rtParent.AddChild rtChild
    Next
End Sub

' 選択範囲の外枠を描く
Private Sub DrawLineStyleForSelection()
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
End Sub

' rt について再帰的に罫線を描く
Public Sub DrawLineStyleForRectTree(ByVal rt As RectTree)
    Dim ws As Worksheet
    Dim child As RectTree
    Dim index As Long

    ' 範囲を選択して画面を描画する
    Set ws = ActiveSheet
    '@Ignore IndexedDefaultMemberAccess
    ws.Range(ws.Cells(rt.Top, rt.Left), ws.Cells(rt.Bottom, rt.Right)).Select
    DoEvents
    Sleep SleepWaitMs
    
    ' 選択範囲の罫線を描く
    DrawLineStyleForSelection
    
    ' 子要素を描く
    For index = 0 To UBound(rt.Children)
        Set child = rt.Children(index)
        DrawLineStyleForRectTree child
    Next
    
    ' 選択しなおす
    '@Ignore IndexedDefaultMemberAccess
    ws.Range(ws.Cells(rt.Top, rt.Left), ws.Cells(rt.Bottom, rt.Right)).Select
End Sub

' 選択範囲の罫線を消す
Private Sub EraseLineStyleForSelection()
    With Selection
        .Borders(xlDiagonalDown).LineStyle = xlNone
        .Borders(xlDiagonalUp).LineStyle = xlNone
        .Borders(xlEdgeLeft).LineStyle = xlNone
        .Borders(xlEdgeTop).LineStyle = xlNone
        .Borders(xlEdgeBottom).LineStyle = xlNone
        .Borders(xlEdgeRight).LineStyle = xlNone
        .Borders(xlInsideVertical).LineStyle = xlNone
        .Borders(xlInsideHorizontal).LineStyle = xlNone
    End With
End Sub

' 選択範囲を親とする親と子孫のアウトラインを描く
' 親と子の定義は次の通り：
' * 親 P に子は存在しないか、あるいは、１個以上の子が存在する。
' * P の子 C1, ..., Ci, ... , Cn (ただし1 <= i <= n)について：
'     * 任意の Ci は、 P と比べて範囲が、等しくなく、かつ、小さい。
'     * 任意の Ci と任意の Cj は、範囲の重複が存在しない。
'     * 任意の Ci は、範囲の最左列は、左上のみ値が存在し、それ以外は値が存在しない。
'     * 任意の Ci の範囲の最右列は、親の最右列と等しい。
'     * 任意の Ci の面積は上記の中で最大とする。
Private Sub DrawOutlineImpl()
    Dim apApp As Application
    Dim rtRoot As RectTree
    Dim rngLeftTop As Range
    Dim rngRightBottom As Range
    Dim cpValueTableOffset As CellPoint
    Dim varValueTable() As Variant
    
    Set apApp = Application
        
    ' 選択範囲を rtRoot に格納する
    Set rtRoot = New RectTree
    '@Ignore IndexedUnboundDefaultMemberAccess
    Set rngLeftTop = apApp.Selection(1)
    '@Ignore IndexedUnboundDefaultMemberAccess
    Set rngRightBottom = apApp.Selection(apApp.Selection.Count)
    rtRoot.Left = rngLeftTop.Column
    rtRoot.Top = rngLeftTop.Row
    rtRoot.Right = rngRightBottom.Column
    rtRoot.Bottom = rngRightBottom.Row
    
    ' 選択範囲を得る
    If Not GetValueTableFromSelection(cpValueTableOffset, varValueTable) Then
        Debug.Print "選択範囲の値を varValueTable に格納できなかった"
        Exit Sub
    End If
    
    ' rtRoot をもとに描く四角形のツリー構造を得る
    AddChildren cpValueTableOffset, varValueTable, rtRoot
    
    ' 選択範囲の罫線を消す
    EraseLineStyleForSelection
    
    ' ツリーの子要素をすべて描く
    DrawLineStyleForRectTree rtRoot
End Sub

' 選択範囲を、最も左のセルに左上に値をもち、それ以外に値を持たない、CellRectの配列を求める。
Private Function GetCellRectArray(ByVal apApp As Application, ByRef sdaCellRectArray As SingleDimArray) As Boolean
    Dim wsSheet As Worksheet
    Dim rngLeftTop As Range
    Dim rngRightBottom As Range
    Dim lngTop As Long
    Dim lngBottom As Long
    
    Set wsSheet = apApp.ActiveSheet
    
    ' 選択範囲が複数存在している場合は処理しない。
    If Selection.Areas.Count <> 1 Then
        GetCellRectArray = False
        Exit Function
    End If
    
    ' 選択範囲を得る。
    '@Ignore IndexedUnboundDefaultMemberAccess
    Set rngLeftTop = apApp.Selection(1)
    '@Ignore IndexedUnboundDefaultMemberAccess
    Set rngRightBottom = apApp.Selection(apApp.Selection.Count)
    
    ' 選択範囲から四角を探してリストに加える。
    Set sdaCellRectArray = CreateSingleDimArray
    lngTop = rngLeftTop.Row
    lngBottom = lngTop
    Do While lngBottom + 1 <= rngRightBottom.Row
        '@Ignore IndexedDefaultMemberAccess
        If wsSheet.Cells(lngBottom + 1, rngLeftTop.Column).Value <> vbNullString Then
            sdaCellRectArray.AddObject CreateCellRect(rngLeftTop.Column, lngTop, rngRightBottom.Column, lngBottom)
            lngTop = lngBottom + 1
        End If
        lngBottom = lngBottom + 1
    Loop
    sdaCellRectArray.AddObject CreateCellRect(rngLeftTop.Column, lngTop, rngRightBottom.Column, rngRightBottom.Row)
    
    GetCellRectArray = True
End Function

'@EntryPoint
'@ExcelHotkey O
Public Sub DrawOutline2()
Attribute DrawOutline2.VB_ProcData.VB_Invoke_Func = "O\n14"
    Dim apApp As Application
    Dim sdaCellRectArray As SingleDimArray
    Dim wsSheet As Worksheet
    Dim crRect As CellRect
    Dim lngIndex As Long
    Dim rngSelection As Range
    
    ' 複数範囲が指定されているなどの場合は、終了する。
    Set apApp = Application
    If Not GetCellRectArray(apApp, sdaCellRectArray) Then
        Exit Sub
    End If
    
    ' 選択範囲を得る。
    '@Ignore IndexedUnboundDefaultMemberAccess
    Set rngSelection = apApp.Selection(1)
    
    ' それぞれについてアウトラインを描く。
    Set wsSheet = apApp.ActiveSheet
    For lngIndex = 0 To sdaCellRectArray.GetLastIndex()
        ' 選択する
        Set crRect = sdaCellRectArray.GetObjectElement(lngIndex)
        '@Ignore IndexedDefaultMemberAccess
        wsSheet.Range(wsSheet.Cells(crRect.Top, crRect.Left), wsSheet.Cells(crRect.Bottom, crRect.Right)).Select
            
        ' アウトラインを描く。
        DrawOutlineImpl
    Next
    
    ' 選択しなおす。
    rngSelection.Select
End Sub
