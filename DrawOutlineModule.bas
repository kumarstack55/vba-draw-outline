Attribute VB_Name = "DrawOutlineModule"
'@Folder "VBAProjectDrawOutline"
'@IgnoreModule UseMeaningfulName, HungarianNotation
Option Explicit

Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Const SleepWaitMs As Long = 0

' �I��͈͂� varValueTable �Ɋi�[����
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
    
    ' �I��͈͂��������݂��Ă���ꍇ�͏������Ȃ�
    If Selection.Areas.Count <> 1 Then
        GetValueTableFromSelection = False
        Exit Function
    End If

    ' �I������Ă���͈͂𓾂�
    '@Ignore IndexedUnboundDefaultMemberAccess
    Set rngLeftTop = app.Selection(1)
    Set cpValueTableOffset = New CellPoint
    cpValueTableOffset.Row = rngLeftTop.Row
    cpValueTableOffset.Column = rngLeftTop.Column
    '@Ignore IndexedUnboundDefaultMemberAccess
    Set rngRightBottom = app.Selection(app.Selection.Count)
    
    ' �e�Z���̒l�𓾂�
    ReDim varValueTable(rngRightBottom.Row - rngLeftTop.Row, rngRightBottom.Column - rngLeftTop.Column)
    For lngRow = rngLeftTop.Row To rngRightBottom.Row
        For lngCol = rngLeftTop.Column To rngRightBottom.Column
            '@Ignore IndexedDefaultMemberAccess
            Set r = ws.Cells(lngRow, lngCol)
            ' �l�̗L�����i�[����
            varValueTable(lngRow - rngLeftTop.Row, lngCol - rngLeftTop.Column) = r.Value <> vbNullString
        Next
    Next
    
    GetValueTableFromSelection = True
End Function

' parent �̎q��Ԃ��B
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
    
    ' ���� + 1 ����A��s����T�����Ă����B
    For lngRowIndex = crParent.Top To crParent.Bottom
        lngColIndex = crParent.Left + 1
        Do While lngColIndex <= lngColEnd
            If CBool(varValueTable( _
                    lngRowIndex - cpValueTableOffset.Row, _
                    lngColIndex - cpValueTableOffset.Column)) Then
                ' ����������A���E��͒T�����Ȃ��B
                lngColEnd = lngColIndex
                
                If Not (crChild Is Nothing) Then
                    ' �q�v�f�̉��s�͊m�肵���B
                    crChild.Bottom = lngRowIndex - 1
                End If
                
                ' �q�v�f�̉��s�ȊO�͊m�肵���B
                Set crChild = CreateCellRect(lngColIndex, lngRowIndex, crParent.Right, -1)
                sdaChildren.AddObject crChild
            End If
        
            lngColIndex = lngColIndex + 1
        Loop
    Next
    If Not (crChild Is Nothing) Then
        ' �q�v�f�̉��s�͊m�肵���B
        crChild.Bottom = crParent.Bottom
    End If
End Sub

Public Sub AddChildren( _
       ByVal cpValueTableOffset As CellPoint, _
       ByRef varValueTable() As Variant, _
       ByVal rtParent As RectTree _
       )
    ' parent �̎q�v�f�� parent �ɒǉ�����
    Dim crChildren As SingleDimArray
    Dim rtChild As RectTree
    Dim lngIndex As Long
    Dim lngLastIndex As Long
    
    ' �q�̍��W�𓾂�
    GetCellRectChildren cpValueTableOffset, varValueTable, rtParent.GetRect, crChildren
    
    ' �q��������
    lngLastIndex = crChildren.GetLastIndex
    For lngIndex = 0 To lngLastIndex
        Set rtChild = New RectTree
        rtChild.SetRect crChildren.GetObjectElement(lngIndex)
        AddChildren cpValueTableOffset, varValueTable, rtChild
        rtParent.AddChild rtChild
    Next
End Sub

' �I��͈͂̊O�g��`��
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

' rt �ɂ��čċA�I�Ɍr����`��
Public Sub DrawLineStyleForRectTree(ByVal rt As RectTree)
    Dim ws As Worksheet
    Dim child As RectTree
    Dim index As Long

    ' �͈͂�I�����ĉ�ʂ�`�悷��
    Set ws = ActiveSheet
    '@Ignore IndexedDefaultMemberAccess
    ws.Range(ws.Cells(rt.Top, rt.Left), ws.Cells(rt.Bottom, rt.Right)).Select
    DoEvents
    Sleep SleepWaitMs
    
    ' �I��͈͂̌r����`��
    DrawLineStyleForSelection
    
    ' �q�v�f��`��
    For index = 0 To UBound(rt.Children)
        Set child = rt.Children(index)
        DrawLineStyleForRectTree child
    Next
    
    ' �I�����Ȃ���
    '@Ignore IndexedDefaultMemberAccess
    ws.Range(ws.Cells(rt.Top, rt.Left), ws.Cells(rt.Bottom, rt.Right)).Select
End Sub

' �I��͈͂̌r��������
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

' �I��͈͂�e�Ƃ���e�Ǝq���̃A�E�g���C����`��
' �e�Ǝq�̒�`�͎��̒ʂ�F
' * �e P �Ɏq�͑��݂��Ȃ����A���邢�́A�P�ȏ�̎q�����݂���B
' * P �̎q C1, ..., Ci, ... , Cn (������1 <= i <= n)�ɂ��āF
'     * �C�ӂ� Ci �́A P �Ɣ�ׂĔ͈͂��A�������Ȃ��A���A�������B
'     * �C�ӂ� Ci �ƔC�ӂ� Cj �́A�͈͂̏d�������݂��Ȃ��B
'     * �C�ӂ� Ci �́A�͈͂̍ō���́A����̂ݒl�����݂��A����ȊO�͒l�����݂��Ȃ��B
'     * �C�ӂ� Ci �͈̔͂̍ŉE��́A�e�̍ŉE��Ɠ������B
'     * �C�ӂ� Ci �̖ʐς͏�L�̒��ōő�Ƃ���B
Private Sub DrawOutlineImpl()
    Dim apApp As Application
    Dim rtRoot As RectTree
    Dim rngLeftTop As Range
    Dim rngRightBottom As Range
    Dim cpValueTableOffset As CellPoint
    Dim varValueTable() As Variant
    
    Set apApp = Application
        
    ' �I��͈͂� rtRoot �Ɋi�[����
    Set rtRoot = New RectTree
    '@Ignore IndexedUnboundDefaultMemberAccess
    Set rngLeftTop = apApp.Selection(1)
    '@Ignore IndexedUnboundDefaultMemberAccess
    Set rngRightBottom = apApp.Selection(apApp.Selection.Count)
    rtRoot.Left = rngLeftTop.Column
    rtRoot.Top = rngLeftTop.Row
    rtRoot.Right = rngRightBottom.Column
    rtRoot.Bottom = rngRightBottom.Row
    
    ' �I��͈͂𓾂�
    If Not GetValueTableFromSelection(cpValueTableOffset, varValueTable) Then
        Debug.Print "�I��͈͂̒l�� varValueTable �Ɋi�[�ł��Ȃ�����"
        Exit Sub
    End If
    
    ' rtRoot �����Ƃɕ`���l�p�`�̃c���[�\���𓾂�
    AddChildren cpValueTableOffset, varValueTable, rtRoot
    
    ' �I��͈͂̌r��������
    EraseLineStyleForSelection
    
    ' �c���[�̎q�v�f�����ׂĕ`��
    DrawLineStyleForRectTree rtRoot
End Sub

' �I��͈͂��A�ł����̃Z���ɍ���ɒl�������A����ȊO�ɒl�������Ȃ��ACellRect�̔z������߂�B
Private Function GetCellRectArray(ByVal apApp As Application, ByRef sdaCellRectArray As SingleDimArray) As Boolean
    Dim wsSheet As Worksheet
    Dim rngLeftTop As Range
    Dim rngRightBottom As Range
    Dim lngTop As Long
    Dim lngBottom As Long
    
    Set wsSheet = apApp.ActiveSheet
    
    ' �I��͈͂��������݂��Ă���ꍇ�͏������Ȃ��B
    If Selection.Areas.Count <> 1 Then
        GetCellRectArray = False
        Exit Function
    End If
    
    ' �I��͈͂𓾂�B
    '@Ignore IndexedUnboundDefaultMemberAccess
    Set rngLeftTop = apApp.Selection(1)
    '@Ignore IndexedUnboundDefaultMemberAccess
    Set rngRightBottom = apApp.Selection(apApp.Selection.Count)
    
    ' �I��͈͂���l�p��T���ă��X�g�ɉ�����B
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
    
    ' �����͈͂��w�肳��Ă���Ȃǂ̏ꍇ�́A�I������B
    Set apApp = Application
    If Not GetCellRectArray(apApp, sdaCellRectArray) Then
        Exit Sub
    End If
    
    ' �I��͈͂𓾂�B
    '@Ignore IndexedUnboundDefaultMemberAccess
    Set rngSelection = apApp.Selection(1)
    
    ' ���ꂼ��ɂ��ăA�E�g���C����`���B
    Set wsSheet = apApp.ActiveSheet
    For lngIndex = 0 To sdaCellRectArray.GetLastIndex()
        ' �I������
        Set crRect = sdaCellRectArray.GetObjectElement(lngIndex)
        '@Ignore IndexedDefaultMemberAccess
        wsSheet.Range(wsSheet.Cells(crRect.Top, crRect.Left), wsSheet.Cells(crRect.Bottom, crRect.Right)).Select
            
        ' �A�E�g���C����`���B
        DrawOutlineImpl
    Next
    
    ' �I�����Ȃ����B
    rngSelection.Select
End Sub
