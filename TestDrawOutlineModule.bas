Attribute VB_Name = "TestDrawOutlineModule"
'@IgnoreModule LineLabelNotUsed, UseMeaningfulName, HungarianNotation
Option Explicit
Option Private Module

'@TestModule
'@Folder "VBAProjectDrawOutline.Tests"

Private Assert As Object
'@Ignore VariableNotUsed
Private Fakes As Object

'@ModuleInitialize
Private Sub ModuleInitialize()
    'this method runs once per module.
    Set Assert = CreateObject("Rubberduck.AssertClass")
    Set Fakes = CreateObject("Rubberduck.FakesProvider")
End Sub

'@ModuleCleanup
Private Sub ModuleCleanup()
    'this method runs once per module.
    Set Assert = Nothing
    Set Fakes = Nothing
End Sub

'@TestInitialize
'@Ignore EmptyMethod
Private Sub TestInitialize()
    'This method runs before every test in the module..
End Sub

'@TestCleanup
'@Ignore EmptyMethod
Private Sub TestCleanup()
    'this method runs after every test in the module.
End Sub

' 文字列から varValueTable を作る
'@Ignore ProcedureCanBeWrittenAsFunction
Private Sub GetValueTableFromString( _
        ByVal strText As String, _
        ByRef varValueTable() As Variant _
        )
    Dim strText2 As String
    Dim strLines As Variant
    Dim strLine As String
    Dim strLine2 As String
    Dim strValues As Variant
    Dim strValue As String
    Dim i As Long
    Dim j As Long
    Dim bInitialized As Boolean
    
    bInitialized = False
    
    strText2 = strText
    If Right$(strText2, 1) = vbLf Then
        strText2 = Left$(strText2, Len(strText2) - 1)
    End If
    
    strLines = Split(strText2, vbLf)
    For i = LBound(strLines) To UBound(strLines)
        strLine = Trim$(strLines(i))
        
        strLine2 = strLine
        If Left$(strLine2, 1) = "|" Then
            strLine2 = Right$(strLine2, Len(strLine2) - 1)
        End If
        If Right$(strLine2, 1) = "|" Then
            strLine2 = Left$(strLine2, Len(strLine2) - 1)
        End If
        
        strValues = Split(strLine2, "|")
        For j = LBound(strValues) To UBound(strValues)
            If Not bInitialized Then
                ReDim varValueTable(UBound(strLines), UBound(strValues))
                bInitialized = True
            End If
            strValue = Trim$(strValues(j))
            varValueTable(i, j) = strValue <> vbNullString
        Next
    Next
End Sub

'@TestMethod("Uncategorized")
Private Sub TestAddChildren_子が親と同じなら子とみなさない()
    Dim cpValueTableOffset As CellPoint
    Dim rtRoot As RectTree
    Dim varValueTable() As Variant
    Dim strText As String
    On Error GoTo TestFail
    
    'Arrange:
    strText = vbNullString
    strText = strText + "|x| |" + vbLf
    strText = strText + "| | |" + vbLf
    GetValueTableFromString strText, varValueTable

    Set cpValueTableOffset = New CellPoint
    cpValueTableOffset.Row = 1000
    cpValueTableOffset.Column = 100
    
    Set rtRoot = New RectTree
    rtRoot.Left = cpValueTableOffset.Column
    rtRoot.Top = cpValueTableOffset.Row
    rtRoot.Right = rtRoot.Left + UBound(varValueTable, 2)
    rtRoot.Bottom = rtRoot.Top + UBound(varValueTable)
    
    'Act:
    AddChildren cpValueTableOffset, varValueTable, rtRoot
    
    'Assert:
    Assert.AreEqual rtRoot.HasChildren, False
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Uncategorized")
Private Sub TestAddChildren_子が親と異なるなら子を追加する()
    Dim cpValueTableOffset As CellPoint
    Dim rtRoot As RectTree
    Dim varValueTable() As Variant
    Dim strText As String
    On Error GoTo TestFail
    
    'Arrange:
    strText = vbNullString
    strText = strText + "| | | |" + vbLf
    strText = strText + "| |x| |" + vbLf
    strText = strText + "| | | |" + vbLf
    GetValueTableFromString strText, varValueTable

    Set cpValueTableOffset = New CellPoint
    cpValueTableOffset.Row = 1000
    cpValueTableOffset.Column = 100
    
    Set rtRoot = New RectTree
    rtRoot.Left = cpValueTableOffset.Column
    rtRoot.Top = cpValueTableOffset.Row
    rtRoot.Right = rtRoot.Left + UBound(varValueTable, 2)
    rtRoot.Bottom = rtRoot.Top + UBound(varValueTable)
    
    'Act:
    AddChildren cpValueTableOffset, varValueTable, rtRoot
    
    'Assert:
    Assert.AreEqual rtRoot.HasChildren, True
    
    Assert.AreEqual CLng(0), UBound(rtRoot.Children)
    
    Assert.AreEqual CLng(1001), rtRoot.Children(0).Top
    Assert.AreEqual CLng(1002), rtRoot.Children(0).Bottom
    Assert.AreEqual CLng(101), rtRoot.Children(0).Left
    Assert.AreEqual CLng(102), rtRoot.Children(0).Right
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Uncategorized")
Private Sub TestAddChildren_最左列に子の兄弟あるときその兄弟を追加する()
    Dim cpValueTableOffset As CellPoint
    Dim rtRoot As RectTree
    Dim varValueTable() As Variant
    Dim strText As String
    On Error GoTo TestFail
    
    'Arrange:
    strText = vbNullString
    strText = strText + "| |x| |" + vbLf
    strText = strText + "| | | |" + vbLf
    strText = strText + "| |x| |" + vbLf
    strText = strText + "| | | |" + vbLf
    GetValueTableFromString strText, varValueTable

    Set cpValueTableOffset = New CellPoint
    cpValueTableOffset.Row = 1000
    cpValueTableOffset.Column = 100
    
    Set rtRoot = New RectTree
    rtRoot.Left = cpValueTableOffset.Column
    rtRoot.Top = cpValueTableOffset.Row
    rtRoot.Right = rtRoot.Left + UBound(varValueTable, 2)
    rtRoot.Bottom = rtRoot.Top + UBound(varValueTable)
    
    'Act:
    AddChildren cpValueTableOffset, varValueTable, rtRoot
    
    'Assert:
    Assert.AreEqual rtRoot.HasChildren, True
    
    Assert.AreEqual CLng(1), UBound(rtRoot.Children)
    
    Assert.AreEqual CLng(1000), rtRoot.Children(0).Top
    Assert.AreEqual CLng(1001), rtRoot.Children(0).Bottom
    Assert.AreEqual CLng(101), rtRoot.Children(0).Left
    Assert.AreEqual CLng(102), rtRoot.Children(0).Right
    
    Assert.AreEqual CLng(1002), rtRoot.Children(1).Top
    Assert.AreEqual CLng(1003), rtRoot.Children(1).Bottom
    Assert.AreEqual CLng(101), rtRoot.Children(1).Left
    Assert.AreEqual CLng(102), rtRoot.Children(1).Right

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Uncategorized")
Private Sub TestAddChildren_最左列に子がなければより右で子を探す()
    Dim cpValueTableOffset As CellPoint
    Dim rtRoot As RectTree
    Dim varValueTable() As Variant
    Dim strText As String
    On Error GoTo TestFail
    
    'Arrange:
    strText = vbNullString
    strText = strText + "| |x| |" + vbLf
    strText = strText + "| | | |" + vbLf
    GetValueTableFromString strText, varValueTable

    Set cpValueTableOffset = New CellPoint
    cpValueTableOffset.Row = 1000
    cpValueTableOffset.Column = 100
    
    Set rtRoot = New RectTree
    rtRoot.Left = cpValueTableOffset.Column
    rtRoot.Top = cpValueTableOffset.Row
    rtRoot.Right = rtRoot.Left + UBound(varValueTable, 2)
    rtRoot.Bottom = rtRoot.Top + UBound(varValueTable)
    
    'Act:
    AddChildren cpValueTableOffset, varValueTable, rtRoot
    
    'Assert:
    Assert.AreEqual rtRoot.HasChildren, True
    
    Assert.AreEqual CLng(0), UBound(rtRoot.Children)
    
    Assert.AreEqual CLng(1000), rtRoot.Children(0).Top
    Assert.AreEqual CLng(1001), rtRoot.Children(0).Bottom
    Assert.AreEqual CLng(101), rtRoot.Children(0).Left
    Assert.AreEqual CLng(102), rtRoot.Children(0).Right
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Uncategorized")
Private Sub TestAddChildren_最左列に子がなければより右で子を探し兄弟を加える()
    Dim cpValueTableOffset As CellPoint
    Dim rtRoot As RectTree
    Dim varValueTable() As Variant
    Dim strText As String
    On Error GoTo TestFail
    
    'Arrange:
    strText = vbNullString
    strText = strText + "| |x| |" + vbLf
    strText = strText + "| | | |" + vbLf
    strText = strText + "| |x| |" + vbLf
    strText = strText + "| | | |" + vbLf
    GetValueTableFromString strText, varValueTable

    Set cpValueTableOffset = New CellPoint
    cpValueTableOffset.Row = 1000
    cpValueTableOffset.Column = 100
    
    Set rtRoot = New RectTree
    rtRoot.Left = cpValueTableOffset.Column
    rtRoot.Top = cpValueTableOffset.Row
    rtRoot.Right = rtRoot.Left + UBound(varValueTable, 2)
    rtRoot.Bottom = rtRoot.Top + UBound(varValueTable)
    
    'Act:
    AddChildren cpValueTableOffset, varValueTable, rtRoot
    
    'Assert:
    Assert.AreEqual rtRoot.HasChildren, True
    
    Assert.AreEqual CLng(1), UBound(rtRoot.Children)
    
    Assert.AreEqual CLng(1000), rtRoot.Children(0).Top
    Assert.AreEqual CLng(1001), rtRoot.Children(0).Bottom
    Assert.AreEqual CLng(101), rtRoot.Children(0).Left
    Assert.AreEqual CLng(102), rtRoot.Children(0).Right
    
    Assert.AreEqual CLng(1002), rtRoot.Children(1).Top
    Assert.AreEqual CLng(1003), rtRoot.Children(1).Bottom
    Assert.AreEqual CLng(101), rtRoot.Children(1).Left
    Assert.AreEqual CLng(102), rtRoot.Children(1).Right
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Uncategorized")
Private Sub TestAddChildren_最左列を子としないならより右で子を探す()
    Dim cpValueTableOffset As CellPoint
    Dim rtRoot As RectTree
    Dim varValueTable() As Variant
    Dim strText As String
    On Error GoTo TestFail
    
    'Arrange:
    strText = vbNullString
    strText = strText + "|x|x| |" + vbLf
    strText = strText + "| | | |" + vbLf
    GetValueTableFromString strText, varValueTable

    Set cpValueTableOffset = New CellPoint
    cpValueTableOffset.Row = 1000
    cpValueTableOffset.Column = 100
    
    Set rtRoot = New RectTree
    rtRoot.Left = cpValueTableOffset.Column
    rtRoot.Top = cpValueTableOffset.Row
    rtRoot.Right = rtRoot.Left + UBound(varValueTable, 2)
    rtRoot.Bottom = rtRoot.Top + UBound(varValueTable)
    
    'Act:
    AddChildren cpValueTableOffset, varValueTable, rtRoot
    
    'Assert:
    Assert.AreEqual rtRoot.HasChildren, True
    
    Assert.AreEqual CLng(0), UBound(rtRoot.Children)
    
    Assert.AreEqual CLng(1000), rtRoot.Children(0).Top
    Assert.AreEqual CLng(1001), rtRoot.Children(0).Bottom
    Assert.AreEqual CLng(101), rtRoot.Children(0).Left
    Assert.AreEqual CLng(102), rtRoot.Children(0).Right
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Uncategorized")
Private Sub TestAddChildren_子を再帰的に追加する()
    Dim cpValueTableOffset As CellPoint
    Dim rtRoot As RectTree
    Dim varValueTable() As Variant
    Dim strText As String
    Dim rtChild As RectTree
    Dim rtChild2 As RectTree
    Dim rtChild3 As RectTree
    On Error GoTo TestFail
    
    'Arrange:
    strText = vbNullString
    strText = strText + "|x| | |" + vbLf
    strText = strText + "| |x| |" + vbLf
    strText = strText + "| | |x|" + vbLf
    strText = strText + "| | |x|" + vbLf
    GetValueTableFromString strText, varValueTable

    Set cpValueTableOffset = New CellPoint
    cpValueTableOffset.Row = 1000
    cpValueTableOffset.Column = 100
    
    Set rtRoot = New RectTree
    rtRoot.Left = cpValueTableOffset.Column
    rtRoot.Top = cpValueTableOffset.Row
    rtRoot.Right = rtRoot.Left + UBound(varValueTable, 2)
    rtRoot.Bottom = rtRoot.Top + UBound(varValueTable)
    
    'Act:
    AddChildren cpValueTableOffset, varValueTable, rtRoot
    
    'Assert:
    Assert.AreEqual True, rtRoot.HasChildren
    
    Assert.AreEqual CLng(0), UBound(rtRoot.Children)
    
    Assert.AreEqual CLng(1001), rtRoot.Children(0).Top
    Assert.AreEqual CLng(1003), rtRoot.Children(0).Bottom
    Assert.AreEqual CLng(101), rtRoot.Children(0).Left
    Assert.AreEqual CLng(102), rtRoot.Children(0).Right
    
    Assert.AreEqual True, rtRoot.Children(0).HasChildren
    
    Assert.AreEqual CLng(1), UBound(rtRoot.Children(0).Children)
    
    Set rtChild = rtRoot.Children(0)
    
    Set rtChild2 = rtChild.Children(0)
    Assert.AreEqual CLng(1002), rtChild2.Top
    Assert.AreEqual CLng(1002), rtChild2.Bottom
    Assert.AreEqual CLng(102), rtChild2.Left
    Assert.AreEqual CLng(102), rtChild2.Right
    
    Set rtChild3 = rtChild.Children(1)
    Assert.AreEqual CLng(1003), rtChild3.Top
    Assert.AreEqual CLng(1003), rtChild3.Bottom
    Assert.AreEqual CLng(102), rtChild3.Left
    Assert.AreEqual CLng(102), rtChild3.Right
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Uncategorized")
Private Sub TestAddChildren_未探索の行があればその範囲から子を探す()
    Dim cpValueTableOffset As CellPoint
    Dim rtRoot As RectTree
    Dim varValueTable() As Variant
    Dim strText As String
    On Error GoTo TestFail
    
    'Arrange:
    strText = vbNullString
    strText = strText + "| | | | |" + vbLf
    strText = strText + "| | |x| |" + vbLf
    strText = strText + "| | | | |" + vbLf
    strText = strText + "| |x| | |" + vbLf
    strText = strText + "| | | | |" + vbLf
    GetValueTableFromString strText, varValueTable

    Set cpValueTableOffset = New CellPoint
    cpValueTableOffset.Row = 1000
    cpValueTableOffset.Column = 100
    
    Set rtRoot = New RectTree
    rtRoot.Left = cpValueTableOffset.Column
    rtRoot.Top = cpValueTableOffset.Row
    rtRoot.Right = rtRoot.Left + UBound(varValueTable, 2)
    rtRoot.Bottom = rtRoot.Top + UBound(varValueTable)
    
    'Act:
    AddChildren cpValueTableOffset, varValueTable, rtRoot
    
    'Assert:
    Assert.AreEqual True, rtRoot.HasChildren
    
    Assert.AreEqual CLng(1), UBound(rtRoot.Children)
    
    Assert.AreEqual CLng(1001), rtRoot.Children(0).Top
    Assert.AreEqual CLng(1002), rtRoot.Children(0).Bottom
    Assert.AreEqual CLng(102), rtRoot.Children(0).Left
    Assert.AreEqual CLng(103), rtRoot.Children(0).Right
    
    Assert.AreEqual CLng(1003), rtRoot.Children(1).Top
    Assert.AreEqual CLng(1004), rtRoot.Children(1).Bottom
    Assert.AreEqual CLng(101), rtRoot.Children(1).Left
    Assert.AreEqual CLng(103), rtRoot.Children(1).Right
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


