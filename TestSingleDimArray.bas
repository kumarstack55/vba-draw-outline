Attribute VB_Name = "TestSingleDimArray"
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

'@TestMethod("Uncategorized")
Private Sub TestCreateSingleDimArray()
    Dim sdaArray As SingleDimArray
    On Error GoTo TestFail
    
    'Arrange:
    
    'Act:
    Set sdaArray = CreateSingleDimArray

    'Assert:
    Assert.AreEqual "SingleDimArray", TypeName(sdaArray)

    '@Ignore LineLabelNotUsed
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Uncategorized")
Private Sub TestGetLastIndex_�v�f�Ȃ���΃}�C�i�X1��Ԃ�()
    Dim sdaArray As SingleDimArray
    On Error GoTo TestFail
    
    'Arrange:

    'Act:
    Set sdaArray = CreateSingleDimArray

    'Assert:
    Assert.AreEqual CLng(-1), sdaArray.GetLastIndex

    '@Ignore LineLabelNotUsed
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Uncategorized")
Private Sub TestGetLength_�v�f�Ȃ���΃[����Ԃ�()
    Dim sdaArray As SingleDimArray
    On Error GoTo TestFail
    
    'Arrange:
    Set sdaArray = CreateSingleDimArray

    'Act:

    'Assert:
    Assert.AreEqual CLng(0), sdaArray.GetLength

    '@Ignore LineLabelNotUsed
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Uncategorized")
Private Sub TestAdd_�v�f��������()
    Dim sdaArray As SingleDimArray
    On Error GoTo TestFail
    
    'Arrange:
    Set sdaArray = CreateSingleDimArray

    'Act:
    sdaArray.Add "a"

    'Assert:
    Assert.Succeed

    '@Ignore LineLabelNotUsed
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Uncategorized")
Private Sub TestGetElement()
    Dim sdaArray As SingleDimArray
    On Error GoTo TestFail
    
    'Arrange:
    Set sdaArray = CreateSingleDimArray

    'Act:
    sdaArray.Add "a"

    'Assert:
    Assert.AreEqual "a", sdaArray.GetElement(0)

    '@Ignore LineLabelNotUsed
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Uncategorized")
Private Sub TestAdd_����������Ă��ێ������()
    Dim sdaArray As SingleDimArray
    On Error GoTo TestFail
    
    'Arrange:
    Set sdaArray = CreateSingleDimArray

    'Act:
    sdaArray.Add "a"
    sdaArray.Add "b"

    'Assert:
    Assert.AreEqual "a", sdaArray.GetElement(0)

    '@Ignore LineLabelNotUsed
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Uncategorized")
Private Sub TestGetLastIndex_�Ō�̓Y������Ԃ�()
    Dim sdaArray As SingleDimArray
    On Error GoTo TestFail
    
    'Arrange:
    Set sdaArray = CreateSingleDimArray

    'Act:
    sdaArray.Add "a"

    'Assert:
    Assert.AreEqual CLng(0), sdaArray.GetLastIndex

    '@Ignore LineLabelNotUsed
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Uncategorized")
Private Sub TestGetLength_������Ԃ�()
    Dim sdaArray As SingleDimArray
    On Error GoTo TestFail
    
    'Arrange:
    Set sdaArray = CreateSingleDimArray

    'Act:
    sdaArray.Add "a"

    'Assert:
    Assert.AreEqual CLng(1), sdaArray.GetLength

    '@Ignore LineLabelNotUsed
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

