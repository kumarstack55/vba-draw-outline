Attribute VB_Name = "FactoryModule"
'@Folder "VBAProjectDrawOutline.Libraries"
'@IgnoreModule ProcedureNotUsed, HungarianNotation
Option Explicit
Option Private Module

Public Function CreateSingleDimArray() As SingleDimArray
    Dim sdaArray As SingleDimArray
    
    Set sdaArray = New SingleDimArray
    sdaArray.Initialize
    
    Set CreateSingleDimArray = sdaArray
End Function

'@Ignore ProcedureNotUsed
Public Function CreateCellPoint(ByVal lngColumn As Long, ByVal lngRow As Long) As CellPoint
    Set CreateCellPoint = New CellPoint
    CreateCellPoint.Column = lngColumn
    CreateCellPoint.Row = lngRow
End Function

Public Function CreateCellRect(ByVal lngLeft As Long, ByVal lngTop As Long, ByVal lngRight As Long, ByVal lngBottom As Long) As CellRect
    Dim crResult As CellRect
    
    Set crResult = New CellRect
    crResult.Left = lngLeft
    crResult.Top = lngTop
    crResult.Right = lngRight
    crResult.Bottom = lngBottom
    
    Set CreateCellRect = crResult
End Function

Public Function CreateCellRectFromRange(ByVal rngRange As Range) As CellRect
    Dim rngRightBottom As Range
    
    '@Ignore IndexedDefaultMemberAccess
    Set rngRightBottom = rngRange(rngRange.Count)
    Set CreateCellRectFromRange = CreateCellRect(rngRange.Row, rngRange.Column, rngRightBottom.Row, rngRightBottom.Column)
End Function
