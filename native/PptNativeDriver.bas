Attribute VB_Name = "PptNativeDriver"
Option Explicit

' Native PowerPoint transport. No radius, protection or layout policy here.
Public Function SelectionRoots() As Collection
    Dim result As New Collection, selected As Object, i As Long
    If Application.Presentations.Count = 0 Then Err.Raise 5, , "Open a presentation first."
    If Application.ActiveWindow.Selection.Type <> 2 Then Err.Raise 5, , "Select shapes first."
    Set selected = Application.ActiveWindow.Selection.ShapeRange
    For i = 1 To selected.Count
        result.Add selected.Item(i)
    Next i
    Set SelectionRoots = result
End Function

Public Function CurrentSlide() As Object
    Set CurrentSlide = Application.ActiveWindow.View.Slide
End Function

Public Function ShapeType(ByVal shape As Object) As Long
    ShapeType = shape.Type
End Function

Public Function GeometryType(ByVal shape As Object) As Long
    If shape.Type = 1 Then GeometryType = shape.AutoShapeType
End Function

Public Function ShapeId(ByVal shape As Object) As Long
    ShapeId = shape.Id
End Function

Public Function IsTopLevel(ByVal slide As Object, ByVal shape As Object) As Boolean
    Dim i As Long
    For i = 1 To slide.Shapes.Count
        If slide.Shapes.Item(i).Id = shape.Id Then
            IsTopLevel = True
            Exit Function
        End If
    Next i
End Function

Public Function ShapeName(ByVal shape As Object) As String
    ShapeName = shape.Name
End Function

Public Function ShortSidePoints(ByVal shape As Object) As Double
    ShortSidePoints = shape.Width
    If shape.Height < ShortSidePoints Then ShortSidePoints = shape.Height
End Function

Public Function ReadFraction(ByVal shape As Object) As Double
    ReadFraction = shape.Adjustments.Item(1)
End Function

Public Sub WriteFraction(ByVal shape As Object, ByVal fraction As Double)
    shape.Adjustments.Item(1) = CSng(fraction)
End Sub

Public Function Children(ByVal shape As Object) As Collection
    Dim result As New Collection, i As Long
    For i = 1 To shape.GroupItems.Count
        result.Add shape.GroupItems.Item(i)
    Next i
    Set Children = result
End Function

Public Function ReadTag(ByVal shape As Object, ByVal key As String) As String
    Dim i As Long
    For i = 1 To shape.Tags.Count
        If StrComp(shape.Tags.Name(i), key, vbTextCompare) = 0 Then
            ReadTag = shape.Tags.Value(i)
            Exit Function
        End If
    Next i
End Function

Public Sub AddTag(ByVal shape As Object, ByVal key As String, ByVal value As String)
    shape.Tags.Add key, value
End Sub

Public Sub DeleteTag(ByVal shape As Object, ByVal key As String)
    shape.Tags.Delete key
End Sub

Public Function SnapshotTags(ByVal shape As Object) As Collection
    Dim result As New Collection, i As Long
    For i = 1 To shape.Tags.Count
        result.Add Array(shape.Tags.Name(i), shape.Tags.Value(i))
    Next i
    Set SnapshotTags = result
End Function

Public Sub RestoreMetadata(ByVal shape As Object, ByVal name As String, ByVal tags As Collection)
    Dim pair As Variant
    shape.Name = name
    For Each pair In tags
        shape.Tags.Add CStr(pair(0)), CStr(pair(1))
    Next pair
End Sub

Public Function Ungroup(ByVal shape As Object) As Collection
    Dim result As New Collection, items As Object, i As Long
    Set items = shape.Ungroup
    For i = 1 To items.Count
        result.Add items.Item(i)
    Next i
    Set Ungroup = result
End Function

Public Function FindShape(ByVal shapes As Collection, ByVal id As Long) As Object
    Dim shape As Object
    For Each shape In shapes
        If shape.Id = id Then
            Set FindShape = shape
            Exit Function
        End If
    Next shape
    Err.Raise 5, , "Shape identity changed during ungroup."
End Function

Private Function RangeFor(ByVal slide As Object, ByVal shapes As Collection) As Object
    Dim indexes() As Variant, shape As Object, i As Long, j As Long, found As Boolean
    ReDim indexes(1 To shapes.Count)
    For Each shape In shapes
        i = i + 1
        found = False
        For j = 1 To slide.Shapes.Count
            If slide.Shapes.Item(j).Id = shape.Id Then
                indexes(i) = j
                found = True
                Exit For
            End If
        Next j
        If Not found Then Err.Raise 5, , "Cannot locate group member on the slide."
    Next shape
    Set RangeFor = slide.Shapes.Range(indexes)
End Function

Public Function Regroup(ByVal slide As Object, ByVal shapes As Collection) As Object
    Set Regroup = RangeFor(slide, shapes).Group
End Function

Public Sub SelectShapes(ByVal slide As Object, ByVal shapes As Collection)
    RangeFor(slide, shapes).Select
End Sub
