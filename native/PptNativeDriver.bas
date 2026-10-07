Attribute VB_Name = "PptNativeDriver"
Option Explicit

' Native PowerPoint transport. No radius, protection or layout policy here.
Public Function HasShapeSelection() As Boolean
    If Application.Presentations.Count = 0 Then Exit Function
    HasShapeSelection = (Application.ActiveWindow.Selection.Type = 2)
End Function

Public Sub ConnectEvents(ByVal listener As RadiusNativeEvents)
    Set listener.Host = Application
End Sub

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

Public Function SelectionObjects() As Collection
    Dim result As New Collection, selected As Object, selection As Object, i As Long
    If Not HasShapeSelection Then Err.Raise 5, , "Select shapes first."
    Set selection = Application.ActiveWindow.Selection
    If selection.HasChildShapeRange Then
        Set selected = selection.ChildShapeRange
    Else
        Set selected = selection.ShapeRange
    End If
    For i = 1 To selected.Count
        result.Add selected.Item(i)
    Next i
    Set SelectionObjects = result
End Function

Public Function CurrentSlide() As Object
    Set CurrentSlide = Application.ActiveWindow.View.Slide
End Function

Public Function CurrentPresentation() As Object
    Set CurrentPresentation = Application.ActivePresentation
End Function

Public Function SlideId(ByVal slide As Object) As Long
    SlideId = slide.SlideId
End Function

Public Function SlideRoots(ByVal slide As Object) As Collection
    Dim result As New Collection, shape As Object
    For Each shape In slide.Shapes
        result.Add shape
    Next shape
    Set SlideRoots = result
End Function

Public Function HasPresentation() As Boolean
    HasPresentation = (Application.Presentations.Count > 0)
End Function

Public Function HasTag(ByVal shape As Object, ByVal key As String) As Boolean
    Dim i As Long
    For i = 1 To shape.Tags.Count
        If StrComp(shape.Tags.Name(i), key, vbTextCompare) = 0 Then
            HasTag = True
            Exit Function
        End If
    Next i
End Function

Public Sub SelectObject(ByVal shape As Object)
    shape.Select
End Sub

Public Sub SelectObjects(ByVal shapes As Collection)
    Dim shape As Object, first As Boolean
    first = True
    For Each shape In shapes
        shape.Select first
        first = False
    Next shape
End Sub

Public Function FirstSlide(ByVal presentation As Object) As Object
    Set FirstSlide = presentation.Slides(1)
End Function

Public Sub ActivatePresentation(ByVal presentation As Object)
    presentation.Windows(1).Activate
End Sub

Public Function CreateSlideCopy(ByVal slide As Object) As Object
    Dim result As Object, pasted As Object
    Dim errorNumber As Long, errorText As String, recoveryText As String
    On Error GoTo Failed
    slide.Copy
    Set result = Application.Presentations.Add
    result.PageSetup.SlideWidth = slide.Parent.PageSetup.SlideWidth
    result.PageSetup.SlideHeight = slide.Parent.PageSetup.SlideHeight
    Set pasted = result.Slides.Paste(1)
    Set CreateSlideCopy = result
    Exit Function
Failed:
    errorNumber = Err.Number
    errorText = Err.Description
    If Not result Is Nothing Then
        On Error Resume Next
        result.Saved = -1
        result.Close
        If Err.Number <> 0 Then recoveryText = " Preview recovery failed: " & Err.Description
        On Error GoTo 0
    End If
    Err.Raise errorNumber, "PptNativeDriver", errorText & recoveryText
End Function

Public Function ShapeBox(ByVal shape As Object) As Variant
    ShapeBox = Array(CDbl(shape.Left), CDbl(shape.Top), CDbl(shape.Width), CDbl(shape.Height))
End Function

Public Sub SetShapeBox(ByVal shape As Object, ByVal box As Variant)
    Dim locked As Long, errorNumber As Long, errorText As String, recoveryText As String
    locked = shape.LockAspectRatio
    On Error GoTo Failed
    shape.LockAspectRatio = 0
    shape.Width = CSng(box(2))
    shape.Height = CSng(box(3))
    shape.Left = CSng(box(0))
    shape.Top = CSng(box(1))
    shape.LockAspectRatio = locked
    Exit Sub
Failed:
    errorNumber = Err.Number
    errorText = Err.Description
    On Error Resume Next
    Err.Clear
    shape.LockAspectRatio = locked
    If Err.Number <> 0 Then recoveryText = " Aspect-ratio recovery failed: " & Err.Description
    On Error GoTo 0
    Err.Raise errorNumber, "PptNativeDriver", errorText & recoveryText
End Sub

Public Function HasTransform(ByVal shape As Object) As Boolean
    HasTransform = (Abs(CDbl(shape.Rotation)) > 0.0001 Or shape.HorizontalFlip <> 0 Or shape.VerticalFlip <> 0)
End Function

Public Function Slides(ByVal presentation As Object) As Collection
    Dim result As New Collection, slide As Object
    For Each slide In presentation.Slides
        result.Add slide
    Next slide
    Set Slides = result
End Function

Public Function IsCurrentSlide(ByVal slide As Object) As Boolean
    Dim active As Object
    If Not HasPresentation Then Exit Function
    Set active = CurrentSlide()
    IsCurrentSlide = (active Is slide)
End Function

Public Sub ClearSelection()
    Application.ActiveWindow.Selection.Unselect
End Sub

Public Sub AddTextBadge(ByVal slide As Object, ByVal text As String, ByVal left As Double, ByVal top As Double, ByVal width As Double, ByVal color As Long)
    Dim badge As Object
    Set badge = slide.Shapes.AddShape(1, CSng(left), CSng(top), CSng(width), 17!)
    badge.Name = "RadiusRelationPreview_" & CStr(badge.Id)
    badge.Fill.ForeColor.RGB = color
    badge.Line.Visible = 0
    With badge.TextFrame
        .MarginLeft = 4!
        .MarginRight = 3!
        .MarginTop = 1!
        .MarginBottom = 0!
        .TextRange.Text = text
        .TextRange.Font.Size = 10!
        .TextRange.Font.Color.RGB = RGB(255, 255, 255)
    End With
End Sub

Public Sub CloseTemporaryPresentation(ByVal presentation As Object)
    ' Only the caller-owned disposable copy may use this operation.
    presentation.Saved = -1
    presentation.Close
End Sub

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
    ' Mac can flatten nested leaves here; safe for read-only preflight.
    ' Ungroup's returned range supplies the direct transaction members.
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
