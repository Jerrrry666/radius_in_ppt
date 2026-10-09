Attribute VB_Name = "PptNativeDriver"
Option Explicit

' Native PowerPoint transport. No radius, protection or layout policy here.
Public Function HasShapeSelection() As Boolean
    If Application.Presentations.Count = 0 Then Exit Function
    HasShapeSelection = (Application.ActiveWindow.Selection.Type = 2)
End Function

Public Function HasChildShapeSelection() As Boolean
    If Not HasShapeSelection Then Exit Function
    HasChildShapeSelection = Application.ActiveWindow.Selection.HasChildShapeRange
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

Public Function PresentationName(ByVal presentation As Object) As String
    PresentationName = presentation.Name
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

Public Function CurrentWindow() As Object
    Set CurrentWindow = Application.ActiveWindow
End Function

Public Function CreateTemporaryPresentation() As Object
    Set CreateTemporaryPresentation = Application.Presentations.Add
End Function

Public Function AddBlankSlide(ByVal presentation As Object) As Object
    Set AddBlankSlide = presentation.Slides.Add(presentation.Slides.Count + 1, 12)
End Function

Public Sub ActivateSlide(ByVal window As Object, ByVal slide As Object)
    Dim windowPresentation As Object, slidePresentation As Object
    Set windowPresentation = window.Presentation
    Set slidePresentation = slide.Parent
    If Not windowPresentation Is slidePresentation Then Err.Raise 5, , "The target slide belongs to another window's presentation."
    window.Activate
    ' Only this explicitly supplied window is prepared for shape operations.
    window.ViewType = 9
    window.View.GotoSlide slide.SlideIndex
End Sub

Public Function AddShape(ByVal slide As Object, ByVal geometryType As Long, ByVal left As Double, ByVal top As Double, ByVal width As Double, ByVal height As Double) As Object
    Set AddShape = slide.Shapes.AddShape(geometryType, CSng(left), CSng(top), CSng(width), CSng(height))
End Function

Public Sub SetRotation(ByVal shape As Object, ByVal degrees As Double)
    shape.Rotation = CSng(degrees)
End Sub

Public Function CaptureWindowContext() As Collection
    Dim result As New Collection, window As Object, presentation As Object, slide As Object
    Dim selection As Object, text As Object, selectedSlides As Object, shapes As New Collection, shape As Object
    Dim kind As Long, viewType As Long, saved As Long, textStart As Long, textLength As Long, childSelection As Boolean
    Dim errorNumber As Long, errorText As String
    On Error GoTo Failed
    If Application.Windows.Count > 0 Then
        Set window = Application.ActiveWindow
        Set presentation = window.Presentation
        viewType = window.ViewType
        saved = presentation.Saved
        Set selection = window.Selection
        kind = selection.Type
        If kind = 2 Then
            childSelection = selection.HasChildShapeRange
            If childSelection Then
                For Each shape In selection.ChildShapeRange
                    shapes.Add shape
                Next shape
            Else
                For Each shape In selection.ShapeRange
                    shapes.Add shape
                Next shape
            End If
        ElseIf kind = 3 Then
            Set text = selection.TextRange
            textStart = text.Start
            textLength = text.Length
        ElseIf kind = 1 Then
            Set selectedSlides = selection.SlideRange
        End If
        ' Slide sorter/master views may not expose View.Slide. They are
        ' never changed by a temporary presentation, so retain that view.
        On Error Resume Next
        Set slide = window.View.Slide
        Err.Clear
        On Error GoTo Failed
    End If
    result.Add window, "Window"
    result.Add presentation, "Presentation"
    result.Add slide, "Slide"
    result.Add viewType, "ViewType"
    result.Add saved, "Saved"
    result.Add kind, "SelectionType"
    result.Add shapes, "Shapes"
    result.Add childSelection, "ChildSelection"
    result.Add text, "TextRange"
    result.Add textStart, "TextStart"
    result.Add textLength, "TextLength"
    result.Add selectedSlides, "SlideRange"
    Set CaptureWindowContext = result
    Exit Function
Failed:
    errorNumber = Err.Number
    errorText = Err.Description
    Err.Raise errorNumber, "PptNativeDriver", "Window context capture failed: " & errorText
End Function

Public Sub RestoreWindowContext(ByVal context As Collection)
    Dim window As Object, presentation As Object, slide As Object, current As Object
    Dim text As Object, selectedSlides As Object, shapes As Collection, selected As Object, i As Long
    Dim errorNumber As Long, errorText As String
    On Error GoTo Failed
    Set window = context("Window")
    If window Is Nothing Then Exit Sub
    Set presentation = context("Presentation")
    Set slide = context("Slide")
    window.Activate
    Set current = Application.ActiveWindow
    If Not current Is window Then Err.Raise 5, , "The exact original window could not be activated."
    ' Do not navigate or switch the original window's view.
    If window.ViewType <> CLng(context("ViewType")) Then Err.Raise 5, , "The original window view changed."
    If Not slide Is Nothing Then
        Set current = window.View.Slide
        If Not current Is slide Then Err.Raise 5, , "The original window slide changed."
    End If
    Select Case CLng(context("SelectionType"))
        Case 0
            window.Selection.Unselect
        Case 1
            Set selectedSlides = context("SlideRange")
            selectedSlides.Select
        Case 2
            Set shapes = context("Shapes")
            SelectObjects shapes
        Case 3
            Set text = context("TextRange")
            text.Select
    End Select
    If window.Selection.Type <> CLng(context("SelectionType")) Then Err.Raise 5, , "The original selection type was not restored."
    If CLng(context("SelectionType")) = 2 Then
        If CBool(window.Selection.HasChildShapeRange) <> CBool(context("ChildSelection")) Then Err.Raise 5, , "The original group-child selection was not restored."
        If CBool(context("ChildSelection")) Then Set selected = window.Selection.ChildShapeRange Else Set selected = window.Selection.ShapeRange
        If selected.Count <> shapes.Count Then Err.Raise 5, , "The original selected shape count changed."
        For i = 1 To shapes.Count
            If selected.Item(i).Id <> shapes(i).Id Then Err.Raise 5, , "The original selected shape IDs changed."
        Next i
    ElseIf CLng(context("SelectionType")) = 3 Then
        Set text = window.Selection.TextRange
        If text.Start <> CLng(context("TextStart")) Or text.Length <> CLng(context("TextLength")) Then Err.Raise 5, , "The original text selection bounds changed."
    ElseIf CLng(context("SelectionType")) = 1 Then
        Set selected = window.Selection.SlideRange
        If selected.Count <> selectedSlides.Count Then Err.Raise 5, , "The original selected slide count changed."
        For i = 1 To selectedSlides.Count
            If selected.Item(i).SlideId <> selectedSlides.Item(i).SlideId Then Err.Raise 5, , "The original selected slide IDs changed."
        Next i
    End If
    If window.ViewType <> CLng(context("ViewType")) Then Err.Raise 5, , "Restoring the selection changed the original view."
    ' Saved is observed only. Never mark the source document as saved.
    If presentation.Saved <> CLng(context("Saved")) Then Err.Raise 5, , "The source presentation Saved state changed."
    Exit Sub
Failed:
    errorNumber = Err.Number
    errorText = Err.Description
    Err.Raise errorNumber, "PptNativeDriver", "Window context restore failed: " & errorText
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
    Dim i As Long, id As Long
    id = shape.Id
    For i = 1 To slide.Shapes.Count
        If slide.Shapes.Item(i).Id = id Then
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
    Dim present As Boolean
    ReadTag = ReadTagState(shape, key, present)
End Function

Public Function ReadTagState(ByVal shape As Object, ByVal key As String, ByRef present As Boolean) As String
    Dim i As Long
    present = False
    For i = 1 To shape.Tags.Count
        If StrComp(shape.Tags.Name(i), key, vbTextCompare) = 0 Then
            ReadTagState = shape.Tags.Value(i)
            present = True
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
    Dim pair As Variant, errorNumber As Long, errorText As String
    ' Try every saved field even when one host property fails. The caller
    ' retains this snapshot and can retry on the already regrouped object.
    On Error Resume Next
    Err.Clear
    shape.Name = name
    If Err.Number <> 0 Then
        errorNumber = Err.Number
        errorText = "Name restore failed: " & Err.Description
    End If
    For Each pair In tags
        Err.Clear
        shape.Tags.Add CStr(pair(0)), CStr(pair(1))
        If Err.Number <> 0 Then
            If errorNumber = 0 Then errorNumber = Err.Number
            errorText = errorText & " Tag " & CStr(pair(0)) & " restore failed: " & Err.Description
        End If
    Next pair
    On Error GoTo 0
    If errorNumber <> 0 Then Err.Raise errorNumber, "PptNativeDriver", errorText
End Sub

Public Function Ungroup(ByVal shape As Object) As Collection
    Dim items As Object
    UngroupForTransaction shape, items
    Set Ungroup = RangeMembers(items)
End Function

Public Sub UngroupForTransaction(ByVal shape As Object, ByRef items As Object)
    ' Capture the destructive operation's result before enumerating it.
    ' A caller can regroup this range if later member reads fail.
    Set items = shape.Ungroup
End Sub

Public Function RangeMembers(ByVal items As Object) As Collection
    Dim result As New Collection, i As Long
    For i = 1 To items.Count
        result.Add items.Item(i)
    Next i
    Set RangeMembers = result
End Function

Private Function RangeFor(ByVal slide As Object, ByVal shapes As Collection) As Object
    Dim indexes() As Variant, positions As New Collection, shape As Object, i As Long, j As Long
    If shapes.Count = 0 Then Err.Raise 5, , "No shapes were provided for the range."
    ReDim indexes(1 To shapes.Count)
    ' This index belongs to this one call; regroup invalidates slide indexes.
    For j = 1 To slide.Shapes.Count
        positions.Add j, "S" & CStr(slide.Shapes.Item(j).Id)
    Next j
    For Each shape In shapes
        i = i + 1
        indexes(i) = PositionFor(positions, shape.Id)
    Next shape
    Set RangeFor = slide.Shapes.Range(indexes)
End Function

Private Function PositionFor(ByVal positions As Collection, ByVal id As Long) As Long
    Dim errorNumber As Long, errorText As String
    On Error GoTo Failed
    PositionFor = CLng(positions("S" & CStr(id)))
    Exit Function
Failed:
    errorNumber = Err.Number
    errorText = Err.Description
    If errorNumber = 5 Then errorText = "Cannot locate group member on the slide: " & CStr(id) & ". " & errorText
    Err.Raise errorNumber, "PptNativeDriver", errorText
End Function

Public Function Regroup(ByVal slide As Object, ByVal shapes As Collection) As Object
    Set Regroup = RangeFor(slide, shapes).Group
End Function

Public Sub SelectShapes(ByVal slide As Object, ByVal shapes As Collection)
    RangeFor(slide, shapes).Select
End Sub
