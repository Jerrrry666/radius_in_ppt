Attribute VB_Name = "RadiusNativeCore"
Option Explicit

Public Const PT_PER_CM As Double = 28.3464566929134
Private Const LOCK_KEY As String = "radiusLock_v1"
Private Const STRICT_KEY As String = "radiusLockStrict_v1"

Public Function ParseNumber(ByVal text As String) As Double
    Dim normalized As String, character As String, dots As Long, digits As Long, i As Long
    normalized = Trim$(Replace(text, ",", "."))
    For i = 1 To Len(normalized)
        character = Mid$(normalized, i, 1)
        If character >= "0" And character <= "9" Then
            digits = digits + 1
        ElseIf character = "." Then
            dots = dots + 1
        Else
            Err.Raise 5, , "Enter a nonnegative number."
        End If
    Next i
    If digits = 0 Or dots > 1 Then Err.Raise 5, , "Enter a nonnegative number."
    ParseNumber = Val(normalized)
    If ParseNumber < 0 Or ParseNumber > 1000000 Then Err.Raise 5, , "Radius is out of range."
End Function

Public Function NumberText(ByVal value As Double) As String
    NumberText = Replace(Format$(value, "0.000000"), ",", ".")
End Function

Public Function StepValue(ByVal value As Double, ByVal unit As String, ByVal direction As Long) As Double
    Dim maximum As Double
    If direction <> 1 And direction <> -1 Then Err.Raise 5, , "Unsupported step direction."
    If value < 0 Or value > 1000000 Then Err.Raise 5, , "Radius is out of range."
    If unit = "cm" Then
        maximum = 1000000
    ElseIf unit = "%" Then
        maximum = 50
    Else
        Err.Raise 5, , "Unsupported unit."
    End If
    StepValue = Round(value + direction * 0.1, 6)
    If StepValue < 0 Then StepValue = 0
    If StepValue > maximum Then StepValue = maximum
End Function

Public Function ComputeTarget(ByVal value As Double, ByVal unit As String, ByVal shortCm As Double) As Double
    If shortCm <= 0 Then Err.Raise 5, , "Shape has zero width or height."
    If value < 0 Then Err.Raise 5, , "Radius must not be negative."
    If unit = "%" Then
        ComputeTarget = value * shortCm / 100#
    ElseIf unit = "cm" Then
        ComputeTarget = value
    Else
        Err.Raise 5, , "Unsupported unit."
    End If
    If ComputeTarget > shortCm / 2# Then ComputeTarget = shortCm / 2#
End Function

Public Function IsRoundRect(ByVal shape As Object) As Boolean
    IsRoundRect = (PptNativeDriver.GeometryType(shape) = 5)
End Function

Public Function CurrentRadius(ByVal shape As Object) As Double
    CurrentRadius = PptNativeDriver.ReadFraction(shape) * PptNativeDriver.ShortSidePoints(shape) / PT_PER_CM
End Function

Private Sub CollectLeaves(ByVal shape As Object, ByVal leaves As Collection, ByVal depth As Long)
    Dim child As Object
    If depth > 64 Then Err.Raise 5, , "Group nesting is too deep."
    If PptNativeDriver.ShapeType(shape) = 6 Then
        For Each child In PptNativeDriver.Children(shape)
            CollectLeaves child, leaves, depth + 1
        Next child
    ElseIf IsRoundRect(shape) Then
        leaves.Add shape
    End If
End Sub

Public Function SelectionLeaves() As Collection
    Dim roots As Collection, leaves As New Collection, shape As Object
    Set roots = PptNativeDriver.SelectionRoots()
    For Each shape In roots
        CollectLeaves shape, leaves, 0
    Next shape
    If leaves.Count = 0 Then Err.Raise 5, , "Select rounded rectangles."
    Set SelectionLeaves = leaves
End Function

Public Sub SelectionSummary(ByRef count As Long, ByRef protectedCount As Long, ByRef editable As Boolean)
    Dim roots As Collection, leaves As New Collection, shape As Object, slide As Object
    count = 0
    protectedCount = 0
    editable = False
    If Not PptNativeDriver.HasShapeSelection() Then Exit Sub
    Set roots = PptNativeDriver.SelectionRoots()
    Set slide = PptNativeDriver.CurrentSlide()
    editable = True
    For Each shape In roots
        If Not PptNativeDriver.IsTopLevel(slide, shape) Then editable = False
        CollectLeaves shape, leaves, 0
    Next shape
    count = leaves.Count
    For Each shape In leaves
        If PptNativeDriver.ReadTag(shape, STRICT_KEY) = "1" Then protectedCount = protectedCount + 1
    Next shape
End Sub

Private Function PrepareNode(ByVal shape As Object, ByVal value As Double, ByVal unit As String, ByVal depth As Long, ByVal action As String) As Collection
    Dim node As New Collection, nodes As New Collection, child As Object, prepared As Collection, eligible As Long
    If depth > 64 Then Err.Raise 5, , "Group nesting is too deep."
    node.Add shape, "Shape"
    node.Add PptNativeDriver.ShapeId(shape), "Id"
    node.Add PptNativeDriver.ShapeName(shape), "Name"
    node.Add PptNativeDriver.SnapshotTags(shape), "Tags"
    node.Add action, "Action"
    node.Add value, "Value"
    node.Add unit, "Unit"
    ' Preserve layout metadata during explicit radius/protection edits.
    If PptNativeDriver.ShapeType(shape) = 6 Then
        For Each child In PptNativeDriver.Children(shape)
            Set prepared = PrepareNode(child, value, unit, depth + 1, action)
            eligible = eligible + CLng(prepared("Eligible"))
            nodes.Add prepared
        Next child
    ElseIf IsRoundRect(shape) Then
        ' Full live preflight happens before ANY write or ungroup operation.
        If action = "radius" And PptNativeDriver.ReadTag(shape, STRICT_KEY) = "1" Then Err.Raise 5, , "Turn off protection explicitly before applying radius."
        If action = "radius" Then
            node.Add ComputeTarget(value, unit, PptNativeDriver.ShortSidePoints(shape) / PT_PER_CM), "Target"
        Else
            node.Add CurrentRadius(shape), "Target"
        End If
        eligible = 1
    End If
    node.Add nodes, "Children"
    node.Add eligible, "Eligible"
    Set PrepareNode = node
End Function

Private Sub SetNodeShape(ByVal node As Collection, ByVal shape As Object)
    node.Remove "Shape"
    node.Add shape, "Shape"
End Sub

Private Function NodeShapes(ByVal nodes As Collection) As Collection
    Dim result As New Collection, node As Collection
    For Each node In nodes
        result.Add node("Shape")
    Next node
    Set NodeShapes = result
End Function

Private Sub WriteNode(ByVal node As Collection, ByVal slide As Object, Optional ByVal depth As Long = 0)
    Dim shape As Object, fresh As Collection, children As Collection, child As Collection, member As Object, directNodes As Collection
    Dim opened As Boolean, mapped As Boolean, errorNumber As Long, errorText As String, restoreError As String
    Dim actualCm As Double
    Set shape = node("Shape")
    Set children = node("Children")
    If CLng(node("Eligible")) = 0 Then Exit Sub
    If depth > 64 Then Err.Raise 5, , "Group nesting is too deep."
    On Error GoTo Failed
    If children.Count > 0 Then
        Set fresh = PptNativeDriver.Ungroup(shape)
        opened = True
        ' GroupItems may flatten nested leaves on Mac. Build the actual
        ' direct hierarchy from the ungroup return, never old group proxies.
        Set directNodes = New Collection
        For Each member In fresh
            directNodes.Add PrepareNode(member, CDbl(node("Value")), CStr(node("Unit")), depth + 1, CStr(node("Action")))
        Next member
        Set children = directNodes
        node.Remove "Children"
        node.Add children, "Children"
        mapped = True
        For Each child In children
            WriteNode child, slide, depth + 1
        Next child
        Set shape = PptNativeDriver.Regroup(slide, NodeShapes(children))
        SetNodeShape node, shape
        opened = False
        PptNativeDriver.RestoreMetadata shape, CStr(node("Name")), node("Tags")
    ElseIf IsRoundRect(shape) Then
        Select Case CStr(node("Action"))
            Case "radius"
                ' Second defense: read the tag again immediately before writing.
                If PptNativeDriver.ReadTag(shape, STRICT_KEY) = "1" Then Err.Raise 5, , "Protected shape; radius was not written."
                actualCm = ComputeTarget(CDbl(node("Value")), CStr(node("Unit")), PptNativeDriver.ShortSidePoints(shape) / PT_PER_CM)
                PptNativeDriver.WriteFraction shape, actualCm * PT_PER_CM / PptNativeDriver.ShortSidePoints(shape)
                If PptNativeDriver.ReadTag(shape, LOCK_KEY) <> "" Then PptNativeDriver.AddTag shape, LOCK_KEY, NumberText(actualCm)
            Case "protect"
                If PptNativeDriver.ReadTag(shape, LOCK_KEY) = "" Then PptNativeDriver.AddTag shape, LOCK_KEY, NumberText(CurrentRadius(shape))
                PptNativeDriver.AddTag shape, STRICT_KEY, "1"
            Case "unprotect"
                PptNativeDriver.DeleteTag shape, STRICT_KEY
            Case Else
                Err.Raise 5, , "Unsupported edit operation."
        End Select
    End If
    Exit Sub
Failed:
    errorNumber = Err.Number
    errorText = Err.Description
    Debug.Print "[NativeWrite] " & errorText
    If opened Then
        On Error Resume Next
        Err.Clear
        If mapped Then
            Set shape = PptNativeDriver.Regroup(slide, NodeShapes(children))
        Else
            Set shape = PptNativeDriver.Regroup(slide, fresh)
        End If
        If Err.Number = 0 Then
            SetNodeShape node, shape
            PptNativeDriver.RestoreMetadata shape, CStr(node("Name")), node("Tags")
        End If
        If Err.Number <> 0 Then restoreError = " Group recovery failed: " & Err.Description
        On Error GoTo 0
    End If
    Err.Raise errorNumber, "RadiusNativeCore", errorText & restoreError
End Sub

Public Sub ApplySelection(ByVal value As Double, ByVal unit As String, Optional ByVal action As String = "radius")
    Dim roots As Collection, leaves As New Collection, nodes As New Collection, shape As Object, node As Collection, slide As Object
    Dim linkedPlan As Collection
    Dim errorNumber As Long, errorText As String
    If RadiusNativeRelations.IsPreviewContext Then Err.Raise 5, , "Close the relation preview before editing the source."
    Set roots = PptNativeDriver.SelectionRoots()
    Set slide = PptNativeDriver.CurrentSlide()
    If action <> "radius" And action <> "protect" And action <> "unprotect" Then Err.Raise 5, , "Unsupported edit operation."
    If action = "radius" Then
        Set linkedPlan = RadiusNativeLayout.LinkedRadiusPlan(roots, value, unit)
        If Not linkedPlan Is Nothing Then
            ApplyShapePlan slide, linkedPlan
            Exit Sub
        End If
    End If
    For Each shape In roots
        If Not PptNativeDriver.IsTopLevel(slide, shape) Then Err.Raise 5, , "Select the complete top-level group for native editing."
        CollectLeaves shape, leaves, 0
        nodes.Add PrepareNode(shape, value, unit, 0, action)
    Next shape
    If leaves.Count = 0 Then Err.Raise 5, , "Select rounded rectangles."
    On Error GoTo Failed
    For Each node In nodes
        WriteNode node, slide
    Next node
    PptNativeDriver.SelectShapes slide, NodeShapes(nodes)
    Exit Sub
Failed:
    errorNumber = Err.Number
    errorText = Err.Description
    On Error Resume Next
    PptNativeDriver.SelectShapes slide, NodeShapes(nodes)
    On Error GoTo 0
    Err.Raise errorNumber, "RadiusNativeCore", errorText
End Sub

Private Function MetadataNode(ByVal shape As Object, ByVal plan As Collection, ByVal depth As Long) As Collection
    Dim node As New Collection, children As New Collection, child As Object, prepared As Collection, change As Variant, eligible As Long
    If depth > 64 Then Err.Raise 5, , "Group nesting is too deep."
    node.Add shape, "Shape"
    node.Add PptNativeDriver.ShapeId(shape), "Id"
    node.Add PptNativeDriver.ShapeName(shape), "Name"
    node.Add PptNativeDriver.SnapshotTags(shape), "Tags"
    If PptNativeDriver.ShapeType(shape) = 6 Then
        For Each child In PptNativeDriver.Children(shape)
            Set prepared = MetadataNode(child, plan, depth + 1)
            eligible = eligible + CLng(prepared("Eligible"))
            children.Add prepared
        Next child
    Else
        For Each change In plan
            If CLng(change(0)) = PptNativeDriver.ShapeId(shape) Then eligible = 1
        Next change
    End If
    node.Add children, "Children"
    node.Add eligible, "Eligible"
    Set MetadataNode = node
End Function

Private Function MetadataRoots(ByVal slide As Object, ByVal plan As Collection) As Collection
    Dim result As New Collection, shape As Object, node As Collection
    For Each shape In PptNativeDriver.SlideRoots(slide)
        Set node = MetadataNode(shape, plan, 0)
        If CLng(node("Eligible")) > 0 Then result.Add node
    Next shape
    Set MetadataRoots = result
End Function

Private Sub WriteMetadataNode(ByVal node As Collection, ByVal slide As Object, ByVal plan As Collection, ByVal depth As Long)
    Dim shape As Object, member As Object, fresh As Collection, children As Collection, child As Collection, change As Variant
    Dim opened As Boolean, mapped As Boolean, errorNumber As Long, errorText As String, recoveryText As String
    If CLng(node("Eligible")) = 0 Then Exit Sub
    If depth > 64 Then Err.Raise 5, , "Group nesting is too deep."
    Set shape = node("Shape")
    On Error GoTo Failed
    If PptNativeDriver.ShapeType(shape) = 6 Then
        Set fresh = PptNativeDriver.Ungroup(shape)
        opened = True
        Set children = New Collection
        For Each member In fresh
            children.Add MetadataNode(member, plan, depth + 1)
        Next member
        node.Remove "Children"
        node.Add children, "Children"
        mapped = True
        For Each child In children
            WriteMetadataNode child, slide, plan, depth + 1
        Next child
        Set shape = PptNativeDriver.Regroup(slide, NodeShapes(children))
        SetNodeShape node, shape
        opened = False
        PptNativeDriver.RestoreMetadata shape, CStr(node("Name")), node("Tags")
    Else
        For Each change In plan
            If CLng(change(0)) = PptNativeDriver.ShapeId(shape) Then
                If CBool(change(3)) Then
                    PptNativeDriver.DeleteTag shape, CStr(change(1))
                Else
                    PptNativeDriver.AddTag shape, CStr(change(1)), CStr(change(2))
                End If
            End If
        Next change
    End If
    Exit Sub
Failed:
    errorNumber = Err.Number
    errorText = Err.Description
    If opened Then
        On Error Resume Next
        Err.Clear
        If mapped Then
            Set shape = PptNativeDriver.Regroup(slide, NodeShapes(children))
        Else
            Set shape = PptNativeDriver.Regroup(slide, fresh)
        End If
        If Err.Number = 0 Then
            SetNodeShape node, shape
            PptNativeDriver.RestoreMetadata shape, CStr(node("Name")), node("Tags")
        End If
        If Err.Number <> 0 Then recoveryText = " Group recovery failed: " & Err.Description
        On Error GoTo 0
    End If
    Err.Raise errorNumber, "RadiusNativeCore", errorText & recoveryText
End Sub

Public Sub ApplyMetadataPlan(ByVal slide As Object, ByVal plan As Collection)
    Dim leaves As Collection, shape As Object, nodes As Collection, node As Collection, change As Variant, rollback As New Collection
    Dim matches As Long, errorNumber As Long, errorText As String, recoveryText As String
    If RadiusNativeRelations.IsPreviewContext Then Err.Raise 5, , "Close the relation preview before editing the source."
    Set leaves = RadiusNativeRelations.SlideLeaves(slide)
    ' Resolve every target and capture old values BEFORE any write or ungroup.
    For Each change In plan
        If StrComp(CStr(change(1)), STRICT_KEY, vbTextCompare) = 0 Or StrComp(CStr(change(1)), LOCK_KEY, vbTextCompare) = 0 Then Err.Raise 5, , "Relationship edits cannot change radius protection."
        matches = 0
        For Each shape In leaves
            If PptNativeDriver.ShapeId(shape) = CLng(change(0)) Then
                matches = matches + 1
                rollback.Add Array(change(0), change(1), PptNativeDriver.ReadTag(shape, CStr(change(1))), Not PptNativeDriver.HasTag(shape, CStr(change(1))))
            End If
        Next shape
        If matches <> 1 Then Err.Raise 5, , "Cannot uniquely resolve a relationship target."
    Next change
    Set nodes = MetadataRoots(slide, plan)
    On Error GoTo Failed
    For Each node In nodes
        WriteMetadataNode node, slide, plan, 0
    Next node
    PptNativeDriver.SelectShapes slide, NodeShapes(nodes)
    Exit Sub
Failed:
    errorNumber = Err.Number
    errorText = Err.Description
    On Error Resume Next
    Err.Clear
    Set nodes = MetadataRoots(slide, rollback)
    If Err.Number = 0 Then
        For Each node In nodes
            WriteMetadataNode node, slide, rollback, 0
            If Err.Number <> 0 Then Exit For
        Next node
    End If
    If Err.Number <> 0 Then recoveryText = " Metadata recovery failed: " & Err.Description
    Err.Clear
    PptNativeDriver.SelectShapes slide, NodeShapes(nodes)
    If Err.Number <> 0 Then recoveryText = recoveryText & " Selection recovery failed: " & Err.Description
    On Error GoTo 0
    Err.Raise errorNumber, "RadiusNativeCore", errorText & recoveryText
End Sub

Public Sub PutRadius(ByVal plan As Collection, ByVal id As Long, ByVal value As Double)
    Dim i As Long, change As Variant
    For i = plan.Count To 1 Step -1
        change = plan(i)
        If CLng(change(0)) = id And CStr(change(1)) = "radius" Then plan.Remove i
    Next i
    plan.Add Array(id, "radius", value)
End Sub

Private Function ShapeForId(ByVal shapes As Collection, ByVal id As Long) As Object
    Dim shape As Object
    For Each shape In shapes
        If PptNativeDriver.ShapeId(shape) = id Then Set ShapeForId = shape: Exit Function
    Next shape
End Function

Private Sub CheckSceneNodes(ByVal nodes As Collection, ByVal spatial As Boolean)
    Dim node As Collection
    For Each node In nodes
        If CLng(node("Eligible")) > 0 Then
            If spatial And PptNativeDriver.HasTransform(node("Shape")) Then Err.Raise 5, , "Layout and linked radius do not support rotated or flipped members/groups."
            CheckSceneNodes node("Children"), spatial
        End If
    Next node
End Sub

Private Sub WriteSceneNode(ByVal node As Collection, ByVal slide As Object, ByVal plan As Collection, ByVal depth As Long)
    Dim shape As Object, fresh As Collection, children As Collection, child As Collection, member As Object, change As Variant
    Dim opened As Boolean, mapped As Boolean, errorNumber As Long, errorText As String, recoveryText As String, actualCm As Double
    If CLng(node("Eligible")) = 0 Then Exit Sub
    If depth > 64 Then Err.Raise 5, , "Group nesting is too deep."
    Set shape = node("Shape")
    On Error GoTo Failed
    If PptNativeDriver.ShapeType(shape) = 6 Then
        Set fresh = PptNativeDriver.Ungroup(shape)
        opened = True
        Set children = New Collection
        For Each member In fresh
            children.Add MetadataNode(member, plan, depth + 1)
        Next member
        node.Remove "Children"
        node.Add children, "Children"
        mapped = True
        For Each child In children
            WriteSceneNode child, slide, plan, depth + 1
        Next child
        Set shape = PptNativeDriver.Regroup(slide, NodeShapes(children))
        SetNodeShape node, shape
        opened = False
        PptNativeDriver.RestoreMetadata shape, CStr(node("Name")), node("Tags")
    Else
        ' Resize before calculating the clamped radius from the fresh shape.
        For Each change In plan
            If CLng(change(0)) = PptNativeDriver.ShapeId(shape) And CStr(change(1)) = "box" Then
                If PptNativeDriver.ReadTag(shape, STRICT_KEY) = "1" Then Err.Raise 5, , "Protected child; layout was not written."
                PptNativeDriver.SetShapeBox shape, change(2)
            End If
        Next change
        For Each change In plan
            If CLng(change(0)) = PptNativeDriver.ShapeId(shape) Then
                Select Case CStr(change(1))
                    Case "radius", "fraction"
                        If PptNativeDriver.ReadTag(shape, STRICT_KEY) = "1" Then Err.Raise 5, , "Protected shape; linked radius was not written."
                        If CStr(change(1)) = "fraction" Then
                            PptNativeDriver.WriteFraction shape, CDbl(change(2))
                        Else
                            actualCm = ComputeTarget(CDbl(change(2)), "cm", PptNativeDriver.ShortSidePoints(shape) / PT_PER_CM)
                            PptNativeDriver.WriteFraction shape, actualCm * PT_PER_CM / PptNativeDriver.ShortSidePoints(shape)
                            If PptNativeDriver.ReadTag(shape, LOCK_KEY) <> "" Then PptNativeDriver.AddTag shape, LOCK_KEY, NumberText(actualCm)
                        End If
                    Case "tag"
                        If StrComp(CStr(change(2)), STRICT_KEY, vbTextCompare) = 0 Then Err.Raise 5, , "Layout cannot change strict protection."
                        If CBool(change(4)) Then
                            PptNativeDriver.DeleteTag shape, CStr(change(2))
                        Else
                            PptNativeDriver.AddTag shape, CStr(change(2)), CStr(change(3))
                        End If
                End Select
            End If
        Next change
    End If
    Exit Sub
Failed:
    errorNumber = Err.Number
    errorText = Err.Description
    If opened Then
        On Error Resume Next
        Err.Clear
        If mapped Then Set shape = PptNativeDriver.Regroup(slide, NodeShapes(children)) Else Set shape = PptNativeDriver.Regroup(slide, fresh)
        If Err.Number = 0 Then
            SetNodeShape node, shape
            PptNativeDriver.RestoreMetadata shape, CStr(node("Name")), node("Tags")
        End If
        If Err.Number <> 0 Then recoveryText = " Group recovery failed: " & Err.Description
        On Error GoTo 0
    End If
    Err.Raise errorNumber, "RadiusNativeCore", errorText & recoveryText
End Sub

Private Sub FindNodeSelection(ByVal nodes As Collection, ByVal id As Long, ByVal result As Collection)
    Dim node As Collection
    For Each node In nodes
        If CLng(node("Id")) = id Then result.Add node("Shape"): Exit Sub
        FindNodeSelection node("Children"), id, result
    Next node
End Sub

Private Sub RestoreSceneSelection(ByVal slide As Object, ByVal nodes As Collection, ByVal ids As Collection)
    Dim shapes As New Collection, id As Variant, roots As Collection, shape As Object, before As Long
    Set roots = PptNativeDriver.SlideRoots(slide)
    For Each id In ids
        before = shapes.Count
        FindNodeSelection nodes, CLng(id), shapes
        If shapes.Count = before Then
            Set shape = ShapeForId(roots, CLng(id))
            If shape Is Nothing Then Set shape = ShapeForId(RadiusNativeRelations.SlideLeaves(slide), CLng(id))
            If Not shape Is Nothing Then shapes.Add shape
        End If
    Next id
    If shapes.Count > 0 Then PptNativeDriver.SelectObjects shapes Else PptNativeDriver.ClearSelection
End Sub

Public Sub ApplyShapePlan(ByVal slide As Object, ByVal plan As Collection)
    Dim leaves As Collection, shape As Object, nodes As Collection, node As Collection, change As Variant, rollback As New Collection, ids As New Collection
    Dim spatial As Boolean, sameSlide As Boolean, errorNumber As Long, errorText As String, recoveryText As String
    If RadiusNativeRelations.IsPreviewContext Then Err.Raise 5, , "Close the relation preview before editing the source."
    If plan.Count = 0 Then Exit Sub
    Set leaves = RadiusNativeRelations.SlideLeaves(slide)
    sameSlide = PptNativeDriver.IsCurrentSlide(slide)
    If sameSlide And PptNativeDriver.HasShapeSelection Then
        For Each shape In PptNativeDriver.SelectionObjects()
            ids.Add PptNativeDriver.ShapeId(shape)
        Next shape
    End If
    ' All targets, live protection, geometry and old values are checked first.
    For Each change In plan
        Set shape = ShapeForId(leaves, CLng(change(0)))
        If shape Is Nothing Then Err.Raise 5, , "Cannot resolve a layout/radius target."
        Select Case CStr(change(1))
            Case "box", "radius"
                spatial = True
                If PptNativeDriver.ReadTag(shape, STRICT_KEY) = "1" Then Err.Raise 5, , "Turn off protection explicitly before layout or linked radius: " & PptNativeDriver.ShapeName(shape)
                If CStr(change(1)) = "box" Then
                    If CDbl(change(2)(2)) <= 0 Or CDbl(change(2)(3)) <= 0 Then Err.Raise 5, , "Layout leaves no space for children."
                    rollback.Add Array(change(0), "box", PptNativeDriver.ShapeBox(shape))
                Else
                    If CDbl(change(2)) < 0 Then Err.Raise 5, , "Linked radius must not be negative."
                    rollback.Add Array(change(0), "fraction", PptNativeDriver.ReadFraction(shape))
                    rollback.Add Array(change(0), "tag", LOCK_KEY, PptNativeDriver.ReadTag(shape, LOCK_KEY), Not PptNativeDriver.HasTag(shape, LOCK_KEY))
                End If
            Case "tag"
                If StrComp(CStr(change(2)), STRICT_KEY, vbTextCompare) = 0 Or StrComp(CStr(change(2)), LOCK_KEY, vbTextCompare) = 0 Then Err.Raise 5, , "Layout metadata cannot change radius protection."
                rollback.Add Array(change(0), "tag", change(2), PptNativeDriver.ReadTag(shape, CStr(change(2))), Not PptNativeDriver.HasTag(shape, CStr(change(2))))
            Case Else: Err.Raise 5, , "Unsupported shape-plan operation."
        End Select
    Next change
    Set nodes = MetadataRoots(slide, plan)
    CheckSceneNodes nodes, spatial
    On Error GoTo Failed
    For Each node In nodes
        WriteSceneNode node, slide, plan, 0
    Next node
    If sameSlide Then RestoreSceneSelection slide, nodes, ids
    Exit Sub
Failed:
    errorNumber = Err.Number
    errorText = Err.Description
    Debug.Print "[NativeScene] " & errorText
    On Error Resume Next
    Err.Clear
    Set nodes = MetadataRoots(slide, rollback)
    If Err.Number = 0 Then
        For Each node In nodes
            WriteSceneNode node, slide, rollback, 0
            If Err.Number <> 0 Then Exit For
        Next node
    End If
    If Err.Number <> 0 Then recoveryText = " Shape recovery failed: " & Err.Description
    Err.Clear
    If sameSlide Then RestoreSceneSelection slide, nodes, ids
    If Err.Number <> 0 Then recoveryText = recoveryText & " Selection recovery failed: " & Err.Description
    On Error GoTo 0
    Err.Raise errorNumber, "RadiusNativeCore", errorText & recoveryText
End Sub

Public Sub SetProtection(ByVal enabled As Boolean)
    If enabled Then
        ApplySelection 0, "cm", "protect"
    Else
        ApplySelection 0, "cm", "unprotect"
    End If
End Sub

Public Sub SelfTest()
    If Abs(ParseNumber("0,30") - 0.3) > 0.000001 Then Err.Raise 5, , "Decimal parsing failed."
    If ComputeTarget(0, "cm", 4) <> 0 Then Err.Raise 5, , "Zero radius failed."
    If ComputeTarget(3, "cm", 4) <> 2 Then Err.Raise 5, , "Radius clamp failed."
    If ComputeTarget(20, "%", 4) <> 0.8 Then Err.Raise 5, , "Percent conversion failed."
    If Abs(StepValue(0.3, "cm", 1) - 0.4) > 0.000001 Then Err.Raise 5, , "Up step failed."
    If Abs(StepValue(0.3, "cm", -1) - 0.2) > 0.000001 Then Err.Raise 5, , "Down step failed."
    If StepValue(0.05, "cm", -1) <> 0 Then Err.Raise 5, , "Step lower bound failed."
    If StepValue(50, "%", 1) <> 50 Then Err.Raise 5, , "Percent step upper bound failed."
    If Abs(StepValue(0.05, "cm", 1) - 0.15) > 0.000001 Then Err.Raise 5, , "Step precision failed."
    If StepValue(1000000, "cm", 1) <> 1000000 Then Err.Raise 5, , "Step upper bound failed."
End Sub
