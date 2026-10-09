Attribute VB_Name = "RadiusNativeCore"
Option Explicit

Public Const PT_PER_CM As Double = 28.3464566929134
Private Const LOCK_KEY As String = "radiusLock_v1"
Private Const STRICT_KEY As String = "radiusLockStrict_v1"
Private Const GROUP_INITIAL As Long = 0
Private Const GROUP_OPEN As Long = 1
Private Const GROUP_REGROUPED As Long = 2
Private Const GROUP_COMPLETE As Long = 3
Private transactionActive As Boolean

Public Function IsTransactionActive() As Boolean
    IsTransactionActive = transactionActive
End Function

Private Sub BeginTransaction()
    If transactionActive Then Err.Raise 5, , "A native transaction is already in progress."
    transactionActive = True
End Sub

Private Sub EndTransaction()
    transactionActive = False
End Sub

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
    Set roots = PptNativeDriver.SelectionObjects()
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
    Set roots = PptNativeDriver.SelectionObjects()
    Set slide = PptNativeDriver.CurrentSlide()
    editable = Not PptNativeDriver.HasChildShapeSelection()
    For Each shape In roots
        If Not PptNativeDriver.IsTopLevel(slide, shape) Then editable = False
        CollectLeaves shape, leaves, 0
    Next shape
    count = leaves.Count
    For Each shape In leaves
        If PptNativeDriver.ReadTag(shape, STRICT_KEY) = "1" Then protectedCount = protectedCount + 1
    Next shape
End Sub

Private Function ObjectForKey(ByVal items As Collection, ByVal key As String) As Object
    Dim errorNumber As Long, errorText As String
    On Error GoTo Missing
    Set ObjectForKey = items(key)
    Exit Function
Missing:
    errorNumber = Err.Number
    errorText = Err.Description
    Err.Clear
    If errorNumber <> 5 Then Err.Raise errorNumber, "RadiusNativeCore", errorText
End Function

Private Function TargetKey(ByVal id As Long) As String
    TargetKey = "S" & CStr(id)
End Function

Private Function PlanBuckets(ByVal plan As Collection) As Collection
    Dim result As New Collection, bucket As Collection, change As Variant, key As String
    For Each change In plan
        key = TargetKey(CLng(change(0)))
        Set bucket = ObjectForKey(result, key)
        If bucket Is Nothing Then
            Set bucket = New Collection
            result.Add bucket, key
        End If
        bucket.Add change
    Next change
    Set PlanBuckets = result
End Function

Private Function ShapeIndex(ByVal shapes As Collection) As Collection
    Dim result As New Collection, shape As Object
    For Each shape In shapes
        result.Add shape, TargetKey(PptNativeDriver.ShapeId(shape))
    Next shape
    Set ShapeIndex = result
End Function

Private Sub SetNodeValue(ByVal node As Collection, ByVal key As String, ByVal value As Variant)
    node.Remove key
    node.Add value, key
End Sub

Private Sub SetNodeObject(ByVal node As Collection, ByVal key As String, ByVal value As Object)
    node.Remove key
    node.Add value, key
End Sub

Private Sub SetNodeShape(ByVal node As Collection, ByVal shape As Object)
    SetNodeObject node, "Shape", shape
    SetNodeValue node, "CurrentId", PptNativeDriver.ShapeId(shape)
End Sub

Private Function TransactionNode(ByVal shape As Object, ByVal buckets As Collection, ByVal depth As Long) As Collection
    Dim node As New Collection, children As New Collection, tags As Collection, child As Object, prepared As Collection, bucket As Collection
    Dim id As Long, eligible As Long, grouped As Boolean, returnedRange As Object
    If depth > 64 Then Err.Raise 5, , "Group nesting is too deep."
    id = PptNativeDriver.ShapeId(shape)
    grouped = (PptNativeDriver.ShapeType(shape) = 6)
    node.Add shape, "Shape"
    node.Add id, "Id"
    node.Add id, "CurrentId"
    node.Add PptNativeDriver.ShapeName(shape), "Name"
    node.Add grouped, "Group"
    node.Add GROUP_INITIAL, "Phase"
    node.Add False, "Mapped"
    node.Add False, "Touched"
    node.Add False, "StrictAttempted"
    node.Add False, "LockAttempted"
    node.Add returnedRange, "OpenRange"
    If grouped Then
        For Each child In PptNativeDriver.Children(shape)
            Set prepared = TransactionNode(child, buckets, depth + 1)
            eligible = eligible + CLng(prepared("Eligible"))
            children.Add prepared
        Next child
    Else
        Set bucket = ObjectForKey(buckets, TargetKey(id))
        If Not bucket Is Nothing Then eligible = 1
    End If
    Set tags = New Collection
    If grouped And eligible > 0 Then Set tags = PptNativeDriver.SnapshotTags(shape)
    node.Add tags, "Tags"
    node.Add children, "Children"
    node.Add eligible, "Eligible"
    Set TransactionNode = node
End Function

Private Function TransactionRoots(ByVal roots As Collection, ByVal buckets As Collection, ByVal eligibleOnly As Boolean) As Collection
    Dim result As New Collection, shape As Object, node As Collection
    For Each shape In roots
        Set node = TransactionNode(shape, buckets, 0)
        If Not eligibleOnly Or CLng(node("Eligible")) > 0 Then result.Add node
    Next shape
    Set TransactionRoots = result
End Function

Private Function NodeShapes(ByVal nodes As Collection) As Collection
    Dim result As New Collection, node As Collection
    For Each node In nodes
        result.Add node("Shape")
    Next node
    Set NodeShapes = result
End Function

Private Sub MapFreshChildren(ByVal node As Collection, ByVal buckets As Collection, ByVal depth As Long)
    Dim previous As New Collection, children As New Collection, member As Object, saved As Collection, child As Collection, members As Collection
    Dim eligible As Long, id As Long
    ' These IDs come from the actual direct tree from the previous ungroup,
    ' or from read-only preflight. They never determine the new hierarchy.
    For Each child In node("Children")
        previous.Add child, TargetKey(CLng(child("CurrentId")))
    Next child
    Set members = PptNativeDriver.RangeMembers(node("OpenRange"))
    For Each member In members
        id = PptNativeDriver.ShapeId(member)
        Set saved = ObjectForKey(previous, TargetKey(id))
        If saved Is Nothing Then
            Set saved = TransactionNode(member, buckets, depth + 1)
        Else
            SetNodeShape saved, member
        End If
        eligible = eligible + CLng(saved("Eligible"))
        children.Add saved
    Next member
    SetNodeObject node, "Children", children
    SetNodeValue node, "Mapped", True
    If eligible <> CLng(node("Eligible")) Then Err.Raise 5, , "Cannot resolve all transaction targets after ungroup."
End Sub

Private Sub CompleteGroupNode(ByVal node As Collection, ByVal slide As Object)
    Dim shape As Object, members As Collection
    If CLng(node("Phase")) = GROUP_OPEN Then
        If CBool(node("Mapped")) Then
            Set members = NodeShapes(node("Children"))
        Else
            Set members = PptNativeDriver.RangeMembers(node("OpenRange"))
        End If
        Set shape = PptNativeDriver.Regroup(slide, members)
        ' Keep the original ID/name/tags and only replace the current object.
        SetNodeObject node, "Shape", shape
        SetNodeValue node, "Phase", GROUP_REGROUPED
    End If
    If CLng(node("Phase")) = GROUP_REGROUPED Then
        SetNodeValue node, "CurrentId", PptNativeDriver.ShapeId(node("Shape"))
        PptNativeDriver.RestoreMetadata node("Shape"), CStr(node("Name")), node("Tags")
        SetNodeValue node, "Phase", GROUP_COMPLETE
    End If
End Sub

Private Function TryCompleteGroup(ByVal node As Collection, ByVal slide As Object) As String
    On Error GoTo Failed
    CompleteGroupNode node, slide
    Exit Function
Failed:
    TryCompleteGroup = " Group recovery failed: " & Err.Description
    Debug.Print "[NativeGroupRecovery] " & Err.Description
End Function

Private Function NeedsRecovery(ByVal node As Collection) As Boolean
    Dim phase As Long
    If CBool(node("Group")) Then
        phase = CLng(node("Phase"))
        If phase = GROUP_OPEN Or phase = GROUP_REGROUPED Then NeedsRecovery = True: Exit Function
        NeedsRecovery = ChildrenNeedRecovery(node)
    Else
        NeedsRecovery = CBool(node("Touched"))
    End If
End Function

Private Function ChildrenNeedRecovery(ByVal node As Collection) As Boolean
    Dim child As Collection
    For Each child In node("Children")
        If NeedsRecovery(child) Then ChildrenNeedRecovery = True: Exit Function
    Next child
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

Private Sub WriteRadiusLeaf(ByVal node As Collection, ByVal value As Double, ByVal unit As String)
    Dim shape As Object, actualCm As Double, shortPoints As Double, hasLock As Boolean
    Set shape = node("Shape")
    shortPoints = PptNativeDriver.ShortSidePoints(shape)
    actualCm = ComputeTarget(value, unit, shortPoints / PT_PER_CM)
    hasLock = (PptNativeDriver.ReadTag(shape, LOCK_KEY) <> "")
    ' The final live protection read is immediately before adjustment write.
    If PptNativeDriver.ReadTag(shape, STRICT_KEY) = "1" Then Err.Raise 5, , "Protected shape; radius was not written."
    SetNodeValue node, "Touched", True
    PptNativeDriver.WriteFraction shape, actualCm * PT_PER_CM / shortPoints
    If hasLock Then PptNativeDriver.AddTag shape, LOCK_KEY, NumberText(actualCm)
End Sub

Private Sub WriteSelectionLeaf(ByVal node As Collection, ByVal bucket As Collection, ByVal action As String)
    Dim shape As Object, change As Variant, radius As Double
    Set shape = node("Shape")
    change = bucket(1)
    Select Case action
        Case "radius": WriteRadiusLeaf node, CDbl(change(2)), CStr(change(3))
        Case "protect"
            If PptNativeDriver.ReadTag(shape, LOCK_KEY) = "" Then
                radius = CurrentRadius(shape)
                SetNodeValue node, "Touched", True
                SetNodeValue node, "LockAttempted", True
                PptNativeDriver.AddTag shape, LOCK_KEY, NumberText(radius)
            End If
            SetNodeValue node, "Touched", True
            SetNodeValue node, "StrictAttempted", True
            PptNativeDriver.AddTag shape, STRICT_KEY, "1"
        Case "unprotect"
            SetNodeValue node, "Touched", True
            SetNodeValue node, "StrictAttempted", True
            PptNativeDriver.DeleteTag shape, STRICT_KEY
        Case Else: Err.Raise 5, , "Unsupported edit operation."
    End Select
End Sub

Private Sub WriteMetadataLeaf(ByVal node As Collection, ByVal bucket As Collection)
    Dim shape As Object, change As Variant
    Set shape = node("Shape")
    For Each change In bucket
        If StrComp(CStr(change(1)), STRICT_KEY, vbTextCompare) = 0 Or StrComp(CStr(change(1)), LOCK_KEY, vbTextCompare) = 0 Then Err.Raise 5, , "Relationship edits cannot change radius protection."
        SetNodeValue node, "Touched", True
        If CBool(change(3)) Then
            PptNativeDriver.DeleteTag shape, CStr(change(1))
        Else
            PptNativeDriver.AddTag shape, CStr(change(1)), CStr(change(2))
        End If
    Next change
End Sub

Private Sub WriteSceneLeaf(ByVal node As Collection, ByVal bucket As Collection, ByVal recovering As Boolean)
    Dim shape As Object, change As Variant
    Set shape = node("Shape")
    ' Restore/resize the box before radius/fraction, using the fresh object.
    For Each change In bucket
        If CStr(change(1)) = "box" Then
            If PptNativeDriver.ReadTag(shape, STRICT_KEY) = "1" Then Err.Raise 5, , "Protected child; layout was not written."
            SetNodeValue node, "Touched", True
            PptNativeDriver.SetShapeBox shape, change(2)
        End If
    Next change
    For Each change In bucket
        Select Case CStr(change(1))
            Case "radius": WriteRadiusLeaf node, CDbl(change(2)), "cm"
            Case "fraction"
                If Not recovering Then Err.Raise 5, , "Fraction writes are reserved for transaction recovery."
                If PptNativeDriver.ReadTag(shape, STRICT_KEY) = "1" Then Err.Raise 5, , "New protection retained; radius recovery was not written."
                PptNativeDriver.WriteFraction shape, CDbl(change(2))
            Case "tag"
                If StrComp(CStr(change(2)), STRICT_KEY, vbTextCompare) = 0 Then Err.Raise 5, , "Layout/recovery cannot change strict protection."
                If Not recovering And StrComp(CStr(change(2)), LOCK_KEY, vbTextCompare) = 0 Then Err.Raise 5, , "Layout metadata cannot change radius protection."
                SetNodeValue node, "Touched", True
                If CBool(change(4)) Then
                    PptNativeDriver.DeleteTag shape, CStr(change(2))
                Else
                    PptNativeDriver.AddTag shape, CStr(change(2)), CStr(change(3))
                End If
            Case "box"
            Case Else: Err.Raise 5, , "Unsupported shape-plan operation."
        End Select
    Next change
End Sub

Private Sub RestoreExplicitProtection(ByVal node As Collection, ByVal bucket As Collection, ByVal action As String)
    Dim shape As Object, saved As Variant, value As String, present As Boolean, original As Boolean, expected As Boolean
    Set shape = node("Shape")
    saved = bucket(1)
    If action <> "protect" And action <> "unprotect" Then Err.Raise 5, , "Only an explicit protection operation may restore protection."
    If CBool(node("StrictAttempted")) Then
        value = PptNativeDriver.ReadTagState(shape, STRICT_KEY, present)
        original = (present = CBool(saved(5)) And value = CStr(saved(4)))
        If action = "protect" Then expected = (present And value = "1") Else expected = Not present
        If Not original Then
            If Not expected Then Err.Raise 5, , "New protection retained; explicit protection recovery could not restore the original tag."
            ' This only undoes the strict write/delete of the user's explicit
            ' protect/unprotect action. Radius/layout recovery never calls it.
            If CBool(saved(5)) Then PptNativeDriver.AddTag shape, STRICT_KEY, CStr(saved(4)) Else PptNativeDriver.DeleteTag shape, STRICT_KEY
        End If
    End If
    If action = "protect" And CBool(node("LockAttempted")) Then
        If CBool(saved(3)) Then PptNativeDriver.AddTag shape, LOCK_KEY, CStr(saved(2)) Else PptNativeDriver.DeleteTag shape, LOCK_KEY
    End If
End Sub

Private Function TryRestoreChange(ByVal node As Collection, ByVal change As Variant, ByVal metadata As Boolean) As String
    Dim singleChange As New Collection
    On Error GoTo Failed
    singleChange.Add change
    If metadata Then WriteMetadataLeaf node, singleChange Else WriteSceneLeaf node, singleChange, True
    Exit Function
Failed:
    TryRestoreChange = " Operation recovery failed: " & Err.Description
End Function

Private Sub RestoreMetadataLeaf(ByVal node As Collection, ByVal bucket As Collection)
    Dim change As Variant, errors As String
    For Each change In bucket
        errors = errors & TryRestoreChange(node, change, True)
    Next change
    If errors <> "" Then Err.Raise 5, , errors
End Sub

Private Sub RestoreSceneLeaf(ByVal node As Collection, ByVal bucket As Collection)
    Dim change As Variant, errors As String, failure As String, geometryRestored As Boolean
    geometryRestored = True
    For Each change In bucket
        If CStr(change(1)) = "box" Then
            failure = TryRestoreChange(node, change, False)
            errors = errors & failure
            If failure <> "" Then geometryRestored = False
        End If
    Next change
    For Each change In bucket
        If CStr(change(1)) = "fraction" Then
            If geometryRestored Then
                failure = TryRestoreChange(node, change, False)
                errors = errors & failure
                If failure <> "" Then geometryRestored = False
            Else
                errors = errors & " Radius recovery skipped after geometry recovery failed."
            End If
        End If
    Next change
    For Each change In bucket
        If CStr(change(1)) = "tag" Then
            If StrComp(CStr(change(2)), LOCK_KEY, vbTextCompare) = 0 And Not geometryRestored Then
                errors = errors & " Fixed radius recovery skipped after geometry/radius recovery failed."
            Else
                errors = errors & TryRestoreChange(node, change, False)
            End If
        End If
    Next change
    If errors <> "" Then Err.Raise 5, , errors
End Sub

Private Function TryRecoverNode(ByVal node As Collection, ByVal slide As Object, ByVal buckets As Collection, ByVal mode As String, ByVal depth As Long) As String
    On Error GoTo Failed
    WriteTransactionNode node, slide, buckets, mode, depth, True
    Exit Function
Failed:
    TryRecoverNode = " Target/group " & CStr(node("Id")) & " recovery failed: " & Err.Description
    Debug.Print "[NativeRecovery] " & Err.Description
End Function

Private Sub WriteTransactionNode(ByVal node As Collection, ByVal slide As Object, ByVal buckets As Collection, ByVal mode As String, ByVal depth As Long, Optional ByVal recovering As Boolean = False)
    Dim shape As Object, returnedRange As Object, child As Collection, bucket As Collection
    Dim errorNumber As Long, errorText As String, recoveryText As String, childErrors As String, opened As Boolean, completing As Boolean
    If CLng(node("Eligible")) = 0 Then Exit Sub
    If recovering Then
        If Not NeedsRecovery(node) Then Exit Sub
    End If
    If depth > 64 Then Err.Raise 5, , "Group nesting is too deep."
    Set shape = node("Shape")
    On Error GoTo Failed
    If CBool(node("Group")) Then
        If recovering Then
            If CLng(node("Phase")) = GROUP_REGROUPED Then
                If Not ChildrenNeedRecovery(node) Then
                    completing = True
                    CompleteGroupNode node, slide
                    Exit Sub
                End If
            End If
        End If
        If CLng(node("Phase")) <> GROUP_OPEN Then
            SetNodeValue node, "Mapped", False
            PptNativeDriver.UngroupForTransaction shape, returnedRange
            opened = True
            SetNodeObject node, "OpenRange", returnedRange
            SetNodeValue node, "Phase", GROUP_OPEN
        End If
        If Not CBool(node("Mapped")) Then MapFreshChildren node, buckets, depth
        For Each child In node("Children")
            If recovering Then
                childErrors = childErrors & TryRecoverNode(child, slide, buckets, mode, depth + 1)
            Else
                WriteTransactionNode child, slide, buckets, mode, depth + 1
            End If
        Next child
        completing = True
        CompleteGroupNode node, slide
        If childErrors <> "" Then Err.Raise 5, , childErrors
    Else
        Set bucket = ObjectForKey(buckets, TargetKey(CLng(node("CurrentId"))))
        If bucket Is Nothing Then Err.Raise 5, , "Cannot resolve transaction operations for a fresh shape."
        If recovering Then
            Select Case mode
                Case "protect", "unprotect": RestoreExplicitProtection node, bucket, mode
                Case "metadata": RestoreMetadataLeaf node, bucket
                Case Else: RestoreSceneLeaf node, bucket
            End Select
            SetNodeValue node, "Touched", False
        Else
            Select Case mode
                Case "radius", "protect", "unprotect": WriteSelectionLeaf node, bucket, mode
                Case "metadata": WriteMetadataLeaf node, bucket
                Case "scene": WriteSceneLeaf node, bucket, False
                Case Else: Err.Raise 5, , "Unsupported transaction operation."
            End Select
        End If
    End If
    Exit Sub
Failed:
    errorNumber = Err.Number
    errorText = Err.Description
    If childErrors <> "" And errorText <> childErrors Then errorText = errorText & childErrors
    Debug.Print "[NativeWrite] " & errorText
    If CBool(node("Group")) Then
        If opened And Not completing Then
            SetNodeObject node, "OpenRange", returnedRange
            SetNodeValue node, "Phase", GROUP_OPEN
        End If
        recoveryText = TryCompleteGroup(node, slide)
    End If
    Err.Raise errorNumber, "RadiusNativeCore", errorText & recoveryText
End Sub

Private Function RecoveryText(ByVal nodes As Collection, ByVal slide As Object, ByVal rollback As Collection, ByVal mode As String) As String
    Dim buckets As Collection, node As Collection
    On Error GoTo Failed
    Set buckets = PlanBuckets(rollback)
    For Each node In nodes
        RecoveryText = RecoveryText & TryRecoverNode(node, slide, buckets, mode, 0)
    Next node
    Exit Function
Failed:
    RecoveryText = RecoveryText & " Recovery preparation failed: " & Err.Description
    Debug.Print "[NativeRecovery] " & Err.Description
End Function

Private Function CaptureSelection(ByVal slide As Object) As Collection
    Dim result As New Collection, shape As Object, saved As Collection
    If PptNativeDriver.IsCurrentSlide(slide) And PptNativeDriver.HasShapeSelection Then
        For Each shape In PptNativeDriver.SelectionObjects()
            Set saved = New Collection
            saved.Add PptNativeDriver.ShapeId(shape), "Id"
            saved.Add shape, "Shape"
            result.Add saved
        Next shape
    End If
    Set CaptureSelection = result
End Function

Private Sub FindNodeSelection(ByVal nodes As Collection, ByVal id As Long, ByVal result As Collection)
    Dim node As Collection
    For Each node In nodes
        If CLng(node("Id")) = id Then result.Add node("Shape"): Exit Sub
        FindNodeSelection node("Children"), id, result
    Next node
End Sub

Private Sub RestoreSceneSelection(ByVal slide As Object, ByVal nodes As Collection, ByVal selection As Collection)
    Dim shapes As New Collection, saved As Collection, before As Long, shape As Object
    For Each saved In selection
        before = shapes.Count
        FindNodeSelection nodes, CLng(saved("Id")), shapes
        If shapes.Count = before Then
            ' Unaffected objects keep their original live reference. Changed
            ' groups are resolved above through original ID -> current Shape.
            Set shape = saved("Shape")
            If PptNativeDriver.ShapeId(shape) <> CLng(saved("Id")) Then Err.Raise 5, , "Cannot restore the original selected shape."
            shapes.Add shape
        End If
    Next saved
    If shapes.Count > 0 Then PptNativeDriver.SelectObjects shapes Else PptNativeDriver.ClearSelection
End Sub

Private Function SelectionRecoveryText(ByVal slide As Object, ByVal nodes As Collection, ByVal selection As Collection) As String
    On Error GoTo Failed
    RestoreSceneSelection slide, nodes, selection
    Exit Function
Failed:
    SelectionRecoveryText = " Selection recovery failed: " & Err.Description
    Debug.Print "[NativeSelectionRecovery] " & Err.Description
End Function

Private Sub CaptureTagRollback(ByVal rollback As Collection, ByVal shape As Object, ByVal id As Long, ByVal key As String, ByVal metadata As Boolean)
    Dim present As Boolean, value As String
    value = PptNativeDriver.ReadTagState(shape, key, present)
    If metadata Then rollback.Add Array(id, key, value, Not present) Else rollback.Add Array(id, "tag", key, value, Not present)
End Sub

Private Sub ApplySelectionInternal(ByVal value As Double, ByVal unit As String, ByVal action As String)
    Dim roots As Collection, leaves As New Collection, nodes As Collection, node As Collection, shape As Object, slide As Object, linkedPlan As Collection
    Dim plan As New Collection, rollback As New Collection, buckets As Collection, selection As Collection
    Dim id As Long, errorNumber As Long, errorText As String, recovery As String, lockValue As String, strictValue As String, hasLock As Boolean, hasStrict As Boolean, target As Double
    If RadiusNativeRelations.IsPreviewContext Then Err.Raise 5, , "Close the relation preview before editing the source."
    If action <> "radius" And action <> "protect" And action <> "unprotect" Then Err.Raise 5, , "Unsupported edit operation."
    If PptNativeDriver.HasChildShapeSelection Then Err.Raise 5, , "Select the complete top-level group for native editing."
    Set roots = PptNativeDriver.SelectionRoots()
    Set slide = PptNativeDriver.CurrentSlide()
    For Each shape In roots
        If Not PptNativeDriver.IsTopLevel(slide, shape) Then Err.Raise 5, , "Select the complete top-level group for native editing."
    Next shape
    If action = "radius" Then
        Set linkedPlan = RadiusNativeLayout.LinkedRadiusPlan(roots, value, unit)
        If Not linkedPlan Is Nothing Then ApplyShapePlanInternal slide, linkedPlan: Exit Sub
    End If
    For Each shape In roots
        CollectLeaves shape, leaves, 0
    Next shape
    If leaves.Count = 0 Then Err.Raise 5, , "Select rounded rectangles."
    ' Capture every old value and validate every target before any mutation.
    For Each shape In leaves
        id = PptNativeDriver.ShapeId(shape)
        If action = "radius" Then
            If PptNativeDriver.ReadTag(shape, STRICT_KEY) = "1" Then Err.Raise 5, , "Turn off protection explicitly before applying radius."
            target = ComputeTarget(value, unit, PptNativeDriver.ShortSidePoints(shape) / PT_PER_CM)
            rollback.Add Array(id, "fraction", PptNativeDriver.ReadFraction(shape))
            CaptureTagRollback rollback, shape, id, LOCK_KEY, False
        Else
            lockValue = PptNativeDriver.ReadTagState(shape, LOCK_KEY, hasLock)
            strictValue = PptNativeDriver.ReadTagState(shape, STRICT_KEY, hasStrict)
            If action = "protect" And lockValue = "" Then target = CurrentRadius(shape)
            rollback.Add Array(id, "protection", lockValue, hasLock, strictValue, hasStrict)
        End If
        plan.Add Array(id, action, value, unit)
    Next shape
    Set buckets = PlanBuckets(plan)
    Set nodes = TransactionRoots(roots, buckets, False)
    Set selection = CaptureSelection(slide)
    On Error GoTo Failed
    For Each node In nodes
        WriteTransactionNode node, slide, buckets, action, 0
    Next node
    PptNativeDriver.SelectShapes slide, NodeShapes(nodes)
    Exit Sub
Failed:
    errorNumber = Err.Number
    errorText = Err.Description
    recovery = RecoveryText(nodes, slide, rollback, action)
    recovery = recovery & SelectionRecoveryText(slide, nodes, selection)
    Err.Raise errorNumber, "RadiusNativeCore", errorText & recovery
End Sub

Public Sub ApplySelection(ByVal value As Double, ByVal unit As String, Optional ByVal action As String = "radius")
    Dim errorNumber As Long, errorText As String
    BeginTransaction
    On Error GoTo Failed
    ApplySelectionInternal value, unit, action
    EndTransaction
    Exit Sub
Failed:
    errorNumber = Err.Number
    errorText = Err.Description
    EndTransaction
    Debug.Print "[NativeSelection] " & errorText
    Err.Raise errorNumber, "RadiusNativeCore", errorText
End Sub

Private Sub ApplyMetadataPlanInternal(ByVal slide As Object, ByVal plan As Collection)
    Dim leaves As Collection, shapes As Collection, shape As Object, nodes As Collection, node As Collection, change As Variant
    Dim rollback As New Collection, buckets As Collection, selection As Collection, sameSlide As Boolean, errorNumber As Long, errorText As String, recovery As String
    If RadiusNativeRelations.IsPreviewContext Then Err.Raise 5, , "Close the relation preview before editing the source."
    If plan.Count = 0 Then Exit Sub
    Set leaves = RadiusNativeRelations.SlideLeaves(slide)
    Set shapes = ShapeIndex(leaves)
    For Each change In plan
        If StrComp(CStr(change(1)), STRICT_KEY, vbTextCompare) = 0 Or StrComp(CStr(change(1)), LOCK_KEY, vbTextCompare) = 0 Then Err.Raise 5, , "Relationship edits cannot change radius protection."
        Set shape = ObjectForKey(shapes, TargetKey(CLng(change(0))))
        If shape Is Nothing Then Err.Raise 5, , "Cannot uniquely resolve a relationship target."
        CaptureTagRollback rollback, shape, CLng(change(0)), CStr(change(1)), True
    Next change
    Set buckets = PlanBuckets(plan)
    Set nodes = TransactionRoots(PptNativeDriver.SlideRoots(slide), buckets, True)
    sameSlide = PptNativeDriver.IsCurrentSlide(slide)
    Set selection = CaptureSelection(slide)
    On Error GoTo Failed
    For Each node In nodes
        WriteTransactionNode node, slide, buckets, "metadata", 0
    Next node
    If sameSlide Then PptNativeDriver.SelectShapes slide, NodeShapes(nodes)
    Exit Sub
Failed:
    errorNumber = Err.Number
    errorText = Err.Description
    recovery = RecoveryText(nodes, slide, rollback, "metadata")
    If sameSlide Then recovery = recovery & SelectionRecoveryText(slide, nodes, selection)
    Err.Raise errorNumber, "RadiusNativeCore", errorText & recovery
End Sub

Public Sub ApplyMetadataPlan(ByVal slide As Object, ByVal plan As Collection)
    Dim errorNumber As Long, errorText As String
    BeginTransaction
    On Error GoTo Failed
    ApplyMetadataPlanInternal slide, plan
    EndTransaction
    Exit Sub
Failed:
    errorNumber = Err.Number
    errorText = Err.Description
    EndTransaction
    Debug.Print "[NativeMetadata] " & errorText
    Err.Raise errorNumber, "RadiusNativeCore", errorText
End Sub

Public Sub PutRadius(ByVal plan As Collection, ByVal id As Long, ByVal value As Double)
    Dim i As Long, change As Variant
    For i = plan.Count To 1 Step -1
        change = plan(i)
        If CLng(change(0)) = id And CStr(change(1)) = "radius" Then plan.Remove i
    Next i
    plan.Add Array(id, "radius", value)
End Sub

Private Sub ApplyShapePlanInternal(ByVal slide As Object, ByVal plan As Collection)
    Dim leaves As Collection, shapes As Collection, shape As Object, nodes As Collection, node As Collection, change As Variant, targetBucket As Collection, item As Variant
    Dim rollback As New Collection, buckets As Collection, selection As Collection, spatial As Boolean, sameSlide As Boolean
    Dim errorNumber As Long, errorText As String, recovery As String, shortPoints As Double, box As Variant, target As Double
    If RadiusNativeRelations.IsPreviewContext Then Err.Raise 5, , "Close the relation preview before editing the source."
    If plan.Count = 0 Then Exit Sub
    Set leaves = RadiusNativeRelations.SlideLeaves(slide)
    Set shapes = ShapeIndex(leaves)
    Set buckets = PlanBuckets(plan)
    For Each change In plan
        Set shape = ObjectForKey(shapes, TargetKey(CLng(change(0))))
        If shape Is Nothing Then Err.Raise 5, , "Cannot resolve a layout/radius target."
        Select Case CStr(change(1))
            Case "box", "radius"
                spatial = True
                If PptNativeDriver.ReadTag(shape, STRICT_KEY) = "1" Then Err.Raise 5, , "Turn off protection explicitly before layout or linked radius: " & PptNativeDriver.ShapeName(shape)
                If CStr(change(1)) = "box" Then
                    If CDbl(change(2)(2)) <= 0 Or CDbl(change(2)(3)) <= 0 Then Err.Raise 5, , "Layout leaves no space for children."
                    rollback.Add Array(change(0), "box", PptNativeDriver.ShapeBox(shape))
                Else
                    shortPoints = PptNativeDriver.ShortSidePoints(shape)
                    Set targetBucket = ObjectForKey(buckets, TargetKey(CLng(change(0))))
                    For Each item In targetBucket
                        If CStr(item(1)) = "box" Then
                            box = item(2)
                            shortPoints = CDbl(box(2))
                            If CDbl(box(3)) < shortPoints Then shortPoints = CDbl(box(3))
                        End If
                    Next item
                    target = ComputeTarget(CDbl(change(2)), "cm", shortPoints / PT_PER_CM)
                    rollback.Add Array(change(0), "fraction", PptNativeDriver.ReadFraction(shape))
                    CaptureTagRollback rollback, shape, CLng(change(0)), LOCK_KEY, False
                End If
            Case "tag"
                If StrComp(CStr(change(2)), STRICT_KEY, vbTextCompare) = 0 Or StrComp(CStr(change(2)), LOCK_KEY, vbTextCompare) = 0 Then Err.Raise 5, , "Layout metadata cannot change radius protection."
                CaptureTagRollback rollback, shape, CLng(change(0)), CStr(change(2)), False
            Case Else: Err.Raise 5, , "Unsupported shape-plan operation."
        End Select
    Next change
    Set nodes = TransactionRoots(PptNativeDriver.SlideRoots(slide), buckets, True)
    CheckSceneNodes nodes, spatial
    sameSlide = PptNativeDriver.IsCurrentSlide(slide)
    Set selection = CaptureSelection(slide)
    On Error GoTo Failed
    For Each node In nodes
        WriteTransactionNode node, slide, buckets, "scene", 0
    Next node
    If sameSlide Then RestoreSceneSelection slide, nodes, selection
    Exit Sub
Failed:
    errorNumber = Err.Number
    errorText = Err.Description
    recovery = RecoveryText(nodes, slide, rollback, "scene")
    If sameSlide Then recovery = recovery & SelectionRecoveryText(slide, nodes, selection)
    Err.Raise errorNumber, "RadiusNativeCore", errorText & recovery
End Sub

Public Sub ApplyShapePlan(ByVal slide As Object, ByVal plan As Collection)
    Dim errorNumber As Long, errorText As String
    BeginTransaction
    On Error GoTo Failed
    ApplyShapePlanInternal slide, plan
    EndTransaction
    Exit Sub
Failed:
    errorNumber = Err.Number
    errorText = Err.Description
    EndTransaction
    Debug.Print "[NativeScene] " & errorText
    Err.Raise errorNumber, "RadiusNativeCore", errorText
End Sub

Public Sub SetProtection(ByVal enabled As Boolean)
    If enabled Then ApplySelection 0, "cm", "protect" Else ApplySelection 0, "cm", "unprotect"
End Sub

Public Function SelfTest() As Long
    Dim passed As Long
    If Abs(ParseNumber("0,30") - 0.3) > 0.000001 Then Err.Raise 5, , "Decimal parsing failed."
    passed = passed + 1
    If ComputeTarget(0, "cm", 4) <> 0 Then Err.Raise 5, , "Zero radius failed."
    passed = passed + 1
    If ComputeTarget(3, "cm", 4) <> 2 Then Err.Raise 5, , "Radius clamp failed."
    passed = passed + 1
    If ComputeTarget(20, "%", 4) <> 0.8 Then Err.Raise 5, , "Percent conversion failed."
    passed = passed + 1
    If Abs(StepValue(0.3, "cm", 1) - 0.4) > 0.000001 Then Err.Raise 5, , "Up step failed."
    passed = passed + 1
    If Abs(StepValue(0.3, "cm", -1) - 0.2) > 0.000001 Then Err.Raise 5, , "Down step failed."
    passed = passed + 1
    If StepValue(0.05, "cm", -1) <> 0 Then Err.Raise 5, , "Step lower bound failed."
    passed = passed + 1
    If StepValue(50, "%", 1) <> 50 Then Err.Raise 5, , "Percent step upper bound failed."
    passed = passed + 1
    If Abs(StepValue(0.05, "cm", 1) - 0.15) > 0.000001 Then Err.Raise 5, , "Step precision failed."
    passed = passed + 1
    If StepValue(1000000, "cm", 1) <> 1000000 Then Err.Raise 5, , "Step upper bound failed."
    passed = passed + 1
    SelfTest = passed
End Function
