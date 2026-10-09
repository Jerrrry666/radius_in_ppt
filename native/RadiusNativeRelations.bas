Attribute VB_Name = "RadiusNativeRelations"
Option Explicit

Private Const RELATION_KEY As String = "radiusRelation_v1"
Private Const ROLE_KEY As String = "radiusRelationRole_v1"
Private Const LEGACY_PARENT As String = "layoutParent_v1"
Private Const LEGACY_CHILD As String = "layoutChild_v1"
Private pendingPresentation As Object
Private pendingSlide As Long
Private pendingId As Long
Private previewPresentation As Object
Private previewSource As Object
Private menuPresentation As Object
Private menuSlide As Long
Private revision As Long

Public Function CaptureTestSession() As Variant
    Dim state(0 To 5) As Variant
    Set state(0) = pendingPresentation
    state(1) = pendingSlide
    state(2) = pendingId
    Set state(3) = menuPresentation
    state(4) = menuSlide
    state(5) = revision
    CaptureTestSession = state
End Function

Public Sub RestoreTestSession(ByVal state As Variant)
    If Not IsArray(state) Then Err.Raise 5, , "Invalid relationship test session."
    If LBound(state) <> 0 Or UBound(state) <> 5 Then Err.Raise 5, , "Invalid relationship test session."
    Set pendingPresentation = state(0)
    pendingSlide = CLng(state(1))
    pendingId = CLng(state(2))
    Set menuPresentation = state(3)
    menuSlide = CLng(state(4))
    revision = CLng(state(5))
End Sub

Public Sub SuspendTestSession()
    ' Preview ownership is never transferred to a test session.
    If IsPreviewOpen Then Err.Raise 5, , "Close the relation preview before running the quick check."
    CancelPending
    Set menuPresentation = Nothing
    menuSlide = 0
    revision = 0
End Sub

Public Function IsPreviewContext() As Boolean
    Dim active As Object
    If previewPresentation Is Nothing Then Exit Function
    If Not PptNativeDriver.HasPresentation Then Exit Function
    Set active = PptNativeDriver.CurrentPresentation()
    IsPreviewContext = (active Is previewPresentation)
End Function

Public Function IsPreviewOpen() As Boolean
    IsPreviewOpen = Not previewPresentation Is Nothing
End Function

Public Sub ForgetPresentation(ByVal presentation As Object)
    If Not pendingPresentation Is Nothing Then
        If presentation Is pendingPresentation Then CancelPending
    End If
    If Not previewPresentation Is Nothing Then
        If presentation Is previewPresentation Then Set previewPresentation = Nothing
    End If
    If Not previewSource Is Nothing Then
        If presentation Is previewSource Then Set previewSource = Nothing
    End If
    If Not menuPresentation Is Nothing Then
        If presentation Is menuPresentation Then Set menuPresentation = Nothing
    End If
End Sub

Private Function FindShape(ByVal leaves As Collection, ByVal id As Long) As Object
    Dim shape As Object
    For Each shape In leaves
        If PptNativeDriver.ShapeId(shape) = id Then
            Set FindShape = shape
            Exit Function
        End If
    Next shape
End Function

Private Sub CollectUnique(ByVal shape As Object, ByVal result As Collection, ByVal depth As Long)
    Dim child As Object, found As Object
    If depth > 64 Then Err.Raise 5, , "Group nesting is too deep."
    If PptNativeDriver.ShapeType(shape) = 6 Then
        For Each child In PptNativeDriver.Children(shape)
            CollectUnique child, result, depth + 1
        Next child
    ElseIf RadiusNativeCore.IsRoundRect(shape) Then
        Set found = FindShape(result, PptNativeDriver.ShapeId(shape))
        If found Is Nothing Then result.Add shape
    End If
End Sub

Public Function SlideLeaves(ByVal slide As Object) As Collection
    Dim result As New Collection, shape As Object
    For Each shape In PptNativeDriver.SlideRoots(slide)
        CollectUnique shape, result, 0
    Next shape
    Set SlideLeaves = result
End Function

Public Function RoundedLeaves(ByVal objects As Collection) As Collection
    Dim result As New Collection, shape As Object
    For Each shape In objects
        CollectUnique shape, result, 0
    Next shape
    Set RoundedLeaves = result
End Function

Private Function SelectedLeaves() As Collection
    Dim result As New Collection
    If PptNativeDriver.HasShapeSelection Then
        Set result = RoundedLeaves(PptNativeDriver.SelectionObjects())
    End If
    Set SelectedLeaves = result
End Function

Private Function PositiveInteger(ByVal value As String) As Long
    Dim i As Long, character As String
    If Len(value) = 0 Or Len(value) > 10 Then Err.Raise 5, , "Invalid relationship identifier."
    For i = 1 To Len(value)
        character = Mid$(value, i, 1)
        If character < "0" Or character > "9" Then Err.Raise 5, , "Invalid relationship identifier."
    Next i
    If CDbl(value) > 2147483647# Then Err.Raise 5, , "Relationship identifier is out of range."
    PositiveInteger = CLng(value)
    If PositiveInteger < 1 Then Err.Raise 5, , "Invalid relationship identifier."
End Function

Private Function GetGroup(ByVal groups As Collection, ByVal key As String) As Collection
    Dim group As Collection
    For Each group In groups
        If CStr(group("Key")) = key Then
            Set GetGroup = group
            Exit Function
        End If
    Next group
End Function

Private Function NewGroup(ByVal key As String, ByVal kind As String) As Collection
    Dim group As New Collection, children As New Collection
    group.Add key, "Key"
    group.Add kind, "Kind"
    group.Add 0&, "ParentId"
    group.Add "", "ParentName"
    group.Add children, "Children"
    group.Add "", "Problem"
    Set NewGroup = group
End Function

Private Sub SetProblem(ByVal group As Collection, ByVal text As String)
    group.Remove "Problem"
    group.Add text, "Problem"
End Sub

Private Sub AddParent(ByVal group As Collection, ByVal shape As Object)
    If CLng(group("ParentId")) > 0 Then Err.Raise 5, , "Duplicate relationship parents: " & CStr(group("Key"))
    group.Remove "ParentId"
    group.Add PptNativeDriver.ShapeId(shape), "ParentId"
    group.Add shape, "Parent"
    group.Remove "ParentName"
    group.Add PptNativeDriver.ShapeName(shape), "ParentName"
End Sub

Private Sub AddChild(ByVal group As Collection, ByVal shape As Object, ByVal ordinal As Long)
    Dim children As Collection, child As New Collection, existing As Collection, i As Long
    Set children = group("Children")
    For Each existing In children
        If CLng(existing("Order")) = ordinal Or PptNativeDriver.ShapeId(existing("Shape")) = PptNativeDriver.ShapeId(shape) Then Err.Raise 5, , "Duplicate relationship child: " & CStr(group("Key"))
    Next existing
    child.Add shape, "Shape"
    child.Add ordinal, "Order"
    For i = 1 To children.Count
        If CLng(children(i)("Order")) > ordinal Then
            children.Add child, Before:=i
            Exit Sub
        End If
    Next i
    children.Add child
End Sub

Private Function LegacyIds(ByVal value As String, ByRef first As Long, ByRef last As Long) As Collection
    Dim result As New Collection, keyPosition As Long, start As Long, finish As Long, items As Variant, item As Variant, token As String
    If Left$(Trim$(value), 1) <> "{" Or Right$(Trim$(value), 1) <> "}" Then Err.Raise 5, , "Invalid legacy layout JSON."
    keyPosition = InStr(1, value, Chr$(34) & "childIds" & Chr$(34), vbTextCompare)
    If keyPosition = 0 Then Err.Raise 5, , "Legacy layout has no childIds array."
    If InStr(keyPosition + 10, value, Chr$(34) & "childIds" & Chr$(34), vbTextCompare) > 0 Then Err.Raise 5, , "Duplicate legacy childIds property."
    start = InStr(keyPosition + 10, value, ":")
    If start = 0 Then Err.Raise 5, , "Invalid legacy childIds property."
    first = InStr(start + 1, value, "[")
    last = InStr(first + 1, value, "]")
    If first = 0 Or last = 0 Then Err.Raise 5, , "Invalid legacy childIds array."
    token = Trim$(Mid$(value, first + 1, last - first - 1))
    If token <> "" Then
        items = Split(token, ",")
        For Each item In items
            token = Trim$(CStr(item))
            If Left$(token, 1) = Chr$(34) And Right$(token, 1) = Chr$(34) Then token = Mid$(token, 2, Len(token) - 2)
            result.Add PositiveInteger(token)
        Next item
    End If
    Set LegacyIds = result
End Function

Public Function GroupsOnSlide(ByVal slide As Object) As Collection
    Set GroupsOnSlide = GroupsFromLeaves(SlideLeaves(slide))
End Function

Public Function GroupsFromLeaves(ByVal leaves As Collection) As Collection
    Dim groups As New Collection, shape As Object, childShape As Object, group As Collection, child As Collection
    Dim key As String, role As String, legacy As String, ids As Collection, id As Variant, first As Long, last As Long, ordinal As Long
    For Each shape In leaves
        key = UCase$(PptNativeDriver.ReadTag(shape, RELATION_KEY))
        role = UCase$(PptNativeDriver.ReadTag(shape, ROLE_KEY))
        If key <> "" Or role <> "" Then
            If Left$(key, 1) <> "G" Then Err.Raise 5, , "Invalid native relationship key."
            ordinal = PositiveInteger(Mid$(key, 2))
            Set group = GetGroup(groups, key)
            If group Is Nothing Then
                Set group = NewGroup(key, "native")
                groups.Add group
            End If
            If role = "P" Then
                AddParent group, shape
            ElseIf Left$(role, 1) = "C" Then
                AddChild group, shape, PositiveInteger(Mid$(role, 2))
            Else
                Err.Raise 5, , "Invalid native relationship role."
            End If
            If PptNativeDriver.ReadTag(shape, LEGACY_PARENT) <> "" Or PptNativeDriver.ReadTag(shape, LEGACY_CHILD) <> "" Then Err.Raise 5, , "Conflicting native and legacy relationship tags."
        End If
    Next shape
    For Each shape In leaves
        legacy = PptNativeDriver.ReadTag(shape, LEGACY_PARENT)
        If legacy <> "" Then
            key = "L" & CStr(PptNativeDriver.ShapeId(shape))
            Set group = NewGroup(key, "legacy")
            AddParent group, shape
            Set ids = LegacyIds(legacy, first, last)
            ordinal = 0
            For Each id In ids
                ordinal = ordinal + 1
                Set childShape = FindShape(leaves, CLng(id))
                If childShape Is Nothing Then
                    SetProblem group, "Missing layout child: " & CStr(id)
                ElseIf PptNativeDriver.ReadTag(childShape, LEGACY_CHILD) <> CStr(PptNativeDriver.ShapeId(shape)) Then
                    SetProblem group, "Inconsistent layout child: " & CStr(id)
                Else
                    AddChild group, childShape, ordinal
                End If
            Next id
            groups.Add group
        End If
    Next shape
    ' Orphaned legacy children remain visible and can be explicitly detached.
    For Each shape In leaves
        legacy = PptNativeDriver.ReadTag(shape, LEGACY_CHILD)
        If legacy <> "" Then
            key = "L" & CStr(PositiveInteger(legacy))
            Set group = GetGroup(groups, key)
            If group Is Nothing Then
                Set group = NewGroup(key, "legacy")
                SetProblem group, "Missing layout parent."
                groups.Add group
            End If
            If Not IsChild(group, PptNativeDriver.ShapeId(shape)) Then
                AddChild group, shape, group("Children").Count + 1
                SetProblem group, "Inconsistent layout relationship."
            End If
        End If
    Next shape
    For Each group In groups
        If CLng(group("ParentId")) = 0 Then SetProblem group, "Missing relationship parent."
        If group("Children").Count = 0 Then SetProblem group, "Relationship has no children."
    Next group
    Set GroupsFromLeaves = groups
End Function

Public Function IsNativeParent(ByVal shape As Object) As Boolean
    Dim key As String, role As String, ignored As Long
    key = UCase$(PptNativeDriver.ReadTag(shape, RELATION_KEY))
    role = UCase$(PptNativeDriver.ReadTag(shape, ROLE_KEY))
    If key = "" And role = "" Then Exit Function
    If Left$(key, 1) <> "G" Then Err.Raise 5, , "Invalid native relationship key."
    ignored = PositiveInteger(Mid$(key, 2))
    If role = "P" Then
        IsNativeParent = True
    ElseIf Left$(role, 1) = "C" Then
        ignored = PositiveInteger(Mid$(role, 2))
    Else
        Err.Raise 5, , "Invalid native relationship role."
    End If
End Function

Private Function IsChild(ByVal group As Collection, ByVal id As Long) As Boolean
    Dim child As Collection
    For Each child In group("Children")
        If PptNativeDriver.ShapeId(child("Shape")) = id Then IsChild = True: Exit Function
    Next child
End Function

Private Function GroupForShape(ByVal groups As Collection, ByVal id As Long) As Collection
    Dim group As Collection, found As Collection
    For Each group In groups
        If CLng(group("ParentId")) = id Or IsChild(group, id) Then
            If Not found Is Nothing Then Err.Raise 5, , "An object belongs to conflicting relationships."
            Set found = group
        End If
    Next group
    Set GroupForShape = found
End Function

Public Sub CancelPending()
    Set pendingPresentation = Nothing
    pendingId = 0
    pendingSlide = 0
End Sub

Public Sub RefreshContext()
    Dim current As Object
    If pendingPresentation Is Nothing Or IsPreviewContext Then Exit Sub
    On Error GoTo Failed
    If Not PptNativeDriver.HasPresentation Then CancelPending: Exit Sub
    Set current = PptNativeDriver.CurrentPresentation()
    If Not current Is pendingPresentation Then CancelPending: Exit Sub
    If PptNativeDriver.SlideId(PptNativeDriver.CurrentSlide()) <> pendingSlide Then CancelPending
    Exit Sub
Failed:
    Debug.Print "[NativeRelationContext] " & Err.Description
    CancelPending
End Sub

Private Function PendingValid() As Boolean
    Dim found As Object
    If pendingPresentation Is Nothing Or IsPreviewContext Then Exit Function
    If Not PptNativeDriver.HasPresentation Then CancelPending: Exit Function
    Set found = PendingParentFrom(SlideLeaves(PptNativeDriver.CurrentSlide()))
    PendingValid = Not found Is Nothing
End Function

Private Function PendingParentFrom(ByVal leaves As Collection) As Object
    Dim current As Object, found As Object
    If pendingPresentation Is Nothing Then Exit Function
    If IsPreviewContext Then Exit Function
    If Not PptNativeDriver.HasPresentation Then CancelPending: Exit Function
    Set current = PptNativeDriver.CurrentPresentation()
    If Not current Is pendingPresentation Then CancelPending: Exit Function
    If PptNativeDriver.SlideId(PptNativeDriver.CurrentSlide()) <> pendingSlide Then CancelPending: Exit Function
    Set found = FindShape(leaves, pendingId)
    If found Is Nothing Then CancelPending: Exit Function
    Set PendingParentFrom = found
End Function

Public Function CanMarkParent() As Boolean
    Dim roots As Collection, groups As Collection
    If IsPreviewContext Or Not PptNativeDriver.HasShapeSelection Then Exit Function
    Set roots = PptNativeDriver.SelectionObjects()
    If roots.Count <> 1 Then Exit Function
    If Not RadiusNativeCore.IsRoundRect(roots(1)) Then Exit Function
    Set groups = GroupsOnSlide(PptNativeDriver.CurrentSlide())
    CanMarkParent = CanMarkParentFrom(roots, groups)
End Function

Private Function CanMarkParentFrom(ByVal roots As Collection, ByVal groups As Collection) As Boolean
    Dim group As Collection, id As Long
    If roots.Count <> 1 Then Exit Function
    If Not RadiusNativeCore.IsRoundRect(roots(1)) Then Exit Function
    id = PptNativeDriver.ShapeId(roots(1))
    Set group = GroupForShape(groups, id)
    If Not group Is Nothing Then
        If CStr(group("Kind")) <> "native" Or CLng(group("ParentId")) <> id Or CStr(group("Problem")) <> "" Then Exit Function
    End If
    CanMarkParentFrom = True
End Function

Public Sub MarkParent()
    Dim roots As Collection
    If Not CanMarkParent Then Err.Raise 5, , "Select one unbound rounded rectangle or an existing native parent. Legacy layouts must be detached explicitly before rebuilding."
    Set roots = PptNativeDriver.SelectionObjects()
    Set pendingPresentation = PptNativeDriver.CurrentPresentation()
    pendingSlide = PptNativeDriver.SlideId(PptNativeDriver.CurrentSlide())
    pendingId = PptNativeDriver.ShapeId(roots(1))
End Sub

Private Sub ValidateBinding(ByRef slide As Object, ByRef parent As Object, ByRef selected As Collection, ByRef groups As Collection)
    Dim problem As String
    If Not PendingValid Then Err.Raise 5, , "Choose a parent on this slide first."
    Set slide = PptNativeDriver.CurrentSlide()
    Set parent = FindShape(SlideLeaves(slide), pendingId)
    Set selected = SelectedLeaves()
    If selected.Count = 0 Then Err.Raise 5, , "Select the child rounded rectangles."
    Set groups = GroupsOnSlide(slide)
    problem = BindingProblem(selected, groups)
    If problem <> "" Then Err.Raise 5, , problem
End Sub

Private Function BindingProblem(ByVal selected As Collection, ByVal groups As Collection) As String
    Dim shape As Object, group As Collection
    If selected.Count = 0 Then BindingProblem = "Select the child rounded rectangles.": Exit Function
    Set group = GroupForShape(groups, pendingId)
    If Not group Is Nothing Then
        If CStr(group("Kind")) <> "native" Or CLng(group("ParentId")) <> pendingId Or CStr(group("Problem")) <> "" Then BindingProblem = "The pending parent has an incompatible existing relationship.": Exit Function
    End If
    For Each shape In selected
        If PptNativeDriver.ShapeId(shape) = pendingId Then BindingProblem = "A parent cannot also be its child. Select only child objects.": Exit Function
        Set group = GroupForShape(groups, PptNativeDriver.ShapeId(shape))
        If Not group Is Nothing Then
            If CStr(group("Kind")) <> "native" Or CLng(group("ParentId")) <> pendingId Or CStr(group("Problem")) <> "" Then BindingProblem = "Object already belongs to " & CStr(group("Key")) & ": " & PptNativeDriver.ShapeName(shape) & ". Detach it explicitly first.": Exit Function
        End If
    Next shape
End Function

Public Function CanBind() As Boolean
    Dim slide As Object, parent As Object, selected As Collection, groups As Collection
    On Error GoTo NotReady
    If Not PendingValid Then Exit Function
    Set selected = SelectedLeaves()
    If selected.Count = 0 Then Exit Function
    ValidateBinding slide, parent, selected, groups
    CanBind = True
    Exit Function
NotReady:
    Debug.Print "[NativeRelationBinding] " & Err.Description
End Function

Private Sub ChangeTag(ByVal plan As Collection, ByVal shape As Object, ByVal key As String, ByVal value As String, Optional ByVal remove As Boolean = False)
    plan.Add Array(PptNativeDriver.ShapeId(shape), key, value, remove)
End Sub

Public Sub BindChildren()
    Dim slide As Object, parent As Object, selected As Collection, groups As Collection, group As Collection, shape As Object, child As Collection
    Dim plan As New Collection, key As String, number As Long, ordinal As Long
    ValidateBinding slide, parent, selected, groups
    Set group = GroupForShape(groups, pendingId)
    If group Is Nothing Then
        number = 1
        Do
            key = "G" & Format$(number, "00")
            Set group = GetGroup(groups, key)
            If group Is Nothing Then Exit Do
            number = number + 1
        Loop
        ChangeTag plan, parent, RELATION_KEY, key
        ChangeTag plan, parent, ROLE_KEY, "P"
    Else
        key = CStr(group("Key"))
        For Each child In group("Children")
            If CLng(child("Order")) > ordinal Then ordinal = CLng(child("Order"))
        Next child
    End If
    For Each shape In selected
        If group Is Nothing Then
            ordinal = ordinal + 1
            ChangeTag plan, shape, RELATION_KEY, key
            ChangeTag plan, shape, ROLE_KEY, "C" & CStr(ordinal)
        ElseIf Not IsChild(group, PptNativeDriver.ShapeId(shape)) Then
            ordinal = ordinal + 1
            ChangeTag plan, shape, RELATION_KEY, key
            ChangeTag plan, shape, ROLE_KEY, "C" & CStr(ordinal)
        End If
    Next shape
    If plan.Count > 0 Then RadiusNativeCore.ApplyMetadataPlan slide, plan
    revision = revision + 1
    CancelPending
End Sub

Public Function SelectionInfo() As Collection
    Dim snapshot As Collection
    Set snapshot = ReadUiSnapshot()
    Set SelectionInfo = snapshot("Info")
End Function

' Only Ribbon display callbacks may reuse this snapshot. Actions read again.
Public Function ReadUiSnapshot() As Collection
    Dim snapshot As New Collection, groups As New Collection, selected As New Collection
    Dim objects As New Collection, leaves As New Collection, parent As Object, group As Collection
    Dim preview As Boolean
    preview = IsPreviewContext
    If Not preview And PptNativeDriver.HasPresentation Then
        Set leaves = SlideLeaves(PptNativeDriver.CurrentSlide())
        Set groups = GroupsFromLeaves(leaves)
        If PptNativeDriver.HasShapeSelection Then Set objects = PptNativeDriver.SelectionObjects()
        Set selected = RoundedLeaves(objects)
        Set parent = PendingParentFrom(leaves)
        Set group = SelectedGroupFrom(groups, selected)
    End If
    snapshot.Add groups, "Groups"
    snapshot.Add selected, "Selected"
    snapshot.Add SelectionInfoFrom(groups, selected, parent, preview), "Info"
    snapshot.Add (Not group Is Nothing), "HasGroup"
    If Not group Is Nothing Then snapshot.Add group, "Group"
    snapshot.Add (Not preview And groups.Count > 0), "CanView"
    snapshot.Add (Not preview And CanMarkParentFrom(objects, groups)), "CanMarkParent"
    snapshot.Add (Not preview And CanBindFrom(parent, selected, groups)), "CanBind"
    snapshot.Add (Not preview And CanDetachFrom(group, selected, False)), "CanDetachChild"
    snapshot.Add (Not preview And Not group Is Nothing), "CanDetachWhole"
    Set ReadUiSnapshot = snapshot
End Function

Private Function CanBindFrom(ByVal parent As Object, ByVal selected As Collection, ByVal groups As Collection) As Boolean
    If parent Is Nothing Or selected.Count = 0 Then Exit Function
    CanBindFrom = (BindingProblem(selected, groups) = "")
End Function

Private Function CanDetachFrom(ByVal group As Collection, ByVal selected As Collection, ByVal whole As Boolean) As Boolean
    Dim shape As Object
    If group Is Nothing Then Exit Function
    If whole Then CanDetachFrom = True: Exit Function
    If selected.Count = 0 Then Exit Function
    For Each shape In selected
        If Not IsChild(group, PptNativeDriver.ShapeId(shape)) Then Exit Function
    Next shape
    CanDetachFrom = True
End Function

Private Function SelectionInfoFrom(ByVal groups As Collection, ByVal selected As Collection, ByVal parent As Object, ByVal preview As Boolean) As Collection
    Dim info As New Collection, shape As Object, group As Collection, chosen As Collection
    Dim state As String, key As String, role As String, name As String, parentName As String, count As Long, groupCount As Long, child As Collection, pendingStatus As String
    Dim keys As String, unboundCount As Long
    state = "empty"
    If preview Then
        state = "preview"
    Else
        groupCount = groups.Count
        count = selected.Count
        For Each shape In selected
            Set group = GroupForShape(groups, PptNativeDriver.ShapeId(shape))
            If Not group Is Nothing Then
                If InStr(1, "|" & keys & "|", "|" & CStr(group("Key")) & "|", vbBinaryCompare) = 0 Then
                    If keys <> "" Then keys = keys & "|"
                    keys = keys & CStr(group("Key"))
                End If
                If chosen Is Nothing Then
                    Set chosen = group
                ElseIf CStr(chosen("Key")) <> CStr(group("Key")) Then
                    state = "mixed"
                End If
            Else
                unboundCount = unboundCount + 1
            End If
        Next shape
        If Not chosen Is Nothing And unboundCount > 0 Then state = "mixed"
        If state <> "mixed" Then
            If count = 0 Then
                state = "empty"
            ElseIf chosen Is Nothing Then
                state = "unbound"
            Else
                state = "bound"
                key = CStr(chosen("Key"))
                If CStr(chosen("Problem")) <> "" Then state = "broken"
                If CLng(chosen("ParentId")) > 0 Then parentName = PptNativeDriver.ShapeName(chosen("Parent"))
                If count = 1 Then
                    name = PptNativeDriver.ShapeName(selected(1))
                    If PptNativeDriver.ShapeId(selected(1)) = CLng(chosen("ParentId")) Then
                        role = "P"
                    Else
                        For Each child In chosen("Children")
                            If PptNativeDriver.ShapeId(child("Shape")) = PptNativeDriver.ShapeId(selected(1)) Then role = "C" & CStr(child("Order"))
                        Next child
                    End If
                End If
            End If
        End If
        If Not parent Is Nothing Then
            state = "pending"
            parentName = PptNativeDriver.ShapeName(parent)
            pendingStatus = "ready"
            If count = 0 Then pendingStatus = "select"
            For Each shape In selected
                If PptNativeDriver.ShapeId(shape) = pendingId Then pendingStatus = "self"
                Set group = GroupForShape(groups, PptNativeDriver.ShapeId(shape))
                If Not group Is Nothing Then
                    If CStr(group("Kind")) <> "native" Or CLng(group("ParentId")) <> pendingId Then pendingStatus = "conflict"
                End If
            Next shape
        End If
    End If
    info.Add state, "State"
    info.Add key, "Key"
    info.Add role, "Role"
    info.Add name, "Name"
    info.Add parentName, "ParentName"
    info.Add count, "Selected"
    info.Add groupCount, "Groups"
    info.Add Replace(keys, "|", " / "), "Keys"
    info.Add unboundCount, "Unbound"
    info.Add pendingStatus, "PendingStatus"
    If chosen Is Nothing Then info.Add 0&, "Children" Else info.Add chosen("Children").Count, "Children"
    Set SelectionInfoFrom = info
End Function

Public Function SelectedGroup() As Collection
    Dim groups As Collection, selected As Collection
    If IsPreviewContext Or Not PptNativeDriver.HasPresentation Then Exit Function
    Set groups = GroupsOnSlide(PptNativeDriver.CurrentSlide())
    Set selected = SelectedLeaves()
    Set SelectedGroup = SelectedGroupFrom(groups, selected)
End Function

Public Function SelectedGroupFrom(ByVal groups As Collection, ByVal selected As Collection) As Collection
    Dim shape As Object, group As Collection, found As Collection
    For Each shape In selected
        Set group = GroupForShape(groups, PptNativeDriver.ShapeId(shape))
        If Not group Is Nothing Then
            If found Is Nothing Then
                Set found = group
            ElseIf CStr(found("Key")) <> CStr(group("Key")) Then
                Exit Function
            End If
        End If
    Next shape
    Set SelectedGroupFrom = found
End Function

Public Function CanDetach(ByVal whole As Boolean) As Boolean
    Dim group As Collection, selected As Collection
    Set group = SelectedGroup()
    If group Is Nothing Then Exit Function
    If whole Then CanDetach = True: Exit Function
    Set selected = SelectedLeaves()
    CanDetach = CanDetachFrom(group, selected, whole)
End Function

Public Function RemovalInfo() As Collection
    Set RemovalInfo = SelectedGroup()
    If RemovalInfo Is Nothing Then Err.Raise 5, , "Select objects belonging to one relationship."
End Function

Public Sub Detach(ByVal whole As Boolean)
    Dim group As Collection, selected As Collection, plan As New Collection, child As Collection, shape As Object, remaining As String, value As String
    Dim first As Long, last As Long, ignored As Collection, remove As Boolean, count As Long
    If Not CanDetach(whole) Then Err.Raise 5, , "Select the child objects or one complete relationship."
    Set group = SelectedGroup()
    Set selected = SelectedLeaves()
    For Each child In group("Children")
        Set shape = child("Shape")
        remove = whole Or Not FindShape(selected, PptNativeDriver.ShapeId(shape)) Is Nothing
        If remove Then
            If CStr(group("Kind")) = "native" Then
                ChangeTag plan, shape, RELATION_KEY, "", True
                ChangeTag plan, shape, ROLE_KEY, "", True
            Else
                ChangeTag plan, shape, LEGACY_CHILD, "", True
            End If
        Else
            count = count + 1
            If remaining <> "" Then remaining = remaining & ","
            remaining = remaining & Chr$(34) & CStr(PptNativeDriver.ShapeId(shape)) & Chr$(34)
        End If
    Next child
    If CLng(group("ParentId")) > 0 Then
        Set shape = group("Parent")
        If CStr(group("Kind")) = "native" Then
            If whole Or count = 0 Then
                ChangeTag plan, shape, RELATION_KEY, "", True
                ChangeTag plan, shape, ROLE_KEY, "", True
                ChangeTag plan, shape, "radiusRelationLayout_v1", "", True
                ChangeTag plan, shape, "radiusRelationBaseline_v1", "", True
            End If
        ElseIf whole Or count = 0 Then
            ChangeTag plan, shape, LEGACY_PARENT, "", True
        Else
            If CStr(group("Problem")) <> "" Then Err.Raise 5, , "Detach the incomplete legacy relationship as a whole to preserve its unresolved IDs."
            value = PptNativeDriver.ReadTag(shape, LEGACY_PARENT)
            Set ignored = LegacyIds(value, first, last)
            ChangeTag plan, shape, LEGACY_PARENT, Left$(value, first) & remaining & Mid$(value, last)
        End If
    End If
    RadiusNativeCore.ApplyMetadataPlan PptNativeDriver.CurrentSlide(), plan
    revision = revision + 1
    CancelPending
End Sub

Public Function BeginMenu() As Collection
    If IsPreviewContext Then Err.Raise 5, , "Close the preview to manage source relationships."
    Set BeginMenu = BeginMenuFrom(GroupsOnSlide(PptNativeDriver.CurrentSlide()))
End Function

Public Function BeginMenuFrom(ByVal groups As Collection) As Collection
    If IsPreviewContext Then Err.Raise 5, , "Close the preview to manage source relationships."
    Set menuPresentation = PptNativeDriver.CurrentPresentation()
    menuSlide = PptNativeDriver.SlideId(PptNativeDriver.CurrentSlide())
    Set BeginMenuFrom = groups
End Function

Public Function MenuToken(ByVal key As String, ByVal mode As String, ByVal id As Long) As String
    MenuToken = CStr(revision) & "|" & CStr(menuSlide) & "|" & key & "|" & mode & "|" & CStr(id)
End Function

Public Sub Locate(ByVal token As String)
    Dim parts As Variant, current As Object, group As Collection, groups As Collection, shapes As New Collection, child As Collection, id As Long
    parts = Split(token, "|")
    If UBound(parts) <> 4 Then Err.Raise 5, , "Invalid relationship menu action."
    If CLng(parts(0)) <> revision Or menuPresentation Is Nothing Then Err.Raise 5, , "Relationship menu is stale. Open it again."
    Set current = PptNativeDriver.CurrentPresentation()
    If Not current Is menuPresentation Then Err.Raise 5, , "Presentation changed. Open the relationship menu again."
    If PptNativeDriver.SlideId(PptNativeDriver.CurrentSlide()) <> menuSlide Or CLng(parts(1)) <> menuSlide Then Err.Raise 5, , "Slide changed. Open the relationship menu again."
    Set groups = GroupsOnSlide(PptNativeDriver.CurrentSlide())
    Set group = GetGroup(groups, CStr(parts(2)))
    If group Is Nothing Then Err.Raise 5, , "Relationship no longer exists."
    If CStr(parts(3)) = "object" Then
        id = CLng(parts(4))
        If CLng(group("ParentId")) = id Then
            PptNativeDriver.SelectObject group("Parent")
        Else
            For Each child In group("Children")
                If PptNativeDriver.ShapeId(child("Shape")) = id Then PptNativeDriver.SelectObject child("Shape"): Exit Sub
            Next child
            Err.Raise 5, , "Relationship member no longer exists."
        End If
    Else
        If CStr(parts(3)) = "family" And CLng(group("ParentId")) > 0 Then shapes.Add group("Parent")
        For Each child In group("Children")
            shapes.Add child("Shape")
        Next child
        If shapes.Count = 0 Then Err.Raise 5, , "No relationship members can be located."
        PptNativeDriver.SelectObjects shapes
    End If
End Sub

Public Function CanView() As Boolean
    Dim groups As Collection
    If IsPreviewContext Or Not PptNativeDriver.HasPresentation Then Exit Function
    Set groups = GroupsOnSlide(PptNativeDriver.CurrentSlide())
    CanView = (groups.Count > 0)
End Function

Public Sub TogglePreview(ByVal banner As String, ByVal parentLabel As String, ByVal childLabel As String)
    Dim copy As Object, source As Object, slide As Object, group As Collection, groups As Collection, child As Collection, box As Variant, text As String
    Dim errorNumber As Long, errorText As String, recoveryText As String
    If Not previewPresentation Is Nothing Then
        Set copy = previewPresentation
        Set source = previewSource
        PptNativeDriver.CloseTemporaryPresentation copy
        Set previewPresentation = Nothing
        If Not source Is Nothing Then PptNativeDriver.ActivatePresentation source
        Set previewSource = Nothing
        Exit Sub
    End If
    If Not CanView Then Err.Raise 5, , "Create or select a relationship first."
    Set source = PptNativeDriver.CurrentPresentation()
    On Error GoTo Failed
    Set copy = PptNativeDriver.CreateSlideCopy(PptNativeDriver.CurrentSlide())
    Set previewSource = source
    Set previewPresentation = copy
    Set slide = PptNativeDriver.FirstSlide(copy)
    Set groups = GroupsOnSlide(slide)
    For Each group In groups
        If CLng(group("ParentId")) > 0 Then
            box = PptNativeDriver.ShapeBox(group("Parent"))
            text = CStr(group("Key")) & " - " & parentLabel
            PptNativeDriver.AddTextBadge slide, text, CDbl(box(0)), CDbl(box(1)), 72#, RGB(36, 82, 166)
        End If
        For Each child In group("Children")
            box = PptNativeDriver.ShapeBox(child("Shape"))
            text = CStr(group("Key")) & " - " & childLabel & CStr(child("Order"))
            PptNativeDriver.AddTextBadge slide, text, CDbl(box(0)), CDbl(box(1)), 82#, RGB(140, 75, 20)
        Next child
    Next group
    PptNativeDriver.AddTextBadge slide, banner, 4#, 2#, 235#, RGB(70, 75, 85)
    Exit Sub
Failed:
    errorNumber = Err.Number
    errorText = Err.Description
    If Not copy Is Nothing Then
        On Error Resume Next
        Err.Clear
        PptNativeDriver.CloseTemporaryPresentation copy
        If Err.Number <> 0 Then recoveryText = " Preview recovery failed: " & Err.Description
        On Error GoTo 0
    End If
    Set previewPresentation = Nothing
    Set previewSource = Nothing
    Err.Raise errorNumber, "RadiusNativeRelations", errorText & recoveryText
End Sub
