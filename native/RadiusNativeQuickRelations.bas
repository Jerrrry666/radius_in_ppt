Attribute VB_Name = "RadiusNativeQuickRelations"
Option Explicit

Private Const RELATION_KEY As String = "radiusRelation_v1"
Private Const ROLE_KEY As String = "radiusRelationRole_v1"
Private Const LOCK_KEY As String = "radiusLock_v1"
Private Const STRICT_KEY As String = "radiusLockStrict_v1"
Private Const SETTINGS_KEY As String = "radiusRelationLayout_v1"
Private Const BASELINE_KEY As String = "radiusRelationBaseline_v1"
Private Const RELATION_TAGS As String = "radiusRelation_v1|radiusRelationRole_v1"
Private Const LAYOUT_TAGS As String = "radiusRelationLayout_v1|radiusRelationBaseline_v1"

' All cases run on a new, caller-owned slide. Expected values below are fixed
' fixture facts, not answers computed by the production grid/link functions.
Public Sub CaseBinding(ByVal slide As Object)
    Dim parent As Object, other As Object, shape As Object, group As Object, members As New Collection
    Dim parentId As Long, childIds(1 To 3) As Variant, arrowId As Long, i As Long
    Dim originals As New Collection, groupState As Collection, before As Collection, snapshot As Collection
    BeginCase slide
    Set parent = RadiusNativeQuickTest.FixtureRound(slide, "QBindParent", 2, 2, 12, 8, 1)
    parentId = PptNativeDriver.ShapeId(parent)
    SeedTag slide, parent, "customParent", "KeepParentCase"
    SeedTag slide, parent, LOCK_KEY, "1.000000"
    originals.Add SnapshotShape(parent)
    For i = 1 To 3
        Set shape = RadiusNativeQuickTest.FixtureRound(slide, "QBindChild" & CStr(i), 2 + i * 4, 4, 3, 2, 0.25)
        childIds(i) = PptNativeDriver.ShapeId(shape)
        SeedTag slide, shape, "customChild", "KeepChildCase" & CStr(i)
        SeedTag slide, shape, LOCK_KEY, "0.250000"
        If i = 3 Then SeedTag slide, shape, STRICT_KEY, "1"
        originals.Add SnapshotShape(shape)
        members.Add shape
    Next i
    Set shape = RadiusNativeQuickTest.FixtureShape(slide, "QBindArrow", 33, 5, 7, 3, 1)
    arrowId = PptNativeDriver.ShapeId(shape)
    SeedTag slide, shape, "customArrow", "KeepArrowCase"
    originals.Add SnapshotShape(shape)
    members.Add shape
    Set group = RadiusNativeQuickTest.FixtureGroup(slide, "QBindGroup", members)
    SeedTag slide, group, "customGroup", "KeepGroupCase"
    Set groupState = SnapshotShape(group)

    SelectId slide, parentId
    RadiusNativeRelations.MarkParent
    Set snapshot = RadiusNativeRelations.ReadUiSnapshot()
    RadiusNativeQuickTest.Check CStr(snapshot("Info")("State")) = "pending", "mark parent enters pending state"
    RadiusNativeRelations.CancelPending
    Set snapshot = RadiusNativeRelations.ReadUiSnapshot()
    RadiusNativeQuickTest.Check CStr(snapshot("Info")("State")) = "unbound", "cancel clears pending parent"
    AssertShapes slide, originals, "cancel leaves fixtures unchanged"

    SelectId slide, parentId
    RadiusNativeRelations.MarkParent
    SelectNamedRoot slide, "QBindGroup"
    Set snapshot = RadiusNativeRelations.ReadUiSnapshot()
    RadiusNativeQuickTest.Check CBool(snapshot("CanBind")), "strict child still allows relation binding"
    RadiusNativeQuickTest.RequireScratch slide
    RadiusNativeRelations.BindChildren
    RadiusNativeQuickTest.AssertGuardReleased
    AssertTag RadiusNativeQuickTest.FindShape(slide, parentId), RELATION_KEY, "G01", "bound parent key"
    AssertTag RadiusNativeQuickTest.FindShape(slide, parentId), ROLE_KEY, "P", "bound parent role"
    For i = 1 To 3
        Set shape = RadiusNativeQuickTest.FindShape(slide, CLng(childIds(i)))
        AssertTag shape, RELATION_KEY, "G01", "batch child key"
        AssertTag shape, ROLE_KEY, "C" & CStr(i), "batch child order"
    Next i
    Set shape = RadiusNativeQuickTest.FindShape(slide, arrowId)
    AssertAbsentTag shape, RELATION_KEY, "non-rounded arrow has no relation"
    AssertAbsentTag shape, ROLE_KEY, "non-rounded arrow has no role"
    AssertShapes slide, originals, "bind preserves leaf state", RELATION_TAGS
    AssertGroup slide, "QBindGroup", groupState, "bind preserves group metadata"
    Set snapshot = RadiusNativeRelations.ReadUiSnapshot()
    RadiusNativeQuickTest.Check CStr(snapshot("Info")("State")) <> "pending", "bind clears pending parent"
    RadiusNativeQuickTest.Check snapshot("Groups").Count = 1, "one batch bind creates one relationship"
    RadiusNativeQuickTest.Check snapshot("Groups")(1)("Children").Count = 3, "one batch bind creates three children"

    Set other = RadiusNativeQuickTest.FixtureRound(slide, "QBindOtherParent", 20, 2, 6, 4, 0.4)
    SelectId slide, PptNativeDriver.ShapeId(other)
    RadiusNativeRelations.MarkParent
    SelectId slide, CLng(childIds(1))
    Set before = SnapshotScene(slide)
    ExpectRejection slide, "bind", "Object already belongs to G01"
    AssertScene slide, before, "existing ownership rejects without writes"
    RadiusNativeRelations.CancelPending

    SelectId slide, parentId
    RadiusNativeRelations.MarkParent
    Set before = SnapshotScene(slide)
    ExpectRejection slide, "bind", "A parent cannot also be its child."
    AssertScene slide, before, "self-binding rejects without writes"
    RadiusNativeRelations.CancelPending

    SelectId slide, CLng(childIds(2))
    RadiusNativeQuickTest.RequireScratch slide
    RadiusNativeRelations.Detach False
    RadiusNativeQuickTest.AssertGuardReleased
    Set shape = RadiusNativeQuickTest.FindShape(slide, CLng(childIds(2)))
    AssertAbsentTag shape, RELATION_KEY, "partial detach removes child relation"
    AssertAbsentTag shape, ROLE_KEY, "partial detach removes child role"
    AssertTag RadiusNativeQuickTest.FindShape(slide, CLng(childIds(1))), ROLE_KEY, "C1", "partial detach keeps first order"
    AssertTag RadiusNativeQuickTest.FindShape(slide, CLng(childIds(3))), ROLE_KEY, "C3", "partial detach keeps surviving order"
    AssertShapes slide, originals, "partial detach preserves leaf state", RELATION_TAGS
    AssertGroup slide, "QBindGroup", groupState, "partial detach preserves group metadata"

    SelectId slide, parentId
    RadiusNativeQuickTest.RequireScratch slide
    RadiusNativeRelations.Detach True
    RadiusNativeQuickTest.AssertGuardReleased
    AssertShapes slide, originals, "whole detach restores original leaf metadata"
    AssertGroup slide, "QBindGroup", groupState, "whole detach preserves group metadata"
    RadiusNativeQuickTest.Check RadiusNativeRelations.GroupsOnSlide(slide).Count = 0, "whole detach leaves no relationship"
End Sub

Public Sub CaseLayout(ByVal slide As Object)
    Dim parentId As Long, childIds As Variant, before As Collection, children As Collection, config As Variant, shape As Object, snapshot As Collection
    BeginCase slide
    CreateFamily slide, "QLayout", 4, parentId, childIds
    SelectId slide, parentId
    Set before = SnapshotScene(slide)
    Stage slide, "rows", "2"
    Stage slide, "padding", "0.5"
    Stage slide, "gap", "0.3"
    RadiusNativeQuickTest.RequireScratch slide
    RadiusNativeLayout.StepParameter "gap", 1
    config = RadiusNativeLayout.SelectedConfig()
    RadiusNativeQuickTest.CheckNear CDbl(config(3)), 0.4, "layout gap step stages value"
    Stage slide, "gap", "0.3"
    AssertScene slide, before, "typed and stepped draft do not write shapes"
    Set snapshot = RadiusNativeLayout.ReadUiSnapshot(RadiusNativeRelations.SelectedGroup())
    RadiusNativeQuickTest.Check CBool(snapshot("Pending")), "layout display marks draft pending"
    RadiusNativeQuickTest.Check RadiusNativeLayout.ChildrenWritable, "unprotected children allow layout"
    ApplyLayout slide
    Set snapshot = RadiusNativeLayout.ReadUiSnapshot(RadiusNativeRelations.SelectedGroup())
    RadiusNativeQuickTest.Check Not CBool(snapshot("Pending")) And CBool(snapshot("Automatic")), "applied layout clears pending and enables automatic state"
    AssertGrid slide, childIds, 2.5, 2.5, 8.15, 6.15, 5.35, 3.35, "same grid"
    AssertChildRadius slide, childIds, 1, "same layout writes child radii and fixed tags"
    AssertTag RadiusNativeQuickTest.FindShape(slide, parentId), SETTINGS_KEY, "1|2|2|0.500000|0.300000|same|1", "serialized same layout"
    RadiusNativeQuickTest.Check RadiusNativeLayout.LinkedSelectionWritable, "unprotected configured parent allows linked radius"

    ApplyRadius slide, 0.5
    AssertParentRadius slide, parentId, 0.5, "linked parent radius"
    AssertChildRadius slide, childIds, 0.5, "parent radius links immediately"
    AssertGrid slide, childIds, 2.5, 2.5, 8.15, 6.15, 5.35, 3.35, "radius edit preserves grid"
    ChangeMode slide, "subtract"
    AssertChildRadius slide, childIds, 0, "subtract clips child radius to zero"
    ApplyRadius slide, 0.8
    AssertParentRadius slide, parentId, 0.8, "subtract parent radius"
    AssertChildRadius slide, childIds, 0.3, "subtract links positive child radius"

    Set children = SnapshotIds(slide, childIds)
    ChangeMode slide, "off"
    ApplyRadius slide, 0.3
    AssertParentRadius slide, parentId, 0.3, "off parent radius"
    AssertShapes slide, children, "off preserves child fraction and fixed tags"
    Set shape = RadiusNativeQuickTest.FindShape(slide, parentId)
    RadiusNativeQuickTest.RequireScratch slide
    PptNativeDriver.SetShapeBox shape, CmBox(3, 3, 14, 10)
    RadiusNativeQuickTest.RequireScratch slide
    RadiusNativeLayout.SyncSlide slide
    RadiusNativeQuickTest.AssertGuardReleased
    RadiusNativeQuickTest.CheckNear PptNativeDriver.ReadFraction(shape), 0.0375, "off resize retains parent fraction", 0.000001
    RadiusNativeQuickTest.CheckNear RadiusNativeCore.CurrentRadius(shape), 0.375, "off resize changes parent radius with short side"
    AssertTag shape, LOCK_KEY, "0.300000", "off resize leaves parent fixed tag unchanged"
    AssertGrid slide, childIds, 3.5, 3.5, 10.15, 8.15, 6.35, 4.35, "off still follows parent geometry"
    AssertShapes slide, children, "off geometry sync preserves fraction and fixed tags", "", False

    ChangeMode slide, "same"
    AssertChildRadius slide, childIds, 0.375, "same uses resized parent radius before disabling automatic layout"
    RadiusNativeQuickTest.RequireScratch slide
    RadiusNativeLayout.SetAutomatic False
    RadiusNativeQuickTest.AssertGuardReleased
    Set children = SnapshotIds(slide, childIds)
    Set shape = RadiusNativeQuickTest.FindShape(slide, parentId)
    RadiusNativeQuickTest.RequireScratch slide
    PptNativeDriver.SetShapeBox shape, CmBox(4, 4, 16, 12)
    RadiusNativeQuickTest.CheckNear PptNativeDriver.ReadFraction(shape), 0.0375, "automatic off resize retains parent fraction", 0.000001
    RadiusNativeQuickTest.CheckNear RadiusNativeCore.CurrentRadius(shape), 0.45, "automatic off resize changes parent radius with short side"
    AssertTag shape, LOCK_KEY, "0.300000", "automatic off resize leaves parent fixed tag unchanged"
    ApplyRadius slide, 0.6
    AssertParentRadius slide, parentId, 0.6, "automatic off permits explicit parent radius"
    RadiusNativeQuickTest.RequireScratch slide
    RadiusNativeLayout.SyncSlide slide
    RadiusNativeQuickTest.AssertGuardReleased
    AssertShapes slide, children, "automatic off keeps children unchanged"

    ' This checks tag serialization/re-reading only, not saving to disk.
    RadiusNativeLayout.ResetPending
    config = RadiusNativeLayout.SelectedConfig()
    RadiusNativeQuickTest.Check CLng(config(0)) = 2 And CLng(config(1)) = 2, "read serialized grid dimensions"
    RadiusNativeQuickTest.CheckNear CDbl(config(2)), 0.5, "read serialized padding"
    RadiusNativeQuickTest.CheckNear CDbl(config(3)), 0.3, "read serialized gap"
    RadiusNativeQuickTest.Check CStr(config(4)) = "same" And Not CBool(config(5)), "read serialized mode and automatic setting"
    AssertTag RadiusNativeQuickTest.FindShape(slide, parentId), SETTINGS_KEY, "1|2|2|0.500000|0.300000|same|0", "automatic setting serialized"
End Sub

Public Sub CaseStrictLayout(ByVal slide As Object)
    Dim parentId As Long, childIds As Variant, before As Collection, snapshot As Collection
    BeginCase slide
    CreateFamily slide, "QStrict", 4, parentId, childIds, 4
    SeedTag slide, RadiusNativeQuickTest.FindShape(slide, parentId), SETTINGS_KEY, "1|2|2|0.500000|0.300000|same|1"
    SelectId slide, parentId
    Set snapshot = RadiusNativeLayout.ReadUiSnapshot(RadiusNativeRelations.SelectedGroup())
    RadiusNativeQuickTest.Check CBool(snapshot("Automatic")), "strict fixture has existing automatic layout"
    RadiusNativeQuickTest.Check Not RadiusNativeLayout.ChildrenWritable, "late strict child makes layout unavailable"
    RadiusNativeQuickTest.Check Not RadiusNativeLayout.LinkedSelectionWritable, "late strict child blocks parent radius eligibility"
    Set before = SnapshotScene(slide)
    ExpectRejection slide, "radius", "before layout or linked radius: QStrictChild4"
    AssertScene slide, before, "late strict rejects linked radius before all writes"
    ExpectRejection slide, "layout", "before layout or linked radius: QStrictChild4"
    AssertScene slide, before, "late strict rejects layout before all writes"
    Stage slide, "rows", "1"
    Stage slide, "padding", "0.6"
    AssertScene slide, before, "strict draft still does not write shapes"
    RadiusNativeLayout.ResetPending
End Sub

Public Sub CaseProtectedParent(ByVal slide As Object)
    Dim parentId As Long, childIds As Variant, parentState As Collection
    BeginCase slide
    CreateFamily slide, "QProtected", 4, parentId, childIds, 0, True
    SelectId slide, parentId
    Set parentState = SnapshotShape(RadiusNativeQuickTest.FindShape(slide, parentId))
    Stage slide, "rows", "2"
    Stage slide, "padding", "0.5"
    Stage slide, "gap", "0.3"
    RadiusNativeQuickTest.Check RadiusNativeLayout.ChildrenWritable, "protected parent still permits unprotected child layout"
    ApplyLayout slide
    AssertGrid slide, childIds, 2.5, 2.5, 8.15, 6.15, 5.35, 3.35, "protected parent child grid"
    AssertChildRadius slide, childIds, 1, "protected parent child radius"
    AssertShape RadiusNativeQuickTest.FindShape(slide, parentId), parentState, "layout retains protected parent state", LAYOUT_TAGS
    AssertTag RadiusNativeQuickTest.FindShape(slide, parentId), STRICT_KEY, "1", "protected parent strict retained"
    AssertTag RadiusNativeQuickTest.FindShape(slide, parentId), LOCK_KEY, "1.000000", "protected parent fixed retained"
End Sub

Public Sub CaseInvalidLayout(ByVal slide As Object)
    Dim parentId As Long, childIds As Variant, before As Collection
    BeginCase slide
    CreateFamily slide, "QInvalid", 4, parentId, childIds
    SelectId slide, parentId
    Stage slide, "rows", "2"
    Stage slide, "padding", "10"
    Set before = SnapshotScene(slide)
    ExpectRejection slide, "layout", "Padding/gap leave no space for children."
    AssertScene slide, before, "insufficient space rejects without writes"
    RadiusNativeLayout.ResetPending
    Stage slide, "rows", "2"
    Stage slide, "padding", "0.5"
    Stage slide, "gap", "0.3"
    RadiusNativeQuickTest.FixtureRotation RadiusNativeQuickTest.FindShape(slide, CLng(childIds(4))), 15
    Set before = SnapshotScene(slide)
    ExpectRejection slide, "layout", "do not support rotated or flipped members/groups."
    AssertScene slide, before, "rotated child rejects without writes"
    RadiusNativeLayout.ResetPending
End Sub

Public Sub CaseIndependentRadius(ByVal slide As Object)
    Dim first As Object, second As Object, child As Object, plain As Object, states As New Collection, before As Collection
    Dim firstId As Long, plainId As Long
    BeginCase slide
    Set first = RadiusNativeQuickTest.FixtureRound(slide, "QBrokenParentA", 2, 2, 6, 4, 0.6)
    Set second = RadiusNativeQuickTest.FixtureRound(slide, "QBrokenParentB", 10, 2, 6, 4, 0.6)
    Set child = RadiusNativeQuickTest.FixtureRound(slide, "QBrokenChild", 3, 3, 3, 2, 0.2)
    Set plain = RadiusNativeQuickTest.FixtureRound(slide, "QIndependent", 18, 2, 4, 3, 0.2)
    firstId = PptNativeDriver.ShapeId(first)
    plainId = PptNativeDriver.ShapeId(plain)
    SeedTag slide, first, RELATION_KEY, "G09"
    SeedTag slide, first, ROLE_KEY, "P"
    SeedTag slide, first, SETTINGS_KEY, "1|1|1|0.300000|0.200000|same|1"
    SeedTag slide, second, RELATION_KEY, "G09"
    SeedTag slide, second, ROLE_KEY, "P"
    SeedTag slide, second, SETTINGS_KEY, "1|1|1|0.300000|0.200000|same|1"
    SeedTag slide, child, RELATION_KEY, "G09"
    SeedTag slide, child, ROLE_KEY, "C1"
    SeedTag slide, plain, "customPlain", "KeepPlainCase"
    SeedTag slide, plain, LOCK_KEY, "0.200000"
    ' A non-parent's leftover settings must not acquire native parent rules.
    SeedTag slide, plain, SETTINGS_KEY, "leftover invalid settings"
    states.Add SnapshotShape(first)
    states.Add SnapshotShape(second)
    states.Add SnapshotShape(child)
    SelectId slide, plainId
    ApplyRadius slide, 0.3
    RadiusNativeQuickTest.CheckNear RadiusNativeCore.CurrentRadius(RadiusNativeQuickTest.FindShape(slide, plainId)), 0.3, "independent radius beside duplicate parents"
    AssertTag RadiusNativeQuickTest.FindShape(slide, plainId), LOCK_KEY, "0.300000", "independent radius updates existing fixed value"
    AssertTag RadiusNativeQuickTest.FindShape(slide, plainId), SETTINGS_KEY, "leftover invalid settings", "non-parent leftover settings retained"
    AssertTag RadiusNativeQuickTest.FindShape(slide, plainId), "customPlain", "KeepPlainCase", "independent custom tag retained"
    AssertBox slide, plainId, 18, 2, 4, 3, "independent geometry retained"
    AssertShapes slide, states, "independent edit leaves broken relationship intact"
    SelectId slide, firstId
    Set before = SnapshotScene(slide)
    ExpectRejection slide, "radius", "Duplicate relationship parents: G09"
    AssertScene slide, before, "broken linked parent rejects without repair or writes"
End Sub

Private Sub BeginCase(ByVal slide As Object)
    RadiusNativeQuickTest.RequireScratch slide
    RadiusNativeQuickTest.Check PptNativeDriver.SlideRoots(slide).Count = 0, "business case starts with a blank scratch slide"
End Sub

Private Sub SeedTag(ByVal slide As Object, ByVal shape As Object, ByVal key As String, ByVal value As String)
    RadiusNativeQuickTest.RequireScratch slide
    PptNativeDriver.AddTag shape, key, value
End Sub

Private Sub SelectId(ByVal slide As Object, ByVal id As Long)
    RadiusNativeQuickTest.RequireScratch slide
    PptNativeDriver.SelectObject RadiusNativeQuickTest.FindShape(slide, id)
End Sub

Private Sub SelectNamedRoot(ByVal slide As Object, ByVal name As String)
    RadiusNativeQuickTest.RequireScratch slide
    PptNativeDriver.SelectObject RootNamed(slide, name)
End Sub

Private Function RootNamed(ByVal slide As Object, ByVal name As String) As Object
    Dim shape As Object
    For Each shape In PptNativeDriver.SlideRoots(slide)
        If PptNativeDriver.ShapeName(shape) = name Then Set RootNamed = shape: Exit Function
    Next shape
    Err.Raise 5, "RadiusNativeQuickRelations", "Cannot find scratch group: " & name
End Function

Private Sub CreateFamily(ByVal slide As Object, ByVal prefix As String, ByVal count As Long, ByRef parentId As Long, ByRef childIds As Variant, Optional ByVal strictChild As Long = 0, Optional ByVal protectedParent As Boolean = False)
    Dim parent As Object, child As Object, selected As New Collection, i As Long
    Set parent = RadiusNativeQuickTest.FixtureRound(slide, prefix & "Parent", 2, 2, 12, 8, 1)
    parentId = PptNativeDriver.ShapeId(parent)
    SeedTag slide, parent, LOCK_KEY, "1.000000"
    SeedTag slide, parent, "customParent", "KeepParentCase"
    If protectedParent Then SeedTag slide, parent, STRICT_KEY, "1"
    ReDim childIds(1 To count)
    For i = 1 To count
        Set child = RadiusNativeQuickTest.FixtureRound(slide, prefix & "Child" & CStr(i), 16 + ((i - 1) Mod 2) * 4, 2 + ((i - 1) \ 2) * 3, 3, 2, 0.2)
        childIds(i) = PptNativeDriver.ShapeId(child)
        SeedTag slide, child, LOCK_KEY, "0.200000"
        SeedTag slide, child, "customChild", "KeepChildCase" & CStr(i)
        If strictChild = i Then SeedTag slide, child, STRICT_KEY, "1"
        selected.Add child
    Next i
    SelectId slide, parentId
    RadiusNativeRelations.MarkParent
    RadiusNativeQuickTest.RequireScratch slide
    PptNativeDriver.SelectShapes slide, selected
    RadiusNativeRelations.BindChildren
    RadiusNativeQuickTest.AssertGuardReleased
    SelectId slide, parentId
End Sub

Private Sub Stage(ByVal slide As Object, ByVal key As String, ByVal text As String)
    RadiusNativeQuickTest.RequireScratch slide
    RadiusNativeLayout.SetParameter key, text
End Sub

Private Sub ApplyLayout(ByVal slide As Object)
    RadiusNativeQuickTest.RequireScratch slide
    RadiusNativeLayout.ApplySelected
    RadiusNativeQuickTest.AssertGuardReleased
End Sub

Private Sub ApplyRadius(ByVal slide As Object, ByVal value As Double)
    RadiusNativeQuickTest.RequireScratch slide
    RadiusNativeCore.ApplySelection value, "cm"
    RadiusNativeQuickTest.AssertGuardReleased
End Sub

Private Sub ChangeMode(ByVal slide As Object, ByVal mode As String)
    RadiusNativeQuickTest.RequireScratch slide
    RadiusNativeLayout.ChangeMode mode
    RadiusNativeQuickTest.AssertGuardReleased
End Sub

Private Sub ExpectRejection(ByVal slide As Object, ByVal action As String, ByVal reason As String)
    Dim errorNumber As Long, errorText As String
    RadiusNativeQuickTest.RequireScratch slide
    On Error GoTo Rejected
    Select Case action
        Case "bind": RadiusNativeRelations.BindChildren
        Case "radius": RadiusNativeCore.ApplySelection 0.5, "cm"
        Case "layout": RadiusNativeLayout.ApplySelected
        Case Else: Err.Raise 5, , "Unknown quick-check rejection action."
    End Select
    On Error GoTo 0
    RadiusNativeQuickTest.Check False, "Expected rejection: " & action & ": " & reason
    Exit Sub
Rejected:
    errorNumber = Err.Number
    errorText = Err.Description
    On Error GoTo 0
    RadiusNativeQuickTest.Check errorNumber <> 0 And InStr(1, errorText, reason, vbBinaryCompare) > 0, "Expected " & reason & "; received: " & errorText
    RadiusNativeQuickTest.AssertGuardReleased
End Sub

Private Function CmBox(ByVal x As Double, ByVal y As Double, ByVal width As Double, ByVal height As Double) As Variant
    CmBox = Array(x * RadiusNativeCore.PT_PER_CM, y * RadiusNativeCore.PT_PER_CM, width * RadiusNativeCore.PT_PER_CM, height * RadiusNativeCore.PT_PER_CM)
End Function

Private Function SnapshotShape(ByVal shape As Object) As Collection
    Dim result As New Collection, fraction As Double, geometry As Long, shapeType As Long
    shapeType = PptNativeDriver.ShapeType(shape)
    If shapeType = 1 Then geometry = PptNativeDriver.GeometryType(shape)
    result.Add PptNativeDriver.ShapeId(shape), "Id"
    result.Add PptNativeDriver.ShapeName(shape), "Name"
    result.Add shapeType, "Type"
    result.Add geometry, "Geometry"
    result.Add PptNativeDriver.ShapeBox(shape), "Box"
    result.Add PptNativeDriver.SnapshotTags(shape), "Tags"
    result.Add PptNativeDriver.HasTransform(shape), "Transform"
    If geometry = 5 Then fraction = PptNativeDriver.ReadFraction(shape)
    result.Add fraction, "Fraction"
    Set SnapshotShape = result
End Function

Private Function SnapshotIds(ByVal slide As Object, ByVal ids As Variant) As Collection
    Dim result As New Collection, id As Variant
    For Each id In ids
        result.Add SnapshotShape(RadiusNativeQuickTest.FindShape(slide, CLng(id)))
    Next id
    Set SnapshotIds = result
End Function

Private Function SnapshotScene(ByVal slide As Object) As Collection
    Dim result As New Collection, shape As Object
    For Each shape In PptNativeDriver.SlideRoots(slide)
        CollectScene shape, result, 0
    Next shape
    Set SnapshotScene = result
End Function

Private Sub CollectScene(ByVal shape As Object, ByVal result As Collection, ByVal depth As Long)
    Dim child As Object, previous As Collection, id As Long
    If depth > 64 Then Err.Raise 5, , "Quick-check snapshot nesting is too deep."
    id = PptNativeDriver.ShapeId(shape)
    Set previous = StoredShape(result, id)
    If Not previous Is Nothing Then Exit Sub
    result.Add SnapshotShape(shape), "S" & CStr(id)
    If PptNativeDriver.ShapeType(shape) = 6 Then
        For Each child In PptNativeDriver.Children(shape)
            CollectScene child, result, depth + 1
        Next child
    End If
End Sub

Private Function StoredShape(ByVal states As Collection, ByVal id As Long) As Collection
    Dim errorNumber As Long, errorText As String
    On Error GoTo Missing
    Set StoredShape = states("S" & CStr(id))
    Exit Function
Missing:
    errorNumber = Err.Number
    errorText = Err.Description
    If errorNumber <> 5 Then Err.Raise errorNumber, "RadiusNativeQuickRelations", errorText
End Function

Private Sub AssertScene(ByVal slide As Object, ByVal before As Collection, ByVal detail As String)
    Dim after As Collection
    Set after = SnapshotScene(slide)
    RadiusNativeQuickTest.Check after.Count = before.Count, detail & ": object count"
    AssertShapes slide, before, detail
End Sub

Private Sub AssertShapes(ByVal slide As Object, ByVal before As Collection, ByVal detail As String, Optional ByVal ignoredTags As String = "", Optional ByVal boxes As Boolean = True)
    Dim state As Collection
    For Each state In before
        AssertShape RadiusNativeQuickTest.FindShape(slide, CLng(state("Id"))), state, detail, ignoredTags, boxes
    Next state
End Sub

Private Sub AssertGroup(ByVal slide As Object, ByVal name As String, ByVal before As Collection, ByVal detail As String)
    AssertShape RootNamed(slide, name), before, detail, "", True, False
End Sub

Private Sub AssertShape(ByVal shape As Object, ByVal before As Collection, ByVal detail As String, Optional ByVal ignoredTags As String = "", Optional ByVal boxes As Boolean = True, Optional ByVal id As Boolean = True)
    Dim oldBox As Variant, actualBox As Variant, i As Long
    If id Then RadiusNativeQuickTest.Check PptNativeDriver.ShapeId(shape) = CLng(before("Id")), detail & ": leaf/object ID"
    RadiusNativeQuickTest.Check PptNativeDriver.ShapeName(shape) = CStr(before("Name")), detail & ": shape name"
    RadiusNativeQuickTest.Check PptNativeDriver.ShapeType(shape) = CLng(before("Type")), detail & ": shape type"
    If CLng(before("Type")) = 1 Then RadiusNativeQuickTest.Check PptNativeDriver.GeometryType(shape) = CLng(before("Geometry")), detail & ": geometry type"
    RadiusNativeQuickTest.Check PptNativeDriver.HasTransform(shape) = CBool(before("Transform")), detail & ": transform state"
    If boxes Then
        oldBox = before("Box")
        actualBox = PptNativeDriver.ShapeBox(shape)
        For i = 0 To 3
            RadiusNativeQuickTest.CheckNear CDbl(actualBox(i)) / RadiusNativeCore.PT_PER_CM, CDbl(oldBox(i)) / RadiusNativeCore.PT_PER_CM, detail & ": geometry " & CStr(i)
        Next i
    End If
    If CLng(before("Geometry")) = 5 Then RadiusNativeQuickTest.CheckNear PptNativeDriver.ReadFraction(shape), CDbl(before("Fraction")), detail & ": fraction", 0.000001
    AssertTags shape, before("Tags"), ignoredTags, detail
End Sub

Private Function IgnoreTag(ByVal key As String, ByVal ignored As String) As Boolean
    IgnoreTag = (InStr(1, "|" & UCase$(ignored) & "|", "|" & UCase$(key) & "|", vbBinaryCompare) > 0)
End Function

Private Function HasSavedTag(ByVal tags As Collection, ByVal key As String, ByVal value As String) As Boolean
    Dim pair As Variant
    For Each pair In tags
        If StrComp(CStr(pair(0)), key, vbTextCompare) = 0 Then HasSavedTag = (CStr(pair(1)) = value): Exit Function
    Next pair
End Function

Private Sub AssertTags(ByVal shape As Object, ByVal before As Collection, ByVal ignored As String, ByVal detail As String)
    Dim pair As Variant, present As Boolean, value As String
    For Each pair In before
        If Not IgnoreTag(CStr(pair(0)), ignored) Then
            value = PptNativeDriver.ReadTagState(shape, CStr(pair(0)), present)
            RadiusNativeQuickTest.Check present And value = CStr(pair(1)), detail & ": retain tag " & CStr(pair(0))
        End If
    Next pair
    For Each pair In PptNativeDriver.SnapshotTags(shape)
        If Not IgnoreTag(CStr(pair(0)), ignored) Then RadiusNativeQuickTest.Check HasSavedTag(before, CStr(pair(0)), CStr(pair(1))), detail & ": no extra/changed tag " & CStr(pair(0))
    Next pair
End Sub

Private Sub AssertTag(ByVal shape As Object, ByVal key As String, ByVal expected As String, ByVal detail As String)
    Dim present As Boolean, value As String
    value = PptNativeDriver.ReadTagState(shape, key, present)
    RadiusNativeQuickTest.Check present And value = expected, detail & ": " & key
End Sub

Private Sub AssertAbsentTag(ByVal shape As Object, ByVal key As String, ByVal detail As String)
    RadiusNativeQuickTest.Check Not PptNativeDriver.HasTag(shape, key), detail
End Sub

Private Sub AssertChildRadius(ByVal slide As Object, ByVal ids As Variant, ByVal radius As Double, ByVal detail As String)
    Dim id As Variant, shape As Object
    For Each id In ids
        Set shape = RadiusNativeQuickTest.FindShape(slide, CLng(id))
        RadiusNativeQuickTest.CheckNear RadiusNativeCore.CurrentRadius(shape), radius, detail
        AssertTag shape, LOCK_KEY, Replace(Format$(radius, "0.000000"), ",", "."), detail & ": fixed"
    Next id
End Sub

Private Sub AssertParentRadius(ByVal slide As Object, ByVal id As Long, ByVal radius As Double, ByVal detail As String)
    Dim shape As Object
    Set shape = RadiusNativeQuickTest.FindShape(slide, id)
    RadiusNativeQuickTest.CheckNear RadiusNativeCore.CurrentRadius(shape), radius, detail
    AssertTag shape, LOCK_KEY, Replace(Format$(radius, "0.000000"), ",", "."), detail & ": fixed"
End Sub

Private Sub AssertBox(ByVal slide As Object, ByVal id As Long, ByVal x As Double, ByVal y As Double, ByVal width As Double, ByVal height As Double, ByVal detail As String)
    Dim box As Variant
    box = PptNativeDriver.ShapeBox(RadiusNativeQuickTest.FindShape(slide, id))
    RadiusNativeQuickTest.CheckNear CDbl(box(0)) / RadiusNativeCore.PT_PER_CM, x, detail & ": left"
    RadiusNativeQuickTest.CheckNear CDbl(box(1)) / RadiusNativeCore.PT_PER_CM, y, detail & ": top"
    RadiusNativeQuickTest.CheckNear CDbl(box(2)) / RadiusNativeCore.PT_PER_CM, width, detail & ": width"
    RadiusNativeQuickTest.CheckNear CDbl(box(3)) / RadiusNativeCore.PT_PER_CM, height, detail & ": height"
End Sub

Private Sub AssertGrid(ByVal slide As Object, ByVal ids As Variant, ByVal x1 As Double, ByVal y1 As Double, ByVal x2 As Double, ByVal y2 As Double, ByVal width As Double, ByVal height As Double, ByVal detail As String)
    Dim i As Long, x As Double, y As Double
    For i = 1 To 4
        x = x1: y = y1
        If i = 2 Or i = 4 Then x = x2
        If i > 2 Then y = y2
        AssertBox slide, CLng(ids(i)), x, y, width, height, detail
    Next i
End Sub
