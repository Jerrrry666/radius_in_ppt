Attribute VB_Name = "RadiusNativeQuickTest"
Option Explicit

' Real-object checks run only in one caller-owned, unsaved presentation.
' The Ribbon holds its UI editing guard while this runner is active.
Private Const FIXTURE_PT As Double = 28.3464566929134
Private Const LOCK_KEY As String = "radiusLock_v1"
Private Const STRICT_KEY As String = "radiusLockStrict_v1"
Private Const CORE_CASE_COUNT As Long = 7
Private Const MEMORY_CASE_COUNT As Long = 13
Private ownedPresentation As Object
Private ownedWindow As Object
Private ownedName As String
Private caseSlide As Object
Private running As Boolean
Private assertionCount As Long
Private passedCases As Long
Private failedCases As Long
Private lastFailures As Long
Private failures As String

Public Function LastFailureCount() As Long
    LastFailureCount = lastFailures
End Function

Public Sub Check(ByVal condition As Boolean, ByVal detail As String)
    assertionCount = assertionCount + 1
    If Not condition Then Err.Raise 5, "RadiusNativeQuickTest", detail
End Sub

Public Sub CheckNear(ByVal actual As Double, ByVal expected As Double, ByVal detail As String, Optional ByVal tolerance As Double = 0.0005)
    Check Abs(actual - expected) <= tolerance, detail & " (actual " & CStr(actual) & ", expected " & CStr(expected) & ")."
End Sub

Private Sub RequireOwned()
    Dim current As Object
    If Not running Or ownedPresentation Is Nothing Then Err.Raise 5, , "No quick-check presentation is owned."
    Set current = PptNativeDriver.CurrentPresentation()
    If Not current Is ownedPresentation Then Err.Raise 5, , "Quick check refused: the active presentation is not its disposable fixture."
End Sub

Public Sub RequireScratch(ByVal slide As Object)
    Dim current As Object
    RequireOwned
    If caseSlide Is Nothing Then Err.Raise 5, , "No quick-check slide is owned."
    If Not slide Is caseSlide Then Err.Raise 5, , "Quick check refused a slide outside the current case."
    Set current = PptNativeDriver.CurrentSlide()
    If Not current Is slide Then Err.Raise 5, , "Quick check refused: the active slide is not the current fixture."
End Sub

Public Sub AssertGuardReleased()
    Check Not RadiusNativeCore.IsTransactionActive(), "The Core transaction guard remained active."
End Sub

Private Function FindInShape(ByVal shape As Object, ByVal id As Long, ByVal depth As Long) As Object
    Dim child As Object, found As Object
    If depth > 64 Then Err.Raise 5, , "Quick-check fixture nesting is too deep."
    If PptNativeDriver.ShapeId(shape) = id Then
        Set FindInShape = shape
    ElseIf PptNativeDriver.ShapeType(shape) = 6 Then
        For Each child In PptNativeDriver.Children(shape)
            Set found = FindInShape(child, id, depth + 1)
            If Not found Is Nothing Then Set FindInShape = found: Exit Function
        Next child
    End If
End Function

Public Function FindShape(ByVal slide As Object, ByVal id As Long) As Object
    Dim shape As Object, found As Object
    For Each shape In PptNativeDriver.SlideRoots(slide)
        Set found = FindInShape(shape, id, 0)
        If Not found Is Nothing Then Set FindShape = found: Exit Function
    Next shape
End Function

Private Sub RequireFixtureShape(ByVal shape As Object)
    Dim found As Object
    RequireScratch caseSlide
    Set found = FindShape(caseSlide, PptNativeDriver.ShapeId(shape))
    If found Is Nothing Then Err.Raise 5, , "A fixture write target is not on the owned slide."
    If Not found Is shape Then Err.Raise 5, , "A fixture write target belongs to another presentation."
End Sub

Public Function FixtureShape(ByVal slide As Object, ByVal name As String, ByVal geometryType As Long, ByVal xCm As Double, ByVal yCm As Double, ByVal wCm As Double, ByVal hCm As Double) As Object
    Dim shape As Object, tags As New Collection
    RequireScratch slide
    If wCm <= 0 Or hCm <= 0 Then Err.Raise 5, , "Fixture dimensions must be positive."
    Set shape = PptNativeDriver.AddShape(slide, geometryType, xCm * FIXTURE_PT, yCm * FIXTURE_PT, wCm * FIXTURE_PT, hCm * FIXTURE_PT)
    PptNativeDriver.RestoreMetadata shape, name, tags
    Set FixtureShape = shape
End Function

Public Function FixtureRound(ByVal slide As Object, ByVal name As String, ByVal xCm As Double, ByVal yCm As Double, ByVal wCm As Double, ByVal hCm As Double, ByVal radiusCm As Double) As Object
    Dim shape As Object, shortCm As Double
    Set shape = FixtureShape(slide, name, 5, xCm, yCm, wCm, hCm)
    shortCm = wCm
    If hCm < shortCm Then shortCm = hCm
    If radiusCm < 0 Or radiusCm > shortCm / 2# Then Err.Raise 5, , "Fixture radius is out of range."
    PptNativeDriver.WriteFraction shape, radiusCm / shortCm
    Set FixtureRound = shape
End Function

Public Function FixtureGroup(ByVal slide As Object, ByVal name As String, ByVal items As Collection) As Object
    Dim shape As Object, grouped As Object, tags As New Collection
    RequireScratch slide
    For Each shape In items
        RequireFixtureShape shape
        If Not PptNativeDriver.IsTopLevel(slide, shape) Then Err.Raise 5, , "Fixture grouping requires top-level members."
    Next shape
    Set grouped = PptNativeDriver.Regroup(slide, items)
    PptNativeDriver.RestoreMetadata grouped, name, tags
    Set FixtureGroup = grouped
End Function

Public Sub FixtureRotation(ByVal shape As Object, ByVal degrees As Double)
    RequireFixtureShape shape
    PptNativeDriver.SetRotation shape, degrees
End Sub

Private Sub SuspendSessions()
    RadiusNativeRelations.SuspendTestSession
    RadiusNativeLayout.SuspendTestSession
End Sub

Private Function ErrorLine(ByVal number As Long, ByVal description As String) As String
    ErrorLine = "Err " & CStr(number) & ": " & description
End Function

Private Sub RecordFailure(ByVal name As String, ByVal number As Long, ByVal description As String, Optional ByVal isCase As Boolean = True)
    Dim line As String
    lastFailures = lastFailures + 1
    If isCase Then failedCases = failedCases + 1
    line = "FAIL " & name & " - " & ErrorLine(number, description)
    If failures <> "" Then failures = failures & vbCrLf
    failures = failures & line
    Debug.Print "[NativeQuickCheck] " & line
End Sub

Private Sub RecordPass(ByVal name As String, ByVal before As Long)
    passedCases = passedCases + 1
    Debug.Print "[NativeQuickCheck] PASS " & name & " (" & CStr(assertionCount - before) & " assertions)."
End Sub

Private Function SessionSame(ByVal left As Variant, ByVal right As Variant) As Boolean
    Dim i As Long, leftObject As Object, rightObject As Object
    If IsObject(left) Or IsObject(right) Then
        If Not IsObject(left) Or Not IsObject(right) Then Exit Function
        Set leftObject = left
        Set rightObject = right
        SessionSame = (leftObject Is rightObject)
    ElseIf IsArray(left) Or IsArray(right) Then
        If Not IsArray(left) Or Not IsArray(right) Then Exit Function
        If LBound(left) <> LBound(right) Or UBound(left) <> UBound(right) Then Exit Function
        For i = LBound(left) To UBound(left)
            If Not SessionSame(left(i), right(i)) Then Exit Function
        Next i
        SessionSame = True
    ElseIf VarType(left) = VarType(right) Then
        If IsEmpty(left) Then SessionSame = True Else SessionSame = (left = right)
    End If
End Function

Private Function RunAlgorithms() As Long
    Dim count As Long, number As Long, description As String
    On Error Resume Next
    Err.Clear
    count = RadiusNativeCore.SelfTest()
    number = Err.Number
    description = Err.Description
    On Error GoTo 0
    If number = 0 Then RunAlgorithms = count Else RecordFailure "Core algorithms", number, description, False
    On Error Resume Next
    Err.Clear
    count = RadiusNativeLayout.SelfTest()
    number = Err.Number
    description = Err.Description
    On Error GoTo 0
    If number = 0 Then RunAlgorithms = RunAlgorithms + count Else RecordFailure "Layout algorithms", number, description, False
    Debug.Print "[NativeQuickCheck] Algorithm checks passed: " & CStr(RunAlgorithms)
End Function

Private Sub SelectCaseShapes(ByVal slide As Object, ByVal shapes As Collection)
    Dim shape As Object
    RequireScratch slide
    For Each shape In shapes
        RequireFixtureShape shape
    Next shape
    PptNativeDriver.SelectObjects shapes
End Sub

Private Sub SeedTag(ByVal shape As Object, ByVal key As String, ByVal value As String)
    RequireFixtureShape shape
    PptNativeDriver.AddTag shape, key, value
End Sub

Private Sub CheckBox(ByVal shape As Object, ByVal expected As Variant, ByVal detail As String)
    Dim actual As Variant, i As Long
    actual = PptNativeDriver.ShapeBox(shape)
    For i = 0 To 3
        CheckNear CDbl(actual(i)), CDbl(expected(i)), detail & " box(" & CStr(i) & ")", 0.02
    Next i
End Sub

Private Sub CheckTags(ByVal shape As Object, ByVal expected As Collection, ByVal detail As String)
    Dim pair As Variant, value As String, present As Boolean, actual As Collection
    Set actual = PptNativeDriver.SnapshotTags(shape)
    Check actual.Count = expected.Count, detail & " tag count changed."
    For Each pair In expected
        value = PptNativeDriver.ReadTagState(shape, CStr(pair(0)), present)
        Check present, detail & " lost tag " & CStr(pair(0)) & "."
        Check value = CStr(pair(1)), detail & " changed tag " & CStr(pair(0)) & "."
    Next pair
End Sub

Private Function CaptureShapeState(ByVal shape As Object) As Collection
    Dim result As New Collection
    result.Add PptNativeDriver.ShapeId(shape), "Id"
    result.Add PptNativeDriver.ShapeName(shape), "Name"
    result.Add PptNativeDriver.ShapeBox(shape), "Box"
    result.Add PptNativeDriver.SnapshotTags(shape), "Tags"
    If PptNativeDriver.ShapeType(shape) <> 6 Then result.Add PptNativeDriver.ReadFraction(shape), "Fraction"
    Set CaptureShapeState = result
End Function

Private Sub CheckShapeState(ByVal shape As Object, ByVal expected As Collection, ByVal detail As String, Optional ByVal fraction As Boolean = True)
    Check PptNativeDriver.ShapeId(shape) = CLng(expected("Id")), detail & " ID changed."
    Check PptNativeDriver.ShapeName(shape) = CStr(expected("Name")), detail & " name changed."
    CheckBox shape, expected("Box"), detail
    CheckTags shape, expected("Tags"), detail
    If fraction And PptNativeDriver.ShapeType(shape) <> 6 Then CheckNear PptNativeDriver.ReadFraction(shape), CDbl(expected("Fraction")), detail & " fraction"
End Sub

Private Function TagsAfterFixed(ByVal before As Collection) As Collection
    Dim result As New Collection, pair As Variant
    For Each pair In before
        If StrComp(CStr(pair(0)), LOCK_KEY, vbTextCompare) = 0 Then
            result.Add Array(CStr(pair(0)), "0.300000")
        Else
            result.Add pair
        End If
    Next pair
    Set TagsAfterFixed = result
End Function

Private Sub ExpectSelectionRejected(ByVal value As Double, ByVal unit As String, ByVal reason As String)
    Dim number As Long, description As String
    RequireScratch caseSlide
    On Error Resume Next
    Err.Clear
    RadiusNativeCore.ApplySelection value, unit
    number = Err.Number
    description = Err.Description
    On Error GoTo 0
    Check number <> 0, "A protected radius batch was accepted."
    Check InStr(1, description, reason, vbTextCompare) > 0, "Unexpected radius rejection: " & ErrorLine(number, description)
    Debug.Print "[NativeQuickCheck] Expected radius rejection: " & ErrorLine(number, description)
    AssertGuardReleased
End Sub

Private Sub ExpectMetadataRejected(ByVal slide As Object, ByVal plan As Collection)
    Dim number As Long, description As String
    RequireScratch slide
    On Error Resume Next
    Err.Clear
    RadiusNativeCore.ApplyMetadataPlan slide, plan
    number = Err.Number
    description = Err.Description
    On Error GoTo 0
    Check number <> 0, "Reserved protection metadata was accepted."
    Check InStr(1, description, "cannot change radius protection", vbTextCompare) > 0, "Unexpected metadata rejection: " & ErrorLine(number, description)
    Debug.Print "[NativeQuickCheck] Expected metadata rejection: " & ErrorLine(number, description)
    AssertGuardReleased
End Sub

Private Sub ExpectShapePlanRejected(ByVal slide As Object, ByVal plan As Collection, ByVal reason As String)
    Dim number As Long, description As String
    RequireScratch slide
    On Error Resume Next
    Err.Clear
    RadiusNativeCore.ApplyShapePlan slide, plan
    number = Err.Number
    description = Err.Description
    On Error GoTo 0
    Check number <> 0, "An invalid or protected shape plan was accepted."
    Check InStr(1, description, reason, vbTextCompare) > 0, "Unexpected shape-plan rejection: " & ErrorLine(number, description)
    Debug.Print "[NativeQuickCheck] Expected shape-plan rejection: " & ErrorLine(number, description)
    AssertGuardReleased
End Sub

Private Sub CaseUnits(ByVal slide As Object)
    Dim first As Object, second As Object, arrow As Object, selected As New Collection, arrowState As Collection
    RequireScratch slide
    Set first = FixtureRound(slide, "QuickUnits_A", 2, 2, 6, 4, 0.2)
    Set second = FixtureRound(slide, "QuickUnits_B", 10, 2, 4, 2, 0.1)
    Set arrow = FixtureShape(slide, "QuickUnits_Arrow", 33, 2, 8, 5, 2)
    PptNativeDriver.WriteFraction arrow, 0.27
    SeedTag arrow, STRICT_KEY, "1"
    SeedTag arrow, "QuickCustom", "Arrow Metadata"
    Set arrowState = CaptureShapeState(arrow)
    selected.Add first
    selected.Add second
    selected.Add arrow
    SelectCaseShapes slide, selected
    RadiusNativeCore.ApplySelection 0.3, "cm"
    AssertGuardReleased
    CheckNear PptNativeDriver.ReadFraction(first), 0.075, "Shared cm on short side 4 cm"
    CheckNear PptNativeDriver.ReadFraction(second), 0.15, "Shared cm on short side 2 cm"
    CheckNear RadiusNativeCore.CurrentRadius(first), 0.3, "Read R after cm write"
    RequireScratch slide
    RadiusNativeCore.ApplySelection 20, "%"
    AssertGuardReleased
    CheckNear PptNativeDriver.ReadFraction(first), 0.2, "Percent on first shape"
    CheckNear PptNativeDriver.ReadFraction(second), 0.2, "Percent on second shape"
    RequireScratch slide
    RadiusNativeCore.ApplySelection 0, "cm"
    AssertGuardReleased
    CheckNear PptNativeDriver.ReadFraction(first), 0, "Zero on first shape"
    CheckNear PptNativeDriver.ReadFraction(second), 0, "Zero on second shape"
    RequireScratch slide
    RadiusNativeCore.ApplySelection 3, "cm"
    AssertGuardReleased
    CheckNear PptNativeDriver.ReadFraction(first), 0.5, "Cm clamp on first shape"
    CheckNear PptNativeDriver.ReadFraction(second), 0.5, "Cm clamp on second shape"
    RequireScratch slide
    RadiusNativeCore.ApplySelection 60, "%"
    AssertGuardReleased
    CheckNear PptNativeDriver.ReadFraction(first), 0.5, "Percent clamp on first shape"
    CheckNear PptNativeDriver.ReadFraction(second), 0.5, "Percent clamp on second shape"
    CheckShapeState arrow, arrowState, "Mixed non-round arrow"
End Sub

Private Sub CaseProtection(ByVal slide As Object)
    Dim first As Object, second As Object, selected As New Collection, fixed As String
    RequireScratch slide
    Set first = FixtureRound(slide, "QuickProtect_A", 2, 2, 6, 4, 0.2)
    Set second = FixtureRound(slide, "QuickProtect_B", 10, 2, 4, 2, 0.1)
    SeedTag first, LOCK_KEY, "0.777000"
    fixed = PptNativeDriver.ReadTag(first, LOCK_KEY)
    selected.Add first
    selected.Add second
    SelectCaseShapes slide, selected
    RadiusNativeCore.SetProtection True
    AssertGuardReleased
    Check PptNativeDriver.ReadTag(first, STRICT_KEY) = "1", "Protection did not set first strict tag."
    Check PptNativeDriver.ReadTag(second, STRICT_KEY) = "1", "Protection did not set second strict tag."
    Check PptNativeDriver.ReadTag(first, LOCK_KEY) = fixed, "Protection replaced an existing fixed radius."
    Check PptNativeDriver.HasTag(second, LOCK_KEY), "Protection did not capture a missing fixed radius."
    CheckNear Val(PptNativeDriver.ReadTag(second, LOCK_KEY)), 0.1, "Captured fixed R"
    CheckNear PptNativeDriver.ReadFraction(first), 0.05, "Protect preserved first fraction"
    CheckNear PptNativeDriver.ReadFraction(second), 0.05, "Protect preserved second fraction"
    RequireScratch slide
    RadiusNativeCore.SetProtection False
    AssertGuardReleased
    Check Not PptNativeDriver.HasTag(first, STRICT_KEY), "Unprotect retained first strict tag."
    Check Not PptNativeDriver.HasTag(second, STRICT_KEY), "Unprotect retained second strict tag."
    Check PptNativeDriver.ReadTag(first, LOCK_KEY) = fixed, "Unprotect changed existing fixed R."
    CheckNear Val(PptNativeDriver.ReadTag(second, LOCK_KEY)), 0.1, "Unprotect preserved captured fixed R"
    RequireScratch slide
    RadiusNativeCore.ApplySelection 0.3, "cm"
    AssertGuardReleased
    Check PptNativeDriver.ReadTag(first, LOCK_KEY) = "0.300000", "Explicit radius did not update first fixed tag."
    Check PptNativeDriver.ReadTag(second, LOCK_KEY) = "0.300000", "Explicit radius did not update second fixed tag."
End Sub

Private Sub CaseLateStrict(ByVal slide As Object)
    Dim first As Object, second As Object, arrow As Object, selected As New Collection
    Dim firstState As Collection, secondState As Collection, arrowState As Collection
    RequireScratch slide
    Set first = FixtureRound(slide, "QuickStrict_First", 2, 2, 6, 4, 0.2)
    Set second = FixtureRound(slide, "QuickStrict_Last", 10, 2, 4, 2, 0.1)
    Set arrow = FixtureShape(slide, "QuickStrict_Arrow", 33, 2, 8, 5, 2)
    SeedTag first, LOCK_KEY, "0.200000"
    selected.Add first
    selected.Add arrow
    selected.Add second
    SelectCaseShapes slide, selected
    ' Added after selection to prove that a cached UI state cannot authorize it.
    SeedTag second, STRICT_KEY, "1"
    Set firstState = CaptureShapeState(first)
    Set secondState = CaptureShapeState(second)
    Set arrowState = CaptureShapeState(arrow)
    ExpectSelectionRejected 0.4, "cm", "protection"
    CheckShapeState first, firstState, "First target after late strict refusal"
    CheckShapeState second, secondState, "Last protected target after refusal"
    CheckShapeState arrow, arrowState, "Arrow after refused batch"
    RequireScratch slide
    RadiusNativeCore.SetProtection False
    AssertGuardReleased
    RadiusNativeCore.ApplySelection 0.4, "cm"
    AssertGuardReleased
    CheckNear PptNativeDriver.ReadFraction(first), 0.1, "A valid action after strict refusal"
    CheckNear PptNativeDriver.ReadFraction(second), 0.2, "Recovered public path after strict refusal"
End Sub

Private Sub CheckTagState(ByVal shape As Object, ByVal key As String, ByVal expectedPresent As Boolean, ByVal expectedValue As String, ByVal detail As String)
    Dim present As Boolean, value As String
    value = PptNativeDriver.ReadTagState(shape, key, present)
    Check present = expectedPresent, detail & " tag presence changed."
    Check value = expectedValue, detail & " tag value changed."
End Sub

Private Sub CheckChangedLeaf(ByVal shape As Object, ByVal state As Collection, ByVal detail As String, ByVal expectedShortCm As Double)
    Dim expectedTags As Collection
    Check PptNativeDriver.ShapeId(shape) = CLng(state("Id")), detail & " leaf ID changed."
    Check PptNativeDriver.ShapeName(shape) = CStr(state("Name")), detail & " leaf name changed."
    CheckBox shape, state("Box"), detail
    CheckNear PptNativeDriver.ReadFraction(shape), 0.3 / expectedShortCm, detail & " requested 0.3 cm R"
    Set expectedTags = TagsAfterFixed(state("Tags"))
    CheckTags shape, expectedTags, detail
End Sub

Private Sub CaseNestedGroup(ByVal slide As Object)
    Dim first As Object, second As Object, arrow As Object, inner As Object, outer As Object, freshInner As Object
    Dim innerItems As New Collection, outerItems As New Collection, selected As New Collection, roots As Collection
    Dim firstState As Collection, secondState As Collection, arrowState As Collection, outerState As Collection, innerTags As Collection
    Dim firstId As Long, secondId As Long, arrowId As Long, innerName As String, box As Variant
    Dim returned As Object, direct As Collection, shape As Object, innerCount As Long, secondCount As Long
    Dim emptyOuter As Boolean, emptyInner As Boolean, emptyLeaf As Boolean, outerValue As String, innerValue As String, leafValue As String
    RequireScratch slide
    Set first = FixtureRound(slide, "QuickGroup_A", 2, 2, 6, 4, 0.2)
    Set second = FixtureRound(slide, "QuickGroup_B", 10, 2, 4, 2, 0.1)
    Set arrow = FixtureShape(slide, "QuickGroup_Arrow", 33, 2, 8, 5, 2)
    firstId = PptNativeDriver.ShapeId(first)
    secondId = PptNativeDriver.ShapeId(second)
    arrowId = PptNativeDriver.ShapeId(arrow)
    SeedTag first, LOCK_KEY, "0.200000"
    SeedTag second, LOCK_KEY, "0.100000"
    SeedTag first, "QuickCustom", "MiXeD Leaf Value"
    SeedTag first, "QuickEmpty", ""
    leafValue = PptNativeDriver.ReadTagState(first, "QuickEmpty", emptyLeaf)
    SeedTag arrow, "QuickCustom", "Arrow Value"
    innerItems.Add first
    innerItems.Add arrow
    Set inner = FixtureGroup(slide, "QuickGroup_Inner", innerItems)
    SeedTag inner, "QuickCustom", "MiXeD Inner Value"
    SeedTag inner, "QuickEmpty", ""
    innerValue = PptNativeDriver.ReadTagState(inner, "QuickEmpty", emptyInner)
    innerName = PptNativeDriver.ShapeName(inner)
    Set innerTags = PptNativeDriver.SnapshotTags(inner)
    outerItems.Add inner
    outerItems.Add second
    Set outer = FixtureGroup(slide, "QuickGroup_Outer", outerItems)
    SeedTag outer, "QuickCustom", "MiXeD Outer Value"
    SeedTag outer, "QuickEmpty", ""
    outerValue = PptNativeDriver.ReadTagState(outer, "QuickEmpty", emptyOuter)
    box = PptNativeDriver.ShapeBox(outer)
    box(0) = 1# * FIXTURE_PT
    box(1) = 1# * FIXTURE_PT
    box(2) = CDbl(box(2)) * 1.35
    box(3) = CDbl(box(3)) * 1.2
    RequireFixtureShape outer
    PptNativeDriver.SetShapeBox outer, box
    Set first = FindShape(slide, firstId)
    Set second = FindShape(slide, secondId)
    Set arrow = FindShape(slide, arrowId)
    Check Not first Is Nothing, "Cannot locate the first scaled nested fixture leaf."
    Check Not second Is Nothing, "Cannot locate the second scaled fixture leaf."
    Check Not arrow Is Nothing, "Cannot locate the scaled nested arrow."
    Set firstState = CaptureShapeState(first)
    Set secondState = CaptureShapeState(second)
    Set arrowState = CaptureShapeState(arrow)
    Set outerState = CaptureShapeState(outer)
    selected.Add outer
    SelectCaseShapes slide, selected
    RadiusNativeCore.ApplySelection 0.3, "cm"
    AssertGuardReleased
    Set roots = PptNativeDriver.SelectionRoots()
    Check roots.Count = 1, "A group transaction did not restore one selected root."
    Set outer = roots(1)
    Check PptNativeDriver.ShapeType(outer) = 6, "The selected transaction result is not a group."
    Check PptNativeDriver.ShapeName(outer) = CStr(outerState("Name")), "Outer group name was lost."
    CheckBox outer, outerState("Box"), "Scaled outer group"
    CheckTags outer, outerState("Tags"), "Outer group"
    CheckTagState outer, "QuickEmpty", emptyOuter, outerValue, "Outer empty tag"
    Set first = FindShape(slide, firstId)
    Set second = FindShape(slide, secondId)
    Set arrow = FindShape(slide, arrowId)
    Check Not first Is Nothing, "The first leaf ID was lost after regroup."
    Check Not second Is Nothing, "The second leaf ID was lost after regroup."
    Check Not arrow Is Nothing, "The arrow ID was lost after regroup."
    CheckChangedLeaf first, firstState, "Scaled nested round leaf", 4.8
    CheckChangedLeaf second, secondState, "Scaled outer round leaf", 2.4
    CheckShapeState arrow, arrowState, "Scaled nested arrow"
    CheckTagState first, "QuickEmpty", emptyLeaf, leafValue, "Leaf empty tag"
    Debug.Print "[NativeQuickCheck] Empty-tag seed presence (outer/inner/leaf): " & CStr(emptyOuter) & "/" & CStr(emptyInner) & "/" & CStr(emptyLeaf)
    ' Last operation of this disposable case: inspect actual direct members.
    ' GroupItems is deliberately not used as a hierarchy oracle on Mac.
    RequireScratch slide
    PptNativeDriver.UngroupForTransaction outer, returned
    Set direct = PptNativeDriver.RangeMembers(returned)
    Check direct.Count = 2, "Fresh outer Ungroup did not return two direct members."
    For Each shape In direct
        If PptNativeDriver.ShapeType(shape) = 6 Then
            innerCount = innerCount + 1
            Set freshInner = shape
        ElseIf PptNativeDriver.ShapeId(shape) = secondId Then
            secondCount = secondCount + 1
            CheckBox shape, Array(11.8 * FIXTURE_PT, 1# * FIXTURE_PT, 5.4 * FIXTURE_PT, 2.4 * FIXTURE_PT), "Fresh outer round geometry"
        Else
            Check False, "Fresh outer Ungroup returned an unexpected direct leaf."
        End If
    Next shape
    Check innerCount = 1 And secondCount = 1, "The outer/inner group hierarchy changed."
    Check PptNativeDriver.ShapeName(freshInner) = innerName, "Inner group name was lost."
    CheckBox freshInner, Array(1# * FIXTURE_PT, 1# * FIXTURE_PT, 8.1 * FIXTURE_PT, 9.6 * FIXTURE_PT), "Fresh inner group geometry"
    CheckTags freshInner, innerTags, "Inner group"
    CheckTagState freshInner, "QuickEmpty", emptyInner, innerValue, "Inner empty tag"
    PptNativeDriver.UngroupForTransaction freshInner, returned
    Set direct = PptNativeDriver.RangeMembers(returned)
    Check direct.Count = 2, "Fresh inner Ungroup did not return two direct leaves."
    innerCount = 0
    secondCount = 0
    For Each shape In direct
        Check PptNativeDriver.ShapeType(shape) <> 6, "An unexpected third nesting level appeared."
        If PptNativeDriver.ShapeId(shape) = firstId Then
            innerCount = innerCount + 1
            CheckBox shape, Array(1# * FIXTURE_PT, 1# * FIXTURE_PT, 8.1 * FIXTURE_PT, 4.8 * FIXTURE_PT), "Fresh nested round geometry"
        End If
        If PptNativeDriver.ShapeId(shape) = arrowId Then
            secondCount = secondCount + 1
            CheckBox shape, Array(1# * FIXTURE_PT, 8.2 * FIXTURE_PT, 6.75 * FIXTURE_PT, 2.4 * FIXTURE_PT), "Fresh nested arrow geometry"
        End If
    Next shape
    Check innerCount = 1 And secondCount = 1, "Fresh inner member IDs changed."
End Sub

Private Sub CaseMetadata(ByVal slide As Object)
    Dim first As Object, second As Object, grouped As Object, members As New Collection, selected As New Collection
    Dim firstState As Collection, secondState As Collection, groupState As Collection, plan As Collection, roots As Collection
    Dim firstId As Long, secondId As Long
    RequireScratch slide
    Set first = FixtureRound(slide, "QuickMeta_A", 2, 2, 6, 4, 0.2)
    Set second = FixtureRound(slide, "QuickMeta_B", 10, 2, 4, 2, 0.1)
    firstId = PptNativeDriver.ShapeId(first)
    secondId = PptNativeDriver.ShapeId(second)
    SeedTag first, "QuickMeta", "OLD"
    SeedTag second, "QuickRemove", "REMOVE"
    SeedTag first, LOCK_KEY, "0.200000"
    SeedTag second, STRICT_KEY, "1"
    members.Add first
    members.Add second
    Set grouped = FixtureGroup(slide, "QuickMeta_Group", members)
    SeedTag grouped, "QuickGroupTag", "GROUP VALUE"
    Set first = FindShape(slide, firstId)
    Set second = FindShape(slide, secondId)
    Set firstState = CaptureShapeState(first)
    Set secondState = CaptureShapeState(second)
    Set groupState = CaptureShapeState(grouped)
    selected.Add grouped
    SelectCaseShapes slide, selected
    Set plan = New Collection
    plan.Add Array(firstId, "QuickMeta", "SHOULD NOT WRITE", False)
    plan.Add Array(secondId, "RADIUSLOCKSTRICT_V1", "", True)
    ExpectMetadataRejected slide, plan
    CheckShapeState first, firstState, "Metadata strict refusal first leaf"
    CheckShapeState second, secondState, "Metadata strict refusal second leaf"
    CheckShapeState grouped, groupState, "Metadata strict refusal group"
    Set plan = New Collection
    plan.Add Array(firstId, "QuickMeta", "SHOULD NOT WRITE", False)
    plan.Add Array(secondId, "RADIUSLOCK_V1", "0.900000", False)
    ExpectMetadataRejected slide, plan
    CheckShapeState first, firstState, "Metadata fixed refusal first leaf"
    CheckShapeState second, secondState, "Metadata fixed refusal second leaf"
    CheckShapeState grouped, groupState, "Metadata fixed refusal group"
    Set plan = New Collection
    plan.Add Array(firstId, "QuickMeta", "NEW", False)
    plan.Add Array(secondId, "QuickRemove", "", True)
    RequireScratch slide
    RadiusNativeCore.ApplyMetadataPlan slide, plan
    AssertGuardReleased
    Set first = FindShape(slide, firstId)
    Set second = FindShape(slide, secondId)
    Check Not first Is Nothing And Not second Is Nothing, "Metadata regroup lost leaf IDs."
    Check PptNativeDriver.ReadTag(first, "QuickMeta") = "NEW", "The metadata plan did not update its tag."
    Check Not PptNativeDriver.HasTag(second, "QuickRemove"), "The metadata plan did not delete its tag."
    Check PptNativeDriver.ReadTag(first, LOCK_KEY) = "0.200000", "Metadata changed fixed radius."
    Check PptNativeDriver.ReadTag(second, STRICT_KEY) = "1", "Metadata removed strict protection."
    CheckNear PptNativeDriver.ReadFraction(first), CDbl(firstState("Fraction")), "Metadata preserved first radius"
    CheckNear PptNativeDriver.ReadFraction(second), CDbl(secondState("Fraction")), "Metadata preserved protected radius"
    CheckBox first, firstState("Box"), "Metadata first leaf"
    CheckBox second, secondState("Box"), "Metadata second leaf"
    Set roots = PptNativeDriver.SelectionRoots()
    Check roots.Count = 1, "Metadata did not restore the selected group."
    Set grouped = roots(1)
    Check PptNativeDriver.ShapeName(grouped) = CStr(groupState("Name")), "Metadata regroup lost group name."
    CheckTags grouped, groupState("Tags"), "Metadata group"
End Sub

Private Sub CaseShapePlan(ByVal slide As Object)
    Dim first As Object, second As Object, selected As New Collection, firstState As Collection, secondState As Collection
    Dim plan As Collection, firstId As Long, secondId As Long, expectedBox As Variant
    RequireScratch slide
    Set first = FixtureRound(slide, "QuickScene_A", 2, 2, 6, 4, 0.2)
    Set second = FixtureRound(slide, "QuickScene_B", 10, 2, 4, 2, 0.1)
    firstId = PptNativeDriver.ShapeId(first)
    secondId = PptNativeDriver.ShapeId(second)
    SeedTag first, LOCK_KEY, "0.200000"
    SeedTag first, "QuickScene", "OLD"
    selected.Add first
    selected.Add second
    SelectCaseShapes slide, selected
    Set firstState = CaptureShapeState(first)
    Set secondState = CaptureShapeState(second)
    expectedBox = Array(3# * FIXTURE_PT, 3# * FIXTURE_PT, 5# * FIXTURE_PT, 3# * FIXTURE_PT)
    Set plan = New Collection
    plan.Add Array(firstId, "tag", "QuickScene", "SHOULD NOT WRITE", False)
    plan.Add Array(firstId, "box", expectedBox)
    plan.Add Array(firstId, "radius", 0.4)
    plan.Add Array(secondId, "box", Array(10# * FIXTURE_PT, 2# * FIXTURE_PT, 0#, 2# * FIXTURE_PT))
    ExpectShapePlanRejected slide, plan, "no space"
    CheckShapeState first, firstState, "Impossible-box first target"
    CheckShapeState second, secondState, "Impossible-box last target"
    SeedTag second, STRICT_KEY, "1"
    Set secondState = CaptureShapeState(second)
    Set plan = New Collection
    plan.Add Array(firstId, "tag", "QuickScene", "SHOULD NOT WRITE", False)
    plan.Add Array(firstId, "box", expectedBox)
    plan.Add Array(firstId, "radius", 0.4)
    plan.Add Array(secondId, "radius", 0.2)
    ExpectShapePlanRejected slide, plan, "protection"
    CheckShapeState first, firstState, "Shape-plan late-strict first target"
    CheckShapeState second, secondState, "Shape-plan late-strict last target"
    RequireScratch slide
    RadiusNativeCore.SetProtection False
    AssertGuardReleased
    Set plan = New Collection
    plan.Add Array(firstId, "tag", "QuickScene", "NEW", False)
    plan.Add Array(firstId, "box", expectedBox)
    plan.Add Array(firstId, "radius", 0.4)
    plan.Add Array(secondId, "radius", 0.2)
    RequireScratch slide
    RadiusNativeCore.ApplyShapePlan slide, plan
    AssertGuardReleased
    CheckBox first, expectedBox, "Public shape-plan geometry"
    CheckNear PptNativeDriver.ReadFraction(first), 0.4 / 3#, "Public shape-plan radius uses new short side"
    CheckNear PptNativeDriver.ReadFraction(second), 0.1, "Public shape-plan second radius"
    Check PptNativeDriver.ReadTag(first, "QuickScene") = "NEW", "Public shape-plan metadata was not written."
    Check PptNativeDriver.ReadTag(first, LOCK_KEY) = "0.400000", "Public shape-plan fixed R was not updated."
End Sub

Private Sub ExpectGroupedRadiusRejected(ByVal slide As Object, ByVal reason As String, ByVal targets As Collection, ByVal states As Collection)
    Dim number As Long, description As String, roots As Collection, i As Long
    RequireScratch slide
    Check Not PptNativeDriver.HasChildShapeSelection(), "The grouped refusal fixture is not a complete group selection."
    On Error Resume Next
    Err.Clear
    RadiusNativeCore.ApplySelection 10, "%"
    number = Err.Number
    description = Err.Description
    On Error GoTo 0
    Check number = 5, "The grouped radius batch was not refused: " & ErrorLine(number, description)
    Check description = reason, "Unexpected grouped radius rejection: " & ErrorLine(number, description)
    AssertGuardReleased
    For i = 1 To targets.Count
        ' A preflight refusal must preserve the current group ID as well as
        ' every leaf, without starting an ungroup/regroup transaction.
        CheckShapeState targets(i), states(i), "Refused grouped radius target " & CStr(i)
    Next i
    Check Not PptNativeDriver.HasChildShapeSelection(), "A refused grouped action changed the selection mode."
    Set roots = PptNativeDriver.SelectionRoots()
    Check roots.Count = 1, "A refused grouped action changed the selected root count."
    Check PptNativeDriver.ShapeId(roots(1)) = CLng(states(5)("Id")), "A refused grouped action changed the selected group ID."
    Debug.Print "[NativeQuickCheck] Expected grouped radius refusal: " & ErrorLine(number, description)
End Sub

Private Sub CaseGroupedLinkedRadius(ByVal slide As Object)
    Const SETTINGS_KEY As String = "radiusRelationLayout_v1"
    Const BASELINE_KEY As String = "radiusRelationBaseline_v1"
    Const CONFIG As String = "1|1|1|0.500000|0.200000|same|1"
    Dim parent As Object, child As Object, other As Object, arrow As Object, grouped As Object, shape As Object
    Dim members As New Collection, targets As New Collection, states As Collection, freshTargets As Collection, expectedTags As Collection, roots As Collection
    Dim parentId As Long, childId As Long, otherId As Long, arrowId As Long, count As Long, protectedCount As Long, editable As Boolean, i As Long
    Dim pair As Variant, box As Variant, baseline As String, expectedFixed As String
    RequireScratch slide
    Set parent = FixtureRound(slide, "QuickGrouped_Parent", 2, 2, 8, 6, 1)
    Set child = FixtureRound(slide, "QuickGrouped_Linked", 12, 2, 3, 2, 0.2)
    Set other = FixtureRound(slide, "QuickGrouped_Other", 16, 2, 3, 2, 0.4)
    Set arrow = FixtureShape(slide, "QuickGrouped_Arrow", 33, 2, 9, 3, 1)
    parentId = PptNativeDriver.ShapeId(parent)
    childId = PptNativeDriver.ShapeId(child)
    otherId = PptNativeDriver.ShapeId(other)
    arrowId = PptNativeDriver.ShapeId(arrow)
    SeedTag parent, "radiusRelation_v1", "G01"
    SeedTag parent, "radiusRelationRole_v1", "P"
    SeedTag parent, SETTINGS_KEY, CONFIG
    SeedTag parent, LOCK_KEY, "1.000000"
    SeedTag parent, "QuickCustom", "Keep Parent"
    SeedTag child, "radiusRelation_v1", "G01"
    SeedTag child, "radiusRelationRole_v1", "C1"
    SeedTag child, LOCK_KEY, "0.200000"
    SeedTag child, "QuickCustom", "Keep Linked"
    SeedTag other, LOCK_KEY, "0.400000"
    SeedTag other, "QuickCustom", "Keep Ordinary"
    SeedTag arrow, "QuickCustom", "Keep Arrow"
    members.Add parent
    members.Add other
    members.Add arrow
    members.Add child
    Set grouped = FixtureGroup(slide, "QuickGrouped_Outer", members)
    SeedTag grouped, "QuickCustom", "Keep Outer"
    Set parent = FindShape(slide, parentId)
    Set child = FindShape(slide, childId)
    Set other = FindShape(slide, otherId)
    Set arrow = FindShape(slide, arrowId)
    targets.Add parent
    targets.Add child
    targets.Add other
    targets.Add arrow
    targets.Add grouped
    PptNativeDriver.SelectObject grouped
    Check Not PptNativeDriver.HasChildShapeSelection(), "The fixture did not select its complete top-level group."
    ' Only the final linked child becomes strict, after selection. The
    ' linked plan must reject all targets before writing or ungrouping.
    SeedTag child, STRICT_KEY, "1"
    Set states = New Collection
    For Each shape In targets
        states.Add CaptureShapeState(shape)
    Next shape
    RadiusNativeCore.SelectionSummary count, protectedCount, editable
    Check count = 3 And protectedCount = 1 And editable, "Complete-group summary did not expose the late strict child."
    ExpectGroupedRadiusRejected slide, "Turn off protection explicitly before layout or linked radius: QuickGrouped_Linked", targets, states

    RequireScratch slide
    RadiusNativeCore.SetProtection False
    AssertGuardReleased
    Set roots = PptNativeDriver.SelectionRoots()
    Check roots.Count = 1, "Complete-group unprotect did not restore one root selection."
    Check Not PptNativeDriver.HasChildShapeSelection(), "Complete-group unprotect changed the selection mode."
    Set grouped = roots(1)
    Check PptNativeDriver.IsTopLevel(slide, grouped), "Complete-group unprotect did not restore a top-level root."
    Check PptNativeDriver.ShapeName(grouped) = CStr(states(5)("Name")), "Complete-group unprotect lost the group name."
    CheckBox grouped, states(5)("Box"), "Complete-group unprotect geometry"
    CheckTags grouped, states(5)("Tags"), "Complete-group unprotect group"
    ' Every successful group operation invalidates old group proxies. Rebind
    ' all leaves by ID and use the returned selected root for the next case.
    Set freshTargets = New Collection
    For i = 1 To 4
        Set shape = FindShape(slide, CLng(states(i)("Id")))
        Check Not shape Is Nothing, "Complete-group unprotect lost a leaf ID."
        Check PptNativeDriver.ShapeId(shape) = CLng(states(i)("Id")), "Complete-group unprotect changed a leaf ID."
        Check PptNativeDriver.ShapeName(shape) = CStr(states(i)("Name")), "Complete-group unprotect changed a leaf name."
        CheckBox shape, states(i)("Box"), "Complete-group unprotect leaf geometry"
        CheckNear PptNativeDriver.ReadFraction(shape), CDbl(states(i)("Fraction")), "Complete-group unprotect retained fraction"
        Set expectedTags = New Collection
        For Each pair In states(i)("Tags")
            If i = 4 Or StrComp(CStr(pair(0)), STRICT_KEY, vbTextCompare) <> 0 Then expectedTags.Add pair
        Next pair
        CheckTags shape, expectedTags, "Complete-group unprotect retained fixed/custom tags"
        freshTargets.Add shape
    Next i
    freshTargets.Add grouped
    Set targets = freshTargets
    Set parent = targets(1)
    Set child = targets(2)
    Set other = targets(3)
    Set arrow = targets(4)

    SeedTag parent, SETTINGS_KEY, "invalid fixture layout"
    Set states = New Collection
    For Each shape In targets
        states.Add CaptureShapeState(shape)
    Next shape
    ExpectGroupedRadiusRejected slide, "Invalid native layout settings.", targets, states
    SeedTag parent, SETTINGS_KEY, CONFIG
    Set states = New Collection
    For Each shape In targets
        states.Add CaptureShapeState(shape)
    Next shape
    ' Baseline serialization uses fixture geometry and independent formatting;
    ' the radius and fraction answers below are fixed numeric expectations.
    box = states(1)("Box")
    For i = 0 To 3
        baseline = baseline & Replace(Format$(CDbl(box(i)), "0.000000"), ",", ".") & "|"
    Next i
    baseline = baseline & "0.600000|" & CStr(childId) & ":1,"
    RequireScratch slide
    RadiusNativeCore.ApplySelection 10, "%"
    AssertGuardReleased
    For i = 1 To 4
        Set shape = FindShape(slide, CLng(states(i)("Id")))
        Check Not shape Is Nothing, "Complete-group linked radius lost a leaf ID."
        If i = 4 Then
            CheckShapeState shape, states(i), "Complete-group linked radius arrow"
        Else
            Check PptNativeDriver.ShapeId(shape) = CLng(states(i)("Id")), "Complete-group linked radius changed a leaf ID."
            Check PptNativeDriver.ShapeName(shape) = CStr(states(i)("Name")), "Complete-group linked radius changed a leaf name."
            CheckBox shape, states(i)("Box"), "Complete-group linked radius leaf geometry"
            Select Case i
                Case 1: CheckNear PptNativeDriver.ReadFraction(shape), 0.1, "Grouped parent 10% gives 0.6 cm"
                Case 2: CheckNear PptNativeDriver.ReadFraction(shape), 0.3, "Grouped same child inherits 0.6 cm instead of its own 10%"
                Case 3: CheckNear PptNativeDriver.ReadFraction(shape), 0.1, "Grouped ordinary member 10% gives 0.2 cm"
            End Select
            If i = 3 Then expectedFixed = "0.200000" Else expectedFixed = "0.600000"
            Set expectedTags = New Collection
            For Each pair In states(i)("Tags")
                If StrComp(CStr(pair(0)), LOCK_KEY, vbTextCompare) = 0 Then
                    expectedTags.Add Array(CStr(pair(0)), expectedFixed)
                Else
                    expectedTags.Add pair
                End If
            Next pair
            If i = 1 Then expectedTags.Add Array(BASELINE_KEY, baseline)
            CheckTags shape, expectedTags, "Complete-group linked radius fixed/settings/custom tags"
        End If
    Next i
    Set roots = PptNativeDriver.SelectionRoots()
    Check roots.Count = 1, "Complete-group linked radius did not restore one root selection."
    Check Not PptNativeDriver.HasChildShapeSelection(), "Complete-group linked radius changed the selection mode."
    Set grouped = roots(1)
    Check PptNativeDriver.IsTopLevel(slide, grouped), "Complete-group linked radius did not restore a top-level root."
    Check PptNativeDriver.ShapeName(grouped) = CStr(states(5)("Name")), "Complete-group linked radius lost the group name."
    CheckBox grouped, states(5)("Box"), "Complete-group linked radius group geometry"
    CheckTags grouped, states(5)("Tags"), "Complete-group linked radius group"
    RadiusNativeCore.SelectionSummary count, protectedCount, editable
    Check count = 3 And protectedCount = 0 And editable, "The complete linked group was not restored as editable."
End Sub

Private Sub ExpectInputRejected(ByVal operation As String, ByVal text As String, ByVal value As Double, ByVal unit As String, ByVal shortCm As Double, ByVal direction As Long, ByVal reason As String, ByVal detail As String)
    Dim ignored As Double, number As Long, description As String
    On Error Resume Next
    Err.Clear
    Select Case operation
        Case "parse": ignored = RadiusNativeCore.ParseNumber(text)
        Case "step": ignored = RadiusNativeCore.StepValue(value, unit, direction)
        Case "target": ignored = RadiusNativeCore.ComputeTarget(value, unit, shortCm)
    End Select
    number = Err.Number
    description = Err.Description
    On Error GoTo 0
    Check number = 5, detail & " was not rejected by validation: " & ErrorLine(number, description)
    Check description = reason, detail & " returned an unexpected validation error: " & ErrorLine(number, description)
    Debug.Print "[NativeQuickCheck] Expected input rejection (" & detail & "): " & ErrorLine(number, description)
End Sub

Private Sub CaseInvalidInputs()
    ExpectInputRejected "parse", "", 0, "cm", 4, 1, "Enter a nonnegative number.", "Empty input"
    ExpectInputRejected "parse", "-1", 0, "cm", 4, 1, "Enter a nonnegative number.", "Negative input"
    ExpectInputRejected "parse", "1.2.3", 0, "cm", 4, 1, "Enter a nonnegative number.", "Multiple decimal separators"
    ExpectInputRejected "parse", "1e2", 0, "cm", 4, 1, "Enter a nonnegative number.", "Exponent input"
    ExpectInputRejected "parse", "1000001", 0, "cm", 4, 1, "Radius is out of range.", "Input above maximum"
    CheckNear RadiusNativeCore.ParseNumber("1000000"), 1000000, "Maximum valid input"
    ExpectInputRejected "step", "", 0.3, "cm", 4, 0, "Unsupported step direction.", "Unsupported direction"
    ExpectInputRejected "step", "", 0.3, "pt", 4, 1, "Unsupported unit.", "Unsupported step unit"
    ExpectInputRejected "step", "", -0.1, "cm", 4, 1, "Radius is out of range.", "Negative step value"
    ExpectInputRejected "step", "", 1000001, "cm", 4, 1, "Radius is out of range.", "Step value above maximum"
    ExpectInputRejected "target", "", 0.3, "pt", 4, 1, "Unsupported unit.", "Unsupported target unit"
    ExpectInputRejected "target", "", 0.3, "cm", 0, 1, "Shape has zero width or height.", "Zero short side"
    ExpectInputRejected "target", "", -0.1, "cm", 4, 1, "Radius must not be negative.", "Negative target value"
    AssertGuardReleased
End Sub

Private Sub RunInputs()
    Dim before As Long, number As Long, description As String
    before = assertionCount
    On Error GoTo Failed
    CaseInvalidInputs
    RecordPass "Input validation", before
    Exit Sub
Failed:
    number = Err.Number
    description = Err.Description
    RecordFailure "Input validation", number, description
End Sub

Private Function CaseName(ByVal index As Long) As String
    Select Case index
        Case 1: CaseName = "Radius units, zero, clamps and mixed arrow"
        Case 2: CaseName = "Protection, fixed radius and explicit unprotect"
        Case 3: CaseName = "Late strict tag rejects a complete radius batch"
        Case 4: CaseName = "Scaled two-level group and fresh direct hierarchy"
        Case 5: CaseName = "Public metadata and reserved protection tags"
        Case 6: CaseName = "Public shape plans, impossible box and late strict"
        Case 7: CaseName = "Complete top-level group: linked %, strict and invalid layout"
        Case 8: CaseName = "Parent binding, conflict and detach"
        Case 9: CaseName = "Grid layout, serialization and radius links"
        Case 10: CaseName = "Strict child blocks a complete layout batch"
        Case 11: CaseName = "Protected parent layout rules"
        Case 12: CaseName = "Invalid grid and transformed layout rejection"
        Case 13: CaseName = "Independent radius with unrelated broken relation"
        Case Else: Err.Raise 5, , "Unknown quick-check case."
    End Select
End Function

Private Sub RunMemoryCase(ByVal index As Long)
    Dim name As String, before As Long, number As Long, description As String, cleanupText As String
    name = CaseName(index)
    before = assertionCount
    On Error GoTo Failed
    RequireOwned
    Set caseSlide = PptNativeDriver.AddBlankSlide(ownedPresentation)
    PptNativeDriver.ActivateSlide ownedWindow, caseSlide
    RequireScratch caseSlide
    SuspendSessions
    Select Case index
        Case 1: CaseUnits caseSlide
        Case 2: CaseProtection caseSlide
        Case 3: CaseLateStrict caseSlide
        Case 4: CaseNestedGroup caseSlide
        Case 5: CaseMetadata caseSlide
        Case 6: CaseShapePlan caseSlide
        Case 7: CaseGroupedLinkedRadius caseSlide
        Case 8: RadiusNativeQuickRelations.CaseBinding caseSlide
        Case 9: RadiusNativeQuickRelations.CaseLayout caseSlide
        Case 10: RadiusNativeQuickRelations.CaseStrictLayout caseSlide
        Case 11: RadiusNativeQuickRelations.CaseProtectedParent caseSlide
        Case 12: RadiusNativeQuickRelations.CaseInvalidLayout caseSlide
        Case 13: RadiusNativeQuickRelations.CaseIndependentRadius caseSlide
    End Select
    RequireScratch caseSlide
    AssertGuardReleased
    GoTo Finished
Failed:
    number = Err.Number
    description = Err.Description
Finished:
    ' A failed case cannot leave pending parent/layout parameters for another.
    On Error Resume Next
    Err.Clear
    RadiusNativeRelations.SuspendTestSession
    If Err.Number <> 0 Then cleanupText = " Relationship case cleanup: " & ErrorLine(Err.Number, Err.Description)
    Err.Clear
    RadiusNativeLayout.SuspendTestSession
    If Err.Number <> 0 Then cleanupText = cleanupText & " Layout case cleanup: " & ErrorLine(Err.Number, Err.Description)
    On Error GoTo 0
    If cleanupText <> "" Then
        If number = 0 Then number = 5
        description = description & cleanupText
    End If
    If number = 0 Then RecordPass name, before Else RecordFailure name, number, description
    Set caseSlide = Nothing
End Sub

Private Function RestoreOneSession(ByVal kind As String, ByVal state As Variant) As Boolean
    Dim number As Long, description As String, current As Variant
    On Error GoTo Failed
    If kind = "Relationship" Then
        RadiusNativeRelations.RestoreTestSession state
        current = RadiusNativeRelations.CaptureTestSession()
    Else
        RadiusNativeLayout.RestoreTestSession state
        current = RadiusNativeLayout.CaptureTestSession()
    End If
    Check SessionSame(state, current), "Original " & kind & " session did not match its snapshot."
    RestoreOneSession = True
    Exit Function
Failed:
    number = Err.Number
    description = Err.Description
    RecordFailure kind & " session restore", number, description, False
End Function

Private Function RestoreSessions(ByVal relationState As Variant, ByVal layoutState As Variant, ByVal hasRelationState As Boolean, ByVal hasLayoutState As Boolean) As Boolean
    Dim restored As Boolean
    RestoreSessions = True
    If hasRelationState Then
        restored = RestoreOneSession("Relationship", relationState)
        If Not restored Then RestoreSessions = False
    End If
    If hasLayoutState Then
        restored = RestoreOneSession("Layout", layoutState)
        If Not restored Then RestoreSessions = False
    End If
End Function

Public Function RunnerRun() As String
    Dim context As Collection, relationState As Variant, layoutState As Variant
    Dim hasRelationState As Boolean, hasLayoutState As Boolean, windowRestored As Boolean, sessionsRestored As Boolean
    Dim closed As Boolean, created As Boolean, algorithmCount As Long, index As Long, started As Double, elapsed As Double
    Dim number As Long, description As String, cleanup As String, result As String
    If running Then Err.Raise 5, , "A quick check is already running."
    If Not ownedPresentation Is Nothing Then
        lastFailures = 1
        RunnerRun = "Quick check refused: disposable presentation '" & ownedName & "' from the previous run could not be closed. Close it and restart PowerPoint before retrying."
        Debug.Print "[NativeQuickCheck] " & RunnerRun
        Exit Function
    End If
    running = True
    assertionCount = 0
    passedCases = 0
    failedCases = 0
    lastFailures = 0
    failures = ""
    ownedName = "Disposable quick-check presentation"
    started = Timer
    On Error GoTo Failed
    If RadiusNativeRelations.IsPreviewOpen Then Err.Raise 5, , "Close the relation preview before running the quick check."
    If RadiusNativeCore.IsTransactionActive Then Err.Raise 5, , "Wait for the native transaction to finish before running the quick check."
    Set context = PptNativeDriver.CaptureWindowContext()
    relationState = RadiusNativeRelations.CaptureTestSession()
    hasRelationState = True
    layoutState = RadiusNativeLayout.CaptureTestSession()
    hasLayoutState = True
    SuspendSessions
    algorithmCount = RunAlgorithms()
    RunInputs
    Set ownedPresentation = PptNativeDriver.CreateTemporaryPresentation()
    created = True
    ownedName = PptNativeDriver.PresentationName(ownedPresentation)
    PptNativeDriver.ActivatePresentation ownedPresentation
    Set ownedWindow = PptNativeDriver.CurrentWindow()
    RequireOwned
    For index = 1 To MEMORY_CASE_COUNT
        RunMemoryCase index
    Next index
    GoTo Cleanup
Failed:
    number = Err.Number
    description = Err.Description
    RecordFailure "Runner setup", number, description, False
Cleanup:
    ' Close exactly the presentation returned to this runner, never the
    ' active presentation. All independent recovery steps are attempted.
    closed = Not created
    If Not ownedPresentation Is Nothing Then
        On Error Resume Next
        Err.Clear
        PptNativeDriver.CloseTemporaryPresentation ownedPresentation
        number = Err.Number
        description = Err.Description
        On Error GoTo 0
        If number = 0 Then closed = True Else RecordFailure "Fixture close (" & ownedName & ")", number, description, False
    End If
    windowRestored = (context Is Nothing)
    If Not context Is Nothing Then
        On Error Resume Next
        Err.Clear
        PptNativeDriver.RestoreWindowContext context
        number = Err.Number
        description = Err.Description
        On Error GoTo 0
        If number = 0 Then windowRestored = True Else RecordFailure "Original window restore", number, description, False
    End If
    ' Restore both sessions even if fixture close or window restore failed.
    sessionsRestored = RestoreSessions(relationState, layoutState, hasRelationState, hasLayoutState)
    If created Then
        On Error Resume Next
        Err.Clear
        AssertGuardReleased
        number = Err.Number
        description = Err.Description
        On Error GoTo 0
        If number <> 0 Then RecordFailure "Final transaction guard", number, description, False
    End If
    elapsed = Timer - started
    If elapsed < 0 Then elapsed = elapsed + 86400#
    If closed And windowRestored And sessionsRestored Then cleanup = "OK; original window/selection and sessions restored." Else cleanup = "FAILED; see recovery errors below."
    If Not created Then cleanup = "No fixture created. " & cleanup
    result = "Algorithm checks: " & CStr(algorithmCount) & vbCrLf
    result = result & "Memory/input cases: " & CStr(passedCases) & " passed, " & CStr(failedCases) & " failed (Core " & CStr(CORE_CASE_COUNT) & "; relations/layout " & CStr(MEMORY_CASE_COUNT - CORE_CASE_COUNT) & "; input 1)." & vbCrLf
    result = result & "Assertions: " & CStr(assertionCount) & "; elapsed: " & Format$(elapsed, "0.00") & " s." & vbCrLf
    result = result & "Cleanup: " & cleanup & vbCrLf
    result = result & "Scope: in-memory production actions; actual child-selection refusals, saved OOXML and injected host write failures require separate checks."
    If failures <> "" Then result = result & vbCrLf & failures
    Debug.Print "[NativeQuickCheck] " & Replace(result, vbCrLf, " | ")
    Set caseSlide = Nothing
    If closed Then
        Set ownedWindow = Nothing
        Set ownedPresentation = Nothing
        ownedName = ""
    End If
    running = False
    RunnerRun = result
End Function
