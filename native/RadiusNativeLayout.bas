Attribute VB_Name = "RadiusNativeLayout"
Option Explicit

Private Const SETTINGS_KEY As String = "radiusRelationLayout_v1"
Private Const BASELINE_KEY As String = "radiusRelationBaseline_v1"
Private pendingPresentation As Object
Private pendingSlide As Long
Private pendingKey As String
Private pendingConfig As Variant
Private hasPending As Boolean
Private lastError As String

Public Function CaptureTestSession() As Variant
    Dim state(0 To 5) As Variant
    Set state(0) = pendingPresentation
    state(1) = pendingSlide
    state(2) = pendingKey
    state(3) = pendingConfig
    state(4) = hasPending
    state(5) = lastError
    CaptureTestSession = state
End Function

Public Sub RestoreTestSession(ByVal state As Variant)
    If Not IsArray(state) Then Err.Raise 5, , "Invalid layout test session."
    If LBound(state) <> 0 Or UBound(state) <> 5 Then Err.Raise 5, , "Invalid layout test session."
    Set pendingPresentation = state(0)
    pendingSlide = CLng(state(1))
    pendingKey = CStr(state(2))
    pendingConfig = state(3)
    hasPending = CBool(state(4))
    lastError = CStr(state(5))
End Sub

Public Sub SuspendTestSession()
    ResetPending
    pendingConfig = Empty
    lastError = ""
End Sub

Public Function GridBoxes(ByVal parentBox As Variant, ByVal rows As Long, ByVal columns As Long, ByVal paddingCm As Double, ByVal gapCm As Double, ByVal count As Long) As Collection
    Dim result As New Collection, width As Double, height As Double, padding As Double, gap As Double, i As Long
    If rows < 1 Or rows > 25 Or columns < 1 Or columns > 25 Then Err.Raise 5, , "Rows and columns must be between 1 and 25."
    If count < 1 Or count > rows * columns Then Err.Raise 5, , "Grid capacity is smaller than the number of children."
    If paddingCm < 0 Or gapCm < 0 Then Err.Raise 5, , "Padding and gap must not be negative."
    padding = paddingCm * RadiusNativeCore.PT_PER_CM
    gap = gapCm * RadiusNativeCore.PT_PER_CM
    width = (CDbl(parentBox(2)) - 2# * padding - (columns - 1) * gap) / columns
    height = (CDbl(parentBox(3)) - 2# * padding - (rows - 1) * gap) / rows
    If width <= 0 Or height <= 0 Then Err.Raise 5, , "Padding/gap leave no space for children. No layout was written."
    For i = 0 To count - 1
        result.Add Array(CDbl(parentBox(0)) + padding + (i Mod columns) * (width + gap), CDbl(parentBox(1)) + padding + (i \ columns) * (height + gap), width, height)
    Next i
    Set GridBoxes = result
End Function

Public Function LinkedRadius(ByVal parentCm As Double, ByVal paddingCm As Double, ByVal mode As String) As Double
    If parentCm < 0 Or paddingCm < 0 Then Err.Raise 5, , "Linked radius must not be negative."
    Select Case mode
        Case "same": LinkedRadius = parentCm
        Case "subtract"
            LinkedRadius = parentCm - paddingCm
            If LinkedRadius < 0 Then LinkedRadius = 0
        Case "off": LinkedRadius = 0
        Case Else: Err.Raise 5, , "Invalid radius-link mode."
    End Select
End Function

Public Sub ResetPending()
    Set pendingPresentation = Nothing
    hasPending = False
    pendingKey = ""
    pendingSlide = 0
End Sub

Public Sub ForgetPresentation(ByVal presentation As Object)
    If Not pendingPresentation Is Nothing Then
        If presentation Is pendingPresentation Then ResetPending
    End If
End Sub

Private Function IntegerValue(ByVal text As String) As Long
    Dim value As Double
    value = RadiusNativeCore.ParseNumber(text)
    If value <> Fix(value) Or value < 1 Or value > 25 Then Err.Raise 5, , "Rows/columns must be whole numbers from 1 to 25."
    IntegerValue = CLng(value)
End Function

Private Function SignedNumber(ByVal text As String) As Double
    If Left$(text, 1) = "-" Then SignedNumber = -RadiusNativeCore.ParseNumber(Mid$(text, 2)) Else SignedNumber = RadiusNativeCore.ParseNumber(text)
End Function

Private Function DefaultConfig(ByVal count As Long) As Variant
    Dim rows As Long, columns As Long
    If count < 1 Or count > 625 Then Err.Raise 5, , "A layout requires between 1 and 625 child objects."
    columns = count
    If columns > 25 Then columns = 25
    rows = (count + columns - 1) \ columns
    DefaultConfig = Array(rows, columns, 0.3, 0.2, "same", True)
End Function

Private Function ReadConfig(ByVal parent As Object, ByVal count As Long) As Variant
    ReadConfig = ReadConfigText(PptNativeDriver.ReadTag(parent, SETTINGS_KEY), count)
End Function

Private Function ReadConfigText(ByVal value As String, ByVal count As Long) As Variant
    Dim fields As Variant, config As Variant
    If value = "" Then ReadConfigText = DefaultConfig(count): Exit Function
    fields = Split(value, "|")
    If UBound(fields) <> 6 Then Err.Raise 5, , "Invalid native layout settings."
    If CStr(fields(0)) <> "1" Then Err.Raise 5, , "Unsupported layout settings version."
    config = Array(IntegerValue(CStr(fields(1))), IntegerValue(CStr(fields(2))), RadiusNativeCore.ParseNumber(CStr(fields(3))), RadiusNativeCore.ParseNumber(CStr(fields(4))), LCase$(CStr(fields(5))), CStr(fields(6)) = "1")
    If CStr(fields(6)) <> "0" And CStr(fields(6)) <> "1" Then Err.Raise 5, , "Invalid automatic-layout setting."
    If CStr(config(4)) <> "same" And CStr(config(4)) <> "subtract" And CStr(config(4)) <> "off" Then Err.Raise 5, , "Invalid radius-link mode."
    ReadConfigText = config
End Function

Private Function ConfigText(ByVal config As Variant) As String
    Dim automatic As String
    automatic = "0"
    If CBool(config(5)) Then automatic = "1"
    ConfigText = "1|" & CStr(config(0)) & "|" & CStr(config(1)) & "|" & RadiusNativeCore.NumberText(CDbl(config(2))) & "|" & RadiusNativeCore.NumberText(CDbl(config(3))) & "|" & CStr(config(4)) & "|" & automatic
End Function

Private Sub FitCapacity(ByRef config As Variant, ByVal count As Long)
    If count < 1 Or count > 625 Then Err.Raise 5, , "A layout requires between 1 and 625 child objects."
    If CLng(config(0)) > count Then config(0) = count
    If CLng(config(1)) > count Then config(1) = count
    If CLng(config(0)) * CLng(config(1)) < count Then config(1) = (count + CLng(config(0)) - 1) \ CLng(config(0))
    If CLng(config(1)) > 25 Then Err.Raise 5, , "This grid needs more than 25 columns. Increase rows first."
End Sub

Private Function EditableGroup() As Collection
    Dim group As Collection
    Set group = RadiusNativeRelations.SelectedGroup()
    If group Is Nothing Then Err.Raise 5, , "Select a parent or child belonging to one relationship."
    If CStr(group("Kind")) <> "native" Then Err.Raise 5, , "Explicitly rebuild legacy relationships before enabling native layout."
    If CStr(group("Problem")) <> "" Then Err.Raise 5, , CStr(group("Problem"))
    Set EditableGroup = group
End Function

Public Function CanConfigure() As Boolean
    Dim group As Collection
    On Error GoTo Failed
    If RadiusNativeRelations.IsPreviewContext Then Exit Function
    Set group = EditableGroup()
    CanConfigure = True
    Exit Function
Failed:
    Debug.Print "[NativeLayoutSelection] " & Err.Description
End Function

Private Function ConfigForGroup(ByVal group As Collection) As Variant
    If HasPendingForGroup(group) Then ConfigForGroup = pendingConfig Else ConfigForGroup = ReadConfig(group("Parent"), group("Children").Count)
End Function

Private Function HasPendingForGroup(ByVal group As Collection) As Boolean
    Dim current As Object
    If hasPending Then
        Set current = PptNativeDriver.CurrentPresentation()
        If Not current Is pendingPresentation Then ResetPending
        If pendingSlide <> PptNativeDriver.SlideId(PptNativeDriver.CurrentSlide()) Or pendingKey <> CStr(group("Key")) Then ResetPending
    End If
    HasPendingForGroup = hasPending
End Function

Public Function SelectedConfig() As Variant
    SelectedConfig = ConfigForGroup(EditableGroup())
End Function

Public Function AutomaticEnabled() As Boolean
    Dim group As Collection, config As Variant
    On Error GoTo Failed
    If Not CanConfigure Then Exit Function
    Set group = EditableGroup()
    If PptNativeDriver.ReadTag(group("Parent"), SETTINGS_KEY) = "" Then Exit Function
    config = ReadConfig(group("Parent"), group("Children").Count)
    AutomaticEnabled = CBool(config(5))
    Exit Function
Failed:
    RememberError Err.Description
End Function

Public Function ChildrenWritable() As Boolean
    Dim group As Collection
    Set group = EditableGroup()
    ChildrenWritable = ChildrenWritableFor(group)
End Function

Private Function ChildrenWritableFor(ByVal group As Collection) As Boolean
    Dim child As Collection
    For Each child In group("Children")
        If PptNativeDriver.ReadTag(child("Shape"), "radiusLockStrict_v1") = "1" Then Exit Function
    Next child
    ChildrenWritableFor = True
End Function

' Display-only queries reuse one already validated relationship model.
Public Function ReadUiSnapshot(ByVal group As Collection) As Collection
    Dim snapshot As New Collection, config As Variant, saved As Variant, text As String
    Dim available As Boolean, writable As Boolean, automatic As Boolean, pending As Boolean
    Dim key As Variant, value As Double, direction As Long, suffix As String, count As Long
    config = Array(1, 1, 0.3, 0.2, "same", False)
    If Not group Is Nothing Then
        If CStr(group("Kind")) = "native" Then
            If CStr(group("Problem")) <> "" Then Err.Raise 5, , CStr(group("Problem"))
            count = group("Children").Count
            text = PptNativeDriver.ReadTag(group("Parent"), SETTINGS_KEY)
            saved = ReadConfigText(text, count)
            pending = HasPendingForGroup(group)
            If pending Then config = pendingConfig Else config = saved
            automatic = (text <> "" And CBool(saved(5)))
            pending = (pending Or text = "")
            writable = ChildrenWritableFor(group)
            available = True
        End If
    End If
    snapshot.Add available, "CanConfigure"
    snapshot.Add config, "Config"
    snapshot.Add writable, "ChildrenWritable"
    snapshot.Add automatic, "Automatic"
    snapshot.Add pending, "Pending"
    For Each key In Array("rows", "columns", "padding", "gap")
        For direction = -1 To 1 Step 2
            If direction = 1 Then suffix = "Up" Else suffix = "Down"
            If available Then
                value = ParameterValue(config, CStr(key))
                snapshot.Add (ParameterStepValue(value, CStr(key), count, direction) <> value), CStr(key) & suffix
            Else
                snapshot.Add False, CStr(key) & suffix
            End If
        Next direction
    Next key
    Set ReadUiSnapshot = snapshot
End Function

Public Sub SetParameter(ByVal key As String, ByVal text As String)
    Dim group As Collection, config As Variant, count As Long, value As Long
    Set group = EditableGroup()
    config = ConfigForGroup(group)
    count = group("Children").Count
    Select Case key
        Case "rows", "columns"
            value = IntegerValue(text)
            If value > count Then value = count
            If key = "rows" Then
                config(0) = value
                config(1) = (count + value - 1) \ value
            Else
                config(1) = value
                config(0) = (count + value - 1) \ value
            End If
        Case "padding": config(2) = RadiusNativeCore.ParseNumber(text)
        Case "gap": config(3) = RadiusNativeCore.ParseNumber(text)
        Case "mode": config(4) = text
        Case Else: Err.Raise 5, , "Unsupported layout parameter."
    End Select
    FitCapacity config, count
    pendingConfig = config
    Set pendingPresentation = PptNativeDriver.CurrentPresentation()
    pendingSlide = PptNativeDriver.SlideId(PptNativeDriver.CurrentSlide())
    pendingKey = CStr(group("Key"))
    hasPending = True
End Sub

Private Function ParameterValue(ByVal config As Variant, ByVal key As String) As Double
    Select Case key
        Case "rows": ParameterValue = CDbl(config(0))
        Case "columns": ParameterValue = CDbl(config(1))
        Case "padding": ParameterValue = CDbl(config(2))
        Case "gap": ParameterValue = CDbl(config(3))
        Case Else: Err.Raise 5, , "Unsupported layout parameter."
    End Select
End Function

Public Function ParameterStepValue(ByVal value As Double, ByVal key As String, ByVal count As Long, ByVal direction As Long) As Double
    Dim minimum As Long, maximum As Long
    If direction <> 1 And direction <> -1 Then Err.Raise 5, , "Unsupported step direction."
    If count < 1 Or count > 625 Then Err.Raise 5, , "A layout requires between 1 and 625 child objects."
    Select Case key
        Case "rows", "columns"
            If value <> Fix(value) Or value < 1 Or value > 25 Then Err.Raise 5, , "Rows/columns must be whole numbers from 1 to 25."
            minimum = (count + 24) \ 25
            maximum = count
            If maximum > 25 Then maximum = 25
            If (direction = 1 And value >= maximum) Or (direction = -1 And value <= minimum) Then ParameterStepValue = value: Exit Function
            ParameterStepValue = value + direction
            If ParameterStepValue < minimum Then ParameterStepValue = minimum
            If ParameterStepValue > maximum Then ParameterStepValue = maximum
        Case "padding", "gap"
            ParameterStepValue = RadiusNativeCore.StepValue(value, "cm", direction)
        Case Else: Err.Raise 5, , "Unsupported layout parameter."
    End Select
End Function

Public Function CanStepParameter(ByVal key As String, ByVal direction As Long) As Boolean
    Dim group As Collection, config As Variant, value As Double
    Set group = EditableGroup()
    config = ConfigForGroup(group)
    value = ParameterValue(config, key)
    CanStepParameter = (ParameterStepValue(value, key, group("Children").Count, direction) <> value)
End Function

Public Sub StepParameter(ByVal key As String, ByVal direction As Long)
    Dim group As Collection, config As Variant, value As Double, nextValue As Double
    Set group = EditableGroup()
    config = ConfigForGroup(group)
    value = ParameterValue(config, key)
    nextValue = ParameterStepValue(value, key, group("Children").Count, direction)
    If nextValue = value Then Exit Sub
    ' Stage through the same validation as typed input; Apply commits the layout.
    SetParameter key, RadiusNativeCore.NumberText(nextValue)
End Sub

Private Function BaselineText(ByVal group As Collection, ByVal radius As Double) As String
    Dim box As Variant, text As String, i As Long, child As Collection
    box = PptNativeDriver.ShapeBox(group("Parent"))
    For i = 0 To 3
        text = text & RadiusNativeCore.NumberText(CDbl(box(i))) & "|"
    Next i
    text = text & RadiusNativeCore.NumberText(radius) & "|"
    For Each child In group("Children")
        text = text & CStr(PptNativeDriver.ShapeId(child("Shape"))) & ":" & CStr(child("Order")) & ","
    Next child
    BaselineText = text
End Function

Private Sub AddSettings(ByVal plan As Collection, ByVal group As Collection, ByVal config As Variant, ByVal radius As Double)
    Dim id As Long
    id = CLng(group("ParentId"))
    plan.Add Array(id, "tag", SETTINGS_KEY, ConfigText(config), False)
    plan.Add Array(id, "tag", BASELINE_KEY, BaselineText(group, radius), False)
End Sub

Private Sub AddGroupPlan(ByVal plan As Collection, ByVal group As Collection, ByVal config As Variant, ByVal boxes As Boolean, ByVal radii As Boolean, ByVal parentRadius As Double)
    Dim targets As Collection, child As Collection, i As Long, id As Long
    If CStr(group("Problem")) <> "" Then Err.Raise 5, , CStr(group("Problem"))
    FitCapacity config, group("Children").Count
    If boxes Then Set targets = GridBoxes(PptNativeDriver.ShapeBox(group("Parent")), CLng(config(0)), CLng(config(1)), CDbl(config(2)), CDbl(config(3)), group("Children").Count)
    For Each child In group("Children")
        i = i + 1
        id = PptNativeDriver.ShapeId(child("Shape"))
        If boxes Then plan.Add Array(id, "box", targets(i))
        If radii And CStr(config(4)) <> "off" Then RadiusNativeCore.PutRadius plan, id, LinkedRadius(parentRadius, CDbl(config(2)), CStr(config(4)))
    Next child
    AddSettings plan, group, config, parentRadius
End Sub

Public Sub ApplySelected(Optional ByVal radiusOnly As Boolean = False)
    Dim group As Collection, config As Variant, plan As New Collection
    Set group = EditableGroup()
    config = ConfigForGroup(group)
    AddGroupPlan plan, group, config, Not radiusOnly, True, RadiusNativeCore.CurrentRadius(group("Parent"))
    RadiusNativeCore.ApplyShapePlan PptNativeDriver.CurrentSlide(), plan
    ResetPending
    lastError = ""
End Sub

Public Sub ChangeMode(ByVal mode As String)
    Dim group As Collection
    Set group = EditableGroup()
    SetParameter "mode", mode
    If PptNativeDriver.ReadTag(group("Parent"), SETTINGS_KEY) <> "" Then ApplySelected True
End Sub

Public Sub SetAutomatic(ByVal enabled As Boolean)
    Dim group As Collection, config As Variant, plan As New Collection
    Set group = EditableGroup()
    config = ConfigForGroup(group)
    config(5) = enabled
    If enabled Then
        AddGroupPlan plan, group, config, True, True, RadiusNativeCore.CurrentRadius(group("Parent"))
    Else
        AddSettings plan, group, config, RadiusNativeCore.CurrentRadius(group("Parent"))
    End If
    RadiusNativeCore.ApplyShapePlan PptNativeDriver.CurrentSlide(), plan
    ResetPending
    lastError = ""
End Sub

Private Function HasConfiguredLayouts(ByVal slide As Object) As Boolean
    Dim shape As Object
    For Each shape In RadiusNativeRelations.SlideLeaves(slide)
        If PptNativeDriver.ReadTag(shape, SETTINGS_KEY) <> "" Then HasConfiguredLayouts = True: Exit Function
    Next shape
End Function

Public Function SelectedLinkedParents(ByVal roots As Collection) As Collection
    Dim parents As New Collection, shape As Object, text As String, config As Variant
    ' Read only selected leaves first. Unrelated broken relationships must not
    ' block an explicit radius edit on an ordinary, unbound rounded rectangle.
    For Each shape In RadiusNativeRelations.RoundedLeaves(roots)
        text = PptNativeDriver.ReadTag(shape, SETTINGS_KEY)
        If text <> "" Then
            If RadiusNativeRelations.IsNativeParent(shape) Then
                config = ReadConfigText(text, 1)
                If CBool(config(5)) And CStr(config(4)) <> "off" Then parents.Add shape
            End If
        End If
    Next shape
    Set SelectedLinkedParents = parents
End Function

Private Function ParentSelected(ByVal parents As Collection, ByVal id As Long) As Boolean
    Dim parent As Object
    For Each parent In parents
        If PptNativeDriver.ShapeId(parent) = id Then ParentSelected = True: Exit Function
    Next parent
End Function

Public Function LinkedRadiusPlan(ByVal roots As Collection, ByVal value As Double, ByVal unit As String) As Collection
    Dim slide As Object, groups As Collection, group As Collection, config As Variant, plan As New Collection, linked As New Collection, leaves As Collection, shape As Object, radius As Double, parents As Collection
    Set parents = SelectedLinkedParents(roots)
    If parents.Count = 0 Then Exit Function
    Set slide = PptNativeDriver.CurrentSlide()
    ' A linked parent requires the complete, current relationship validation.
    Set groups = RadiusNativeRelations.GroupsOnSlide(slide)
    For Each group In groups
        If CStr(group("Kind")) = "native" And CLng(group("ParentId")) > 0 Then
            If PptNativeDriver.ReadTag(group("Parent"), SETTINGS_KEY) <> "" Then
                config = ReadConfig(group("Parent"), group("Children").Count)
                If CBool(config(5)) And CStr(config(4)) <> "off" And ParentSelected(parents, CLng(group("ParentId"))) Then linked.Add group
            End If
        End If
    Next group
    If linked.Count = 0 Then Exit Function
    For Each shape In roots
        If Not PptNativeDriver.IsTopLevel(slide, shape) Then Err.Raise 5, , "Select the complete top-level group for radius editing."
    Next shape
    Set leaves = RadiusNativeCore.SelectionLeaves()
    For Each shape In leaves
        RadiusNativeCore.PutRadius plan, PptNativeDriver.ShapeId(shape), RadiusNativeCore.ComputeTarget(value, unit, PptNativeDriver.ShortSidePoints(shape) / RadiusNativeCore.PT_PER_CM)
    Next shape
    For Each group In linked
        config = ReadConfig(group("Parent"), group("Children").Count)
        radius = RadiusNativeCore.ComputeTarget(value, unit, PptNativeDriver.ShortSidePoints(group("Parent")) / RadiusNativeCore.PT_PER_CM)
        AddGroupPlan plan, group, config, False, True, radius
    Next group
    Set LinkedRadiusPlan = plan
End Function

Public Function LinkedSelectionWritable() As Boolean
    Dim roots As Collection, parents As Collection, groups As Collection, shape As Object
    LinkedSelectionWritable = True
    If Not PptNativeDriver.HasShapeSelection Then Exit Function
    Set roots = PptNativeDriver.SelectionRoots()
    Set parents = SelectedLinkedParents(roots)
    If parents.Count = 0 Then Exit Function
    Set groups = RadiusNativeRelations.GroupsOnSlide(PptNativeDriver.CurrentSlide())
    If Not LinkedParentsWritable(parents, groups) Then LinkedSelectionWritable = False: Exit Function
    For Each shape In RadiusNativeRelations.RoundedLeaves(roots)
        If PptNativeDriver.ReadTag(shape, "radiusLockStrict_v1") = "1" Then LinkedSelectionWritable = False: Exit Function
    Next shape
End Function

Public Function LinkedParentsWritable(ByVal parents As Collection, ByVal groups As Collection) As Boolean
    Dim group As Collection, config As Variant
    LinkedParentsWritable = True
    For Each group In groups
        If CStr(group("Kind")) = "native" And CLng(group("ParentId")) > 0 Then
            If ParentSelected(parents, CLng(group("ParentId"))) Then
                If CStr(group("Problem")) <> "" Then Err.Raise 5, , CStr(group("Problem"))
                config = ReadConfig(group("Parent"), group("Children").Count)
                FitCapacity config, group("Children").Count
                If CBool(config(5)) And CStr(config(4)) <> "off" Then
                    If Not ChildrenWritableFor(group) Then LinkedParentsWritable = False: Exit Function
                End If
            End If
        End If
    Next group
End Function

Public Sub SyncSlide(ByVal slide As Object)
    Dim groups As Collection, group As Collection, config As Variant, baseline As String, old As Variant, current As Variant, boxChanged As Boolean, radiusChanged As Boolean, i As Long, radius As Double, plan As New Collection
    If Not HasConfiguredLayouts(slide) Then Exit Sub
    Set groups = RadiusNativeRelations.GroupsOnSlide(slide)
    For Each group In groups
        If CStr(group("Kind")) = "native" And CLng(group("ParentId")) > 0 Then
            If PptNativeDriver.ReadTag(group("Parent"), SETTINGS_KEY) <> "" Then
                config = ReadConfig(group("Parent"), group("Children").Count)
                If CBool(config(5)) Then
                    radius = RadiusNativeCore.CurrentRadius(group("Parent"))
                    baseline = PptNativeDriver.ReadTag(group("Parent"), BASELINE_KEY)
                    current = Split(BaselineText(group, radius), "|")
                    boxChanged = False
                    radiusChanged = False
                    If baseline = "" Then
                        boxChanged = True
                    Else
                        old = Split(baseline, "|")
                        If UBound(old) <> 5 Then Err.Raise 5, , "Invalid native layout baseline."
                        For i = 0 To 3
                            If Abs(SignedNumber(CStr(old(i))) - SignedNumber(CStr(current(i)))) > 0.002 Then boxChanged = True
                        Next i
                        If CStr(old(5)) <> CStr(current(5)) Then boxChanged = True
                        radiusChanged = (Abs(RadiusNativeCore.ParseNumber(CStr(old(4))) - radius) > 0.0001)
                    End If
                    If boxChanged Or radiusChanged Then AddGroupPlan plan, group, config, boxChanged, True, radius
                End If
            End If
        End If
    Next group
    If plan.Count > 0 Then RadiusNativeCore.ApplyShapePlan slide, plan
End Sub

Public Sub SyncCurrent()
    lastError = ""
    If RadiusNativeRelations.IsPreviewContext Or Not PptNativeDriver.HasPresentation Then Exit Sub
    SyncSlide PptNativeDriver.CurrentSlide()
End Sub

Public Sub SyncPresentation(ByVal presentation As Object)
    Dim slide As Object
    If RadiusNativeRelations.IsPreviewContext Then Exit Sub
    lastError = ""
    For Each slide In PptNativeDriver.Slides(presentation)
        SyncSlide slide
    Next slide
End Sub

Public Sub RememberError(ByVal text As String)
    lastError = text
    Debug.Print "[NativeLayout] " & text
End Sub

Public Function Status(ByVal enabledText As String, ByVal disabledText As String, ByVal pendingText As String, ByVal errorText As String, ByVal emptyText As String) As String
    Status = StatusFromSnapshot(ReadUiSnapshot(RadiusNativeRelations.SelectedGroup()), enabledText, disabledText, pendingText, errorText, emptyText)
End Function

Public Function StatusFromSnapshot(ByVal snapshot As Collection, ByVal enabledText As String, ByVal disabledText As String, ByVal pendingText As String, ByVal errorText As String, ByVal emptyText As String) As String
    Dim config As Variant
    If lastError <> "" Then StatusFromSnapshot = errorText & lastError: Exit Function
    If Not CBool(snapshot("CanConfigure")) Then StatusFromSnapshot = emptyText: Exit Function
    config = snapshot("Config")
    If CBool(snapshot("Pending")) Then
        StatusFromSnapshot = pendingText
    ElseIf CBool(config(5)) Then
        StatusFromSnapshot = enabledText
    Else
        StatusFromSnapshot = disabledText
    End If
    StatusFromSnapshot = StatusFromSnapshot & " " & CStr(config(0)) & " x " & CStr(config(1))
End Function

Public Function SelfTest() As Long
    Dim boxes As Collection, box As Variant
    Dim passed As Long
    Set boxes = GridBoxes(Array(0#, 0#, 12# * RadiusNativeCore.PT_PER_CM, 8# * RadiusNativeCore.PT_PER_CM), 2, 2, 0.5, 0.3, 4)
    box = boxes(1)
    If Abs(CDbl(box(2)) / RadiusNativeCore.PT_PER_CM - 5.35) > 0.000001 Then Err.Raise 5, , "Grid width failed."
    passed = passed + 1
    If Abs(CDbl(box(3)) / RadiusNativeCore.PT_PER_CM - 3.35) > 0.000001 Then Err.Raise 5, , "Grid height failed."
    passed = passed + 1
    box = boxes(4)
    If Abs(CDbl(box(0)) / RadiusNativeCore.PT_PER_CM - 6.15) > 0.000001 Then Err.Raise 5, , "Grid column position failed."
    passed = passed + 1
    If Abs(CDbl(box(1)) / RadiusNativeCore.PT_PER_CM - 4.15) > 0.000001 Then Err.Raise 5, , "Grid row position failed."
    passed = passed + 1
    If LinkedRadius(1#, 0.3, "same") <> 1# Then Err.Raise 5, , "Same-radius link failed."
    passed = passed + 1
    If Abs(LinkedRadius(1#, 0.3, "subtract") - 0.7) > 0.000001 Then Err.Raise 5, , "Inset-radius link failed."
    passed = passed + 1
    If LinkedRadius(0.2, 0.3, "subtract") <> 0 Then Err.Raise 5, , "Inset lower bound failed."
    passed = passed + 1
    If LinkedRadius(1#, 0.3, "off") <> 0 Then Err.Raise 5, , "Radius-link off failed."
    passed = passed + 1
    If ParameterStepValue(1#, "rows", 4, 1) <> 2# Then Err.Raise 5, , "Row step failed."
    passed = passed + 1
    If ParameterStepValue(4#, "columns", 4, -1) <> 3# Then Err.Raise 5, , "Column step failed."
    passed = passed + 1
    If ParameterStepValue(2#, "rows", 30, -1) <> 2# Then Err.Raise 5, , "Grid-capacity lower bound failed."
    passed = passed + 1
    If ParameterStepValue(25#, "columns", 625, 1) <> 25# Then Err.Raise 5, , "Grid step upper bound failed."
    passed = passed + 1
    If ParameterStepValue(1#, "rows", 1, 1) <> 1# Or ParameterStepValue(1#, "rows", 1, -1) <> 1# Then Err.Raise 5, , "Single-child grid bound failed."
    passed = passed + 1
    If Abs(ParameterStepValue(0.123456, "padding", 4, 1) - 0.223456) > 0.000001 Then Err.Raise 5, , "Padding-step precision failed."
    passed = passed + 1
    If Abs(ParameterStepValue(0.2, "gap", 4, -1) - 0.1) > 0.000001 Then Err.Raise 5, , "Gap step failed."
    passed = passed + 1
    If ParameterStepValue(0.05, "gap", 4, -1) <> 0# Then Err.Raise 5, , "Spacing-step lower bound failed."
    passed = passed + 1
    SelfTest = passed
End Function
