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
    If PptNativeDriver.ShapeType(shape) = 6 Then
        If PptNativeDriver.ReadTag(shape, "layoutParent_v1") <> "" Or PptNativeDriver.ReadTag(shape, "layoutChild_v1") <> "" Then Err.Raise 5, , "Use the Office.js edition for tagged layouts."
        For Each child In PptNativeDriver.Children(shape)
            Set prepared = PrepareNode(child, value, unit, depth + 1, action)
            eligible = eligible + CLng(prepared("Eligible"))
            nodes.Add prepared
        Next child
    ElseIf IsRoundRect(shape) Then
        ' Full live preflight happens before ANY write or ungroup operation.
        If action = "radius" And PptNativeDriver.ReadTag(shape, STRICT_KEY) = "1" Then Err.Raise 5, , "Turn off protection explicitly before applying radius."
        If PptNativeDriver.ReadTag(shape, "layoutParent_v1") <> "" Or PptNativeDriver.ReadTag(shape, "layoutChild_v1") <> "" Then Err.Raise 5, , "Use the Office.js edition for tagged layouts."
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

Private Sub WriteNode(ByVal node As Collection, ByVal slide As Object)
    Dim shape As Object, fresh As Collection, children As Collection, child As Collection
    Dim opened As Boolean, mapped As Boolean, errorNumber As Long, errorText As String, restoreError As String
    Dim actualCm As Double
    Set shape = node("Shape")
    Set children = node("Children")
    If CLng(node("Eligible")) = 0 Then Exit Sub
    On Error GoTo Failed
    If children.Count > 0 Then
        Set fresh = PptNativeDriver.Ungroup(shape)
        opened = True
        For Each child In children
            SetNodeShape child, PptNativeDriver.FindShape(fresh, CLng(child("Id")))
        Next child
        mapped = True
        For Each child In children
            WriteNode child, slide
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
    Dim errorNumber As Long, errorText As String
    Set roots = PptNativeDriver.SelectionRoots()
    Set slide = PptNativeDriver.CurrentSlide()
    If action <> "radius" And action <> "protect" And action <> "unprotect" Then Err.Raise 5, , "Unsupported edit operation."
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
End Sub
