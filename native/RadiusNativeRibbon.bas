Attribute VB_Name = "RadiusNativeRibbon"
Option Explicit

Private ribbon As Object
Private events As RadiusNativeEvents
Private radiusValue As Double
Private radiusUnit As String
Private editing As Boolean

Public Sub NativeOnLoad(ByVal ui As Object)
    On Error GoTo Failed
    Set ribbon = ui
    radiusValue = 0.3
    radiusUnit = "cm"
    Set events = New RadiusNativeEvents
    PptNativeDriver.ConnectEvents events
    Exit Sub
Failed:
    ReportError Err.Description
End Sub

Public Sub InvalidateSelection()
    ' Ungroup/regroup generates intermediate selection events. Read only
    ' after the core transaction has restored the final group and selection.
    If editing Then Exit Sub
    RadiusNativeRelations.RefreshContext
    If Not ribbon Is Nothing Then ribbon.Invalidate
End Sub

Public Sub NativeHostChanged()
    On Error GoTo Failed
    If editing Then Exit Sub
    editing = True
    RadiusNativeLayout.SyncCurrent
    editing = False
    InvalidateSelection
    Exit Sub
Failed:
    RadiusNativeLayout.RememberError Err.Description
    editing = False
    InvalidateSelection
End Sub

Public Sub NativeBeforeSave(ByVal presentation As Object)
    Dim errorText As String
    On Error GoTo Failed
    If editing Then Exit Sub
    editing = True
    RadiusNativeLayout.SyncPresentation presentation
    editing = False
    InvalidateSelection
    Exit Sub
Failed:
    errorText = Err.Description
    RadiusNativeLayout.RememberError errorText
    editing = False
    InvalidateSelection
    ReportError "Automatic layout/radius synchronization stopped: " & errorText
End Sub

Private Function ReadSummary(ByRef count As Long, ByRef protectedCount As Long, ByRef editable As Boolean) As Boolean
    On Error GoTo Failed
    If editing Then Exit Function
    RadiusNativeCore.SelectionSummary count, protectedCount, editable
    ReadSummary = True
    Exit Function
Failed:
    Debug.Print "[NativeSelection] " & Err.Description
    count = 0
    protectedCount = 0
    editable = False
End Function

Public Sub NativeGetValue(ByVal control As Object, ByRef returnedValue)
    returnedValue = Format$(radiusValue, "0.00####")
End Sub

Public Sub NativeChangeValue(ByVal control As Object, ByVal text As String)
    On Error GoTo Failed
    radiusValue = RadiusNativeCore.ParseNumber(text)
    InvalidateSelection
    Exit Sub
Failed:
    InvalidateSelection
    ReportError Err.Description
End Sub

Public Sub NativeGetUnit(ByVal control As Object, ByRef returnedValue)
    If radiusUnit = "%" Then returnedValue = 1 Else returnedValue = 0
End Sub

Public Sub NativeChangeUnit(ByVal control As Object, ByVal selectedId As String, ByVal selectedIndex As Integer)
    If selectedIndex = 1 Then radiusUnit = "%" Else radiusUnit = "cm"
    InvalidateSelection
End Sub

Public Sub NativeGetWriteEnabled(ByVal control As Object, ByRef returnedValue)
    Dim count As Long, protectedCount As Long, editable As Boolean
    On Error GoTo Failed
    returnedValue = False
    If RadiusNativeRelations.IsPreviewContext Then Exit Sub
    If ReadSummary(count, protectedCount, editable) Then
        returnedValue = (count > 0 And editable And protectedCount = 0)
        If returnedValue Then returnedValue = RadiusNativeLayout.LinkedSelectionWritable
    End If
    Exit Sub
Failed:
    returnedValue = False
    RadiusNativeLayout.RememberError Err.Description
End Sub

Public Sub NativeGetStepEnabled(ByVal control As Object, ByRef returnedValue)
    NativeGetWriteEnabled control, returnedValue
    If Not returnedValue Then Exit Sub
    If control.Tag = "down" Then
        returnedValue = (radiusValue > 0)
    ElseIf radiusUnit = "%" Then
        returnedValue = (radiusValue < 50)
    Else
        returnedValue = (radiusValue < 1000000)
    End If
End Sub

Public Sub NativeGetReadEnabled(ByVal control As Object, ByRef returnedValue)
    Dim count As Long, protectedCount As Long, editable As Boolean
    returnedValue = False
    If ReadSummary(count, protectedCount, editable) Then returnedValue = (count > 0)
End Sub

Public Sub NativeGetProtectionEnabled(ByVal control As Object, ByRef returnedValue)
    Dim count As Long, protectedCount As Long, editable As Boolean
    returnedValue = False
    If RadiusNativeRelations.IsPreviewContext Then Exit Sub
    If Not ReadSummary(count, protectedCount, editable) Then Exit Sub
    returnedValue = (count > 0 And editable)
    If control.Tag = "off" Then returnedValue = (returnedValue And protectedCount > 0)
End Sub

Private Function XmlText(ByVal text As String) As String
    text = Replace(text, "&", "&amp;")
    text = Replace(text, "<", "&lt;")
    text = Replace(text, ">", "&gt;")
    text = Replace(text, Chr$(34), "&quot;")
    XmlText = Replace(text, Chr$(39), "&apos;")
End Function

Private Function ShortLabel(ByVal text As String) As String
    If Len(text) > 20 Then ShortLabel = Left$(text, 19) & "..." Else ShortLabel = text
End Function

Public Sub NativeRelGetEnabled(ByVal control As Object, ByRef returnedValue)
    Dim text As Variant, info As Collection
    On Error GoTo Failed
    returnedValue = False
    If editing Then Exit Sub
    If CStr(control.Id) = "NativeRelView" Then
        returnedValue = RadiusNativeRelations.CanView
        Exit Sub
    End If
    text = Split(CStr(control.Tag), "|")
    Select Case CStr(text(0))
        Case "parent": returnedValue = RadiusNativeRelations.CanMarkParent
        Case "bind": returnedValue = RadiusNativeRelations.CanBind
        Case "view": returnedValue = RadiusNativeRelations.CanView
        Case "preview"
            If RadiusNativeRelations.IsPreviewOpen Then
                returnedValue = True
            Else
                returnedValue = RadiusNativeRelations.CanView
            End If
        Case "child": returnedValue = RadiusNativeRelations.CanDetach(False)
        Case "whole": returnedValue = RadiusNativeRelations.CanDetach(True)
        Case "cancel"
            Set info = RadiusNativeRelations.SelectionInfo()
            returnedValue = (CStr(info("State")) = "pending")
    End Select
    Exit Sub
Failed:
    returnedValue = False
    Debug.Print "[NativeRelationUI] " & Err.Description
End Sub

Public Sub NativeRelGetStatus(ByVal control As Object, ByRef returnedValue)
    Dim info As Collection, text As Variant, state As String, role As String
    text = Split(CStr(control.Tag), "|")
    On Error GoTo Failed
    Set info = RadiusNativeRelations.SelectionInfo()
    state = CStr(info("State"))
    Select Case state
        Case "empty": returnedValue = text(5)
        Case "unbound": returnedValue = text(0) & text(1)
        Case "mixed"
            returnedValue = text(6) & CStr(info("Keys"))
            If CLng(info("Unbound")) > 0 Then returnedValue = returnedValue & " | " & text(11) & CStr(info("Unbound"))
        Case "preview": returnedValue = text(7)
        Case "pending": returnedValue = text(4) & ShortLabel(CStr(info("ParentName")))
        Case "broken": returnedValue = CStr(info("Key")) & " | " & text(9)
        Case "bound"
            role = CStr(info("Role"))
            If role = "P" Then
                returnedValue = CStr(info("Key")) & " | " & text(2) & " " & ShortLabel(CStr(info("Name")))
            ElseIf Left$(role, 1) = "C" Then
                returnedValue = CStr(info("Key")) & " | " & text(3) & Mid$(role, 2) & " " & ShortLabel(CStr(info("Name")))
            Else
                returnedValue = CStr(info("Key")) & " | " & CStr(info("Selected")) & text(10)
            End If
    End Select
    Exit Sub
Failed:
    returnedValue = text(8)
    Debug.Print "[NativeRelationStatus] " & Err.Description
End Sub

Public Sub NativeRelGetInfo(ByVal control As Object, ByRef returnedValue)
    Dim info As Collection, text As Variant, state As String
    text = Split(CStr(control.Tag), "|")
    On Error GoTo Failed
    Set info = RadiusNativeRelations.SelectionInfo()
    state = CStr(info("State"))
    If state = "pending" Then
        Select Case CStr(info("PendingStatus"))
            Case "self": returnedValue = text(5)
            Case "conflict": returnedValue = text(6)
            Case "ready": returnedValue = text(3) & CStr(info("Selected"))
            Case Else: returnedValue = text(4)
        End Select
    ElseIf state = "preview" Then
        returnedValue = text(7)
    ElseIf state = "bound" Or state = "broken" Then
        returnedValue = text(0) & ShortLabel(CStr(info("ParentName"))) & " | " & text(1) & CStr(info("Children"))
    Else
        returnedValue = text(2) & CStr(info("Groups"))
    End If
    Exit Sub
Failed:
    returnedValue = "-"
    Debug.Print "[NativeRelationInfo] " & Err.Description
End Sub

Private Sub RunRelation(ByVal action As String)
    Dim errorNumber As Long, errorText As String
    If editing Then Err.Raise 5, , "An edit is already in progress."
    On Error GoTo Failed
    editing = True
    Select Case action
        Case "parent": RadiusNativeRelations.MarkParent
        Case "bind": RadiusNativeRelations.BindChildren
        Case "cancel": RadiusNativeRelations.CancelPending
        Case "child": RadiusNativeRelations.Detach False
        Case "whole": RadiusNativeRelations.Detach True
        Case Else: Err.Raise 5, , "Unsupported relationship operation."
    End Select
    editing = False
    NativeHostChanged
    Exit Sub
Failed:
    errorNumber = Err.Number
    errorText = Err.Description
    editing = False
    InvalidateSelection
    Err.Raise errorNumber, "RadiusNativeRibbon", errorText
End Sub

Public Sub NativeRelAction(ByVal control As Object)
    Dim group As Collection, message As String, text As Variant, parentName As String
    On Error GoTo Failed
    text = Split(CStr(control.Tag), "|")
    If CStr(text(0)) = "whole" Then
        Set group = RadiusNativeRelations.RemovalInfo()
        parentName = CStr(group("ParentName"))
        message = Replace(CStr(text(1)), "{key}", CStr(group("Key")))
        message = Replace(message, "{parent}", parentName)
        message = Replace(message, "{count}", CStr(group("Children").Count))
        If MsgBox(message, vbQuestion Or vbYesNo, "R-corner Native") <> vbYes Then Exit Sub
    End If
    RunRelation CStr(text(0))
    Exit Sub
Failed:
    ReportError Err.Description
End Sub

Public Sub NativeRelGetMenu(ByVal control As Object, ByRef returnedValue)
    Dim groups As Collection, group As Collection, child As Collection, text As Variant, xml As String, key As String, name As String, id As Long
    text = Split(CStr(control.Tag), "|")
    On Error GoTo Failed
    Set groups = RadiusNativeRelations.BeginMenu()
    xml = "<menu xmlns=" & Chr$(34) & "http://schemas.microsoft.com/office/2006/01/customui" & Chr$(34) & ">"
    For Each group In groups
        key = CStr(group("Key"))
        name = text(5)
        If CLng(group("ParentId")) > 0 Then name = CStr(group("ParentName"))
        If CStr(group("Problem")) <> "" Then name = name & " [!]"
        If groups.Count > 1 Then xml = xml & "<menu id=" & Chr$(34) & "RelMenu_" & key & Chr$(34) & " label=" & Chr$(34) & XmlText(key & " | " & ShortLabel(name)) & Chr$(34) & ">"
        If CLng(group("ParentId")) > 0 Then xml = xml & MemberButton(key & "_P", text(0) & name, RadiusNativeRelations.MenuToken(key, "object", CLng(group("ParentId"))))
        For Each child In group("Children")
            id = PptNativeDriver.ShapeId(child("Shape"))
            xml = xml & MemberButton(key & "_C" & CStr(id), text(1) & CStr(child("Order")) & ": " & PptNativeDriver.ShapeName(child("Shape")), RadiusNativeRelations.MenuToken(key, "object", id))
        Next child
        xml = xml & "<menuSeparator id=" & Chr$(34) & "RelSep_" & key & Chr$(34) & "/>"
        xml = xml & MemberButton(key & "_All", text(2), RadiusNativeRelations.MenuToken(key, "family", 0))
        xml = xml & MemberButton(key & "_Children", text(3), RadiusNativeRelations.MenuToken(key, "children", 0))
        If groups.Count > 1 Then xml = xml & "</menu>"
    Next group
    If groups.Count = 0 Then xml = xml & "<button id=" & Chr$(34) & "RelMenu_Empty" & Chr$(34) & " label=" & Chr$(34) & XmlText(text(4)) & Chr$(34) & " enabled=" & Chr$(34) & "false" & Chr$(34) & "/>"
    returnedValue = xml & "</menu>"
    Exit Sub
Failed:
    Debug.Print "[NativeRelationMenu] " & Err.Description
    returnedValue = "<menu xmlns=" & Chr$(34) & "http://schemas.microsoft.com/office/2006/01/customui" & Chr$(34) & "><button id=" & Chr$(34) & "RelMenu_Error" & Chr$(34) & " label=" & Chr$(34) & XmlText(text(6)) & Chr$(34) & " enabled=" & Chr$(34) & "false" & Chr$(34) & "/></menu>"
End Sub

Private Function MemberButton(ByVal id As String, ByVal label As String, ByVal token As String) As String
    MemberButton = "<button id=" & Chr$(34) & "RelItem_" & id & Chr$(34) & " label=" & Chr$(34) & XmlText(label) & Chr$(34) & " onAction=" & Chr$(34) & "NativeRelLocate" & Chr$(34) & " tag=" & Chr$(34) & XmlText(token) & Chr$(34) & "/>"
End Function

Public Sub NativeRelLocate(ByVal control As Object)
    On Error GoTo Failed
    RadiusNativeRelations.Locate CStr(control.Tag)
    InvalidateSelection
    Exit Sub
Failed:
    ReportError Err.Description
End Sub

Public Sub NativeRelPreviewPressed(ByVal control As Object, ByRef returnedValue)
    returnedValue = RadiusNativeRelations.IsPreviewOpen
End Sub

Public Sub NativeRelPreview(ByVal control As Object, ByVal pressed As Boolean)
    Dim text As Variant, errorNumber As Long, errorText As String
    On Error GoTo Failed
    If editing Then Exit Sub
    text = Split(CStr(control.Tag), "|")
    editing = True
    RadiusNativeRelations.TogglePreview CStr(text(1)), CStr(text(2)), CStr(text(3))
    editing = False
    InvalidateSelection
    Exit Sub
Failed:
    errorNumber = Err.Number
    errorText = Err.Description
    editing = False
    InvalidateSelection
    ReportError errorText
End Sub

Public Sub NativeLayoutGetEnabled(ByVal control As Object, ByRef returnedValue)
    On Error GoTo Failed
    returnedValue = False
    If editing Then Exit Sub
    returnedValue = RadiusNativeLayout.CanConfigure
    If Not returnedValue Then Exit Sub
    If control.Id = "NativeLayoutApply" Then returnedValue = RadiusNativeLayout.ChildrenWritable
    If control.Id = "NativeLayoutAuto" Then
        If Not RadiusNativeLayout.AutomaticEnabled Then returnedValue = RadiusNativeLayout.ChildrenWritable
    End If
    Exit Sub
Failed:
    returnedValue = False
    RadiusNativeLayout.RememberError Err.Description
End Sub

Public Sub NativeLayoutGetText(ByVal control As Object, ByRef returnedValue)
    Dim config As Variant
    On Error GoTo Failed
    returnedValue = "-"
    If Not RadiusNativeLayout.CanConfigure Then Exit Sub
    config = RadiusNativeLayout.SelectedConfig()
    Select Case CStr(control.Tag)
        Case "rows": returnedValue = CStr(config(0))
        Case "columns": returnedValue = CStr(config(1))
        Case "padding": returnedValue = Format$(CDbl(config(2)), "0.00####")
        Case "gap": returnedValue = Format$(CDbl(config(3)), "0.00####")
    End Select
    Exit Sub
Failed:
    RadiusNativeLayout.RememberError Err.Description
End Sub

Public Sub NativeLayoutParameter(ByVal control As Object, ByVal text As String)
    Dim errorText As String
    On Error GoTo Failed
    RadiusNativeLayout.SetParameter CStr(control.Tag), text
    InvalidateSelection
    Exit Sub
Failed:
    errorText = Err.Description
    RadiusNativeLayout.RememberError errorText
    InvalidateSelection
    ReportError errorText
End Sub

Private Function LayoutStepDirection(ByVal tag As String, ByRef key As String) As Long
    Dim fields As Variant
    fields = Split(tag, "|")
    If UBound(fields) <> 1 Then Err.Raise 5, , "Invalid layout step control."
    key = CStr(fields(0))
    Select Case CStr(fields(1))
        Case "up": LayoutStepDirection = 1
        Case "down": LayoutStepDirection = -1
        Case Else: Err.Raise 5, , "Unsupported step direction."
    End Select
End Function

Public Sub NativeLayoutGetStepEnabled(ByVal control As Object, ByRef returnedValue)
    Dim key As String, direction As Long
    On Error GoTo Failed
    returnedValue = False
    If editing Then Exit Sub
    If Not RadiusNativeLayout.CanConfigure Then Exit Sub
    direction = LayoutStepDirection(CStr(control.Tag), key)
    returnedValue = RadiusNativeLayout.CanStepParameter(key, direction)
    Exit Sub
Failed:
    returnedValue = False
    RadiusNativeLayout.RememberError Err.Description
End Sub

Public Sub NativeLayoutStep(ByVal control As Object)
    Dim key As String, direction As Long, errorText As String
    On Error GoTo Failed
    If editing Then Err.Raise 5, , "An edit is already in progress."
    If Not RadiusNativeLayout.CanConfigure Then Exit Sub
    direction = LayoutStepDirection(CStr(control.Tag), key)
    RadiusNativeLayout.StepParameter key, direction
    InvalidateSelection
    Exit Sub
Failed:
    errorText = Err.Description
    RadiusNativeLayout.RememberError errorText
    InvalidateSelection
    ReportError errorText
End Sub

Public Sub NativeLayoutGetMode(ByVal control As Object, ByRef returnedValue)
    Dim config As Variant
    On Error GoTo Failed
    returnedValue = 0
    If Not RadiusNativeLayout.CanConfigure Then Exit Sub
    config = RadiusNativeLayout.SelectedConfig()
    If CStr(config(4)) = "subtract" Then returnedValue = 1
    If CStr(config(4)) = "off" Then returnedValue = 2
    Exit Sub
Failed:
    RadiusNativeLayout.RememberError Err.Description
End Sub

Private Sub RunLayout(ByVal action As String, ByVal value As String)
    Dim errorNumber As Long, errorText As String
    If editing Then Err.Raise 5, , "An edit is already in progress."
    On Error GoTo Failed
    editing = True
    Select Case action
        Case "apply": RadiusNativeLayout.ApplySelected
        Case "mode": RadiusNativeLayout.ChangeMode value
        Case "auto": RadiusNativeLayout.SetAutomatic (value = "1")
        Case Else: Err.Raise 5, , "Unsupported layout operation."
    End Select
    editing = False
    InvalidateSelection
    Exit Sub
Failed:
    errorNumber = Err.Number
    errorText = Err.Description
    RadiusNativeLayout.ResetPending
    RadiusNativeLayout.RememberError errorText
    editing = False
    InvalidateSelection
    Err.Raise errorNumber, "RadiusNativeRibbon", errorText
End Sub

Public Sub NativeLayoutApply(ByVal control As Object)
    On Error GoTo Failed
    RunLayout "apply", ""
    Exit Sub
Failed:
    ReportError Err.Description
End Sub

Public Sub NativeLayoutMode(ByVal control As Object, ByVal selectedId As String, ByVal selectedIndex As Integer)
    Dim mode As String
    On Error GoTo Failed
    mode = "same"
    If selectedIndex = 1 Then mode = "subtract"
    If selectedIndex = 2 Then mode = "off"
    RunLayout "mode", mode
    Exit Sub
Failed:
    ReportError Err.Description
End Sub

Public Sub NativeLayoutAutoPressed(ByVal control As Object, ByRef returnedValue)
    returnedValue = RadiusNativeLayout.AutomaticEnabled
End Sub

Public Sub NativeLayoutAuto(ByVal control As Object, ByVal pressed As Boolean)
    Dim value As String
    On Error GoTo Failed
    value = "0"
    If pressed Then value = "1"
    RunLayout "auto", value
    Exit Sub
Failed:
    ReportError Err.Description
End Sub

Public Sub NativeLayoutStatus(ByVal control As Object, ByRef returnedValue)
    Dim text As Variant
    text = Split(CStr(control.Tag), "|")
    On Error GoTo Failed
    returnedValue = RadiusNativeLayout.Status(CStr(text(0)), CStr(text(1)), CStr(text(2)), CStr(text(3)), CStr(text(4)))
    Exit Sub
Failed:
    returnedValue = text(3) & Err.Description
    RadiusNativeLayout.RememberError Err.Description
End Sub

Public Sub NativeGetProtectionPressed(ByVal control As Object, ByRef returnedValue)
    Dim count As Long, protectedCount As Long, editable As Boolean
    returnedValue = False
    If ReadSummary(count, protectedCount, editable) Then returnedValue = (count > 0 And protectedCount = count)
End Sub

Public Sub NativeGetStatusVisible(ByVal control As Object, ByRef returnedValue)
    Dim count As Long, protectedCount As Long, editable As Boolean, state As String
    state = "error"
    If ReadSummary(count, protectedCount, editable) Then
        If count = 0 Then
            state = "empty"
        ElseIf Not editable Then
            state = "group"
        ElseIf protectedCount = 0 Then
            state = "off"
        ElseIf protectedCount = count Then
            state = "on"
        Else
            state = "mixed"
        End If
    End If
    returnedValue = (control.Tag = state)
End Sub

Public Sub NativeGetCountLabel(ByVal control As Object, ByRef returnedValue)
    Dim count As Long, protectedCount As Long, editable As Boolean
    returnedValue = "R: - | -/-"
    If ReadSummary(count, protectedCount, editable) Then
        returnedValue = "R: " & CStr(count) & " | " & CStr(protectedCount) & "/" & CStr(count)
    End If
End Sub

Private Sub RunEdit(ByVal value As Double, ByVal unit As String, ByVal action As String)
    Dim errorNumber As Long, errorText As String
    If editing Then Err.Raise 5, , "An edit is already in progress."
    On Error GoTo Failed
    editing = True
    RadiusNativeCore.ApplySelection value, unit, action
    If action = "radius" Then
        radiusValue = value
        radiusUnit = unit
    End If
    editing = False
    NativeHostChanged
    Exit Sub
Failed:
    errorNumber = Err.Number
    errorText = Err.Description
    editing = False
    InvalidateSelection
    Err.Raise errorNumber, "RadiusNativeRibbon", errorText
End Sub

Public Sub NativeApply(ByVal control As Object)
    On Error GoTo Failed
    RunEdit radiusValue, radiusUnit, "radius"
    Exit Sub
Failed:
    ReportError Err.Description
End Sub

Public Sub NativeStep(ByVal control As Object)
    Dim direction As Long, nextValue As Double
    On Error GoTo Failed
    If control.Tag = "up" Then
        direction = 1
    ElseIf control.Tag = "down" Then
        direction = -1
    Else
        Err.Raise 5, , "Unsupported step direction."
    End If
    nextValue = RadiusNativeCore.StepValue(radiusValue, radiusUnit, direction)
    If nextValue = radiusValue Then Exit Sub
    ' Commit the displayed value only after the live protected write succeeds.
    RunEdit nextValue, radiusUnit, "radius"
    Exit Sub
Failed:
    ReportError Err.Description
End Sub

Public Sub NativePreset(ByVal control As Object)
    On Error GoTo Failed
    RunEdit RadiusNativeCore.ParseNumber(control.Tag), "cm", "radius"
    Exit Sub
Failed:
    ReportError Err.Description
End Sub

Public Sub NativeRead(ByVal control As Object)
    Dim leaves As Collection
    On Error GoTo Failed
    Set leaves = RadiusNativeCore.SelectionLeaves()
    radiusValue = RadiusNativeCore.CurrentRadius(leaves(1))
    radiusUnit = "cm"
    InvalidateSelection
    Exit Sub
Failed:
    ReportError Err.Description
End Sub

Public Sub NativeToggleProtect(ByVal control As Object, ByVal pressed As Boolean)
    On Error GoTo Failed
    If pressed Then
        RunEdit radiusValue, radiusUnit, "protect"
    Else
        RunEdit radiusValue, radiusUnit, "unprotect"
    End If
    Exit Sub
Failed:
    ReportError Err.Description
End Sub

Public Sub NativeProtect(ByVal control As Object)
    On Error GoTo Failed
    If control.Tag = "on" Then
        RunEdit radiusValue, radiusUnit, "protect"
    Else
        RunEdit radiusValue, radiusUnit, "unprotect"
    End If
    Exit Sub
Failed:
    ReportError Err.Description
End Sub

Public Sub NativeSelfTest(ByVal control As Object)
    On Error GoTo Failed
    RadiusNativeCore.SelfTest
    RadiusNativeLayout.SelfTest
    MsgBox "Native radius/layout cores: 26 checks passed.", vbInformation, "R-corner Native"
    Exit Sub
Failed:
    ReportError Err.Description
End Sub

Private Sub ReportError(ByVal message As String)
    Debug.Print "[NativeRibbon] " & message
    MsgBox message, vbExclamation, "R-corner Native"
End Sub
