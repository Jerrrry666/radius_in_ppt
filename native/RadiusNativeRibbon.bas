Attribute VB_Name = "RadiusNativeRibbon"
Option Explicit

Private ribbon As Object
Private events As RadiusNativeEvents
Private radiusValue As Double
Private radiusUnit As String
Private editing As Boolean

' Display snapshots last only until the next selection/edit invalidation.
' Action callbacks never pass cached shapes to a write operation.
Private summaryReady As Boolean
Private summaryCount As Long
Private summaryProtected As Long
Private summaryEditable As Boolean
Private summaryError As String
Private relationsReady As Boolean
Private relationsSnapshot As Collection
Private relationsError As String
Private layoutReady As Boolean
Private layoutSnapshot As Collection
Private layoutError As String
Private writeReady As Boolean
Private writeAllowed As Boolean
Private writeError As String

Private Function UiReadBlocked() As Boolean
    UiReadBlocked = (editing Or RadiusNativeCore.IsTransactionActive)
End Function

Private Sub ClearUiSnapshots()
    summaryReady = False
    summaryError = ""
    relationsReady = False
    Set relationsSnapshot = Nothing
    relationsError = ""
    layoutReady = False
    Set layoutSnapshot = Nothing
    layoutError = ""
    writeReady = False
    writeAllowed = False
    writeError = ""
End Sub

Private Function ReadUiRelations() As Boolean
    If UiReadBlocked Then Exit Function
    If Not relationsReady Then
        relationsReady = True
        On Error GoTo Failed
        Set relationsSnapshot = RadiusNativeRelations.ReadUiSnapshot()
    End If
    ReadUiRelations = (relationsError = "")
    Exit Function
Failed:
    relationsError = Err.Description
    Debug.Print "[NativeRelationUI] " & relationsError
End Function

Private Function ReadUiLayout() As Boolean
    Dim group As Collection
    If UiReadBlocked Then Exit Function
    If Not layoutReady Then
        layoutReady = True
        If Not ReadUiRelations Then layoutError = relationsError: Exit Function
        On Error GoTo Failed
        If CBool(relationsSnapshot("HasGroup")) Then Set group = relationsSnapshot("Group")
        Set layoutSnapshot = RadiusNativeLayout.ReadUiSnapshot(group)
    End If
    ReadUiLayout = (layoutError = "")
    Exit Function
Failed:
    layoutError = Err.Description
    Debug.Print "[NativeLayoutUI] " & layoutError
End Function

Private Function ReadUiWrite() As Boolean
    Dim count As Long, protectedCount As Long, editable As Boolean, parents As Collection
    If UiReadBlocked Then Exit Function
    If Not writeReady Then
        writeReady = True
        On Error GoTo Failed
        If Not ReadSummary(count, protectedCount, editable) Then writeError = summaryError: Exit Function
        If Not RadiusNativeRelations.IsPreviewContext And count > 0 And editable And protectedCount = 0 Then
            Set parents = RadiusNativeLayout.SelectedLinkedParents(PptNativeDriver.SelectionRoots())
            If parents.Count = 0 Then
                writeAllowed = True
            Else
                If Not ReadUiRelations Then writeError = relationsError: Exit Function
                writeAllowed = RadiusNativeLayout.LinkedParentsWritable(parents, relationsSnapshot("Groups"))
            End If
        End If
    End If
    ReadUiWrite = (writeError = "")
    Exit Function
Failed:
    writeError = Err.Description
    Debug.Print "[NativeRadiusUI] " & writeError
End Function

Public Sub NativeOnLoad(ByVal ui As Object)
    On Error GoTo Failed
    Set ribbon = ui
    radiusValue = 0.3
    radiusUnit = "cm"
    ClearUiSnapshots
    Set events = New RadiusNativeEvents
    PptNativeDriver.ConnectEvents events
    Exit Sub
Failed:
    ReportError Err.Description
End Sub

Public Sub InvalidateSelection()
    ' Ungroup/regroup generates intermediate selection events. Read only
    ' after the core transaction has restored the final group and selection.
    ClearUiSnapshots
    If UiReadBlocked Then Exit Sub
    RadiusNativeRelations.RefreshContext
    If Not ribbon Is Nothing Then ribbon.Invalidate
End Sub

Public Sub NativeHostChanged()
    On Error GoTo Failed
    If UiReadBlocked Then Exit Sub
    ClearUiSnapshots
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
    If UiReadBlocked Then Exit Sub
    ClearUiSnapshots
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
    count = 0
    protectedCount = 0
    editable = False
    If UiReadBlocked Then Exit Function
    If Not summaryReady Then
        summaryReady = True
        On Error GoTo Failed
        RadiusNativeCore.SelectionSummary summaryCount, summaryProtected, summaryEditable
    End If
    If summaryError <> "" Then Exit Function
    count = summaryCount
    protectedCount = summaryProtected
    editable = summaryEditable
    ReadSummary = True
    Exit Function
Failed:
    summaryError = Err.Description
    Debug.Print "[NativeSelection] " & summaryError
End Function

Public Sub NativeGetValue(ByVal control As Object, ByRef returnedValue)
    returnedValue = Format$(radiusValue, "0.00####")
End Sub

Public Sub NativeChangeValue(ByVal control As Object, ByVal text As String)
    On Error GoTo Failed
    If UiReadBlocked Then Err.Raise 5, , "An edit is already in progress."
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
    If UiReadBlocked Then Exit Sub
    If selectedIndex = 1 Then radiusUnit = "%" Else radiusUnit = "cm"
    InvalidateSelection
End Sub

Public Sub NativeGetWriteEnabled(ByVal control As Object, ByRef returnedValue)
    returnedValue = False
    If ReadUiWrite Then returnedValue = writeAllowed
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
    If UiReadBlocked Then Exit Sub
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
    If UiReadBlocked Then Exit Sub
    If CStr(control.Id) = "NativeRelPreview" And RadiusNativeRelations.IsPreviewOpen Then returnedValue = True: Exit Sub
    If Not ReadUiRelations Then Exit Sub
    If CStr(control.Id) = "NativeRelView" Then
        returnedValue = relationsSnapshot("CanView")
        Exit Sub
    End If
    text = Split(CStr(control.Tag), "|")
    Select Case CStr(text(0))
        Case "parent": returnedValue = relationsSnapshot("CanMarkParent")
        Case "bind": returnedValue = relationsSnapshot("CanBind")
        Case "view": returnedValue = relationsSnapshot("CanView")
        Case "preview"
            If RadiusNativeRelations.IsPreviewOpen Then
                returnedValue = True
            Else
                returnedValue = relationsSnapshot("CanView")
            End If
        Case "child": returnedValue = relationsSnapshot("CanDetachChild")
        Case "whole": returnedValue = relationsSnapshot("CanDetachWhole")
        Case "cancel"
            Set info = relationsSnapshot("Info")
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
    returnedValue = "-"
    If UiReadBlocked Then Exit Sub
    If Not ReadUiRelations Then returnedValue = text(8) & ": " & relationsError: Exit Sub
    Set info = relationsSnapshot("Info")
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
    returnedValue = text(8) & ": " & Err.Description
    Debug.Print "[NativeRelationStatus] " & Err.Description
End Sub

Public Sub NativeRelGetInfo(ByVal control As Object, ByRef returnedValue)
    Dim info As Collection, text As Variant, state As String
    text = Split(CStr(control.Tag), "|")
    On Error GoTo Failed
    returnedValue = "-"
    If UiReadBlocked Then Exit Sub
    If Not ReadUiRelations Then returnedValue = relationsError: Exit Sub
    Set info = relationsSnapshot("Info")
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
    returnedValue = Err.Description
    Debug.Print "[NativeRelationInfo] " & Err.Description
End Sub

Private Sub RunRelation(ByVal action As String)
    Dim errorNumber As Long, errorText As String
    If UiReadBlocked Then Err.Raise 5, , "An edit is already in progress."
    On Error GoTo Failed
    ClearUiSnapshots
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
    If UiReadBlocked Then Err.Raise 5, , "An edit is already in progress."
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
    returnedValue = "<menu xmlns=" & Chr$(34) & "http://schemas.microsoft.com/office/2006/01/customui" & Chr$(34) & "/>"
    If UiReadBlocked Then Exit Sub
    If Not ReadUiRelations Then Err.Raise 5, , relationsError
    Set groups = RadiusNativeRelations.BeginMenuFrom(relationsSnapshot("Groups"))
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
    returnedValue = "<menu xmlns=" & Chr$(34) & "http://schemas.microsoft.com/office/2006/01/customui" & Chr$(34) & "><button id=" & Chr$(34) & "RelMenu_Error" & Chr$(34) & " label=" & Chr$(34) & XmlText(text(6)) & Chr$(34) & " supertip=" & Chr$(34) & XmlText(Err.Description) & Chr$(34) & " enabled=" & Chr$(34) & "false" & Chr$(34) & "/></menu>"
End Sub

Private Function MemberButton(ByVal id As String, ByVal label As String, ByVal token As String) As String
    MemberButton = "<button id=" & Chr$(34) & "RelItem_" & id & Chr$(34) & " label=" & Chr$(34) & XmlText(label) & Chr$(34) & " onAction=" & Chr$(34) & "NativeRelLocate" & Chr$(34) & " tag=" & Chr$(34) & XmlText(token) & Chr$(34) & "/>"
End Function

Public Sub NativeRelLocate(ByVal control As Object)
    On Error GoTo Failed
    If UiReadBlocked Then Err.Raise 5, , "An edit is already in progress."
    ClearUiSnapshots
    RadiusNativeRelations.Locate CStr(control.Tag)
    InvalidateSelection
    Exit Sub
Failed:
    ReportError Err.Description
End Sub

Public Sub NativeRelPreviewPressed(ByVal control As Object, ByRef returnedValue)
    returnedValue = False
    If UiReadBlocked Then Exit Sub
    returnedValue = RadiusNativeRelations.IsPreviewOpen
End Sub

Public Sub NativeRelPreview(ByVal control As Object, ByVal pressed As Boolean)
    Dim text As Variant, errorNumber As Long, errorText As String
    On Error GoTo Failed
    If UiReadBlocked Then Exit Sub
    ClearUiSnapshots
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
    If Not ReadUiLayout Then Exit Sub
    returnedValue = layoutSnapshot("CanConfigure")
    If Not returnedValue Then Exit Sub
    If control.Id = "NativeLayoutApply" Then returnedValue = layoutSnapshot("ChildrenWritable")
    If control.Id = "NativeLayoutAuto" Then
        If Not CBool(layoutSnapshot("Automatic")) Then returnedValue = layoutSnapshot("ChildrenWritable")
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
    If Not ReadUiLayout Then Exit Sub
    If Not CBool(layoutSnapshot("CanConfigure")) Then Exit Sub
    config = layoutSnapshot("Config")
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
    If UiReadBlocked Then Err.Raise 5, , "An edit is already in progress."
    ClearUiSnapshots
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
    Dim key As String, direction As Long, suffix As String
    On Error GoTo Failed
    returnedValue = False
    If Not ReadUiLayout Then Exit Sub
    If Not CBool(layoutSnapshot("CanConfigure")) Then Exit Sub
    direction = LayoutStepDirection(CStr(control.Tag), key)
    If direction = 1 Then suffix = "Up" Else suffix = "Down"
    returnedValue = layoutSnapshot(key & suffix)
    Exit Sub
Failed:
    returnedValue = False
    RadiusNativeLayout.RememberError Err.Description
End Sub

Public Sub NativeLayoutStep(ByVal control As Object)
    Dim key As String, direction As Long, errorText As String
    On Error GoTo Failed
    If UiReadBlocked Then Err.Raise 5, , "An edit is already in progress."
    ClearUiSnapshots
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
    If Not ReadUiLayout Then Exit Sub
    If Not CBool(layoutSnapshot("CanConfigure")) Then Exit Sub
    config = layoutSnapshot("Config")
    If CStr(config(4)) = "subtract" Then returnedValue = 1
    If CStr(config(4)) = "off" Then returnedValue = 2
    Exit Sub
Failed:
    RadiusNativeLayout.RememberError Err.Description
End Sub

Private Sub RunLayout(ByVal action As String, ByVal value As String)
    Dim errorNumber As Long, errorText As String
    If UiReadBlocked Then Err.Raise 5, , "An edit is already in progress."
    On Error GoTo Failed
    ClearUiSnapshots
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
    returnedValue = False
    If ReadUiLayout Then returnedValue = layoutSnapshot("Automatic")
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
    returnedValue = "-"
    If UiReadBlocked Then Exit Sub
    If Not ReadUiLayout Then returnedValue = text(3) & layoutError: Exit Sub
    If Not ReadUiWrite Then returnedValue = text(3) & writeError: Exit Sub
    returnedValue = RadiusNativeLayout.StatusFromSnapshot(layoutSnapshot, CStr(text(0)), CStr(text(1)), CStr(text(2)), CStr(text(3)), CStr(text(4)))
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
    ElseIf summaryError <> "" Then
        returnedValue = "R: - | " & summaryError
    End If
End Sub

Private Sub RunEdit(ByVal value As Double, ByVal unit As String, ByVal action As String)
    Dim errorNumber As Long, errorText As String
    If UiReadBlocked Then Err.Raise 5, , "An edit is already in progress."
    On Error GoTo Failed
    ClearUiSnapshots
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
    If UiReadBlocked Then Err.Raise 5, , "An edit is already in progress."
    ClearUiSnapshots
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

Public Sub NativeSelfTestEnabled(ByVal control As Object, ByRef returnedValue)
    returnedValue = False
    On Error GoTo Failed
    returnedValue = Not UiReadBlocked And Not RadiusNativeRelations.IsPreviewOpen
    Exit Sub
Failed:
    Debug.Print "[NativeQuickTestEnabled] " & Err.Description
End Sub

Public Sub NativeSelfTest(ByVal control As Object)
    Dim report As String, started As Boolean, errorText As String, icon As Long
    On Error GoTo Failed
    If UiReadBlocked Then Err.Raise 5, , "An edit is already in progress."
    If RadiusNativeRelations.IsPreviewOpen Then Err.Raise 5, , "Close the relation preview before running the quick test."
    ClearUiSnapshots
    editing = True
    started = True
    report = RadiusNativeQuickTest.RunnerRun()
    editing = False
    started = False
    InvalidateSelection
    icon = vbInformation
    If RadiusNativeQuickTest.LastFailureCount() > 0 Then icon = vbExclamation
    ShowQuickTestReport report, icon
    Exit Sub
Failed:
    errorText = Err.Description
    If started Then
        editing = False
        InvalidateSelection
    End If
    ShowQuickTestReport errorText, vbExclamation
End Sub

Private Sub ShowQuickTestReport(ByVal report As String, ByVal icon As Long)
    Dim position As Long, count As Long, text As String, boundary As Long
    Debug.Print "[NativeQuickTest] " & report
    position = 1
    ' MsgBox prompts have a practical length limit. Keep every error visible.
    Do While position <= Len(report)
        count = 900
        text = Mid$(report, position, count)
        If position + Len(text) <= Len(report) Then
            boundary = InStrRev(text, vbCrLf)
            If boundary > 1 Then text = Left$(text, boundary + 1)
        End If
        MsgBox text, icon, "R-corner Native quick test"
        position = position + Len(text)
    Loop
End Sub

Private Sub ReportError(ByVal message As String)
    Debug.Print "[NativeRibbon] " & message
    MsgBox message, vbExclamation, "R-corner Native"
End Sub
