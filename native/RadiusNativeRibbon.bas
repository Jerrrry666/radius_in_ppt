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
    If Not ribbon Is Nothing Then ribbon.Invalidate
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
    returnedValue = False
    If ReadSummary(count, protectedCount, editable) Then
        returnedValue = (count > 0 And editable And protectedCount = 0)
    End If
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
    If Not ReadSummary(count, protectedCount, editable) Then Exit Sub
    returnedValue = (count > 0 And editable)
    If control.Tag = "off" Then returnedValue = (returnedValue And protectedCount > 0)
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
    InvalidateSelection
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
    MsgBox "Native radius core: 10 checks passed.", vbInformation, "R-corner Native"
    Exit Sub
Failed:
    ReportError Err.Description
End Sub

Private Sub ReportError(ByVal message As String)
    Debug.Print "[NativeRibbon] " & message
    MsgBox message, vbExclamation, "R-corner Native"
End Sub
