Attribute VB_Name = "RadiusNativeEvents"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit

Public WithEvents Host As PowerPoint.Application

Private Sub Host_WindowSelectionChange(ByVal Sel As PowerPoint.Selection)
    RadiusNativeRibbon.NativeHostChanged
End Sub

Private Sub Host_WindowActivate(ByVal Pres As PowerPoint.Presentation, ByVal Wn As PowerPoint.DocumentWindow)
    RadiusNativeRibbon.NativeHostChanged
End Sub

Private Sub Host_SlideSelectionChanged(ByVal SldRange As PowerPoint.SlideRange)
    RadiusNativeRibbon.NativeHostChanged
End Sub

Private Sub Host_AfterShapeSizeChange(ByVal shp As PowerPoint.Shape)
    ' PowerPoint raises this after resizing completes; Ribbon/Core guards ignore writes during transactions.
    RadiusNativeRibbon.NativeHostChanged
End Sub

Private Sub Host_PresentationBeforeSave(ByVal Pres As PowerPoint.Presentation, Cancel As Boolean)
    RadiusNativeRibbon.NativeBeforeSave Pres
End Sub

Private Sub Host_PresentationClose(ByVal Pres As PowerPoint.Presentation)
    RadiusNativeRelations.ForgetPresentation Pres
    RadiusNativeLayout.ForgetPresentation Pres
    RadiusNativeRibbon.InvalidateSelection
End Sub
