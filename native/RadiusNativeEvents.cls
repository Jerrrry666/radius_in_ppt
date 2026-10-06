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
    RadiusNativeRibbon.InvalidateSelection
End Sub
