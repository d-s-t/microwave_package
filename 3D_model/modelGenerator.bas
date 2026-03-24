Dim swApp As Object
Dim Part As Object
Dim boolstatus As Boolean
Dim templatePath As String
Dim featBox As Object, sketchBox As Object
Dim featWings As Object, sketchWings As Object
Dim featPCB As Object, sketchPCB As Object
Dim featChip As Object, sketchChip As Object

Sub main()
    Set swApp = Application.SldWorks
    
    ' Fetch the dynamic default template path
    templatePath = swApp.GetUserPreferenceStringValue(8) 
    If templatePath = "" Then
        MsgBox "No default part template found in SolidWorks settings.", vbCritical
        Exit Sub
    End If
    
    Set Part = swApp.NewDocument(templatePath, 0, 0, 0)
    If Part Is Nothing Then
        MsgBox "Failed to open part. Check template paths.", vbCritical
        Exit Sub
    End If

    ' --- 1. BASE BOX ---
    boolstatus = Part.Extension.SelectByID2("Top Plane", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
    If Not boolstatus Then Part.FeatureByPosition(2).Select2 False, 0
    
    Part.SketchManager.InsertSketch True
    Part.SketchManager.CreateCenterRectangle 0, 0, 0, 0.02, 0.02, 0 
    
    ' Capture the Extrude object, then ask it for its parent Sketch object
    Set featBox = Part.FeatureManager.FeatureExtrusion3(True, False, False, 0, 0, 0.02, 0.02, False, False, False, False, 0, 0, False, False, False, False, True, True, True, 0, 0, False)
    Part.SelectionManager.EnableContourSelection = False
    Set sketchBox = featBox.GetFirstSubFeature()
    
    ' Rename dynamically
    Part.Parameter("D1@" & sketchBox.Name).Name = "Box_Width"
    Part.Parameter("D2@" & sketchBox.Name).Name = "Box_Length"
    Part.Parameter("D1@" & featBox.Name).Name = "Box_Thickness"
    
    ' --- 2. MOUNTING WINGS ---
    Part.ClearSelection2 True
    boolstatus = Part.Extension.SelectByID2("Top Plane", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
    If Not boolstatus Then Part.FeatureByPosition(2).Select2 False, 0

    Part.SketchManager.InsertSketch True
    Part.SketchManager.CreateCenterRectangle 0, 0, 0, 0.03, 0.02, 0 
    
    Set featWings = Part.FeatureManager.FeatureExtrusion3(True, False, False, 0, 0, 0.005, 0.005, False, False, False, False, 0, 0, False, False, False, False, True, True, True, 0, 0, False)
    Part.SelectionManager.EnableContourSelection = False
    Set sketchWings = featWings.GetFirstSubFeature()

    Part.Parameter("D1@" & sketchWings.Name).Name = "Total_Wing_Span"
    Part.Parameter("D2@" & sketchWings.Name).Name = "Wing_Length"
    Part.Parameter("D1@" & featWings.Name).Name = "Wing_Thickness"

    ' --- 3. PCB CAVITY ---
    Part.ClearSelection2 True
    boolstatus = Part.Extension.SelectByRay(0, 0, 0.025, 0, 0, -1, 0.001, 2, False, 0, 0)
    Part.SketchManager.InsertSketch True
    Part.SketchManager.CreateCenterRectangle 0, 0, 0, 0.015, 0.015, 0 
    
    Set featPCB = Part.FeatureManager.FeatureCut4(True, False, False, 1, 0, 0.005, 0.005, False, False, False, False, 0, 0, False, False, False, False, False, True, True, True, True, False, 0, 0, False, False)
    Set sketchPCB = featPCB.GetFirstSubFeature()

    Part.Parameter("D1@" & sketchPCB.Name).Name = "PCB_Cavity_Width"
    Part.Parameter("D2@" & sketchPCB.Name).Name = "PCB_Cavity_Length"
    Part.Parameter("D1@" & featPCB.Name).Name = "PCB_Cavity_Depth"
    
    ' --- 4. CHIP CAVITY ---
    Part.ClearSelection2 True
    boolstatus = Part.Extension.SelectByRay(0, 0, 0.0175, 0, 0, -1, 0.001, 2, False, 0, 0)
    Part.SketchManager.InsertSketch True
    Part.SketchManager.CreateCenterRectangle 0, 0, 0, 0.00315, 0.00315, 0 
    
    Set featChip = Part.FeatureManager.FeatureCut4(True, False, False, 1, 0, 0.001, 0.001, False, False, False, False, 0, 0, False, False, False, False, False, True, True, True, True, False, 0, 0, False, False)
    Set sketchChip = featChip.GetFirstSubFeature()

    Part.Parameter("D1@" & sketchChip.Name).Name = "Chip_Cavity_Width"
    Part.Parameter("D1@" & sketchChip.Name).SystemValue = 0.0063
    Part.Parameter("D2@" & sketchChip.Name).Name = "Chip_Cavity_Length"
    Part.Parameter("D2@" & sketchChip.Name).SystemValue = 0.0063
    Part.Parameter("D1@" & featChip.Name).Name = "Chip_Cavity_Depth"
    Part.Parameter("D1@" & featChip.Name).SystemValue = 0.001
    
    Part.ClearSelection2 True
    Part.ForceRebuild3 True

End Sub
