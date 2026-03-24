Dim swApp As Object
Dim Part As Object
Dim boolstatus As Boolean
Dim templatePath As String

Dim myDim As Object
Dim feat As Object

Sub main()
    Set swApp = Application.SldWorks
    
    ' 1. Open Document using dynamic template
    templatePath = swApp.GetUserPreferenceStringValue(8)
    Set Part = swApp.NewDocument(templatePath, 0, 0, 0)
    If Part Is Nothing Then Exit Sub

    ' --- 1. BASE BOX ---
    ' Select Top Plane (Fallback to Index 2 if localized language)
    boolstatus = Part.Extension.SelectByID2("Top Plane", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
    If Not boolstatus Then Part.FeatureByPosition(2).Select2 False, 0
    
    Part.SketchManager.InsertSketch True
    Part.SketchManager.CreateCenterRectangle 0, 0, 0, 0.02, 0.02, 0
    
    ' Explicitly select the top line and Add Dimension Object (Width)
    Part.ClearSelection2 True
    Part.Extension.SelectByID2 "", "EXTSKETCHSEGMENT", 0, 0.02, 0, False, 0, Nothing, 0
    Set myDim = Part.AddDimension2(0, 0.025, 0)
    If Not myDim Is Nothing Then myDim.GetDimension2(0).Name = "Box_Width"
    
    ' Explicitly select the right line and Add Dimension Object (Length)
    Part.ClearSelection2 True
    Part.Extension.SelectByID2 "", "EXTSKETCHSEGMENT", 0.02, 0, 0, False, 0, Nothing, 0
    Set myDim = Part.AddDimension2(0.025, 0, 0)
    If Not myDim Is Nothing Then myDim.GetDimension2(0).Name = "Box_Length"
    
    ' Extrude Box and grab the depth dimension directly from the feature object
    Set feat = Part.FeatureManager.FeatureExtrusion3(True, False, False, 0, 0, 0.02, 0.02, False, False, False, False, 0, 0, False, False, False, False, True, True, True, 0, 0, False)
    Set myDim = feat.GetFirstDisplayDimension()
    If Not myDim Is Nothing Then myDim.GetDimension2(0).Name = "Box_Thickness"
    
    ' --- 2. MOUNTING WINGS ---
    Part.ClearSelection2 True
    boolstatus = Part.Extension.SelectByID2("Top Plane", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
    If Not boolstatus Then Part.FeatureByPosition(2).Select2 False, 0

    Part.SketchManager.InsertSketch True
    Part.SketchManager.CreateCenterRectangle 0, 0, 0, 0.03, 0.02, 0
    
    Part.ClearSelection2 True
    Part.Extension.SelectByID2 "", "EXTSKETCHSEGMENT", 0, 0.02, 0, False, 0, Nothing, 0
    Set myDim = Part.AddDimension2(0, 0.03, 0)
    If Not myDim Is Nothing Then myDim.GetDimension2(0).Name = "Total_Wing_Span"
    
    Part.ClearSelection2 True
    Part.Extension.SelectByID2 "", "EXTSKETCHSEGMENT", 0.03, 0, 0, False, 0, Nothing, 0
    Set myDim = Part.AddDimension2(0.04, 0, 0)
    If Not myDim Is Nothing Then myDim.GetDimension2(0).Name = "Wing_Length"
    
    Set feat = Part.FeatureManager.FeatureExtrusion3(True, False, False, 0, 0, 0.005, 0.005, False, False, False, False, 0, 0, False, False, False, False, True, True, True, 0, 0, False)
    Set myDim = feat.GetFirstDisplayDimension()
    If Not myDim Is Nothing Then myDim.GetDimension2(0).Name = "Wing_Thickness"

    ' --- 3. PCB CAVITY ---
    Part.ClearSelection2 True
    Part.Extension.SelectByRay 0, 0, 0.025, 0, 0, -1, 0.001, 2, False, 0, 0
    Part.SketchManager.InsertSketch True
    Part.SketchManager.CreateCenterRectangle 0, 0, 0, 0.015, 0.015, 0
    
    Part.ClearSelection2 True
    Part.Extension.SelectByID2 "", "EXTSKETCHSEGMENT", 0, 0.015, 0, False, 0, Nothing, 0
    Set myDim = Part.AddDimension2(0, 0.02, 0)
    If Not myDim Is Nothing Then myDim.GetDimension2(0).Name = "PCB_Cavity_Width"
    
    Part.ClearSelection2 True
    Part.Extension.SelectByID2 "", "EXTSKETCHSEGMENT", 0.015, 0, 0, False, 0, Nothing, 0
    Set myDim = Part.AddDimension2(0.02, 0, 0)
    If Not myDim Is Nothing Then myDim.GetDimension2(0).Name = "PCB_Cavity_Length"
    
    Set feat = Part.FeatureManager.FeatureCut4(True, False, False, 1, 0, 0.005, 0.005, False, False, False, False, 0, 0, False, False, False, False, False, True, True, True, True, False, 0, 0, False, False)
    Set myDim = feat.GetFirstDisplayDimension()
    If Not myDim Is Nothing Then myDim.GetDimension2(0).Name = "PCB_Cavity_Depth"
    
    ' --- 4. CHIP CAVITY ---
    Part.ClearSelection2 True
    Part.Extension.SelectByRay 0, 0, 0.0175, 0, 0, -1, 0.001, 2, False, 0, 0
    Part.SketchManager.InsertSketch True
    Part.SketchManager.CreateCenterRectangle 0, 0, 0, 0.00315, 0.00315, 0
    
    Part.ClearSelection2 True
    Part.Extension.SelectByID2 "", "EXTSKETCHSEGMENT", 0, 0.00315, 0, False, 0, Nothing, 0
    Set myDim = Part.AddDimension2(0, 0.005, 0)
    If Not myDim Is Nothing Then myDim.GetDimension2(0).Name = "Chip_Cavity_Width"
    
    Part.ClearSelection2 True
    Part.Extension.SelectByID2 "", "EXTSKETCHSEGMENT", 0.00315, 0, 0, False, 0, Nothing, 0
    Set myDim = Part.AddDimension2(0.005, 0, 0)
    If Not myDim Is Nothing Then myDim.GetDimension2(0).Name = "Chip_Cavity_Length"
    
    Set feat = Part.FeatureManager.FeatureCut4(True, False, False, 1, 0, 0.001, 0.001, False, False, False, False, 0, 0, False, False, False, False, False, True, True, True, True, False, 0, 0, False, False)
    Set myDim = feat.GetFirstDisplayDimension()
    If Not myDim Is Nothing Then myDim.GetDimension2(0).Name = "Chip_Cavity_Depth"
    
    Part.ClearSelection2 True
    Part.ForceRebuild3 True

End Sub