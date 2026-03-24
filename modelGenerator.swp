Dim swApp As Object
Dim Part As Object
Dim boolstatus As Boolean
Dim savePath As String

Sub main()
    Set swApp = Application.SldWorks
    ' Create a new part using the default template
    Set Part = swApp.NewDocument("C:\ProgramData\SOLIDWORKS\SOLIDWORKS 2023\templates\Part.prtdot", 0, 0, 0)
    
    If Part Is Nothing Then
        MsgBox "Failed to open part. You may need to update the default template path in the code."
        Exit Sub
    End If

    ' --- 1. BASE BOX (The Aluminum/Copper Body) ---
    boolstatus = Part.Extension.SelectByID2("Top Plane", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
    Part.SketchManager.InsertSketch True
    Part.SketchManager.CreateCenterRectangle 0, 0, 0, 0.02, 0.02, 0 ' Default 40x40 mm
    Part.FeatureManager.FeatureExtrusion3 True, False, False, 0, 0, 0.02, 0.02, False, False, False, False, 0, 0, False, False, False, False, True, True, True, 0, 0, False
    Part.SelectionManager.EnableContourSelection = False
    
    Part.Parameter("D1@Sketch1").Name = "Box_Width"
    Part.Parameter("D2@Sketch1").Name = "Box_Length"
    Part.Parameter("D1@Boss-Extrude1").Name = "Box_Thickness"
    
    ' --- 2. MOUNTING WINGS ---
    Part.ClearSelection2 True
    boolstatus = Part.Extension.SelectByID2("Top Plane", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
    Part.SketchManager.InsertSketch True
    ' Create one large 60x40mm rectangle. The 10mm overhangs on each side become the wings.
    Part.SketchManager.CreateCenterRectangle 0, 0, 0, 0.03, 0.02, 0 
    Part.FeatureManager.FeatureExtrusion3 True, False, False, 0, 0, 0.005, 0.005, False, False, False, False, 0, 0, False, False, False, False, True, True, True, 0, 0, False
    Part.SelectionManager.EnableContourSelection = False
    
    Part.Parameter("D1@Sketch2").Name = "Total_Wing_Span"
    Part.Parameter("D2@Sketch2").Name = "Wing_Length"
    Part.Parameter("D1@Boss-Extrude2").Name = "Wing_Thickness"

    ' --- 3. PCB CAVITY ---
    ' Raycast to select the top face of the box (Z = 20mm pointing down)
    Part.ClearSelection2 True
    boolstatus = Part.Extension.SelectByRay(0, 0, 0.025, 0, 0, -1, 0.001, 2, False, 0, 0)
    Part.SketchManager.InsertSketch True
    Part.SketchManager.CreateCenterRectangle 0, 0, 0, 0.015, 0.015, 0 ' Default 30x30 mm
    Part.FeatureManager.FeatureCut4 True, False, False, 1, 0, 0.005, 0.005, False, False, False, False, 0, 0, False, False, False, False, False, True, True, True, True, False, 0, 0, False, False
    
    Part.Parameter("D1@Sketch3").Name = "PCB_Cavity_Width"
    Part.Parameter("D2@Sketch3").Name = "PCB_Cavity_Length"
    Part.Parameter("D1@Cut-Extrude1").Name = "PCB_Cavity_Depth"
    
    ' --- 4. CHIP CAVITY (6.3 x 6.3 x 1.0 mm) ---
    ' Raycast to select the floor of the PCB cavity (Z = 15mm pointing down)
    Part.ClearSelection2 True
    boolstatus = Part.Extension.SelectByRay(0, 0, 0.0175, 0, 0, -1, 0.001, 2, False, 0, 0)
    Part.SketchManager.InsertSketch True
    Part.SketchManager.CreateCenterRectangle 0, 0, 0, 0.00315, 0.00315, 0 ' 6.3x6.3 mm
    Part.FeatureManager.FeatureCut4 True, False, False, 1, 0, 0.001, 0.001, False, False, False, False, 0, 0, False, False, False, False, False, True, True, True, True, False, 0, 0, False, False
    
    Part.Parameter("D1@Sketch4").Name = "Chip_Cavity_Width"
    Part.Parameter("D1@Sketch4").SystemValue = 0.0063
    Part.Parameter("D2@Sketch4").Name = "Chip_Cavity_Length"
    Part.Parameter("D2@Sketch4").SystemValue = 0.0063
    Part.Parameter("D1@Cut-Extrude2").Name = "Chip_Cavity_Depth"
    Part.Parameter("D1@Cut-Extrude2").SystemValue = 0.001
    
    Part.ClearSelection2 True
    Part.ForceRebuild3 True

End Sub
