Dim swApp As Object
Dim swModel As Object
Dim fileNum As Integer
Dim filePath As String
Dim lineData As String
Dim rowData() As String
Dim headers() As String
Dim i As Integer
Dim configName As String
Dim dimName As String
Dim dimValue As Double
Dim swDim As Object
Dim boolstatus As Boolean

Sub main()
    Set swApp = Application.SldWorks
    Set swModel = swApp.ActiveDoc
    
    ' Check if a part is open
    If swModel Is Nothing Then
        MsgBox "Please open the generated microwave package part first.", vbExclamation
        Exit Sub
    End If
    
    ' Prompt for the CSV file location
    filePath = InputBox("Enter the full path to your configurations.csv file:" & vbCrLf & vbCrLf & "(e.g., C:\Github\microwave_package\configurations.csv)", "Load CSV Data")
    
    If filePath = "" Or Dir(filePath) = "" Then
        MsgBox "Operation cancelled or file not found.", vbExclamation
        Exit Sub
    End If
    
    fileNum = FreeFile
    Open filePath For Input As #fileNum
    
    ' 1. Read the Header Row
    If Not EOF(fileNum) Then
        Line Input #fileNum, lineData
        headers = Split(lineData, ",")
    End If
    
    ' 2. Loop through the Data Rows
    Do While Not EOF(fileNum)
        Line Input #fileNum, lineData
        If Trim(lineData) <> "" Then
            rowData = Split(lineData, ",")
            configName = Trim(rowData(0))
            
            ' Check if configuration exists; if not, create it
            boolstatus = swModel.ShowConfiguration2(configName)
            If boolstatus = False Then
                swModel.AddConfiguration3 configName, "", "", 256 ' swConfigOption_DontActivate
                swModel.ShowConfiguration2 configName
            End If
            
            ' Apply dimensions to the active configuration
            For i = 1 To UBound(headers)
                dimName = Trim(headers(i))
                dimValue = Val(rowData(i)) / 1000 ' Convert mm from CSV to meters for SolidWorks
                
                Set swDim = swModel.Parameter(dimName)
                If Not swDim Is Nothing Then
                    ' Set the value specifically for the current configuration
                    boolstatus = swDim.SetSystemValue3(dimValue, 1, Nothing) ' 1 = swSetValue_InThisConfiguration
                End If
            Next i
        End If
    Loop
    
    Close #fileNum
    
    ' Rebuild the model to show changes
    swModel.ForceRebuild3 True
    MsgBox "Configurations successfully generated and updated from CSV!", vbInformation
End Sub