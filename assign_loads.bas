Attribute VB_Name = "Module1"

Sub assignLoads()

'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
'\  Create surfaces in LUSAS based on coordinates
'\  Author :    Shem Noremberg (szcz1360)
'\              Oxford; January 2020
'\  Summary :   Loops through columns to read assignment information_
'\              then uses LPI to assing a load attribute to objects
'\
'\              A LUSAS project must be open, otherwise it will crash.
'\
'\ AUTHOR TAKES NO RESPONISBILITY IF THIS SCRIPT IS USED INCORRECTLY
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Dim wb As Workbook
Dim ws As Worksheet

Dim lastCol As Long
Dim startRow As Long
Dim loadID As Long
Dim LCID As Long
Dim objIDs As String
Dim loadFactor As Double
Dim i As Integer

Debug.Print ("Starting Assingmnet")

Set wb = ThisWorkbook
Set ws = Sheet1

startCol = 3
lastCol = 24

startRow = ws.Range("macroStart").Row

For i = startCol To lastCol
    
    loadID = ws.Cells(startRow + 1, i)
    LCID = ws.Cells(startRow + 2, i)
    objIDs = ws.Cells(startRow + 3, i)
    loadFactor = Round(ws.Cells(startRow + 4, i), 3)
    
    Call assignment.setAllDefaults
    Call assignment.setLoadset(LCID)
    Call assignment.setLoadFactor(loadFactor)
    
    Call database.getAttribute("Loading", loadID).assignTo(newObjectSet.Add("Line", objIDs), assignment)

Next i

Debug.Print ("Assingment complete")

End Sub
