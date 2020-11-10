Attribute VB_Name = "Module1"
Option Explicit
Sub getresults()


Set Group = database.getGroupByName("P_SUPPORTS")
objArray1 = Group.getObjects("Points")
For j = 0 To UBound(objArray1)
    Set obj1 = objArray1(j)
   
    nodesArray = obj1.getNodes()
    For i = 0 To UBound(nodesArray)
    Set f = nodesArray(i)
    
    
    getTextWindow().writeLine ("Point " & obj1.getID() & " (Node " & f.getID() & ")" & " reaction FZ=" & f.getresults("Reaction", "FZ"))
    Next
Next

End Sub


Sub getReactions()


Dim oGroup As IFGroup
Dim objArr() As Variant
Dim nodeArr() As Variant
Dim point As IFPoint
Dim node As IFNode

Dim wb As Workbook
Dim ws As Worksheet
Dim rPoint As Range
Dim rFZ As Range
Dim lastrow As Long

Dim i As Long
Dim j As Long

Set wb = ThisWorkbook
Set ws = Sheet2
Set rPoint = ws.Range("B3")
Set rFZ = ws.Range("C3")
lastrow = ws.Cells(ws.Rows.Count, "B").End(xlUp).Row

'For i = rPoint.row + 1 To lastrow

Set oGroup = database.getGroupByName("P_SUPPORTS")
objArr = oGroup.getObjects("Points")
'
For i = 0 To UBound(objArr)
    Set point = objArr(i) 'database.getPointByName(ws.Cells(i, rPoint.Column).Value)
    nodeArr = point.getNodes()
    
    For j = 0 To UBound(nodeArr)
        Set node = nodeArr(j)
        ws.Cells(i + 4, rFZ.Column).Value = node.getresults("Reaction", "FZ")
    Next j
Next i

End Sub

Sub alignPoints2Line()
'aligns points to a line selected by user, ignoring the Z coordinate




Dim selec As IFSelection
Dim pointC As IFPoint 'point C to be moved
Dim line As IFLine

Dim Acoord As Variant
Dim Bcoord As Variant
Dim Ccoord As Variant
Dim Dcoord As Variant
Dim Vx As Variant
Dim VxUnit As Variant
Dim Vy As Variant
Dim VyProject As Double
Dim Vtrans As Variant

Dim transform As IFTransformationAttr
Dim returnedSetA As Variant
Dim i As Long

Set selec = LusasWinApp.getSelection

'check selection for min 1 point and line
If selec.countPoints < 1 And selec.countLines < 1 Then
    MsgBox ("At least one line and one point must be selected")
    Exit Sub
End If

Set line = selec.getLine(0)

Acoord = line.getStartPosition
Bcoord = line.getEndPosition
Vx = subtractVs(Bcoord, Acoord)
VxUnit = unitV(Bcoord, Acoord)

For i = 0 To selec.countPoints - 1
    'get the offset point
    Set pointC = selec.getPoint(i)
    Ccoord = pointC.getPosition
    
    'get vector to offset point
    Vy = subtractVs(Ccoord, Acoord)
    
    'get projection of Vy on Unit vector of Vx
    VyProject = dotProd(Vy, VxUnit)
    
    'get point on line perpendicular to point C
    Dcoord = addVs(Acoord, multiplyVsByScalar(VxUnit, VyProject))
    
    'get the translation Vector
    Vtrans = subtractVs(Dcoord, Ccoord, "Z")
    
    'create a trasnformation attribute
    Set transform = database.createTranslationTransAttr("Trn3", Vtrans)
    
    'Set the transformation attribute to geometry data
    Call geometryData.setAllDefaults
    Call geometryData.setTransformation(transform)
    
    'move PointC
    Set returnedSetA = pointC.Move(geometryData)

Next i
    'delete the transformation attrribute
    Call database.deleteAttribute(transform)
End Sub

Sub alignPoints2Point()
'aligns point to a single point mantaining the Z coordinate



Dim selec As IFSelection
Dim pointC As IFPoint 'point C to be moved
Dim line As IFLine

Dim pointCoord As Variant
Dim BPCoord As Variant
Dim Ccoord As Variant
Dim Dcoord As Variant
Dim Vx As Variant
Dim VxUnit As Variant
Dim Vy As Variant
Dim VyProject As Double
Dim Vtrans As Variant

Dim basePoint As IFPoint
Dim basePointID As Long
Dim points As Variant

Dim transform As IFTransformationAttr
Dim returnedSetA As Variant
Dim i As Long

Set selec = LusasWinApp.getSelection

'check some points are selected
If selec.countPoints < 1 Then
    MsgBox ("At least one point must be selected")
    Exit Sub
End If

'get base point ID from user
basePointID = InputBox("Provide basepoint ID")
Set basePoint = database.getPointByNumber(basePointID)

'select base point
pointCoord = selec.getPoint(0).getPosition
BPCoord = basePoint.getPosition
Vtrans = subtractVs(BPCoord, pointCoord, "Z")

'create a trasnformation attribute
Set transform = database.createTranslationTransAttr("Trn3", Vtrans)
Call geometryData.setAllDefaults
Call geometryData.setTransformation(transform)

'move Points
selec.Move geometryData

'delete the transformation attrribute
Call database.deleteAttribute(transform)

End Sub

Sub correctPoint()
'rounds point position to 6 decimal places


Dim selec As IFSelection
Dim pointC As IFPoint 'point C to be moved
Dim orgPosition As Variant
Dim correction As Variant
Dim transform As IFTransformationAttr

Set selec = LusasWinApp.getSelection
Set pointC = selec.getPoint(0)
orgPosition = pointC.getPosition
correction = roundCoordinates(orgPosition, 6)

'create a trasnformation attribute
Set transform = database.createTranslationTransAttr("Trn3", correction)

'Set the transformation attribute to geometry data
Call geometryData.setAllDefaults
Call geometryData.setTransformation(transform)

'move PointC
pointC.Move geometryData

'delete the transformation attrribute
Call database.deleteAttribute(transform)

End Sub

Function roundCoordinates(origin As Variant, precision As Integer, Optional Ignore As String) As Variant

Dim i As Long
Dim newV(2) As Double

'subtract each vector component by scalar
For i = 0 To 2
    newV(i) = Round(origin(i), precision) - origin(i)
Next i

If Ignore <> "" Then
    Select Case Ignore
    Case "X"
        newV(0) = 0
    Case "Y"
        newV(1) = 0
    Case "Z"
        newV(2) = 0
    End Select
End If

roundCoordinates = newV

End Function

Function subtractVs(V1 As Variant, V2 As Variant, Optional Ignore As String) As Variant

Dim i As Long
Dim newV(2) As Double

'subtract each vector component by scalar
For i = 0 To 2
    newV(i) = V1(i) - V2(i)
Next i

If Ignore <> "" Then
    Select Case Ignore
    Case "X"
        newV(0) = 0
    Case "Y"
        newV(1) = 0
    Case "Z"
        newV(2) = 0
    End Select
End If

subtractVs = newV

End Function

Function addVs(V1 As Variant, V2 As Variant) As Variant

Dim i As Long
Dim newV(2) As Double

'add each vector component by scalar
For i = 0 To 2
    newV(i) = V1(i) + V2(i)
Next i

addVs = newV

End Function

Function divideVsByScalar(V As Variant, S As Double) As Variant

Dim i As Long
Dim newV(2) As Double

'divide each vector component by scalar
For i = 0 To 2
    newV(i) = V(i) / S
Next i

divideVsByScalar = newV

End Function

Function multiplyVsByScalar(V As Variant, S As Double) As Variant

Dim i As Long
Dim newV(2) As Double

'divide each vector component by scalar
For i = 0 To 2
    newV(i) = V(i) * S
Next i

multiplyVsByScalar = newV

End Function

Function VLen(P1 As Variant, P2 As Variant) As Double

'Vector = (X2-X1), (Y2-Y1), (Z2-Z1)
'Length = (X^2, Y^2, Z^2)^0.5

'get the Vector length from points
VLen = (((P2(0) - P1(0)) ^ 2) + ((P2(1) - P1(1)) ^ 2) + ((P2(2) - P1(2)) ^ 2)) ^ (0.5)

End Function

Function unitV(P1 As Variant, P2 As Variant) As Variant

'Unit Vector = VectorX/ Length of VectorX

Dim i As Long
Dim V(2) As Double
Dim VLength As Double

'get the vector from points
For i = 0 To 2
    V(i) = P2(i) - P1(i)
Next i

'calcualte length of Vector
VLength = VLen(P1, P2)

'calculate unit Vector
unitV = divideVsByScalar(V, VLength)

End Function

Function dotProd(V1 As Variant, V2 As Variant) As Variant

'Vector = (X2-X1), (Y2-Y1), (Z2-Z1)
'Dot Product = X1*X2 + Y1*Y2 + Z1*Z2

dotProd = V1(0) * V2(0) + V1(1) * V2(1) + V1(2) * V2(2)

End Function


