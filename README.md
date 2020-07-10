# AutoCAD-Macro
 
This is my first VBA macro for AutoCAD.

## List of features:
I have made a VBA module form which has those features

+ Create a simple line (segment) AB
+ Create a simple point C
+ Create a perpendicular projection made by that isolated point C with line AB
+ After above features has been done successfully, the form will appear some extra features
+ Write to a log text file to Desktop

## Demo image:
<a href="https://ibb.co/KN4Mkr3"><img src="https://i.ibb.co/cbfGHr7/Untitled.png" alt="Untitled" border="0"></a>

## Core functions:

Create a line:
``` vb
Public Sub CreateLine(x1, y1, z1, x2, y2, z2)
    Dim moSpace As AcadModelSpace
    Set moSpace = ThisDrawing.ModelSpace
    Dim startPoint(0 To 2) As Double, endPoint(0 To 2) As Double
    Dim LineObj As AcadLine
    startPoint(0) = x1: startPoint(1) = y1: startPoint(2) = z1
    endPoint(0) = x2: endPoint(1) = y2: endPoint(2) = x2
    Set LineObj = moSpace.AddLine(startPoint, endPoint)
End Sub
```

Create a point:
``` vb

Public Sub CreatePoints(x, y, z)
    Dim moSpace As AcadModelSpace
    Set moSpace = ThisDrawing.ModelSpace
    Dim Point(0 To 2) As Double
    Dim PointObj As AcadPoint
    Point(0) = x: Point(1) = y: Point(2) = z
    Set PointObj = moSpace.AddPoint(Point)
End Sub
```

Create a circle:
``` vb
Public Sub CreateCircle(x1, y1, z1, x2, y2, z2)
    Dim circleObj As AcadCircle
    Dim centerPoint(0 To 2) As Double
    Dim radius As Double
    
    centerPoint(0) = x1: centerPoint(1) = y1: centerPoint(2) = z1
    radius = Distance2Points(x1, y1, x2, y2)
    
    Set circleObj = ThisDrawing.ModelSpace.AddCircle(centerPoint, radius)
    ZoomAll
End Sub
```