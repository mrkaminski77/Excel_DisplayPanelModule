Attribute VB_Name = "DisplayPanelModule"
Option Explicit


Public Sub DisplayPanel( _
    Target As Worksheet, _
    list As Object, _
    padding As Integer, _
    labelWidth As Integer, _
    ResultWidth As Integer, _
    ResultHeight As Integer, _
    frameTop As Integer, _
    frameLeft As Integer, _
    color As Long, _
    gradient As Double _
    )
    
    Dim frameWidth As Integer, frameHeight As Integer
    frameWidth = labelWidth + ResultWidth + (padding * 2)
    frameHeight = (2 * padding) + (ResultHeight * list.Count) + (padding * list.Count)

    On Error Resume Next
    Target.Shapes("ParamsFrame").Delete
    On Error GoTo 0

    Target.Shapes.AddShape(msoShapeRectangularCallout, frameLeft, frameTop, frameWidth, frameHeight).Name = "StatusFrame"
    With Target.Shapes("StatusFrame")
        .Line.Visible = msoFalse
        .Fill.ForeColor.RGB = color
        .Fill.OneColorGradient msoGradientVertical, 1, gradient
        .Placement = xlFreeFloating
    End With
    Dim p As Variant, i As Integer
    i = 0
    For Each p In list.Keys

        Target.Shapes.AddLabel( _
            msoTextOrientationHorizontal, _
            frameLeft + padding, _
            frameTop + padding + (i * padding) + (ResultHeight * i), _
            labelWidth, _
            ResultHeight).Name = p & "Label"
        With Target.Shapes(p & "Label")
            .TextFrame2.TextRange.Characters = p
            .TextFrame2.TextRange.Font.Fill.ForeColor.RGB = rgbWhite
            .Placement = xlFreeFloating
        End With

        Target.Shapes.AddLabel( _
            msoTextOrientationHorizontal, _
            frameLeft + 20 + labelWidth, _
            frameTop + padding + (i * padding) + (ResultHeight * i), _
            ResultWidth, _
            ResultHeight).Name = p & "Result"
        With Target.Shapes(p & "Result")
            .TextFrame2.TextRange.Characters = list.Item(p)
            .TextFrame2.TextRange.Font.Fill.ForeColor.RGB = rgbWhite
            .Placement = xlFreeFloating
        End With
        i = i + 1
    Next
    Target.Shapes.Range(GroupArray(list)).Group().Name = "ParamsFrame"
    Target.Shapes("ParamsFrame").Placement = xlFreeFloating
End Sub



Private Function GroupArray(list As Object) As String()
    Dim result() As String
    ReDim result(list.Count * 2 + 1)
    Dim i As Integer, p As Variant
    i = 0
    For Each p In list.Keys
        result(i) = p & "Label"
        result(i + 1) = p & "Result"
    i = i + 2
    Next
    result(i) = "StatusFrame"
    GroupArray = result
End Function


