Sub InsertArrow()
    ActiveSheet.Shapes.AddConnector(msoConnectorStraight, 248.25, 118.5, 417, 234). _
        Select
        Selection.ShapeRange.Line.EndArrowheadStyle = msoArrowheadTriangle
    With Selection.ShapeRange.Line
        .Visible = msoTrue
        .ForeColor.RGB = RGB(0, 176, 240)
        .Transparency = 0
    End With
    With Selection.ShapeRange.Line
        .Visible = msoTrue
        .Weight = 3
    End With
    Range("L23").Select
End Sub