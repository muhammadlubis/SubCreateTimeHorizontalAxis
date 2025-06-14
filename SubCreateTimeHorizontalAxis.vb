Sub CreateTimeAxis()
    Dim slide As slide
    Dim startX As Single
    Dim startY As Single
    Dim tickSpacing As Single
    Dim numTicks As Integer
    Dim i As Integer
    Dim tickHeight As Single
    Dim tick As Shape
    Dim label As Shape
    Dim timeLabel As String

    ' Configuration
    Set slide = ActivePresentation.Slides(1) ' Adjust if you want another slide
    startX = 100 ' Starting horizontal position
    startY = 300 ' Vertical position of the axis
    tickSpacing = 50 ' Distance between ticks (you can change to represent time units)
    numTicks = 10 ' Number of time ticks
    tickHeight = 10 ' Height of each tick mark

    ' Draw base line (the axis)
    slide.Shapes.AddLine startX, startY, startX + tickSpacing * numTicks, startY

    ' Loop to create tick marks and labels
    For i = 0 To numTicks
        ' Draw tick
        Set tick = slide.Shapes.AddLine( _
            startX + i * tickSpacing, startY - tickHeight / 2, _
            startX + i * tickSpacing, startY + tickHeight / 2)

        ' Create time label (e.g., Hour 0, Hour 1, etc.)
        timeLabel = "T+" & i

        ' Add label under the tick
        Set label = slide.Shapes.AddTextbox( _
            Orientation:=msoTextOrientationHorizontal, _
            Left:=startX + i * tickSpacing - 10, _
            Top:=startY + tickHeight, _
            Width:=30, Height:=15)

        label.TextFrame.TextRange.Text = timeLabel
        label.TextFrame.TextRange.Font.Size = 10
        label.TextFrame.HorizontalAnchor = msoAnchorCenter
        label.TextFrame.VerticalAnchor = msoAnchorTop
    Next i
End Sub

