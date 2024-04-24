Attribute VB_Name = "Module6"

Sub change_wllevel()
Attribute change_wllevel.VB_ProcData.VB_Invoke_Func = " \n14"
Dim create_string As String

Call take_DaLe
create_string = "ÉtÉåÉìÉhÉäÅ[É}ÉbÉ`Å@" + s_level + "Å@èüîsï\"

    Selection.ShapeRange(1).TextFrame2.TextRange.Characters.Text = "create_string"
    With Selection.ShapeRange(1).TextFrame2.TextRange.Characters(1, 20). _
        ParagraphFormat
        .TextDirection = msoTextDirectionLeftToRight
        .FirstLineIndent = 0
        .Alignment = msoAlignCenter
    End With
    With Selection.ShapeRange(1).TextFrame2.TextRange.Characters(1, 10).Font
        .NameComplexScript = "+mn-cs"
        .NameFarEast = "HGSënâpäpŒﬂØÃﬂëÃ"
        .Fill.Visible = msoTrue
        .Fill.ForeColor.RGB = RGB(255, 255, 0)
        .Fill.Transparency = 0
        .Fill.Solid
        .Size = 36
        .Kerning = 0.1000000015
        .Line.Visible = msoTrue
        .Line.ForeColor.RGB = RGB(0, 255, 0)
        .Line.Transparency = 0
        .Line.Weight = 0.75
        .Line.DashStyle = msoLineSolid
        .Line.Style = msoLineSingle
        .Name = "HGSënâpäpŒﬂØÃﬂëÃ"
        .Spacing = 0
    End With
    With Selection.ShapeRange(1).TextFrame2.TextRange.Characters(11, 4).Font
        .NameComplexScript = "+mn-cs"
        .NameFarEast = "HGSënâpäpŒﬂØÃﬂëÃ"
        .Fill.Visible = msoTrue
        .Fill.ForeColor.RGB = RGB(255, 255, 0)
        .Fill.Transparency = 0
        .Fill.Solid
        .Size = 36
        .Kerning = 0.1000000015
        .Line.Visible = msoTrue
        .Line.ForeColor.RGB = RGB(0, 255, 0)
        .Line.Transparency = 0
        .Line.Weight = 0.75
        .Line.DashStyle = msoLineSolid
        .Line.Style = msoLineSingle
        .Name = "HGSënâpäpŒﬂØÃﬂëÃ"
        .Spacing = 0
    End With
    With Selection.ShapeRange(1).TextFrame2.TextRange.Characters(15, 6).Font
        .NameComplexScript = "+mn-cs"
        .NameFarEast = "HGSënâpäpŒﬂØÃﬂëÃ"
        .Fill.Visible = msoTrue
        .Fill.ForeColor.RGB = RGB(255, 255, 0)
        .Fill.Transparency = 0
        .Fill.Solid
        .Size = 36
        .Kerning = 0.1000000015
        .Line.Visible = msoTrue
        .Line.ForeColor.RGB = RGB(0, 255, 0)
        .Line.Transparency = 0
        .Line.Weight = 0.75
        .Line.DashStyle = msoLineSolid
        .Line.Style = msoLineSingle
        .Name = "HGSënâpäpŒﬂØÃﬂëÃ"
        .Spacing = 0
    End With

End Sub
