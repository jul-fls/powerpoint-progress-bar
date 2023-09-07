Sub BarreDeProgression()
'Génère une barre de progression

'Valeurs à adapter selon besoin
Const Longueur As Single = 1    'Total bar length (% of slide length) (0.25 =25%))
Const Hauteur As Single = 0.02     'Total bar height (% of slide height)
Const PositionX As Single = 0       'X position of bar (% of slide length from left)
Const PositionY As Single = 0.021   'Y position of bar (% of slide height from left)

Const Dernier_diapo As Integer = 65 'Stop bar at slide number

'Information retrieval
Set Pres = ActivePresentation
H = Pres.PageSetup.SlideHeight
W = Pres.PageSetup.SlideWidth * Longueur
nb = Dernier_diapo - 1
counter = 1

'For each Slide
For Each SLD In Pres.Slides
    If counter <> 1 And counter < (Dernier_diapo + 1) Then
        'Removes old progress bar
        nbShape = SLD.Shapes.Count
        del = 0
        For a = 1 To nbShape
            If Left(SLD.Shapes.Item(a - del).Name, 2) = "PB" Then
                SLD.Shapes.Item(a - del).Delete
                del = del + 1
            End If
        Next
        
        'install the new progress bar
        For i = 0 To nb - 1
            Set OBJ = SLD.Shapes.AddShape(msoShapeRectangle, (W * i / nb) + W / nb * (PositionX / 2), H * (1 - PositionY) + 2, (W / nb) * (1 - PositionX), 5)
            OBJ.Name = "PB" & i
            OBJ.Line.Visible = msoFalse
            If (i + 2) = counter Then
                OBJ.Fill.ForeColor.RGB = RGB(164, 88, 255) 'Color of the slider rectangle
            Else
                OBJ.Fill.ForeColor.RGB = RGB(53, 68, 255) 'Background color in rgb format
            End If
        Next
    End If
    
    counter = counter + 1
Next
    
End Sub


