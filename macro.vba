Sub BarreDeProgression()
    'Génère une barre de progression + numéro de diapo (Montserrat) en haut à droite de la barre

    '=== Paramètres ===
    Const Longueur As Single = 1       'longueur totale de la barre (% de la largeur diapo)
    Const Hauteur As Single = 0.02     'hauteur totale de la barre (% de la hauteur diapo)
    Const PositionX As Single = 0      'position X (% de la largeur diapo depuis la gauche)
    Const PositionY As Single = 0.021  'position Y (% de la hauteur diapo depuis le haut)

    Const Dernier_diapo As Integer = 35 'la barre s’arrête à cette diapo incluse

    '--- Apparence ---
    Const BarColorActiveR As Integer = 255
    Const BarColorActiveG As Integer = 171
    Const BarColorActiveB As Integer = 64

    Const BarColorBgR As Integer = 0
    Const BarColorBgG As Integer = 22
    Const BarColorBgB As Integer = 51

    Const LabelHeightPt As Single = 8  'hauteur de la zone de texte (points)
    Const LabelWidthPt As Single = 50   'largeur de la zone de texte (points)
    Const LabelOffsetPt As Single = 6   'décalage vertical au-dessus de la barre (points)
    Const LabelRightPaddingPt As Single = 5 'petit retrait par rapport au bord droit

    Dim Pres As Presentation
    Dim SLD As Slide
    Dim OBJ As Shape
    Dim H As Single, W As Single
    Dim nb As Integer, counter As Integer
    Dim nbShape As Long, del As Long, a As Long, i As Integer

    Set Pres = ActivePresentation
    H = Pres.PageSetup.SlideHeight
    W = Pres.PageSetup.SlideWidth * Longueur
    nb = Dernier_diapo - 1
    counter = 1

    'Pour chaque diapo
    For Each SLD In Pres.Slides
        If counter <> 1 And counter < (Dernier_diapo + 1) Then
            'Supprime les anciennes barres/labels
            nbShape = SLD.Shapes.Count
            del = 0
            For a = 1 To nbShape
                With SLD.Shapes.Item(a - del)
                    If Left(.Name, 2) = "PB" Or Left(.Name, 6) = "PB_TXT" Then
                        .Delete
                        del = del + 1
                    End If
                End With
            Next a

            'Ajoute la barre de progression (segments)
            For i = 0 To nb - 1
                Set OBJ = SLD.Shapes.AddShape( _
                    Type:=msoShapeRectangle, _
                    Left:=(W * i / nb) + W / nb * (PositionX / 2), _
                    Top:=H * (1 - PositionY) + 2, _
                    Width:=(W / nb) * (1 - PositionX), _
                    Height:=5)
                OBJ.Name = "PB" & i
                OBJ.Line.Visible = msoFalse
                If (i + 2) = counter Then
                    OBJ.Fill.ForeColor.RGB = RGB(BarColorActiveR, BarColorActiveG, BarColorActiveB)
                Else
                    OBJ.Fill.ForeColor.RGB = RGB(BarColorBgR, BarColorBgG, BarColorBgB)
                End If
            Next i

            '=== Ajoute le label du numéro de diapo ===
            Dim containerLeft As Single, containerRight As Single, containerTop As Single
            containerLeft = (W / nb) * (PositionX / 2)
            containerRight = containerLeft + W
            containerTop = H * (1 - PositionY) + 2

            Dim lbl As Shape
            Set lbl = SLD.Shapes.AddTextbox( _
                        Orientation:=msoTextOrientationHorizontal, _
                        Left:=containerRight - LabelRightPaddingPt - LabelWidthPt, _
                        Top:=containerTop - LabelOffsetPt - LabelHeightPt, _
                        Width:=LabelWidthPt, _
                        Height:=LabelHeightPt)
            lbl.Name = "PB_TXT" & SLD.SlideIndex

            With lbl.TextFrame2
                .TextRange.Text = CStr(SLD.SlideIndex) 'numéro de diapo
                .TextRange.Font.Name = "Montserrat"
                .TextRange.Font.Size = 9
                .TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
                .TextRange.ParagraphFormat.Alignment = msoAlignRight
                .AutoSize = msoAutoSizeNone
                .MarginLeft = 0
                .MarginRight = 0
                .MarginTop = 0
                .MarginBottom = 0
            End With

            'verrouille la sélection visuelle
            lbl.Line.Visible = msoFalse
            lbl.Fill.Visible = msoFalse
        End If

        counter = counter + 1
    Next SLD
End Sub


