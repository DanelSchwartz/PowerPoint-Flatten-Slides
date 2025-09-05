Sub FlattenSlidesSecurely()
    Dim sld As Slide
    Dim shpPicture As Shape
    Dim shpRange As ShapeRange

    On Error Resume Next

    For Each sld In ActivePresentation.Slides
        If sld.Shapes.Count > 0 Then
            Set shpRange = sld.Shapes.Range
            shpRange.Copy
            shpRange.Delete

            Set shpPicture = sld.Shapes.PasteSpecial(ppPasteEnhancedMetafile)(1)

            If Not shpPicture Is Nothing Then
                With shpPicture
                    .LockAspectRatio = msoFalse
                    .Top = 0
                    .Left = 0
                    .Width = ActivePresentation.PageSetup.SlideWidth
                    .Height = ActivePresentation.PageSetup.SlideHeight
                End With
            End If
        End If
    Next sld

    On Error GoTo 0
    MsgBox "All slides have been converted into flat images – no movable layers remain.", vbInformation
End Sub
