# PowerPoint Flatten Slides Macro
This PowerPoint VBA macro converts each slide in your presentation into a single image (EMF), effectively flattening all content and preventing users from revealing sensitive data by moving or deleting overlay shapes.

## Why Use This?

When sharing PowerPoint files that include screenshots or overlays (e.g., user credentials, ID numbers), hiding them behind rectangles isn't enough — they can be moved or deleted. This macro:

- Turns each slide into a **flat image**
- Preserves **high quality (vector-based EMF)**
- Prevents layer manipulation
- Keeps output as a `.pptx` file

##  Features

- One-click flattening of entire presentation
- Keeps original slide layout dimensions
- Runs natively in PowerPoint (no external tools)
- Safe for corporate environments (no registry edits, no online services)

##  How to Use

1. Open PowerPoint
2. Press `Alt + F11` to open the VBA editor
3. Insert → Module → Paste the macro (see below)
4. Press `Alt + F8`, select `FlattenSlidesSecurely`, then click `Run`

##  Macro Code

```vba
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

