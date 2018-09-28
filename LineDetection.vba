Private Sub SL_Test_LocateAlternatives(ByVal pXDoc As CASCADELib.CscXDocument, ByVal pLocator As CASCADELib.CscXDocField)

   Dim i As Long
   Dim h As Long
   Dim img As New CscImage
   Dim info As String
   Dim alt As CscXDocFieldAlternative

   ' load the first page of the xdoc
   img.Load(pXDoc.CDoc.SourceFiles(0).FileName)

   Debug.Clear

   Dim linesDetection As New CscLinesDetection
   ' detect all lines on the image
   linesDetection.DetectLines(img, 0, 0, img.Width, img.Height)
   For i = 0 To linesDetection.HorLineCount - 1
      With linesDetection.GetHorLine(i)

         info = String_Format("line #{0} @ {1}:{2} to {3}:{4} - type: {5}", _
            .StartX, _
            .StartY, _
            .EndX, _
            .EndY, _
            .LineType)
         Debug.Print info

         ' visualize them in the locator
         Set alt = pLocator.Alternatives.Create
         h = 10 ' amount of pixels to append to the bounding box
         alt.PageIndex = 0
         alt.Left = .StartX - h
         alt.Top = .StartY - h
         alt.Width = Abs(.StartX - .EndX) + 2*h
         alt.Height = Abs(.StartY - .EndY) + 2*h

      End With

   Next

End Sub



Public Function String_Format(mask As String, ParamArray tokens()) As String
   ' formats a string similar to the .net method
    Dim i As Long
    For i = 0 To UBound(tokens)
        mask = Replace$(mask, "{" & i & "}", tokens(i))
    Next
    Return mask

End Function
