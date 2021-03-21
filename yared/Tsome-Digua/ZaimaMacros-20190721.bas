Attribute VB_Name = "ZaimaMacros"
Sub ShowProperties(control As IRibbonControl)
    Dim myShape As Shape
    Dim Message As String
    Dim Items As Long
    Items = Selection.ShapeRange.count
    If Items = 0 Then
        MsgBox "No objects selected."
        Exit Sub
    End If
    For i = 1 To Items
        Set myShape = Selection.ShapeRange(i)
        Message = Message & "Name: " & myShape.Name & vbCrLf _
            & "Top: " & myShape.top & vbCrLf _
            & "Bottom: " & myShape.top + myShape.height & vbCrLf _
            & "Left: " & myShape.left & vbCrLf _
            & "Right: " & myShape.left + myShape.width & vbCrLf _
            & "Height: " & myShape.height & vbCrLf _
            & "Width: " & myShape.width & vbCrLf _
            & "Radius: " & myShape.Adjustments(1) & vbCrLf _
            & "R*H: " & myShape.height * myShape.Adjustments(1) & vbCrLf _
            & "RelVP: " & myShape.RelativeVerticalPosition & vbCrLf _
            & "Background: " & myShape.Fill.ForeColor.RGB & vbCrLf
        If Items <> 1 And i <> Items Then
            Message = Message & "===============================" & vbCrLf
        End If
    Next
    MsgBox Message
End Sub

Sub CheckShapes()
    Dim Changes As Integer: Changes = 0
    For i = 1 To ActiveDocument.Shapes.count
        Set myShape = ActiveDocument.Shapes(i)
        If myShape.AutoShapeType = msoShapeRoundedRectangle Then
            If (myShape.height > 20) Then
                If (myShape.top <> 1.5) And (myShape.top <> -6.5) Then
                    myShape.Fill.ForeColor.RGB = RGB(255, 255, 0)
                    Changes = Changes + 1
                End If
                If (myShape.height <> 24) And (myShape.height <> 32) Then
                    myShape.Fill.ForeColor.RGB = RGB(255, 255, 0)
                    Changes = Changes + 1
                End If
                Debug.Print "Radius: " & ActiveDocument.Shapes(i).Adjustments(1)
                If (myShape.height = 24) And (ActiveDocument.Shapes(i).Adjustments(1) <> 0.10417) Then
                    myShape.Fill.ForeColor.RGB = RGB(255, 255, 0)
                    Changes = Changes + 1
                End If
               If (myShape.height = 32) And (ActiveDocument.Shapes(i).Adjustments(1) <> 0.10417) Then
                    myShape.Fill.ForeColor.RGB = RGB(255, 255, 0)
                    Changes = Changes + 1
                End If
            Else
                ' Check outline height & position
                If (myShape.top <> 10.5) Then
                    myShape.Fill.ForeColor.RGB = RGB(255, 255, 0)
                    myShape.Fill.Visible = msoTrue
                    Changes = Changes + 1
                End If
                If (myShape.height <> 15) Then
                    myShape.Fill.ForeColor.RGB = RGB(255, 255, 0)
                    myShape.Fill.Visible = msoTrue
                    Changes = Changes + 1
                End If
            End If
        End If
    Next
    Debug.Print "Total Changes Made: " & Changes
End Sub

Sub SetRadiusOld(control As IRibbonControl)
    Dim myShape As Shape
    If Selection.ShapeRange.count = 0 Then
        MsgBox "No objects selected."
        Exit Sub
    End If
    For i = 1 To Selection.ShapeRange.count
        Set myShape = Selection.ShapeRange(i)
        ' myShape.Adjustments(1) = 3.841743 / myShape.width
        myShape.Adjustments(1) = 0.16667
    Next
End Sub

Function isAligned(background As Shape, outline As Shape) As Boolean
    Dim backgroundRight As Single
    Dim outlineRight As Single
    
    backgroundRight = background.left + background.width
    outlineRight = outline.left + outline.width
    If (background.left = outline.left) Or (backgroundRight = outlineRight) Then
        isAligned = True
    Else
        isAligned = False
    End If
End Function

Sub SetRadius(control As IRibbonControl)
    Dim background As Shape
    Dim outline As Shape
    Dim backgroundIndex As Integer
    If Selection.ShapeRange.count = 0 Then
        MsgBox "No objects selected."
        Exit Sub
    End If
    backgroundIndex = 1
    For i = 1 To Selection.ShapeRange.count
        If (Selection.ShapeRange(i).height > Selection.ShapeRange(backgroundIndex).height) Then
            backgroundIndex = i
        End If
    Next
    
    Set background = Selection.ShapeRange(backgroundIndex)
    For i = 1 To Selection.ShapeRange.count
        If (i <> backgroundIndex) Then
            Set outline = Selection.ShapeRange(i)
            If (isAligned(background, outline)) Then
                Dim outlineLongSide As Single: outlineLongSide = outline.height
                If (outline.height > outline.width) Then
                    outlineLongSide = outline.width
                End If
                Dim backgroundLongSide As Single: backgroundLongSide = background.height
                If (background.height > background.width) Then
                    backgroundLongSide = background.width
                End If
                outline.Adjustments(1) = background.Adjustments(1) * (backgroundLongSide / outlineLongSide)
            Else
                ' Standard is R = H/6  (Assumes H < W , otherwise R = W/6 )
                outline.Adjustments(1) = 0.16667
            End If
        End If
    Next
End Sub

Sub SetRelativeVerticalPosition()
    Dim Source As Shape
    Dim target As Shape
    Set Source = Selection.ShapeRange(1)
    Set target = Selection.ShapeRange(2)
    target.RelativeVerticalPosition = Source.RelativeVerticalPosition
    target.top = Source.top
    target.height = Source.height
End Sub

Sub SetWidth(control As IRibbonControl)
    Dim myShape As Shape
    Dim OldWidth As Single
    Dim NewWidth As Single
    If Selection.ShapeRange.count = 0 Then
        MsgBox "No objects selected."
        Exit Sub
    End If
    For i = 1 To Selection.ShapeRange.count
        Set myShape = Selection.ShapeRange(i)
        OldWidth = myShape.width
        UserValue = InputBox(Prompt:="Enter New Width", _
            Title:="Enter New Width", Default:=OldWidth)
        If UserValue = Blank Then Exit Sub
        NewWidth = CSng(UserValue)
        myShape.width = NewWidth
    Next
End Sub

Sub SetHeight(control As IRibbonControl)
    Dim myShape As Shape
    Dim OldHeight As Single
    Dim NewHeight As Single
    If Selection.ShapeRange.count = 0 Then
        MsgBox "No objects selected."
        Exit Sub
    End If
    For i = 1 To Selection.ShapeRange.count
        Set myShape = Selection.ShapeRange(i)
        OldHeight = myShape.height
        UserValue = InputBox(Prompt:="Enter New Height", _
            Title:="Enter New Height", Default:=OldHeight)
        If UserValue = Blank Then Exit Sub
        NewHeight = CSng(UserValue)
        myShape.height = NewHeight
    Next
End Sub

Sub ResetBackground(control As IRibbonControl)
    If Selection.ShapeRange.count = 0 Then
        MsgBox "No objects selected."
        Exit Sub
    End If
    For i = 1 To Selection.ShapeRange.count
        Selection.ShapeRange(i).height = 24#
        Selection.ShapeRange(i).top = 1.5
        Selection.ShapeRange(i).Adjustments(1) = 0.10417
        Selection.ShapeRange(i).Fill.ForeColor.RGB = RGB(238, 236, 225)
    Next
End Sub

Sub ResetOutline(control As IRibbonControl)
    If Selection.ShapeRange.count = 0 Then
        MsgBox "No objects selected."
        Exit Sub
    End If
    For i = 1 To Selection.ShapeRange.count
        Selection.ShapeRange(i).height = 15#
        Selection.ShapeRange(i).top = 10.5
        Selection.ShapeRange(i).Fill.Visible = msoFalse
    Next
End Sub

Sub AlignLeft(control As IRibbonControl)
    If Selection.ShapeRange.count = 0 Then
        MsgBox "No objects selected."
        Exit Sub
    End If
    Dim LeftMost As Single: LeftMost = Selection.ShapeRange(1).left
    For i = 2 To Selection.ShapeRange.count
        If Selection.ShapeRange(i).left < LeftMost Then LeftMost = Selection.ShapeRange(i).left
    Next
    For i = 1 To Selection.ShapeRange.count
        Selection.ShapeRange(i).left = LeftMost
    Next
End Sub

Sub AlignRight(control As IRibbonControl)
    If Selection.ShapeRange.count = 0 Then
        MsgBox "No objects selected."
        Exit Sub
    End If
    Dim RightSide As Single: RightSide = (Selection.ShapeRange(1).left + Selection.ShapeRange(1).width)
    Dim RightMost As Integer: RightMost = RightSide
    For i = 2 To Selection.ShapeRange.count
        RightSide = (Selection.ShapeRange(i).left + Selection.ShapeRange(i).width)
        If RightSide > RightMost Then RightMost = RightSide
    Next
    For i = 1 To Selection.ShapeRange.count
        RightSide = (Selection.ShapeRange(i).left + Selection.ShapeRange(i).width)
        Selection.ShapeRange(i).left = Selection.ShapeRange(i).left + (RightMost - RightSide)
    Next
End Sub

Sub AlignBottom(control As IRibbonControl)
    If Selection.ShapeRange.count = 0 Then
        MsgBox "No objects selected."
        Exit Sub
    End If
    Dim Bottom As Single: Bottom = (Selection.ShapeRange(1).top + Selection.ShapeRange(1).height)
    Dim BottomMost As Single: BottomMost = Bottom
    For i = 2 To Selection.ShapeRange.count
        Bottom = (Selection.ShapeRange(i).top + Selection.ShapeRange(i).height)
        If Bottom > BottomMost Then BottomMost = Bottom
    Next
    For i = 1 To Selection.ShapeRange.count
        Selection.ShapeRange(i).top = (BottomMost - Selection.ShapeRange(i).height)
    Next
End Sub

Sub SetLeft(control As IRibbonControl)
    Dim myShape As Shape
    Dim OldLeft As Single
    Dim NewLeft As Single
    If Selection.ShapeRange.count = 0 Then
        MsgBox "No objects selected."
        Exit Sub
    End If
    For i = 1 To Selection.ShapeRange.count
        Set myShape = Selection.ShapeRange(i)
        OldLeft = myShape.left
        UserValue = InputBox(Prompt:="Enter New Left Position", _
            Title:="Enter New Left Position", Default:=OldLeft)
        If UserValue = Blank Then Exit Sub
        NewLeft = CSng(UserValue)
        myShape.left = NewLeft
    Next
End Sub

Sub SetTop(control As IRibbonControl)
    Dim myShape As Shape
    Dim OldTop As Single
    Dim NewTop As Single
    If Selection.ShapeRange.count = 0 Then
        MsgBox "No objects selected."
        Exit Sub
    End If
    For i = 1 To Selection.ShapeRange.count
        Set myShape = Selection.ShapeRange(i)
        OldTop = myShape.top
        UserValue = InputBox(Prompt:="Enter New Top Position", _
            Title:="Enter New Top Position", Default:=OldTop)
        If UserValue = Blank Then Exit Sub
        NewTop = CSng(UserValue)
        myShape.top = NewTop
    Next
End Sub

Sub SetBackgroundHeight(control As IRibbonControl)
    Dim myShape As Shape
    If Selection.ShapeRange.count = 0 Then
        MsgBox "No objects selected."
        Exit Sub
    End If
    For i = 1 To Selection.ShapeRange.count
        Selection.ShapeRange(i).height = 24
    Next
End Sub

Sub SetOutlineHeight(control As IRibbonControl)
    Dim myShape As Shape
    If Selection.ShapeRange.count = 0 Then
        MsgBox "No objects selected."
        Exit Sub
    End If
    For i = 1 To Selection.ShapeRange.count
        Selection.ShapeRange(i).height = 15
    Next
End Sub

Sub FindBadHeights(control As IRibbonControl)
    MsgBox "Not Yet Reimplemented"
End Sub

Sub ResetTops(control As IRibbonControl)
    MsgBox "Not Yet Reimplemented"
End Sub

Sub NumberPagesOld()
    Dim PageNumber As Integer
    PageNumber = 10
    For Each oSection In ActiveDocument.Sections
        For Each oFoot In oSection.Footers
            If oFoot.index = wdHeaderFooterPrimary Then
            With oFoot
                .Range.Text = "Page (" & PageNumber & ")"
                .PageNumbers.Add FirstPage:=True
            End With
            PageNumber = PageNumber + 1
            End If
        Next
    Next
    ' With ActiveDocument.Sections(1)
    '    .Footers(wdHeaderFooterPrimary).Range.Text = vbTab & "Page "
    '    .Footers(wdHeaderFooterPrimary).PageNumbers.Add FirstPage:=True
    ' End With
End Sub

Sub DuplicateShape(control As IRibbonControl)
    If Selection.ShapeRange.count = 0 Then
        MsgBox "No objects selected."
        Exit Sub
    End If
    
    Dim SectionNumber As Integer: SectionNumber = Selection.Information(wdActiveEndSectionNumber) + 1
    Dim OriginalName As String: OriginalName = Selection.ShapeRange.Name
    Dim oShape As Shape
    
    Set oShape = ActiveDocument.Shapes(OriginalName)
    Dim top As Single: top = oShape.top
    Dim left As Single: left = oShape.left
    Dim width As Single: width = oShape.width
    Dim height As Single: height = oShape.height
    
    For i = SectionNumber To ActiveDocument.Sections.count
        Dim newShape As Shape
        Set newShape = oShape.Duplicate
        newShape.Select
        Selection.Cut
        ActiveDocument.Sections(i).Range.Paste
        newShape.top = top
        newShape.left = left
    Next

End Sub

Sub SetHeaders(control As IRibbonControl)
    Dim SectionNumber As Integer
    ' Book Title
    ' Dim SecondHeader As String: SecondHeader = ChrW(&H1273) & ChrW(&H120B) & ChrW(&H1241) & ChrW(&H1361) & ChrW(&H12A2) & ChrW(&H1275) & ChrW(&H12EE) & ChrW(&H1335) & ChrW(&H12EB) & ChrW(&H12CA) & ChrW(&H1361) & ChrW(&H120A) & ChrW(&H1245) & ChrW(&H1361) & ChrW(&H1245) & ChrW(&H12F1) & ChrW(&H1235) & ChrW(&H1361) & ChrW(&H12EB) & ChrW(&H122C) & ChrW(&H12F5) & ChrW(&H1293) & ChrW(&H1361) & ChrW(&H12E8) & ChrW(&H12DC) & ChrW(&H121B) & ChrW(&H12CD) & ChrW(&H1361) & ChrW(&H1273) & ChrW(&H122A) & ChrW(&H12AD) & ChrW(&H1361) & ChrW(&H12A8) & ChrW(&H1290) & ChrW(&H121D) & ChrW(&H120D) & ChrW(&H12AD) & ChrW(&H1271) & ChrW(&H1362)
    ' Moges
    Dim FirstHeader As String: FirstHeader = ChrW(&H120A) & ChrW(&H1240) & ChrW(&H1361) & ChrW(&H1218) & ChrW(&H12D8) & ChrW(&H121D) & ChrW(&H122B) & ChrW(&H1295) & ChrW(&H1361) & ChrW(&H121E) & ChrW(&H1308) & ChrW(&H1235) & ChrW(&H1361) & ChrW(&H1225) & ChrW(&H12E9) & ChrW(&H121D) & ChrW(&H1361)

    ' Me'eraf
    ' Dim SecondHeader As String: SecondHeader = ChrW(&H1260) & ChrW(&H121D) & ChrW(&H12D5) & ChrW(&H122B) & ChrW(&H134D) & ChrW(&H1361) & ChrW(&H12CD) & ChrW(&H1235) & ChrW(&H1325) & ChrW(&H1361) & ChrW(&H12E8) & ChrW(&H121A) & ChrW(&H1308) & ChrW(&H1299) & ChrW(&H1361) & ChrW(&H12E8) & ChrW(&H1225) & ChrW(&H1228) & ChrW(&H12ED) & ChrW(&H1361) & ChrW(&H121D) & ChrW(&H120D) & ChrW(&H12AD) & ChrW(&H1276) & ChrW(&H127D) & ChrW(&H1293) & ChrW(&H1361) & ChrW(&H1270) & ChrW(&H1218) & ChrW(&H1223) & ChrW(&H1223) & ChrW(&H12EE) & ChrW(&H127B) & ChrW(&H1278) & ChrW(&H12CD) & ChrW(&H1362)
    ' Tsome Dugua
    ' Dim SecondHeader As String: SecondHeader = ChrW(&H1260) & ChrW(&H133E) & ChrW(&H1218) & ChrW(&H1361) & ChrW(&H12F5) & ChrW(&H1313) & ChrW(&H1361) & ChrW(&H12CD) & ChrW(&H1235) & ChrW(&H1325) & ChrW(&H1361) & ChrW(&H12E8) & ChrW(&H121A) & ChrW(&H1308) & ChrW(&H1299) & ChrW(&H1275) & ChrW(&H1295) & ChrW(&H1361) & ChrW(&H12E8) & ChrW(&H1225) & ChrW(&H1228) & ChrW(&H12ED) & ChrW(&H1361) & ChrW(&H121D) & ChrW(&H120D) & ChrW(&H12AD) & ChrW(&H1276) & ChrW(&H127D) & ChrW(&H1361) & ChrW(&H12A8) & ChrW(&H1225) & ChrW(&H122D) & ChrW(&H1361) & ChrW(&H12E8) & ChrW(&H1270) & ChrW(&H1220) & ChrW(&H1218) & ChrW(&H1228) & ChrW(&H1263) & ChrW(&H1278) & ChrW(&H12CD) & ChrW(&H1295) & ChrW(&H1361) & ChrW(&H1270) & ChrW(&H1218) & ChrW(&H1208) & ChrW(&H12A8) & ChrW(&H1275) & ChrW(&H1362)
    ' Dugua
    Dim SecondHeader As String: SecondHeader = ChrW(&H1260) & ChrW(&H1218) & ChrW(&H133D) & ChrW(&H1210) & ChrW(&H1348) & ChrW(&H1361) & ChrW(&H12F5) & ChrW(&H1313) & ChrW(&H1361) & ChrW(&H12CD) & ChrW(&H1235) & ChrW(&H1325) & ChrW(&H1361) & ChrW(&H12E8) & ChrW(&H121A) & ChrW(&H1308) & ChrW(&H1299) & ChrW(&H1361) & ChrW(&H12E8) & ChrW(&H1225) & ChrW(&H1228) & ChrW(&H12ED) & ChrW(&H1361) & ChrW(&H121D) & ChrW(&H120D) & ChrW(&H12AD) & ChrW(&H1276) & ChrW(&H127D) & ChrW(&H1293) & ChrW(&H1361) & ChrW(&H1270) & ChrW(&H1218) & ChrW(&H1223) & ChrW(&H1223) & ChrW(&H12EE) & ChrW(&H127B) & ChrW(&H1278) & ChrW(&H12CD) & ChrW(&H1362)
        
    SectionNumber = InputBox(Prompt:="Enter Starting Section", _
                    Title:="Enter Staring Section", Default:=1)
    If SectionNumber = Blank Then Exit Sub
    
    'FirstHeader = InputBox(Prompt:="Enter First Header", _
    '              Title:="Enter First Header")
    'If FirstHeader = Blank Then Exit Sub
    
    'SecondHeader = InputBox(Prompt:="Enter Second Header", _
    '               Title:="Enter Second Header", Default:=FirstHeader)
    'If SecondHeader = Blank Then Exit Sub
    
    
    Dim counter As Integer: counter = 0
    
    For i = SectionNumber To ActiveDocument.Sections.count
        Set oSection = ActiveDocument.Sections(i)
        If ((counter Mod 2) = 0) Then
            HeaderText = FirstHeader
        Else
            HeaderText = SecondHeader
        End If
        Debug.Print "Section: " & i
        With oSection.Headers(wdHeaderFooterPrimary)
            .Range.Text = HeaderText
            .Range.Font.Name = "Abyssinica SIL"
            .Range.Font.Italic = True
            .Range.ParagraphFormat.Alignment = wdAlignParagraphCenter
        End With
        counter = counter + 1
    Next
End Sub


Sub NumberPages(control As IRibbonControl)
    Dim PageNumber As Integer
    Dim SectionNumber As Integer
    PageNumber = InputBox(Prompt:="Enter Starting Number", _
                 Title:="Enter Staring Number", Default:=1)
    If PageNumber = Blank Then Exit Sub
    SectionNumber = InputBox(Prompt:="Enter Starting Section", _
                    Title:="Enter Staring Section", Default:=1)
    If SectionNumber = Blank Then Exit Sub
    
    Dim pageNumberAlign As WdPageNumberAlignment
    Dim pageNumberAlignEth As WdParagraphAlignment

    Dim Box As Shape


    For i = SectionNumber To ActiveDocument.Sections.count
        Debug.Print "Page " & i & " : " & (PageNumber Mod 2)
        Set oSection = ActiveDocument.Sections(i)

        pageNumberAlign = wdAlignPageNumberRight
        pageNumberAlignEth = wdAlignParagraphLeft
        pageNumberPositionEth = 70#
        If ((PageNumber Mod 2) = 0) Then
            pageNumberAlign = wdAlignPageNumberLeft
            pageNumberAlignEth = wdAlignParagraphRight
            pageNumberPositionEth = 460#
        End If
        With oSection.Headers(wdHeaderFooterPrimary)
            .PageNumbers.RestartNumberingAtSection = True
            .PageNumbers.StartingNumber = PageNumber
            .PageNumbers.Add PageNumberAlignment:=pageNumberAlign, FirstPage:=True
        End With

        ActiveDocument.Sections(i).Headers(wdHeaderFooterPrimary).Range.Paragraphs(1).FirstLineIndent = 0
        
        Set Box = oSection.Headers(wdHeaderFooterPrimary).Shapes.AddTextbox( _
        Orientation:=msoTextOrientationHorizontal, _
        left:=pageNumberPositionEth, top:=30, width:=80, height:=30)

        'The solution for you:
        Box.TextFrame.TextRange.Text = ArabicToEthiopic(PageNumber)
        Box.TextFrame.TextRange.Font.Name = "Abyssinica SIL"
        Box.TextFrame.TextRange.ParagraphFormat.Alignment = pageNumberAlignEth
        Box.Line.Visible = False
        Box.Fill.Visible = False
        Box.TextFrame.MarginLeft = 0
        Box.TextFrame.MarginRight = 0
        Box.TextFrame.TextRange.Paragraphs(1).FirstLineIndent = 0
        
        PageNumber = PageNumber + 1
    Next
End Sub

Sub NumberPagesFooter(control As IRibbonControl)
    Dim PageNumber As Integer
    Dim SectionNumber As Integer
    PageNumber = InputBox(Prompt:="Enter Starting Number", _
                 Title:="Enter Staring Number", Default:=1)
    If PageNumber = Blank Then Exit Sub
    SectionNumber = InputBox(Prompt:="Enter Starting Section", _
                    Title:="Enter Staring Section", Default:=1)
    If SectionNumber = Blank Then Exit Sub
    
    Dim pageNumberAlign As WdPageNumberAlignment
    Dim pageNumberAlignEth As WdParagraphAlignment
    ' For Each oSection In ActiveDocument.Sections
    For i = SectionNumber To ActiveDocument.Sections.count
        Debug.Print "Page " & i & " : " & (PageNumber Mod 2)
        Set oSection = ActiveDocument.Sections(i)
        pageNumberAlign = wdAlignPageNumberRight
        pageNumberAlignEth = wdAlignParagraphLeft
        If ((PageNumber Mod 2) = 0) Then
            pageNumberAlign = wdAlignPageNumberLeft
            pageNumberAlignEth = wdAlignParagraphRight
        End If
        With oSection.Footers(wdHeaderFooterPrimary)
            .Range.Text = ArabicToEthiopic(PageNumber)
            .Range.Font.Name = "Abyssinica SIL"
            .Range.ParagraphFormat.Alignment = pageNumberAlignEth
            .PageNumbers.RestartNumberingAtSection = True
            .PageNumbers.StartingNumber = PageNumber
            .PageNumbers.Add PageNumberAlignment:=pageNumberAlign, FirstPage:=True
        End With
       PageNumber = PageNumber + 1
    Next
End Sub

Sub FindReplace(find As String, replace As String)
    With ActiveDocument.Range.find
      .Text = find
      .Replacement.Text = replace
      .Replacement.ClearFormatting
      .Replacement.Font.Italic = False
      .Forward = True
      .Wrap = wdFindContinue
      .Format = False
      .MatchCase = False
      .MatchWholeWord = False
      .MatchWildcards = False
      .MatchSoundsLike = False
      .MatchAllWordForms = False
      .Execute replace:=wdReplaceAll
    End With
End Sub

Sub FindReplacePageReferences()
    Dim oldReference As String
    Dim newReference As String
    Dim oldPage As Integer
    Dim newPage As Integer
        

    ' For oldPage = 218 To 181 Step -1 ' Digua
    ' For oldPage = 174 To 141 Step -1 ' Tsome Digua
    For oldPage = 127 To 67 Step -1 ' Me'eraf
        oldReference = ChrW(&H1363) & ChrW(&H1308) & ChrW(&H133D) & "/" & ArabicToEthiopic(oldPage) & ChrW(&H1361) & "(" & oldPage & ")" & ChrW(&H1363)
        ' newPage = oldPage + 5  ' Digua
        ' newPage = oldPage + 4  ' Tsome Digua
        newPage = oldPage + 6  ' Me'eraf
        newReference = ChrW(&H1363) & ChrW(&H1308) & ChrW(&H133D) & "/" & ArabicToEthiopic(newPage) & ChrW(&H1361) & "(" & newPage & ")" & ChrW(&H1363)
        FindReplace oldReference, newReference
    Next
End Sub

Sub TestConversion()
    Dim num123 As String
    num123 = ArabicToEthiopic(5)
    Debug.Print "Number is " & num123
End Sub

Function ArabicToEthiopic(number As Integer) As String
    Dim numberString As String: numberString = CStr(number)
    Dim n As Integer: n = Len(numberString) - 1
        
    If ((n Mod 2) = 0) Then
        numberString = "0" & numberString
        n = n + 1
    End If
         
    Dim ETHIOPIC_ONE As String: ETHIOPIC_ONE = ChrW(&H1369)
    Dim ETHIOPIC_TEN As String: ETHIOPIC_TEN = ChrW(&H1372)
    Dim ETHIOPIC_HUNDRED As String: ETHIOPIC_HUNDRED = ChrW(&H137B)
    Dim ETHIOPIC_TEN_THOUSAD As String: ETHIOPIC_TEN_THOUSAND = ChrW(&H137C)
    Dim ethioNumberString As String: ethioNumberString = ""

    Dim asciiOne, asciiTen, ethioOne, ethioTen, sep As String
    Dim pos, index As Integer
     
    For place = n To 0 Step -1
            asciiOne = ""
            asciiTen = ""
            ethioOne = ""
            ethioTen = ""

            index = n - place
            asciiTen = Strings.Mid(numberString, (n - place + 1), 1)
            place = place - 1
            asciiOne = Mid(numberString, (n - place + 1), 1)
            
            If (asciiOne <> "0") Then
                ethioAddr = AscW(asciiOne)
                oneAddr = Asc("1")
                testAddr = AscW(ETHIOPIC_ONE) + ethioAddr - oneAddr
                ethioOne = ChrW(AscW(ETHIOPIC_ONE) + (AscW(asciiOne) - AscW("1")))
            End If
                
            If (asciiTen <> "0") Then
               ethioTen = ChrW(AscW(ETHIOPIC_TEN) + (AscW(asciiTen) - AscW("1")))
            End If
            
            pos = (place Mod 4) / 2
            
            sep = ""
            If (place <> 0) Then
                If (pos = 0) Then
                    sep = ETHIOPIC_TEN_THOUSAND
                ElseIf ((ethioOne <> "") Or (ethioTen <> "")) Then
                    sep = ETHIOPIC_HUNDRED
                End If
            End If

            If ((ethioOne = ETHIOPIC_ONE) And (ethioTen = "") And (n > 1)) Then
                If ((sep = ETHIOPIC_HUNDRED) Or ((place + 1) = n)) Then
                    ethioOne = ""
                End If
            End If
            
            If (ethioTen <> "") Then ethioNumberString = ethioNumberString & ethioTen
            If (ethioOne <> "") Then ethioNumberString = ethioNumberString & ethioOne
            If (sep <> "") Then ethioNumberString = ethioNumberString & sep
    Next
         
    ArabicToEthiopic = ethioNumberString
End Function

Sub RemoveHeaderTextBox()
    Dim sh As Shape

    For i = 1 To ActiveDocument.Sections.count
        For Each sh In ActiveDocument.Sections(i).Headers(wdHeaderFooterPrimary).Shapes
            If sh.Type = msoTextBox Then
                sh.Delete
            End If
        Next sh
    Next
End Sub


Sub ClearFooter()
    For i = 1 To ActiveDocument.Sections.count
        ActiveDocument.Sections(i).Footers(wdHeaderFooterPrimary).Range.Delete
    Next
End Sub

Sub SetColumnWidths()
        ' For i = 1 To ActiveDocument.Tables.count
        For i = 56 To 56

            Set myTable = ActiveDocument.Tables(i)
            Debug.Print myTable.Columns.count
            
            myTable.AutoFitBehavior wdAutoFitFixed
            myTable.PreferredWidthType = wdPreferredWidthPoints
            myTable.PreferredWidth = 477#
            myTable.Rows(1).Cells(4).width = 29#
            myTable.Rows(1).Cells(4).LeftPadding = 4#
            myTable.Rows(1).Cells(4).RightPadding = 4#
            cell2Width = myTable.Rows(2).Cells(2).width
            myTable.Rows(1).Cells(1).width = cell2Width + 25#
            If myTable.Columns.count = 5 Then
                For j = 2 To myTable.Rows.count
                myTable.Rows(j).Cells(1).Range.Shading.BackgroundPatternColor = wdColorYellow
                    myTable.Rows(j).Cells(1).width = 25#
                    myTable.Rows(1).Cells(2).LeftPadding = 2.9
                    myTable.Rows(1).Cells(4).LeftPadding = 2.9
                    myTable.Rows(j).Cells(5).width = 29#
                Next
            End If
            myTable.PreferredWidth = 477#
        Next
End Sub

Sub CheckTableCells()
    Dim oCell As Cell
    Dim oRow As Row

    For i = 1 To ActiveDocument.Tables.count
        Set myTable = ActiveDocument.Tables(i)
        For Each oRow In myTable.Rows
            For Each oCell In oRow.Cells
                If oCell.Range.Text = Chr(13) & Chr(7) Then
                    oCell.Range.Shading.BackgroundPatternColor = wdColorYellow
                    ' MsgBox oCell.RowIndex & " " & oCell.ColumnIndex & " is empty."
                End If
            Next oCell
        Next oRow
    Next

End Sub

Sub FixMargins()
        For i = 1 To ActiveDocument.Tables.count
            Set myTable = ActiveDocument.Tables(i)
            myTable.PreferredWidth = 477#
            myTable.Rows(1).Cells(4).LeftPadding = 4#
            myTable.Rows(1).Cells(4).RightPadding = 4#
        Next
End Sub

Sub SetNormalMiliket(control As IRibbonControl)
    Selection.Range.PhoneticGuide Raise:=0, FontSize:=5.5
End Sub

Sub SetSereyulMiliket(control As IRibbonControl)
    Selection.Range.PhoneticGuide Raise:=5, FontSize:=8
End Sub

Sub ReadPara()
    Dim docSource As Word.Document
    Dim docOutline As Word.Document
    Dim rng As Word.Range
    Dim strText As String
    Dim count As Integer

    Set docSource = ActiveDocument
    Set docOutline = Documents.Add
    Set rng = docOutline.Content
    Dim DocPara As Paragraph
    count = 1
    ' Dim StartHeader As String: StartHeader = ChrW(&H1299) & ChrW(&H1361)
    For Each DocPara In docSource.Paragraphs
        If (DocPara.Range.Style Like "Zaima Heading*") And (Len(DocPara.Range.Text) > 1) Then
            strText = Trim(DocPara.Range.Text)
            ' If InStr(Left(strText, 5), StartHeader) > 0 Then
            If (count Mod 2) = 1 Then
                ' rng.InsertAfter strText
                rng.InsertAfter left(strText, Len(strText) - 1)
            Else
                rng.InsertAfter strText
                'rng.InsertAfter Left(strText, Len(strText) - 1)
            End If


        ' Set the style of the selected range and
        ' then collapse the range for the next entry.
       ' rng.Style = "Heading " & intLevel
            rng.Collapse wdCollapseEnd
            count = count + 1
        End If
    Next

End Sub


Sub MakeDictionary()
    Dim docSource As Word.Document
    Dim docOutline As Word.Document
    Dim rng As Word.Range
    Dim strText As String
    Dim count As Integer

    Set docSource = ActiveDocument
    ' Set docOutline = Documents.Add
    ' Set rng = docOutline.Content
    ' Dim DocPara As Paragraph
    count = 1
    Dim t As Table, r As Long, c As Long, rw As Row
    Set t = docSource.Tables(1)

    For r = 1 To t.Rows.count
     Set rw = t.Rows(r)
        For c = 1 To rw.Cells.count
            Debug.Print "In row " & r & " cell " & c
        Next c
    Next r

End Sub

Sub ReadAllRows()
    Dim NbRows As Integer
    Dim NbColumns As Integer
    Dim i, j As Integer
    Dim SplitStr() As String
    Dim Silt As String
    Dim Silt2 As String
    Dim Source As String
    Dim Page As String
    
    'note : my table here is a public value that i get thanks to bookmarks
    Dim t As Table
    Set t = ActiveDocument.Tables(1)
    NbRows = t.Rows.count
    NbColumns = t.Columns.count
    
    Set docOut = Documents.Add
    Set Content = docOut.Content


    For i = 3 To NbRows
            On Error GoTo ErrorHandler

            t.Cell(i, 3).Range.Copy
            Content.Paragraphs.Last.Range.Select
        
            Selection.PasteAndFormat wdPasteDefault
            Selection.Collapse Direction:=wdCollapseEnd
            Selection.MoveLeft
            Selection.InsertAfter vbTab
            Selection.Collapse Direction:=wdCollapseEnd
            Silt = t.Cell(i, 6).Range.Text
            Silt2 = Right(Silt, Len(Silt) - 1)
            Selection.InsertAfter "(" & left$(Silt2, Len(Silt2) - 3) & ") "
            Selection.Collapse Direction:=wdCollapseEnd
             
            t.Cell(i, 4).Range.Copy
            Selection.PasteAndFormat wdSingleCellText
            Selection.MoveLeft
            Selection.MoveLeft
            Selection.Delete
            Selection.Collapse Direction:=wdCollapseEnd
            Source = t.Cell(i, 5).Range.Text
            Page = t.Cell(i, 2).Range.Text
            Selection.InsertAfter ChrW(&H1364) & left$(Source, Len(Source) - 3) _
            & ChrW(&H1364) & ChrW(&H1308) & ChrW(&H133D) & " " & left$(Page, Len(Page) - 2) & ChrW(&H1362)
         
NextRow:


        'We have here all the values of the line
    Next i

'This Error handler will skip the whole Select Case and thus will proceed towards next cell
ErrorHandler:
If Err.number = 5941 Then
    Err.Clear
    Resume NextRow
End If

End Sub

