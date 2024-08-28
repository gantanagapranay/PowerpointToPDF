Sub CreatePresentation()
    ' Declare variables
    Dim pptApp As Object
    Dim pptPres As Object
    Dim slideIndex As Integer
    Dim slide As Object
    Dim slideTitle As String
    Dim slideContent As String
    
    ' Create a new PowerPoint application
    Set pptApp = CreateObject("PowerPoint.Application")
    pptApp.Visible = True
    
    ' Create a new presentation
    Set pptPres = pptApp.Presentations.Add
    
    ' Slide 1: Title Slide
    slideIndex = slideIndex + 1
    Set slide = pptPres.Slides.Add(slideIndex, 1)
    slideTitle = "Difference Between Web Application and Mobile Application"
    slide.Shapes.Title.TextFrame.TextRange.Text = slideTitle
    
    ' Slide 2: Introduction Slide
    slideIndex = slideIndex + 1
    Set slide = pptPres.Slides.Add(slideIndex, 1)
    slideTitle = "Introduction"
    slideContent = "This presentation will cover the key differences between web applications and mobile applications."
    slide.Shapes.Title.TextFrame.TextRange.Text = slideTitle
    slide.Shapes.Placeholders(2).TextFrame.TextRange.Text = slideContent
    
    ' Slide 3: Web Applications
    slideIndex = slideIndex + 1
    Set slide = pptPres.Slides.Add(slideIndex, 2)
    slideTitle = "Web Applications"
    slideContent = "• Accessible through web browsers\n• Cross-platform compatibility\n• Requires internet connection\n• Easier to update and maintain\n• Limited access to device features"
    slide.Shapes.Title.TextFrame.TextRange.Text = slideTitle
    slide.Shapes(2).TextFrame.TextRange.Text = slideContent
    
    ' Slide 4: Mobile Applications
    slideIndex = slideIndex + 1
    Set slide = pptPres.Slides.Add(slideIndex, 2)
    slideTitle = "Mobile Applications"
    slideContent = "• Installed on mobile devices\n• Platform-specific (iOS, Android)\n• Can work offline\n• Direct access to device features (camera, GPS, etc.)\n• Regular updates through app stores"
    slide.Shapes.Title.TextFrame.TextRange.Text = slideTitle
    slide.Shapes(2).TextFrame.TextRange.Text = slideContent
    
    ' Slide 5: Key Differences
    slideIndex = slideIndex + 1
    Set slide = pptPres.Slides.Add(slideIndex, 2)
    slideTitle = "Key Differences"
    slideContent = "• Accessibility: Web apps via browsers; Mobile apps via installation\n• Internet Dependency: Web apps need it; Mobile apps may not\n• Platform: Web apps are cross-platform; Mobile apps are platform-specific"
    slide.Shapes.Title.TextFrame.TextRange.Text = slideTitle
    slide.Shapes(2).TextFrame.TextRange.Text = slideContent
    
    ' Slide 6: Conclusion
    slideIndex = slideIndex + 1
    Set slide = pptPres.Slides.Add(slideIndex, 2)
    slideTitle = "Conclusion"
    slideContent = "Both web and mobile applications have their unique advantages and use cases. The choice depends on the target audience and the specific features required."
    slide.Shapes.Title.TextFrame.TextRange.Text = slideTitle
    slide.Shapes(2).TextFrame.TextRange.Text = slideContent

    ' Finalize the presentation
    pptPres.SaveAs "Difference_Between_Web_and_Mobile_Applications.pptx"
    pptApp.Quit

    MsgBox "Presentation Created Successfully!"
End Sub
