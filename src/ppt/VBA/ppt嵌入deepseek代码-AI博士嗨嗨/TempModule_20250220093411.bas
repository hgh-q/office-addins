' Module: TempModule_20250220093411

Public Sub CenterCurrentSlideContent()
    ' 获取当前幻灯片
    Dim oSlide As slide
    Set oSlide = Application.ActiveWindow.View.slide
    
    ' 居中所有文本框内容
    Dim oShape As Shape
    For Each oShape In oSlide.Shapes
        ' 只处理文本框和文本效果
        If oShape.Type = msoTextBox Or oShape.Type = msoTextEffect Then
            ' 居中对齐（水平和垂直）
            oShape.Left = (Application.ActivePresentation.PageSetup.SlideWidth - oShape.Width) / 2
            oShape.Top = (Application.ActivePresentation.PageSetup.SlideHeight - oShape.Height) / 2
        End If
    Next oShape
End Sub