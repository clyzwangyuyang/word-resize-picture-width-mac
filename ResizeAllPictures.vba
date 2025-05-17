Sub ResizeAllPictures()

    Dim oInlineShape As InlineShape
    Dim oShape As Shape
    Dim desiredWidthCm As Single ' 期望的宽度，单位：厘米
    Dim desiredWidthPoints As Single ' 期望的宽度，单位：磅 (1 厘米 = 28.3465 磅)

    ' --- 在这里设置你期望的图片宽度 (单位：厘米) ---
    desiredWidthCm = 8
    ' ---------------------------------------------

    ' 将厘米转换为磅
    desiredWidthPoints = desiredWidthCm * 28.3465

    ' 调整嵌入型图片
    For Each oInlineShape In ActiveDocument.InlineShapes
        If oInlineShape.Type = wdInlineShapePicture Or oInlineShape.Type = wdInlineShapeLinkedPicture Then
            ' 保持图片的原始宽高比
            Dim originalAspectRatio As Single
            originalAspectRatio = oInlineShape.Height / oInlineShape.Width

            oInlineShape.LockAspectRatio = msoTrue ' 锁定宽高比
            oInlineShape.Width = desiredWidthPoints
            ' 如果需要同时调整高度以保持比例，则不需要下面这行，因为 LockAspectRatio = msoTrue 会自动处理
            ' oInlineShape.Height = desiredWidthPoints * originalAspectRatio
        End If
    Next oInlineShape

    ' 调整浮动型图片
    For Each oShape In ActiveDocument.Shapes
        If oShape.Type = msoPicture Or oShape.Type = msoLinkedPicture Then
            ' 保持图片的原始宽高比
            Dim originalAspectRatioShape As Single
            originalAspectRatioShape = oShape.Height / oShape.Width

            oShape.LockAspectRatio = msoTrue ' 锁定宽高比
            oShape.Width = desiredWidthPoints
            ' 如果需要同时调整高度以保持比例，则不需要下面这行
            ' oShape.Height = desiredWidthPoints * originalAspectRatioShape
        End If
    Next oShape

    MsgBox "所有图片的宽度已调整完毕！", vbInformation

End Sub
