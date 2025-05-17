<img width="1400" alt="image" src="https://github.com/user-attachments/assets/07c2e050-ce25-49c5-b8f7-d4d40cd310d0" /># word调整图片宽度-mac版
**VBA 代码：**

这里提供一个基础的 VBA 代码，可以将文档中所有图片的宽度调整为一个设定的值。

```vba
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
```

**如何使用 VBA 代码：**

1.  **打开 Word 文档：** 打开您想要调整图片宽度的 Word 文档。
2.  **打开 VBA 编辑器：**
      * 点击顶部菜单栏的“工具 (Tools)”。
      * 选择“宏 (Macro)”。
      * 选择“Visual Basic 编辑器 (Visual Basic Editor)”。 或者，您也可以直接使用快捷键 `Option + F11` (在某些 Mac 键盘上可能是 `Fn + Option + F11`)。
3.  **插入模块：**
      * 在 VBA 编辑器中，您会看到一个“项目 (Project)”窗口（通常在左侧）。
      * 找到您的文档项目（例如，“Project (YourDocumentName)”）。
      * 右键点击您的文档项目，或者点击顶部菜单栏的“插入 (Insert)”。
      * 选择“模块 (Module)”。
4.  **粘贴代码：**
5.  <img width="1400" alt="image" src="https://github.com/user-attachments/assets/e8c603e5-fb5b-402a-b6e8-bf1fc9ff7189" />

      * 一个新的模块窗口会打开（通常在右侧）。
      * 将上面提供的 VBA 代码复制并粘贴到这个模块窗口中。
6.  **修改期望宽度 (可选)：**
      * 在代码中找到这一行：`desiredWidthCm = 8`
      * 将数字 `8` 修改为您期望的图片宽度值（单位是厘米）。例如，如果您希望图片宽度为 5 厘米，就改成 `desiredWidthCm = 5`。
7.  **运行宏：**
      * 确保光标在 `Sub ResizeAllPictures()` 和 `End Sub` 之间的任何位置。
      * 点击 VBA 编辑器工具栏上的“运行 (Run)”按钮（一个绿色的三角形播放按钮），或者按 `F5` 键。
      * 或者，您可以关闭 VBA 编辑器，返回到 Word 文档。然后：
          * 点击顶部菜单栏的“工具 (Tools)”。
          * 选择“宏 (Macro)”。
          * 再选择“宏 (Macros)...” (或者快捷键 `Option + F8`)。
          * 在弹出的“宏”对话框中，选择名为 `ResizeAllPictures` 的宏。
          * 点击“运行 (Run)”。
8.  **查看结果：**
      * 宏运行完毕后，文档中所有图片的宽度应该已经被调整。您会看到一个提示框显示“所有图片的宽度已调整完毕！”。

