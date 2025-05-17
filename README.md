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
3.  **粘贴VBA代码：**

  <img width="1400" alt="image" src="https://github.com/user-attachments/assets/e8c603e5-fb5b-402a-b6e8-bf1fc9ff7189" />

4.  **修改期望宽度 (可选)**
      * 在代码中找到这一行：`desiredWidthCm = 8`
      * 将数字 `8` 修改为您期望的图片宽度值（单位是厘米）。例如，如果您希望图片宽度为 5 厘米，就改成 `desiredWidthCm = 5`。

5. **control+S保存,选择删除宏并保存**
<img width="1109" alt="image" src="https://github.com/user-attachments/assets/33fd68cb-8c26-47fa-835b-71d3da5f1061" />

6.  **运行宏：**
      <img width="1405" alt="image" src="https://github.com/user-attachments/assets/b26c545c-4b3f-40be-9746-60c5210466c5" />

7.  **查看结果：**
      * 宏运行完毕后，文档中所有图片的宽度应该已经被调整。您会看到一个提示框显示“所有图片的宽度已调整完毕！”。

