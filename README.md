# docx-diff

本项目用于比较两个docx文件的差异研究

This project is used to compare the differences between two docx files

## 我们的目标

虽然名字是比较两个docx文件的差异，但是根据实用需求，暂时的目标是对于两个基本由文本组成的Word文件，其中一个文件在另一个文件的基础上进行修改。需要得到这两个Word文件的文本差异，包括增加，修改，删除等等，并且将它标注出来，生成一个新的文件。

## 可以参考的资料

可以搜到一些比较两个文件的项目，例如<https://github.com/wooneusean/DocxDiff>（C#）和<https://github.com/Nopom/difflib_docx>（Python），后者使用difflib，并且使用如下的标注方案：

| 标注 | 类型 |
|---|---|
| 红色 | 删除 |
| 绿色 | 增加 |
| 黄色 | 修改 |
| 灰色 | 全部改变 |

## 过程

在实验之前，提供一个简单的样例，`test/raw.docx`和`test/raw_mod1.docx`，其中包含了一些设计哈基米的不堪入目的内容，后者在前者的基础上修改了一些文字内容

### 首先看看Microsoft Word的比较文档功能吧

在Word中打开这个功能，选择两个文件开始比较就可以了。Word很人性化地生成了一个修订版本。把它保存下来就可以了。

其实这个功能的成功运行可以直接把问题转化为“如何解析一个修订版本的Word文件并且把它转换为标注”

### 再看看搜到的difflib_docx的效果吧

