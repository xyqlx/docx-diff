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

虽然这1000+的Python代码行数有些让人望而却步……

## 探索过程

在实验之前，提供一个简单的样例，`test/raw.docx`和`test/raw_mod1.docx`，其中包含了一些涉及哈基米的不堪入目的内容，后者在前者的基础上修改了一些文字内容

### 首先看看Microsoft Word的比较文档功能吧

在Word中打开这个功能，选择两个文件开始比较就可以了。Word很人性化地生成了一个修订版本。把它保存下来就可以了。

其实这个功能的成功运行可以直接把问题转化为“如何解析一个修订版本的Word文件并且把它转换为标注”

### 再看看搜到的difflib_docx的效果吧

简单看了下代码，按照说明把两个docx文件放在compareRobot文件夹下运行compareWithContract.py即可

然而结果并不是很理想，Word无法打开生成的文件，提示“Word在试图打开文件时出现错误，请尝试下列方法……”

经过了一些弱智转换、比较与测试，发现**解压docx文件然后再压缩**可以解决问题，这……可能是编码原因吗？

既然能够打开文件，那么是时候看看效果了

……遗憾的是，这个工具感觉不如……Word的比较功能，虽然它给的样例中能够识别出增加和更改，但是在我们的样例中，它把所有的内容变更都识别成了增加

不过，这个代码还是有可以学习的地方的。比如说，查看xml的话，可以猜测它是通过给文字内容加上Highlight标签实现标识的

## 阶段性结论与下一轮检索

那么，答案就是——能找到一款将Word修订转换为高亮标注的工具吗？

这个需求听起来很常见，应该不会那么难找吧——但是我应该使用什么检索关键词呢？

### 开始搜索解决方案

首先xyq用中文检索，比如说“word 修订转高亮”

首先检索到cnblog上的一篇博文[Word修订内容批量标红](https://www.cnblogs.com/geoli91/p/16618266.html)，感谢这位博主，为后来者奏了不少弯路：

1. python-docx：不支持修订，需要操作xml
2. aspose-words：付费
3. 用Python操作win32com：写了挺长一段，然而“在处理到有图篇插入和分节符插入等相关的修订时，后面的修订都不会再处理”
4. VBA：写了几行，成功了

这位博主的代码如下：

```VBA
Sub Set_Revisions_Red()
'关闭修订模式
ActiveDocument.TrackRevisions = Flase
 
'迭代每一个修订，改为红色并接受修订
For n = 1 To ActiveDocument.Revisions.Count
    '移动至下一个修订
    Selection.NextRevision (True)
    '设置修订内容字体颜色为红色
    Selection.Font.Color = wdColorRed
    '接受当前修订
    Selection.Range.Revisions.AcceptAll
    Next n
End Sub
```

xyq也用英文检索了，比如说“Convert Word track changes to highlighted annotations”

检索到了一个提问[Convert tracked changes to highlighted](https://superuser.com/questions/813428/convert-tracked-changes-to-highlighted)

这里的回答给出的VBA代码如下：

```VBA
Sub tracked_to_highlighted()           
    tempState = ActiveDocument.TrackRevisions
    ActiveDocument.TrackRevisions = False    
    For Each Change In ActiveDocument.Revisions        
        Set myRange = Change.Range
        myRange.Revisions.AcceptAll
        myRange.HighlightColorIndex = wdGreen            
    Next    
    ActiveDocument.TrackRevisions = tempState
End Sub
```

下面的回复还给出了一个想使用不同颜色时可以参考的网页<https://word.tips.net/T000253_Changing_Character_Color.html>

### 虽然但是还是想先试试C\#

QAQ

VS，启动！

因为要处理不同的revision类型，[官网文档](https://learn.microsoft.com/en-us/dotnet/api/microsoft.office.interop.word.wdrevisiontype?view=word-pia)可以查看到共有21种

这里我们暂时先试试只处理wdRevisionDelete，wdRevisionInsert，wdRevisionMovedFrom和wdRevisionMovedTo

怎么说呢，调试过程还挺长，以下是xyq遇到的雷点：

1. wdRevisionMovedFrom和wdRevisionMovedTo只要Accept了一个就会自动Accept另一个，所以在其中一个分支处理就可以了，这里可以用MovedRange属性获取另一个的Range
2. _Document.Revisions每次迭代的时候都会重新计算（有可能是记录当前的字符位置），所以谨慎在迭代过程中修改文档（虽然xyq确实这样干了）
3. Range.InsertBefore(String)这个方法会修改Range的起始位置，修改后还会导致这个Range对应的Revision无法访问。这里xyq采用的方案是记录Range的起始位置，在Accept Revision之后创建新的Range，然后把原来的Range的内容复制到新的Range中

最后的代码如下：

```csharp
using Word = Microsoft.Office.Interop.Word;

namespace Revision2Highlight
{
    internal class Program
    {
        static void Main(string[] args)
        {
            // open the document from the command line
            var wordApp = new Word.Application();
            // check args
            if (args.Length == 0)
            {
                Console.WriteLine("Please specify a document to open.");
                return;
            }
            else
            {
                if (!File.Exists(args[0]))
                {
                    Console.WriteLine("File does not exist.");
                    return;
                }
                if (Path.GetExtension(args[0]) != ".docx")
                {
                    Console.WriteLine("File is not a .docx file.");
                    return;
                }
            }
            var doc = wordApp.Documents.Open(args[0], ReadOnly: true);
            // copy to a new document
            var newDoc = wordApp.Documents.Add();
            doc.Content.Copy();
            newDoc.Content.Paste();
            // close
            doc.Close();
            // stop tracking revisions
            newDoc.TrackRevisions = false;
            // highlight revisions
            foreach (Word.Revision revision in newDoc.Revisions)
            {
                Console.WriteLine(revision.Type);
                // highlight deletions in red
                if (revision.Type == Word.WdRevisionType.wdRevisionDelete)
                {
                    var text = revision.Range.Text;
                    var start = revision.Range.Start;
                    revision.Accept();
                    var newRange = newDoc.Range(start, start);
                    newRange.InsertBefore(text);
                    newRange.Font.ColorIndex = Word.WdColorIndex.wdRed;
                }
                // highlight insertions in green
                else if (revision.Type == Word.WdRevisionType.wdRevisionInsert)
                {
                    revision.Range.Font.ColorIndex = Word.WdColorIndex.wdGreen;
                    revision.Accept();
                }
                else if (revision.Type == Word.WdRevisionType.wdRevisionMovedFrom)
                {
                    var text = revision.Range.Text;
                    var start = revision.Range.Start;
                    revision.MovedRange.Font.ColorIndex = Word.WdColorIndex.wdGreen;
                    revision.Accept();
                    var newRange = newDoc.Range(start, start);
                    newRange.InsertBefore(text);
                    newRange.Font.ColorIndex = Word.WdColorIndex.wdRed;
                }
                else if (revision.Type == Word.WdRevisionType.wdRevisionMovedTo)
                {
                    // processed in wdRevisionMovedFrom
                }
                // highlight other revisions in yellow
                else
                {
                    revision.Range.HighlightColorIndex = Word.WdColorIndex.wdYellow;
                    revision.Accept();
                }
            }

            // save new document in the same folder as the original document
            string newDocPath = Path.Combine(Path.GetDirectoryName(args[0]) ?? Environment.CurrentDirectory, Path.GetFileNameWithoutExtension(args[0]) + "_highlighted.docx");
            newDoc.SaveAs2(newDocPath);
            newDoc.Close();
            // press any key to continue
            Console.WriteLine("Press any key to continue...");
            Console.ReadKey();
        }
    }
}
```

没错，可以注意到其实标记的方式并不是highlight……实际上代码可以调整的空间很大，在使用的时候肯定大概率要根据具体的文档定制……不过，能自动化就已经是胜利了
