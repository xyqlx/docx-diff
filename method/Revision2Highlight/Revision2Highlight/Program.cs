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
